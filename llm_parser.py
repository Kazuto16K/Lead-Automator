"""Turn the manager's natural-language message into structured lead JSON using Gemini."""
import difflib
import json
import re
import time

from schema import fields_for_steps

DEFAULT_MODEL = "gemini-flash-latest"  # alias that always points at the current Flash model
FALLBACK_MODELS = ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite"]  # tried if the first is busy/gone

ALIASES = {  # common phrasings -> allowed value (applied before fuzzy matching)
    "no website": "None", "none": "None", "no site": "None", "poor": "Outdated / Poor",
    "outdated": "Outdated / Poor", "not working": "Broken", "doesn't load": "Broken",
}


def _field_spec(steps, opts):
    lines = []
    for key, (_col, label, _s, ftype) in fields_for_steps(steps).items():
        if ftype.startswith("enum:"):
            allowed = opts[ftype.split(":")[1]]
            lines.append(f'- "{key}" ({label}): one of {json.dumps(allowed)}')
        elif ftype == "float":
            lines.append(f'- "{key}" ({label}): number')
        elif ftype == "int":
            lines.append(f'- "{key}" ({label}): integer')
        else:
            lines.append(f'- "{key}" ({label}): string')
    return "\n".join(lines)


def build_prompt(text, niche, steps, opts):
    must = dict(zip(opts["A"], opts.get("B", []))).get(niche, "")
    return f"""You extract lead-research data for a local-business lead tracker (Kolkata, India).
The marketing manager wrote the note below. Niche: "{niche}". It covers step(s): {sorted(steps)}.

Return ONLY JSON of the form {{"leads": [ {{...}}, ... ]}} - one object per business mentioned.
Use exactly these keys (omit a key or use null if the note does not say; NEVER guess or invent values):
{_field_spec(steps, opts)}

Rules:
- "business_name" is only the business's own name. Drop asides meant for the reader such as
  "(NOT the one on main road)", "the one near the station", "I think" - put those in "notes".
- Enumerated fields must use one of the allowed values exactly.
- "top_gaps": short specific observable gaps, formatted "1) ... 2) ... 3) ...".
- Rating is 1.0-5.0; reviews is a plain integer ("1.2k" -> 1200).
- If a business has no website, website_status is "None" and mobile_friendly is "No website".
- "stage" only if the manager explicitly says qualified / disqualified (or new / audited).
- "must_have_present": this niche needs online: "{must}". Judge whether the business currently offers it,
  from what the note says (Yes / Partly / No). Leave it out if the note gives no evidence either way.
- Put anything useful that fits no field (opening days, best time to call, owner remarks) in "notes".
- Understand English, Hinglish (Hindi/Bengali in English letters), spelled-out numbers ("three hundred and
  fifty" = 350, "4 point 5" = 4.5) and phone numbers written with spaces or dashes.
- Negations matter: "not that X has no website - they do have one" means they HAVE a website.
- If the manager is unsure ("maybe", "couldn't confirm"), use "Unknown" where allowed, else leave it out.
- Keep the manager's wording for names, areas, phone numbers and links.

Manager's note:
\"\"\"{text}\"\"\""""


def _norm_enum(value, allowed):
    if value in (None, ""):
        return None
    v = str(value).strip()
    for a in allowed:
        if a.casefold() == v.casefold():
            return a
    alias = ALIASES.get(v.casefold())
    if alias in allowed:
        return alias
    for a in allowed:
        if v.casefold() in a.casefold() or a.casefold() in v.casefold():
            return a
    hit = difflib.get_close_matches(v, allowed, 1, 0.6)
    return hit[0] if hit else None


def normalize_lead(lead, steps, opts):
    """Coerce a parsed lead to valid types/dropdown values. Returns (lead, warnings)."""
    out, warns = {}, []
    for key, (_c, label, _s, ftype) in fields_for_steps(steps).items():
        val = lead.get(key)
        if val in (None, ""):
            continue
        if ftype == "float":
            m = re.search(r"\d+(\.\d+)?", str(val))
            num = float(m.group()) if m else None
            if num is None or not 1 <= num <= 5:
                warns.append(f"{label}: '{val}' is not a valid rating")
                continue
            out[key] = num
        elif ftype == "int":
            s = str(val).lower().replace(",", "").strip()
            m = re.search(r"\d+(\.\d+)?", s)
            if not m:
                warns.append(f"{label}: '{val}' is not a number")
                continue
            out[key] = int(float(m.group()) * (1000 if s.endswith("k") else 1))
        elif ftype.startswith("enum:"):
            fixed = _norm_enum(val, opts[ftype.split(":")[1]])
            if fixed is None:
                warns.append(f"{label}: '{val}' did not match any allowed option")
            else:
                out[key] = fixed
        else:
            out[key] = str(val).strip()
    if out.get("website_status") == "None" and "mobile_friendly" not in out and 2 in steps:
        out["mobile_friendly"] = "No website"
    if out.get("stage") == "Disqualified" and "disqualify_reason" not in out and 3 in steps:
        out["disqualify_reason"] = "Other (see notes)"
    return out, warns


def parse_with_gemini(text, niche, steps, opts, api_key, model=DEFAULT_MODEL):
    """Returns (list_of_leads, warnings). Raises on API / JSON failure."""
    from google import genai
    from google.genai import errors, types

    client = genai.Client(api_key=api_key)
    prompt = build_prompt(text, niche, steps, opts)
    config = types.GenerateContentConfig(response_mime_type="application/json", temperature=0)
    resp, errs = None, []
    for m in [model] + [x for x in FALLBACK_MODELS if x != model]:
        for attempt in range(3):
            try:
                resp = client.models.generate_content(model=m, contents=prompt, config=config)
                break
            except errors.APIError as e:
                errs.append(e)
                if e.code in (429, 500, 503) and attempt < 2:
                    time.sleep(2 * (attempt + 1))  # busy: back off and retry, then try the next model
                    continue
                break  # 404 (model gone) or other error: next model
        if resp is not None:
            break
    if resp is None:
        # prefer an error that is not just "model retired", it says what really went wrong
        raise next((e for e in reversed(errs) if e.code != 404), errs[-1])
    data = json.loads(resp.text)
    raw = data["leads"] if isinstance(data, dict) and "leads" in data else (data if isinstance(data, list) else [data])
    leads, warns = [], []
    for item in raw:
        lead, w = normalize_lead(item, steps, opts)
        leads.append(lead)
        warns += w
    return leads, warns
