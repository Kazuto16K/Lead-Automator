"""Lead Automator - Streamlit front end.  Run:  streamlit run app.py"""
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import auth
import excel_io as xl
import github_sync as gh
from llm_parser import DEFAULT_MODEL, normalize_lead, parse_with_gemini
from schema import FIELDS, STEP_DETAILS, STEP_HINTS, STEPS, fields_for_steps

st.set_page_config(page_title="Lead Automator", layout="wide", initial_sidebar_state="collapsed")

# Neon look & feel (base colours also set in .streamlit/config.toml)
st.markdown(f"<style>{(Path(__file__).parent / 'neon.css').read_text(encoding='utf-8')}</style>",
            unsafe_allow_html=True)

user = auth.require_login()  # nothing below runs (no data, no GitHub sync) until someone signs in


def _get_key():
    """Key comes from Streamlit secrets (hosted), else env var / local .env. Never shown in the UI."""
    try:
        if st.secrets.get("GEMINI_API_KEY"):
            return st.secrets["GEMINI_API_KEY"]
    except Exception:  # no secrets file locally
        pass
    if os.environ.get("GEMINI_API_KEY"):
        return os.environ["GEMINI_API_KEY"]
    env = Path(__file__).parent / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("GEMINI_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


api_key, model = _get_key(), DEFAULT_MODEL
gh_cfg = gh.config(st.secrets)  # None unless GITHUB_TOKEN / GITHUB_REPO secrets are set


@st.cache_resource
def _sync_on_start():
    """Runs once per app start: GitHub holds the master copy, so restore it over the (possibly fresh) local file."""
    if not gh_cfg:
        return "off"
    state = gh.pull(gh_cfg, xl.EXCEL_PATH)
    if state == "missing":  # first ever run: seed the data repo with the current file
        gh.push(gh_cfg, xl.EXCEL_PATH, "Initial upload")
        return "initialised"
    return state


def sync_push(message):
    """Upload after a change. Returns a warning text, or '' when fine / sync is off."""
    if not gh_cfg:
        return ""
    ok, text = gh.push(gh_cfg, xl.EXCEL_PATH, message[:200])
    return "" if ok else text


sync_state = _sync_on_start()

try:
    opts, _pts, _must = xl.read_lists()
    grid = xl.read_grid()
except PermissionError:
    st.error("The Excel file is open in another program. Close it and refresh.")
    st.stop()

st.title("Lead Automator")
who, out = st.columns([4, 1])
who.caption(f"Signed in as **{user}**. Type what you found in plain English and we fill the Excel tracker.")
if out.button("Sign out"):
    auth.sign_out()
    st.rerun()

banner = st.container()  # fixed slot above the tabs, so a new message does not reset the selected tab
with banner:
    if sync_state.startswith("error"):
        st.warning(f"Could not load the latest data from GitHub, showing the local copy. ({sync_state})")
    if st.session_state.get("err_msg"):
        st.error(st.session_state.pop("err_msg"))
    if st.session_state.get("saved_msg"):
        st.success(st.session_state.pop("saved_msg"))
    if st.session_state.get("sync_warn"):
        st.warning(st.session_state.pop("sync_warn") + ". Your change is saved here, so download the Excel as a precaution.")

tab_add, tab_view, tab_dash = st.tabs(["Add", "View", "Stats"])

# ------------------------------------------------------------ add / update
with tab_add:
    st.session_state.setdefault("researched_by", user)  # the signed-in user, still editable
    researched_by = st.text_input("Your name", key="researched_by", help="Saved as 'Researched By' on new leads")
    niche = st.selectbox("Which niche?", opts["A"])
    st.markdown("**What kind of details are you giving?**")
    picked = st.pills("Select all that apply", list(STEPS.values()), selection_mode="multi",
                      default=[STEPS[1]], label_visibility="collapsed")
    steps = {k for k, v in STEPS.items() if v in (picked or [])}
    if steps:
        st.markdown("".join(
            f'<div class="step-card s{n}"><div class="step-num">{n}</div><div>'
            f'<div class="step-title">{STEPS[n]}</div>'
            f'<div class="step-fields">Include: {STEP_DETAILS[n]}</div>'
            f'<div class="step-eg">{STEP_HINTS[n].replace("*", "")}</div></div></div>'
            for n in sorted(steps)), unsafe_allow_html=True)
    if not steps:
        st.info("Tap at least one option above to continue.")
        st.stop()

    selected_id = None
    if 1 not in steps:
        existing = xl.existing_leads(grid, niche)
        choice = st.selectbox("Which lead is this about?",
                              ["Auto-detect from the business name in my text"] +
                              [f"{i} · {n}" for i, n in existing])
        if " · " in choice:
            selected_id = choice.split(" · ")[0]
        if not existing:
            st.warning(f"No {niche} leads yet. Add them with Step 1 first (or include Step 1 here).")

    st.markdown("**Write your note**")
    # A form submits the text and the click together (a bare text_area only commits on blur,
    # which swallowed the first click on Parse).
    with st.form("note_form", border=False):
        text = st.text_area("Your note", height=170, label_visibility="collapsed",
                            placeholder="Write in plain English. You can describe more than one business.")
        submitted = st.form_submit_button("Parse my note", type="primary")

    if submitted:
        if not text.strip():
            st.warning("Write your note first.")
        elif not api_key:
            st.error("The app is not configured yet (missing GEMINI_API_KEY secret). Contact the admin.")
        else:
            try:
                with st.spinner("Reading your note..."):
                    leads, warns = parse_with_gemini(text, niche, steps, opts, api_key, model)
                st.session_state["parsed"] = {"leads": leads, "warns": warns, "niche": niche,
                                              "steps": steps, "selected_id": selected_id}
            except Exception as e:  # noqa: BLE001 - show any API/JSON problem to the user
                st.error(f"Could not parse: {e}")

    parsed = st.session_state.get("parsed")
    if parsed:
        st.divider()
        st.subheader("Check what we understood")
        st.caption("Edit any cell if something is wrong. Nothing is written to Excel until you confirm.")
        for w in parsed["warns"]:
            st.warning(w)
        shown = ["business_name"] + [k for k in fields_for_steps(parsed["steps"]) if k != "business_name"]
        cfg = {}
        for k in shown:
            _col, label, _s, ftype = FIELDS[k]
            if ftype.startswith("enum:"):
                cfg[k] = st.column_config.SelectboxColumn(label, options=opts[ftype.split(":")[1]])
            elif ftype in ("float", "int"):
                cfg[k] = st.column_config.NumberColumn(label)
            else:
                cfg[k] = st.column_config.TextColumn(label)
        df = pd.DataFrame(parsed["leads"]).reindex(columns=shown)
        for k in shown:  # an empty column would default to float and break text/select editing
            df[k] = (pd.to_numeric(df[k], errors="coerce") if FIELDS[k][3] in ("float", "int")
                     else df[k].astype(object).where(df[k].notna(), None))
        edited = st.data_editor(df, column_config=cfg, hide_index=True, width="stretch",
                                num_rows="dynamic", key="editor")
        leads = [{k: v for k, v in r.items() if pd.notna(v)} for r in edited.to_dict("records")]
        leads = [normalize_lead(l, parsed["steps"], opts)[0] for l in leads]

        plan = xl.plan_rows(grid, parsed["niche"], leads, parsed["selected_id"], allow_new=1 in parsed["steps"])
        for lead, p in zip(leads, plan):
            name = lead.get("business_name") or "(no name)"
            if p["action"] == "nomatch":
                st.error(f"{name}: no existing lead with this name. Choose the lead in the list above, "
                         f"or include Step 1 to add it as a new lead.")
            elif p["action"] == "noname":
                st.error("A lead has no business name. Add one in the table above.")
            elif p["action"] == "extend":
                st.success(f"**{name}** - new lead **{p['lead_id']}** (all {parsed['niche']} slots are full, "
                           f"so an extra row is added)")
            elif p["action"] == "new":
                st.success(f"**{name}** - new lead **{p['lead_id']}**")
            else:
                st.info(f"**{name}** - update existing **{p['lead_id']}**")

        if st.button("Confirm & write to Excel", type="primary"):
            try:
                msgs = xl.write_leads(parsed["niche"], leads, parsed["steps"], researched_by,
                                      parsed["selected_id"])
                st.session_state.pop("parsed")
                st.session_state["sync_warn"] = sync_push("Update leads: " + "; ".join(msgs))
                st.session_state["saved_msg"] = "Saved to Excel.  \n" + "  \n".join(msgs)
                st.rerun()  # reload so the View tab and Download button show the new data
            except PermissionError:
                st.error("The Excel file is open in another program. Close it and try again.")

# ------------------------------------------------------------ view
df_all = xl.leads_dataframe(grid)


def _color(v):
    return {"Hot": "background-color:rgba(255,43,109,.22);color:#ff2b6d;font-weight:600",
            "Warm": "background-color:rgba(255,179,0,.20);color:#ffb300;font-weight:600",
            "Cold": "background-color:rgba(0,240,255,.16);color:#00f0ff;font-weight:600"}.get(v, "")


with tab_view:
    f1, f2 = st.columns(2)
    niche_f = f1.selectbox("Niche", ["All"] + opts["A"], key="vn")
    only_filled = f2.checkbox("Only leads already started", value=True)
    view = df_all if niche_f == "All" else df_all[df_all["Niche"] == niche_f]
    if only_filled:
        view = view[view["Business Name"].notna()]
    all_cols = st.toggle("Show all columns", value=False, help="Off = a compact view that fits a phone screen.")
    compact = ["Lead ID", "Business Name", "Area / Locality", "Priority", "Opportunity Score (/10)", "Stage"]
    view = view if all_cols else view[compact]
    if view.empty:
        st.info("No leads yet. Add your first one in the Add tab.")
    else:
        st.dataframe(view.style.map(_color, subset=["Priority"]), hide_index=True,
                     width="stretch", height=min(480, 38 * (len(view) + 1) + 3))
    st.markdown("**Delete a lead**")
    st.caption("The leads after it in the same niche move up one place, so there are no gaps.")
    with st.container(border=True):  # not an expander: those fold shut on every reload
        started = df_all[df_all["Business Name"].notna()]
        if started.empty:
            st.caption("No leads to delete yet.")
        else:
            options = {f"{r['Lead ID']} · {r['Business Name']}": r["Lead ID"] for _, r in started.iterrows()}

            def _delete():  # callback: runs before the page reloads, so it can reset the checkbox
                try:
                    msg = xl.delete_lead(options[st.session_state["del_pick"]])
                    st.session_state["sync_warn"] = sync_push(msg)
                    st.session_state["saved_msg"] = msg
                except PermissionError:
                    st.session_state["err_msg"] = "The Excel file is open in another program, nothing was deleted."
                st.session_state["del_sure"] = False

            st.selectbox("Which lead?", list(options), key="del_pick")
            sure = st.checkbox("Yes, delete this lead permanently", key="del_sure")
            st.button("Delete lead", disabled=not sure, on_click=_delete)
    st.download_button("Download Excel", data=xl.EXCEL_PATH.read_bytes(), file_name=xl.EXCEL_PATH.name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="Always the latest saved version.", type="primary")
    st.caption(f"{len(view)} rows · Score, Priority and Audit Status are calculated the same way as in Excel. "
               + ("Every change is also backed up to GitHub." if gh_cfg else "GitHub backup is off."))

with tab_dash:
    dash = xl.dashboard_dataframe(df_all)
    total = dash[dash["Niche"] == "TOTAL"].iloc[0]
    m = st.columns(2) + st.columns(2)  # 2x2 grid: reads well on a phone
    m[0].metric("Leads captured", f"{total['Captured']} / {total['Target']}")
    m[1].metric("Audits complete", total["Audit Complete"])
    m[2].metric("Hot leads", total["Hot"])
    m[3].metric("Qualified", total["Qualified"])
    st.dataframe(dash, hide_index=True, width="stretch")

auth.apply_cookie_changes()  # writes the "stay signed in" cookie after a login (no-op otherwise)
