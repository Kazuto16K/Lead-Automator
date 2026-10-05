"""Field definitions: maps JSON keys <-> Excel columns on the Leads sheet, grouped by step."""

# key -> (Excel column letter, label, step, type)
# type: text | float | int | enum:<Lists column letter>
FIELDS = {
    # Step 1 - FIND
    "business_name":      ("C",  "Business Name",            1, "text"),
    "area":               ("D",  "Area / Locality",          1, "text"),
    "maps_link":          ("E",  "Google Maps Link",         1, "text"),
    "phone":              ("F",  "Phone",                    1, "text"),
    "rating":             ("G",  "Google Rating",            1, "float"),
    "reviews":            ("H",  "No. of Reviews",           1, "int"),
    # Step 2 - AUDIT
    "website_url":        ("K",  "Website URL",              2, "text"),
    "website_status":     ("L",  "Website Status",           2, "enum:C"),
    "mobile_friendly":    ("M",  "Mobile Friendly",          2, "enum:D"),
    "instagram_link":     ("N",  "Instagram Link",           2, "text"),
    "instagram_activity": ("O",  "Instagram Activity",       2, "enum:E"),
    "whatsapp":           ("P",  "WhatsApp Available?",      2, "enum:F"),
    "online_booking":     ("Q",  "Online Booking / Enquiry?", 2, "enum:F"),
    "must_have_present":  ("S",  "Must-Have Present?",       2, "enum:G"),
    "top_gaps":           ("T",  "Top 3 Gaps",               2, "text"),
    # Step 3 - QUALIFY
    "high_ticket":        ("V",  "High-Ticket Service?",     3, "enum:F"),
    "decision_maker":     ("W",  "Decision-Maker Name / Role", 3, "text"),
    "reachable":          ("X",  "Decision-Maker Reachable?", 3, "enum:H"),
    "stage":              ("AA", "Stage",                    3, "enum:I"),
    "disqualify_reason":  ("AB", "Disqualify Reason",        3, "enum:J"),
    "notes":              ("AC", "Notes",                    3, "text"),
}

STEPS = {
    1: "Step 1 · Find",
    2: "Step 2 · Audit",
    3: "Step 3 · Qualify",
}
STEP_DETAILS = {
    1: "business name, area, Maps link, phone, rating, reviews",
    2: "website, mobile, Instagram, WhatsApp, booking, gaps",
    3: "high-ticket, decision-maker, stage, notes",
}

STEP_HINTS = {
    1: "e.g. *Glow Studio Unisex Salon, Salt Lake Sector V, phone 98300 00000, 4.6 stars with 412 reviews, maps link https://maps.app.goo.gl/xxxx*",
    2: "e.g. *Glow Studio has no website. Instagram glowstudio is active, they have WhatsApp but no online booking. Prices only in Instagram highlights. Gaps: no website, scattered prices, can't book online.*",
    3: "e.g. *Bridal packages so high ticket. Owner is Rina, reachable. Mark as qualified. Best time to call is evenings.*",
}

# Other columns the app fills automatically
COL_NICHE, COL_ID, COL_DATE, COL_BY, COL_NAME = "B", "A", "I", "J", "C"
# Stage ordering used to avoid moving a lead backwards
STAGE_RANK = {"New": 1, "Audited": 2, "Qualified": 3, "Disqualified": 3}
AUTO_STAGE = {1: "New", 2: "Audited", 3: "Audited"}

FIRST_ROW, LAST_ROW = 3, 52


def fields_for_steps(steps):
    # business_name (matches a note to its lead) and notes (opening days, call times...) are always included
    return {k: v for k, v in FIELDS.items() if v[2] in steps or k in ("business_name", "notes")}
