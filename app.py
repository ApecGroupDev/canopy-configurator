"""
Canopy Pricing Configurator — Streamlit app (v6)
Reads all rates, cheat sheets, and defaults from canopy_config.xlsx (same folder).
Spec: 14 corrections through 2026-05-10. v6 adds GEO Canopies logo swap.

Run locally:
    streamlit run canopy_configurator.py
"""

import math
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

CONFIG_PATH = Path(__file__).parent / "canopy_config.xlsx"
APEC_LOGO   = Path(__file__).parent / "Apec Imaging Logo.jpg"
GEO_LOGO    = Path(__file__).parent / "GEO Canopies logo.jpg"
PASSWORD = "cheap"

US_STATES = [
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
    "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
    "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
    "TX","UT","VT","VA","WA","WV","WI","WY","DC",
]


@st.cache_data(ttl=60)
def load_config():
    wb = load_workbook(CONFIG_PATH, data_only=True)
    c = {}

    s = wb["Settings"]
    c["canopy_height_default_ft"]    = s["B3"].value
    c["default_dispensers"]          = s["B4"].value
    c["default_overhang_ft"]         = s["B5"].value
    c["default_distance_dive_in"]    = s["B6"].value
    c["default_distance_stacked"]    = s["B7"].value
    c["default_distance_in_line"]    = s["B8"].value
    c["depth_dive_in"]               = s["B9"].value
    c["depth_stacked"]               = s["B10"].value
    c["depth_in_line"]               = s["B11"].value
    c["acm_panel_width_ft"]          = s["B12"].value
    c["tax_default_rate"]            = s["B13"].value
    c["laborer_count"]               = s["B14"].value
    c["hours_per_day"]               = s["B15"].value
    c["hourly_rate"]                 = s["B16"].value
    raw = s["B17"].value
    c["labor_daily_rate"] = raw if raw is not None else (
        c["laborer_count"] * c["hours_per_day"] * c["hourly_rate"]
    )
    c["branded_labor_adder"]         = s["B18"].value
    c["two_disp_labor_adder"]        = s["B19"].value
    c["shipping_handling"]           = s["B20"].value

    m = wb["Material_Rates"]
    c["acm_per_panel"]   = m["B3"].value
    c["steel_primary"]   = m["B4"].value
    c["steel_secondary"] = m["B5"].value
    c["decking_per_col"] = m["B6"].value
    c["light_unit_price"]= m["B7"].value

    mt = wb["MISC_Tiers"]
    c["misc_tiers"] = [
        (mt.cell(r, 1).value, mt.cell(r, 2).value, mt.cell(r, 3).value)
        for r in range(3, 6) if mt.cell(r, 1).value is not None
    ]

    ld = wb["Labor_Days"]
    c["labor_days"] = {}
    for r in range(3, 12):
        d, days = ld.cell(r, 1).value, ld.cell(r, 2).value
        if d is not None and days is not None:
            c["labor_days"][int(d)] = int(days)

    bi = wb["Brand_Imaging"]
    c["brand_imaging"] = {}
    for r in range(3, 10):
        brand = bi.cell(r, 1).value
        if not brand: continue
        prices = {}
        for disp in range(2, 11):
            v = bi.cell(r, disp).value
            if v is not None:
                prices[disp] = v
        c["brand_imaging"][brand] = prices

    mp = wb["MID_PriceSign"]
    c["mid_prices"] = {}
    for r in range(3, 8):
        brand, price = mp.cell(r, 1).value, mp.cell(r, 2).value
        if brand and price is not None:
            c["mid_prices"][brand] = price

    ml = wb["MID_Labor"]
    c["mid_labor"] = {}
    for r in range(3, 5):
        st_code, amt = ml.cell(r, 1).value, ml.cell(r, 2).value
        if st_code and amt is not None:
            c["mid_labor"][st_code] = amt

    sp = wb["Salespeople"]
    c["salespeople"] = {}
    for r in range(3, 6):
        name = sp.cell(r, 1).value
        if name:
            c["salespeople"][name] = {"phone": sp.cell(r, 2).value, "email": sp.cell(r, 3).value}

    ipl = wb["Internal_PL"]
    c["gp_apec"] = ipl["B6"].value
    c["gp_geo"]  = ipl["B7"].value
    return c


def auto_dimensions(canopy_type, dispensers, distance, overhang, cfg):
    if canopy_type == "Dive-in":
        return 2*overhang + (dispensers-1)*distance, cfg["depth_dive_in"]
    if canopy_type == "Stacked":
        return 2*overhang + (dispensers//2 - 1)*distance, cfg["depth_stacked"]
    if canopy_type == "In-line":
        return 2*overhang + (dispensers-1)*distance, cfg["depth_in_line"]
    return 0, 0


def column_count(canopy_type, dispensers, double_col):
    return 2*dispensers if (canopy_type == "Dive-in" and double_col) else dispensers


def compute_acm(width, depth, cfg):
    a1 = 2*width + 2*depth
    a2 = math.ceil(a1 / cfg["acm_panel_width_ft"])
    return a2 * cfg["acm_per_panel"]


def lookup_misc(columns, cfg):
    for lo, hi, price in cfg["misc_tiers"]:
        if lo <= columns <= hi:
            return price
    return 0


def lookup_brand_imaging(brand, dispensers, cfg):
    table = cfg["brand_imaging"].get(brand, {})
    if dispensers in table:
        return table[dispensers], True, ""
    if not table:
        return 0, False, f"Brand imaging for {brand} is not included (jobber-supplied)."
    return 0, False, f"Brand imaging for {brand} not in cheat sheet at {dispensers} dispensers."


def compute_quote(*, dispensers, canopy_type, double_col, branded, brand_name,
                  include_mid, mid_brand, mid_labor_amt, width, depth,
                  tax_rate, gp_rate, cfg):
    columns = column_count(canopy_type, dispensers, double_col)
    items_not_included = []

    acm = 0 if branded else compute_acm(width, depth, cfg)
    steel_primary = dispensers * cfg["steel_primary"]
    decking = columns * cfg["decking_per_col"]
    lights = dispensers * 4 * cfg["light_unit_price"]
    misc = lookup_misc(columns, cfg)
    marked_up_material = acm + steel_primary + decking + lights + misc

    steel_secondary = (dispensers * cfg["steel_secondary"]
                       if (canopy_type == "Dive-in" and double_col) else 0)
    shipping = cfg["shipping_handling"] if branded else 0
    brand_imaging_amt = 0
    if branded:
        amt, ok, msg = lookup_brand_imaging(brand_name, dispensers, cfg)
        brand_imaging_amt = amt
        if not ok:
            items_not_included.append(msg)
    else:
        items_not_included.append("Brand imaging is not included.")

    mid_material = 0
    if include_mid:
        if mid_brand and mid_brand in cfg["mid_prices"]:
            mid_material = cfg["mid_prices"][mid_brand]
        else:
            items_not_included.append(f"MID/Price sign material for {mid_brand} not in cheat sheet.")
    else:
        items_not_included.append("MID / Price sign is not included.")

    at_cost_material = steel_secondary + shipping + brand_imaging_amt + mid_material

    days = cfg["labor_days"].get(int(dispensers), 0)
    base_labor = days * cfg["labor_daily_rate"]
    marked_up_labor = base_labor

    branded_add = cfg["branded_labor_adder"] if branded else 0
    two_disp_add = cfg["two_disp_labor_adder"] if dispensers == 2 else 0
    mid_labor_total = mid_labor_amt if include_mid else 0
    at_cost_labor = branded_add + two_disp_add + mid_labor_total

    retail_material = marked_up_material * (1 + gp_rate) + at_cost_material
    retail_labor    = marked_up_labor * (1 + gp_rate) + at_cost_labor
    tax = retail_material * tax_rate

    material_cost = marked_up_material + at_cost_material
    labor_cost    = marked_up_labor + at_cost_labor
    final = retail_material + retail_labor + tax

    return {
        "material_cost":    material_cost,
        "labor_cost":       labor_cost,
        "cost_total":       material_cost + labor_cost,
        "retail_material":  retail_material,
        "retail_labor":     retail_labor,
        "retail_total":     retail_material + retail_labor,
        "tax":              tax,
        "final":            final,
        "columns":          columns,
        "labor_days":       days,
        "items_not_included": items_not_included,
    }


# ════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Canopy Configurator", page_icon="⛽", layout="centered")

st.markdown("""
<style>
body { font-family: 'Segoe UI', sans-serif; }
.apec-banner { background: linear-gradient(135deg, #0d1b2a 0%, #1b3a5c 60%, #1a78c2 100%);
    border-radius: 12px; padding: 28px 24px 18px 24px; text-align: center; margin-bottom: 24px; }
.apec-logo-text { font-size: 2.6rem; font-weight: 900; letter-spacing: 3px; color: #f5a623;
    text-shadow: 2px 2px 6px rgba(0,0,0,0.5); }
.apec-sub { font-size: 1.05rem; color: #cde4f7; letter-spacing: 1px; margin-top: 4px; }
.section-title { font-size: 1.05rem; font-weight: 700; color: #1b3a5c;
    border-left: 4px solid #1a78c2; padding-left: 10px; margin: 22px 0 10px 0; }
.auto-readout { background: #f4f8fd; border: 1px dashed #b3d4f0; border-radius: 6px;
    padding: 8px 12px; font-size: 0.88rem; color: #1b3a5c; margin-top: 4px; }
.cost-card { background: #f4f8fd; border: 1px solid #b3d4f0; border-radius: 10px;
    padding: 16px 22px; margin: 12px 0; }
.cost-card .label { font-size: 0.85rem; color: #1b3a5c; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.5px; }
.cost-card .value { font-size: 1.4rem; color: #1b3a5c; font-weight: 700; }
.retail-card { background: linear-gradient(135deg, #1b3a5c 0%, #2a5a8c 100%);
    border-radius: 12px; padding: 24px 28px; margin: 14px 0; color: #fff; }
.retail-card .label { font-size: 1rem; color: #cde4f7; font-weight: 700;
    text-transform: uppercase; letter-spacing: 1px; }
.retail-card .value { font-size: 2.2rem; font-weight: 900; color: #f5a623; }
.summary-row { padding: 6px 0; border-bottom: 1px dotted #ccd9e8; font-size: 0.92rem; }
.summary-row .k { color: #6a7a8c; }
.summary-row .v { color: #1b3a5c; font-weight: 600; }
.disclaimer { background: #fff8e1; border-left: 4px solid #f5a623; padding: 12px 16px;
    font-size: 0.85rem; color: #5a4a1a; margin-top: 18px; border-radius: 6px; }
</style>
""", unsafe_allow_html=True)

try:
    cfg = load_config()
except FileNotFoundError:
    st.error(f"⚠️ `canopy_config.xlsx` not found at {CONFIG_PATH}.")
    st.stop()

brand_company = st.radio("Is this quote for APEC Canopies or GEO Canopies?",
                         ["APEC Canopies", "GEO Canopies"], horizontal=True)
company_key = "APEC" if "APEC" in brand_company else "GEO"
headline = ("APEC Imaging & Canopies — Price Configurator"
            if company_key == "APEC" else "GEO Canopies — Price Configurator")

active_logo = APEC_LOGO if company_key == "APEC" else GEO_LOGO
if active_logo.exists():
    st.image(str(active_logo), width=260)
st.markdown(
    f'<div class="apec-banner">'
    f'<div class="apec-logo-text">{"APEC" if company_key=="APEC" else "GEO"}</div>'
    f'<div class="apec-sub">{headline}</div></div>',
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">Customer Information</div>', unsafe_allow_html=True)
ca, cb = st.columns(2)
with ca:
    cust_company = st.text_input("Company Name")
    cust_name    = st.text_input("Customer Name")
    cust_phone   = st.text_input("Cell Phone")
    cust_email   = st.text_input("Email")
with cb:
    cust_street  = st.text_input("Street Address")
    cust_city    = st.text_input("City")
    cust_state   = st.selectbox("State", US_STATES, index=US_STATES.index("GA"))
    cust_zip     = st.text_input("Zip")

st.markdown('<div class="section-title">Sales Person</div>', unsafe_allow_html=True)
sales_person = st.selectbox("Sales Person", list(cfg["salespeople"].keys()))
sales_info = cfg["salespeople"][sales_person]
st.markdown(
    f'<div class="auto-readout"><b>Cell:</b> {sales_info["phone"]} &nbsp;|&nbsp; '
    f'<b>Email:</b> {sales_info["email"]}</div>',
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">Canopy Height</div>', unsafe_allow_html=True)
height_choice = st.radio("Canopy height", ["16.5 ft (standard)", "20 ft"], horizontal=True, index=0)
if height_choice == "20 ft":
    st.error("⚠️ This configurator does not yet support 20 ft canopies.")
    st.stop()

st.markdown('<div class="section-title">Canopy Configuration</div>', unsafe_allow_html=True)
canopy_type = st.radio("Canopy type", ["Dive-in", "Stacked", "In-line"], horizontal=True, index=0)
dispensers = int(st.number_input("Number of dispensers",
                                 min_value=2, max_value=10,
                                 value=int(cfg["default_dispensers"]), step=1))

if canopy_type == "Stacked" and dispensers % 2 != 0:
    st.error("⚠️ Stacked canopies require an EVEN number of dispensers (2 rows).")
    st.stop()

double_col = False
if canopy_type == "Dive-in":
    col_arr = st.radio("Column arrangement",
                       ["Single column", "Double column"], horizontal=True, index=0)
    double_col = (col_arr == "Double column")

columns = column_count(canopy_type, dispensers, double_col)
st.markdown(f'<div class="auto-readout">Columns: <b>{columns}</b> (auto-calculated)</div>',
            unsafe_allow_html=True)

if canopy_type == "In-line":
    default_distance = cfg["default_distance_in_line"]
elif canopy_type == "Stacked":
    default_distance = cfg["default_distance_stacked"]
else:
    default_distance = cfg["default_distance_dive_in"]

override_distance = st.checkbox(f"Override distance (default {default_distance} ft)", value=False)
distance = (st.number_input("Distance (ft)", 1.0, 100.0, float(default_distance), 1.0)
            if override_distance else float(default_distance))

override_overhang = st.checkbox(f"Override overhang (default {cfg['default_overhang_ft']} ft)", value=False)
overhang = (st.number_input("Overhang (ft)", 1.0, 50.0, float(cfg["default_overhang_ft"]), 1.0)
            if override_overhang else float(cfg["default_overhang_ft"]))

auto_w, auto_d = auto_dimensions(canopy_type, dispensers, distance, overhang, cfg)
override_width = st.checkbox(f"Override width (auto: {auto_w} ft)", value=False)
width = (st.number_input("Width (ft)", 1.0, 500.0, float(auto_w), 1.0)
         if override_width else float(auto_w))
override_depth = st.checkbox(f"Override depth (auto: {auto_d} ft)", value=False)
depth = (st.number_input("Depth (ft)", 1.0, 500.0, float(auto_d), 1.0)
         if override_depth else float(auto_d))

labor_days_preview = cfg["labor_days"].get(int(dispensers), 0)
st.markdown(
    f'<div class="auto-readout">Labor days: <b>{labor_days_preview}</b> '
    f'&nbsp;|&nbsp; Square footage: <b>{int(width*depth):,}</b></div>',
    unsafe_allow_html=True,
)

st.markdown('<div class="section-title">Branded Site</div>', unsafe_allow_html=True)
branded_choice = st.radio("Is this a branded site?", ["No (Unbranded)", "Yes (Branded)"],
                          horizontal=True, index=0)
branded = (branded_choice == "Yes (Branded)")
brand_name = None
if branded:
    brand_name = st.selectbox("Brand", list(cfg["brand_imaging"].keys()), index=0)
    st.caption("Branded site → adds brand imaging at cost, $3,500 shipping at cost, "
               "$2,800 labor at cost. ACM NOT charged (covered in brand imaging).")

# ── MID / Price Sign — brand auto-derived from canopy brand (correction #14) ──
st.markdown('<div class="section-title">MID / Price Sign</div>', unsafe_allow_html=True)
include_mid = st.radio("Include MID (Main ID Sign / Price Sign)?",
                       ["No", "Yes"], horizontal=True, index=0) == "Yes"
mid_brand = None
mid_labor_amt = 0
mid_jobber_note = None

if include_mid:
    # MID brand DERIVED from canopy brand — branded canopy + branded sign must match;
    # unbranded canopy can only have an Unbranded sign.
    if branded and brand_name:
        if brand_name == "Chevron Level I":
            chevron_sign = st.radio("Chevron sign style", ["C45", "C60"],
                                    horizontal=True, index=0,
                                    help="Chevron offers two sign sizes — pick which the customer wants.")
            mid_brand = f"Chevron {chevron_sign}"
            st.info(f"MID brand: **{mid_brand}** (matches canopy brand).")
        elif brand_name in ("Shell", "Valero"):
            mid_brand = brand_name
            st.info(f"MID brand auto-set to **{mid_brand}** (matches canopy brand).")
        else:
            # Jobber-supplied brands (Exxon / BP / Marathon / Citgo) — no MID in cheat sheet
            mid_brand = None
            st.warning(f"MID/Price sign for **{brand_name}** is provided by the jobber "
                       "— not in our cheat sheet, not included in this quote.")
            mid_jobber_note = (f"MID / Price sign for {brand_name} is provided by the jobber "
                               "(not included in this quote).")
    else:
        # Unbranded canopy → must be Unbranded MID
        mid_brand = "Unbranded"
        st.info("Canopy is Unbranded → MID brand auto-set to **Unbranded** "
                "(cannot pair a branded sign with an unbranded canopy).")

    # Labor only if mid_brand resolved
    if mid_brand is not None:
        if cust_state in cfg["mid_labor"]:
            st.caption(f"MID labor for {cust_state}: ${cfg['mid_labor'][cust_state]:,} (auto-applied)")
            mid_labor_amt = cfg["mid_labor"][cust_state]
            if st.checkbox("Override MID labor amount", value=False):
                mid_labor_amt = int(st.number_input("Custom MID labor ($)",
                                                     min_value=0, value=mid_labor_amt, step=500))
        else:
            st.warning(f"State {cust_state} not in MID labor cheat sheet — enter custom amount:")
            mid_labor_amt = int(st.number_input("Custom MID labor ($)",
                                                 min_value=0,
                                                 value=int(cfg["mid_labor"].get("GA", 11000)),
                                                 step=500))

tax_rate_pct = st.number_input("Tax rate (%)", 0.0, 20.0,
                                float(cfg["tax_default_rate"])*100, 0.25)
tax_rate = tax_rate_pct / 100.0

st.markdown('<div class="section-title">GP Override (optional)</div>', unsafe_allow_html=True)
default_gp = cfg["gp_apec"] if company_key == "APEC" else cfg["gp_geo"]
gp = default_gp
if st.checkbox("Override GP %", value=False):
    pwd = st.text_input("Password", type="password")
    if pwd == PASSWORD:
        gp_pct = st.number_input("Custom GP (%)", 0.0, 200.0, default_gp*100, 1.0)
        gp = gp_pct / 100.0
    elif pwd:
        st.error("Incorrect password — using standard GP.")

st.markdown('<br>', unsafe_allow_html=True)
calc = st.button("⚡ Calculate Canopy Price", type="primary", use_container_width=True)

if calc:
    include_mid_effective = include_mid and (mid_brand is not None)
    result = compute_quote(
        dispensers=dispensers, canopy_type=canopy_type, double_col=double_col,
        branded=branded, brand_name=brand_name,
        include_mid=include_mid_effective, mid_brand=mid_brand, mid_labor_amt=mid_labor_amt,
        width=width, depth=depth, tax_rate=tax_rate, gp_rate=gp, cfg=cfg,
    )
    if mid_jobber_note:
        result["items_not_included"].append(mid_jobber_note)

    st.markdown('<div class="section-title">Quote Result</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="cost-card">'
        f'<div class="label">Canopy Cost (internal)</div>'
        f'<div class="value">Material: ${result["material_cost"]:,.2f} '
        f'&nbsp;|&nbsp; Labor: ${result["labor_cost"]:,.2f}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div class="retail-card">'
        f'<div class="label">Retail Price</div>'
        f'<div class="value">${result["retail_total"]:,.2f}</div>'
        f'<div style="margin-top:8px;font-size:0.95rem;color:#cde4f7;">'
        f'Material: ${result["retail_material"]:,.2f} &nbsp;|&nbsp; '
        f'Labor: ${result["retail_labor"]:,.2f} &nbsp;|&nbsp; '
        f'Tax ({tax_rate_pct:.2f}%): ${result["tax"]:,.2f}</div>'
        f'<div style="margin-top:14px;font-size:1.1rem;font-weight:700;">'
        f'Final Price for Customer: ${result["final"]:,.2f}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Project Summary</div>', unsafe_allow_html=True)
    col_arr_label = ("Double column" if (canopy_type == "Dive-in" and double_col) else "Single column")
    summary = [
        ("Quote for", brand_company),
        ("Customer", f"{cust_company} — {cust_name}"),
        ("Customer contact", f"{cust_phone} | {cust_email}"),
        ("Site address", f"{cust_street}, {cust_city}, {cust_state} {cust_zip}"),
        ("Sales person", f"{sales_person} ({sales_info['phone']} | {sales_info['email']})"),
        ("Canopy type", f"{canopy_type} ({col_arr_label})" if canopy_type == "Dive-in" else canopy_type),
        ("Branded?", f"Yes — {brand_name}" if branded else "No"),
        ("MID / Price sign?", f"Yes — {mid_brand}" if (include_mid and mid_brand) else ("Yes — jobber-supplied" if include_mid else "No")),
        ("Dispensers / Columns", f"{dispensers} / {result['columns']}"),
        ("Width × Depth", f"{width:.0f} ft × {depth:.0f} ft"),
        ("Square footage", f"{int(width*depth):,} sq ft"),
        ("Labor days", result["labor_days"]),
        ("Tax rate", f"{tax_rate_pct:.2f}%"),
    ]
    for k, v in summary:
        st.markdown(
            f'<div class="summary-row"><span class="k">{k}:</span> <span class="v">{v}</span></div>',
            unsafe_allow_html=True,
        )

    if result["items_not_included"]:
        st.markdown('<div class="section-title">Items Not Included</div>', unsafe_allow_html=True)
        for item in result["items_not_included"]:
            st.markdown(f"- {item}")

    st.markdown(
        '<div class="disclaimer">'
        '<b>Disclaimer:</b> Price does not include permit, electrical work, '
        'canopy piers/footers, and bricking.<br><br>'
        '<i>This is an experimental configurator and not to be used in '
        'quoting real jobs at this time.</i>'
        '</div>',
        unsafe_allow_html=True,
    )
