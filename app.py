"""
Canopy Pricing Configurator — Streamlit app (v9 — Google Sheets quote tracker)
Reads all rates, cheat sheets, and defaults from canopy_config.xlsx (same folder).
Spec: 14 corrections through 2026-05-10. v9 swaps the local CSV tracker for a
Google Sheets backend (Streamlit Community Cloud has ephemeral storage that
wipes the CSV on every restart).

Run locally:
    streamlit run canopy_configurator.py
"""

import csv
import datetime as dt
import io
import math
import os
import tempfile
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

from proposal_writer import build_proposal

# Google Sheets backend (production). Falls back to local CSV when creds aren't
# configured (e.g., local dev), so the app still runs without google deps.
try:
    import gspread
    from google.oauth2.service_account import Credentials
    _HAS_GSPREAD = True
except ImportError:
    _HAS_GSPREAD = False

CONFIG_PATH = Path(__file__).parent / "canopy_config.xlsx"
APEC_LOGO   = Path(__file__).parent / "Apec Imaging Logo.jpg"
GEO_LOGO    = Path(__file__).parent / "GEO Canopies logo.jpg"
PASSWORD = "cheap"          # GP-override password (existing)
# Admin password — gates the Profitability Tracker view at the bottom of the
# app. CHANGE THIS to whatever you want. Keep it different from PASSWORD so
# GP-override leakage doesn't expose margins.
ADMIN_PASSWORD = "profit_tracker"

# ─── Quote tracker storage ─────────────────────────────────────────────────
# Production: Google Sheet on Ali's personal Drive.
# Local dev: falls back to CSV next to the app if no Streamlit secrets present.
SHEET_ID         = "1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w"
WORKSHEET_NAME   = "Tracker"
TRACKER_PATH     = Path(__file__).parent / "quote_tracker.csv"  # local-dev fallback
TRACKER_COLUMNS  = ["Date", "Quote No", "Customer Name", "City",
                    "Sales Rep", "Grand Total", "Profitability"]
GSHEETS_SCOPES   = ["https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"]


@st.cache_resource(show_spinner=False)
def _get_tracker_sheet():
    """Return a gspread worksheet handle, or None if creds aren't available.
    When None is returned, callers fall back to the local CSV path."""
    if not _HAS_GSPREAD:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
    except Exception:
        return None  # No secrets configured → CSV fallback
    try:
        creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEETS_SCOPES)
        gc    = gspread.authorize(creds)
        sh    = gc.open_by_key(SHEET_ID)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME,
                                  rows=1000, cols=len(TRACKER_COLUMNS))
            ws.append_row(TRACKER_COLUMNS, value_input_option="USER_ENTERED")
        # Ensure header row is present (in case sheet was wiped manually)
        if not ws.row_values(1):
            ws.append_row(TRACKER_COLUMNS, value_input_option="USER_ENTERED")
        return ws
    except Exception as e:
        # Surface auth/permission errors so the rep notices, but don't crash
        # the app — fall back to CSV so quotes still get logged somewhere.
        st.warning(f"Google Sheets tracker unavailable ({e}); using local CSV fallback.")
        return None

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

    def _num(cell, fallback):
        """Return cell value if it's numeric (cost column), else fallback (selling).
        Guards against placeholder strings like '(value is cost)' in unfilled cells."""
        v = cell.value if hasattr(cell, "value") else cell
        return v if isinstance(v, (int, float)) and v is not False else fallback

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
    # True costs for at-cost adders + shipping (fall back to selling = no profit)
    c["branded_labor_adder_cost"]    = _num(s["C18"], s["B18"].value)
    c["two_disp_labor_adder_cost"]   = _num(s["C19"], s["B19"].value)
    c["shipping_handling_cost"]      = _num(s["C20"], s["B20"].value)
    c["labor_daily_rate_cost"]       = _num(s["C17"], c["labor_daily_rate"])

    m = wb["Material_Rates"]
    c["acm_per_panel"]   = m["B3"].value
    c["steel_primary"]   = m["B4"].value
    c["steel_secondary"] = m["B5"].value
    c["decking_per_col"] = m["B6"].value
    c["light_unit_price"]= m["B7"].value
    # True-cost mirrors of the above (fall back to selling if Excel column C empty)
    c["acm_per_panel_cost"]    = _num(m["C3"], m["B3"].value)
    c["steel_primary_cost"]    = _num(m["C4"], m["B4"].value)
    c["steel_secondary_cost"]  = _num(m["C5"], m["B5"].value)
    c["decking_per_col_cost"]  = _num(m["C6"], m["B6"].value)
    c["light_unit_price_cost"] = _num(m["C7"], m["B7"].value)

    mt = wb["MISC_Tiers"]
    # selling tiers (col C) and parallel cost tiers (col D, fall back to col C)
    c["misc_tiers"] = [
        (mt.cell(r, 1).value, mt.cell(r, 2).value, mt.cell(r, 3).value)
        for r in range(3, 6) if mt.cell(r, 1).value is not None
    ]
    c["misc_tiers_cost"] = [
        (mt.cell(r, 1).value, mt.cell(r, 2).value,
         _num(mt.cell(r, 4), mt.cell(r, 3).value))
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
    c["mid_prices"]      = {}
    c["mid_prices_cost"] = {}
    for r in range(3, 8):
        brand, price = mp.cell(r, 1).value, mp.cell(r, 2).value
        if brand and price is not None:
            c["mid_prices"][brand] = price
            c["mid_prices_cost"][brand] = _num(mp.cell(r, 3), price)

    ml = wb["MID_Labor"]
    c["mid_labor"]      = {}
    c["mid_labor_cost"] = {}
    for r in range(3, 5):
        st_code, amt = ml.cell(r, 1).value, ml.cell(r, 2).value
        if st_code and amt is not None:
            c["mid_labor"][st_code] = amt
            c["mid_labor_cost"][st_code] = _num(ml.cell(r, 3), amt)

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


def lookup_misc_cost(columns, cfg):
    for lo, hi, cost in cfg["misc_tiers_cost"]:
        if lo <= columns <= hi:
            return cost
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

    # ── True-cost & profitability ──────────────────────────────────────
    acm_cost = 0 if branded else compute_acm(width, depth, cfg) * (
        cfg["acm_per_panel_cost"] / cfg["acm_per_panel"]
        if cfg["acm_per_panel"] else 0
    )
    steel_primary_cost = dispensers * cfg["steel_primary_cost"]
    decking_cost       = columns * cfg["decking_per_col_cost"]
    lights_cost        = dispensers * 4 * cfg["light_unit_price_cost"]
    misc_cost          = lookup_misc_cost(columns, cfg)
    steel_secondary_cost = (dispensers * cfg["steel_secondary_cost"]
                            if (canopy_type == "Dive-in" and double_col) else 0)
    shipping_cost      = cfg["shipping_handling_cost"] if branded else 0
    # Brand imaging: pure pass-through, selling = true cost
    brand_imaging_cost = brand_imaging_amt
    mid_material_cost  = (cfg["mid_prices_cost"].get(mid_brand, mid_material)
                          if include_mid and mid_brand else 0)
    base_labor_cost    = days * cfg["labor_daily_rate_cost"]
    branded_labor_add_cost  = cfg["branded_labor_adder_cost"] if branded else 0
    two_disp_labor_add_cost = cfg["two_disp_labor_adder_cost"] if dispensers == 2 else 0
    # MID labor: cost lookup by state where available; otherwise fall back to
    # the actual labor amount applied (preserves correctness for custom states)
    if include_mid:
        cust_state_default = None  # configurator passes labor amount directly
        mid_labor_cost = mid_labor_amt  # treat MID labor as at-cost pass-through
    else:
        mid_labor_cost = 0

    true_cost_material = (acm_cost + steel_primary_cost + decking_cost + lights_cost
                          + misc_cost + steel_secondary_cost + shipping_cost
                          + brand_imaging_cost + mid_material_cost)
    true_cost_labor    = (base_labor_cost + branded_labor_add_cost
                          + two_disp_labor_add_cost + mid_labor_cost)
    true_cost_total    = true_cost_material + true_cost_labor

    # Revenue excludes tax (pass-through to taxing authority)
    revenue = retail_material + retail_labor
    profit  = revenue - true_cost_total

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
        # Per-component costs — consumed by proposal_builder for retail line items.
        "acm":                acm,
        "steel":              steel_primary,
        "decking":            decking,
        "lights":             lights,
        "misc":               misc,
        "steel_secondary":    steel_secondary,
        "shipping":           shipping,
        "brand_imaging":      brand_imaging_amt,
        "mid_material":       mid_material,
        "base_labor":         base_labor,
        "branded_labor_add":  branded_add,
        "two_disp_labor_add": two_disp_add,
        "mid_labor":          mid_labor_total,
        # Profitability fields
        "revenue":            revenue,
        "true_cost":          true_cost_total,
        "profit":             profit,
    }


# ════════════════════════════════════════════════════════════════════════════
# Proposal generation helpers
# ════════════════════════════════════════════════════════════════════════════

def _build_proposal_data(q):
    """Translate Streamlit-session quote snapshot → proposal_builder data dict."""
    r = q["result"]
    return {
        "company_key":  q["company_key"],
        "quote_number": q["quote_number"],
        "quote_date":   q["quote_date"],
        "customer": {
            "company": q["cust_company"], "name":  q["cust_name"],
            "phone":   q["cust_phone"],   "email": q["cust_email"],
            "street":  q["cust_street"],  "city":  q["cust_city"],
            "state":   q["cust_state"],   "zip":   q["cust_zip"],
        },
        "sales_person": {
            "name":  q["sales_person"],
            "phone": q["sales_info"]["phone"],
            "email": q["sales_info"]["email"],
        },
        "canopy": {
            "type":        q["canopy_type"],
            "branded":     q["branded"],
            "brand_name":  q["brand_name"] or "",
            "double_col":  q["double_col"],
            "dispensers":  q["dispensers"],
            "columns":     r["columns"],
            "width":       q["width"],
            "depth":       q["depth"],
            "distance":    q["distance"],
            "overhang":    q["overhang"],
            "labor_days":  r["labor_days"],
        },
        "mid": {
            "include": q["include_mid"],
            "brand":   q["mid_brand"],
        },
        "pricing": {
            "gp_rate":            q["gp_rate"],
            "tax_rate":           q["tax_rate"],
            "acm":                r["acm"],
            "steel":              r["steel"],
            "decking":            r["decking"],
            "lights":             r["lights"],
            "misc":               r["misc"],
            "steel_secondary":    r["steel_secondary"],
            "shipping":           r["shipping"],
            "brand_imaging":      r["brand_imaging"],
            "mid_material":       r["mid_material"],
            "base_labor":         r["base_labor"],
            "branded_labor_add":  r["branded_labor_add"],
            "two_disp_labor_add": r["two_disp_labor_add"],
            "mid_labor":          r["mid_labor"],
            "items_not_included": r["items_not_included"],
        },
    }


def _docx_to_bytes(data):
    """Run build_proposal() to a temp file, return the docx bytes."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
        tmp = tf.name
    try:
        build_proposal(data, tmp)
        with open(tmp, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def _next_quote_number():
    """Return next sequential quote number based on existing tracker rows.
    Format: Q-YYYYMMDD-NNN with NNN = max existing today + 1 (else 001).
    Reads from Google Sheet in production, falls back to CSV locally."""
    today  = dt.date.today().strftime("%Y%m%d")
    prefix = f"Q-{today}-"
    max_seq = 0

    ws = _get_tracker_sheet()
    if ws is not None:
        try:
            quote_col = ws.col_values(2)[1:]  # column B = "Quote No", skip header
            for v in quote_col:
                if v and v.startswith(prefix):
                    try:
                        max_seq = max(max_seq, int(v.split("-")[-1]))
                    except (ValueError, IndexError):
                        pass
        except Exception:
            pass  # sheet unreachable mid-session → start fresh today
    elif TRACKER_PATH.exists():
        try:
            with open(TRACKER_PATH, encoding="utf-8") as f:
                reader = csv.reader(f)
                next(reader, None)  # skip header
                for row in reader:
                    if len(row) >= 2 and row[1].startswith(prefix):
                        try:
                            max_seq = max(max_seq, int(row[1].split("-")[-1]))
                        except (ValueError, IndexError):
                            pass
        except Exception:
            pass  # corrupt tracker → start fresh
    return f"{prefix}{max_seq + 1:03d}"


def _log_quote_to_tracker(q):
    """Append a single row to the profitability tracker.
    Uses Google Sheet in production, falls back to local CSV when no creds."""
    r = q["result"]
    row = [
        q["quote_date"],
        q["quote_number"],
        q["cust_company"],
        q["cust_city"],
        q["sales_person"],
        f'{r["final"]:.2f}',
        f'{r["profit"]:.2f}',
    ]
    ws = _get_tracker_sheet()
    if ws is not None:
        ws.append_row(row, value_input_option="USER_ENTERED")
        return
    # CSV fallback
    is_new = not TRACKER_PATH.exists()
    with open(TRACKER_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if is_new:
            w.writerow(TRACKER_COLUMNS)
        w.writerow(row)


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

    # Snapshot inputs + result so display and proposal both see the SAME data.
    # Quote number is NOT assigned here — only when Generate Proposal is clicked
    # (sequential, server-wide; pulled from the tracker file at that moment).
    st.session_state["last_quote"] = {
        "result":         result,
        "company_key":    company_key,
        "brand_company":  brand_company,
        "cust_company":   cust_company, "cust_name":  cust_name,
        "cust_phone":     cust_phone,   "cust_email": cust_email,
        "cust_street":    cust_street,  "cust_city":  cust_city,
        "cust_state":     cust_state,   "cust_zip":   cust_zip,
        "sales_person":   sales_person, "sales_info": sales_info,
        "canopy_type":    canopy_type,  "branded":    branded,
        "brand_name":     brand_name,   "double_col": double_col,
        "include_mid":    include_mid,  "mid_brand":  mid_brand,
        "dispensers":     dispensers,   "width":      width,
        "depth":          depth,        "distance":   distance,
        "overhang":       overhang,
        "tax_rate":       tax_rate,     "tax_rate_pct": tax_rate_pct,
        "gp_rate":        gp,
        "quote_number":   None,  # assigned by _next_quote_number() on Generate
        "quote_date":     dt.date.today().strftime("%B %d, %Y"),
    }
    # New calc invalidates any previously generated proposal + quote-logged flag.
    st.session_state.pop("proposal_bytes", None)
    st.session_state.pop("proposal_filename", None)
    st.session_state.pop("quote_logged", None)


# ── Display + Generate Proposal — driven by the snapshot ────────────────────
if "last_quote" in st.session_state:
    q = st.session_state["last_quote"]
    result = q["result"]

    st.markdown('<div class="section-title">Quote Result</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="cost-card">'
        f'<div class="label">Canopy Cost (internal)</div>'
        f'<div class="value">Material: ${result["material_cost"]:,.2f} '
        f'&nbsp;|&nbsp; Installation: ${result["labor_cost"]:,.2f}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div class="retail-card">'
        f'<div class="label">Retail Price</div>'
        f'<div class="value">${result["retail_total"]:,.2f}</div>'
        f'<div style="margin-top:8px;font-size:0.95rem;color:#cde4f7;">'
        f'Material: ${result["retail_material"]:,.2f} &nbsp;|&nbsp; '
        f'Installation: ${result["retail_labor"]:,.2f} &nbsp;|&nbsp; '
        f'Tax ({q["tax_rate_pct"]:.2f}%): ${result["tax"]:,.2f}</div>'
        f'<div style="margin-top:14px;font-size:1.1rem;font-weight:700;">'
        f'Final Price for Customer: ${result["final"]:,.2f}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Project Summary</div>', unsafe_allow_html=True)
    col_arr_label = ("Double column" if (q["canopy_type"] == "Dive-in" and q["double_col"])
                     else "Single column")
    summary = [
        ("Quote for", q["brand_company"]),
        ("Customer", f'{q["cust_company"]} — {q["cust_name"]}'),
        ("Customer contact", f'{q["cust_phone"]} | {q["cust_email"]}'),
        ("Site address", f'{q["cust_street"]}, {q["cust_city"]}, {q["cust_state"]} {q["cust_zip"]}'),
        ("Sales person", f'{q["sales_person"]} ({q["sales_info"]["phone"]} | {q["sales_info"]["email"]})'),
        ("Canopy type", f'{q["canopy_type"]} ({col_arr_label})' if q["canopy_type"] == "Dive-in" else q["canopy_type"]),
        ("Branded?", f'Yes — {q["brand_name"]}' if q["branded"] else "No"),
        ("Price sign?", f'Yes — {q["mid_brand"]}' if (q["include_mid"] and q["mid_brand"]) else ("Yes — jobber-supplied" if q["include_mid"] else "No")),
        ("Dispensers / Columns", f'{q["dispensers"]} / {result["columns"]}'),
        ("Width × Depth", f'{q["width"]:.0f} ft × {q["depth"]:.0f} ft'),
        ("Square footage", f'{int(q["width"]*q["depth"]):,} sq ft'),
        ("Installation days", result["labor_days"]),
        ("Tax rate", f'{q["tax_rate_pct"]:.2f}%'),
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

    # ── Generate Proposal ────────────────────────────────────────────────
    st.markdown('<div class="section-title">Generate Proposal</div>', unsafe_allow_html=True)
    gen = st.button("📄 Generate Word Proposal", use_container_width=True)
    if gen:
        # First Generate on a fresh calc → assign quote# and log to tracker.
        # Subsequent Generate clicks on the same calc reuse the same number
        # and skip logging (no duplicate rows).
        if not st.session_state.get("quote_logged"):
            q["quote_number"] = _next_quote_number()
            st.session_state["last_quote"] = q
            try:
                _log_quote_to_tracker(q)
                st.session_state["quote_logged"] = True
            except Exception as e:
                st.warning(f"Quote tracker append failed (proposal still generated): {e}")
        data = _build_proposal_data(q)
        try:
            st.session_state["proposal_bytes"] = _docx_to_bytes(data)
            st.session_state["proposal_filename"] = f'Proposal_{q["quote_number"]}.docx'
            st.success(f'Proposal generated — {q["quote_number"]}.')
        except Exception as e:
            st.error(f"Could not generate proposal: {e}")

    if "proposal_bytes" in st.session_state:
        st.download_button(
            "⬇️ Download Proposal (.docx)",
            data=st.session_state["proposal_bytes"],
            file_name=st.session_state["proposal_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )



# ════════════════════════════════════════════════════════════════════════════
# Profitability Tracker — admin view at bottom of page (password-gated).
# ════════════════════════════════════════════════════════════════════════════
st.markdown("---")
with st.expander("🔒 Profitability Tracker (admin)"):
    admin_pw = st.text_input("Admin password", type="password", key="admin_pw")
    if admin_pw == ADMIN_PASSWORD:
        ws = _get_tracker_sheet()
        rows = []
        source_label = ""
        try:
            if ws is not None:
                rows = ws.get_all_values()
                source_label = "Google Sheets"
                st.markdown(
                    f"**Source:** [Open tracker in Google Sheets ↗]"
                    f"(https://docs.google.com/spreadsheets/d/{SHEET_ID})"
                )
            elif TRACKER_PATH.exists():
                with open(TRACKER_PATH, encoding="utf-8") as _f:
                    rows = list(csv.reader(_f))
                source_label = "Local CSV (fallback — Google Sheets not configured)"
                st.markdown(f"**Source:** {source_label}")
        except Exception as e:
            st.error(f"Tracker read error: {e}")
            rows = []

        if len(rows) >= 2:
            header, data_rows = rows[0], rows[1:]
            # Newest on top
            display_rows = list(reversed(data_rows))
            table_data = [dict(zip(header, r)) for r in display_rows]
            st.markdown(f"**{len(data_rows)} quote(s) logged.**")
            st.dataframe(table_data, use_container_width=True, hide_index=True)
            # Build CSV bytes for download (chronological order preserved)
            buf = io.StringIO()
            w = csv.writer(buf)
            w.writerow(header)
            for r_ in data_rows:
                w.writerow(r_)
            st.download_button(
                "⬇️ Download Tracker (.csv)",
                data=buf.getvalue().encode("utf-8"),
                file_name=f"quote_tracker_{dt.date.today()}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.info("No quotes logged yet — tracker will populate after the "
                    "first proposal is generated.")
    elif admin_pw:
        st.error("Incorrect admin password.")
