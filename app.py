import streamlit as st
import math

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="APEC Canopy Configurator",
    page_icon="⛽",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    body { font-family: 'Segoe UI', sans-serif; }

    /* ---- header banner ---- */
    .apec-banner {
        background: linear-gradient(135deg, #0d1b2a 0%, #1b3a5c 60%, #1a78c2 100%);
        border-radius: 12px;
        padding: 28px 24px 18px 24px;
        text-align: center;
        margin-bottom: 24px;
    }
    .apec-logo-text {
        font-size: 2.6rem;
        font-weight: 900;
        letter-spacing: 3px;
        color: #f5a623;
        text-shadow: 2px 2px 6px rgba(0,0,0,0.5);
    }
    .apec-sub {
        font-size: 1.05rem;
        color: #cde4f7;
        letter-spacing: 1px;
        margin-top: 4px;
    }
    .apec-tagline {
        font-size: 0.95rem;
        color: #f5a623;
        margin-top: 8px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }

    /* ---- section headers ---- */
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #1b3a5c;
        border-left: 4px solid #1a78c2;
        padding-left: 10px;
        margin: 22px 0 10px 0;
    }

    /* ---- info pill ---- */
    .info-pill {
        background: #eef5fc;
        border: 1px solid #b3d4f0;
        border-radius: 8px;
        padding: 10px 18px;
        display: inline-block;
        margin: 4px 6px 4px 0;
        font-size: 0.92rem;
        color: #1b3a5c;
    }

    /* ---- auto-calc readout ---- */
    .auto-readout {
        background: #f4f8fd;
        border: 1px dashed #b3d4f0;
        border-radius: 6px;
        padding: 8px 12px;
        font-size: 0.88rem;
        color: #1b3a5c;
        margin-top: 4px;
    }

    /* ---- result cards ---- */
    .cost-card {
        background: #e8f2fb;
        border-left: 6px solid #1a78c2;
        border-radius: 10px;
        padding: 16px 20px;
        margin: 14px 0;
    }
    .cost-card .card-label {
        font-size: 0.85rem;
        color: #5580a0;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    .cost-card .card-amount {
        font-size: 1.9rem;
        font-weight: 800;
        color: #0d3b6e;
    }

    .retail-card {
        background: #eef5fc;
        border-left: 6px solid #1b3a5c;
        border-radius: 10px;
        padding: 16px 20px;
        margin: 14px 0;
    }
    .retail-card .card-label {
        font-size: 0.85rem;
        color: #5580a0;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    .retail-card .card-amount {
        font-size: 1.9rem;
        font-weight: 800;
        color: #0d3b6e;
    }

    .tax-card {
        background: #fff4e6;
        border-left: 6px solid #f5a623;
        border-radius: 10px;
        padding: 14px 20px;
        margin: 14px 0;
    }
    .tax-card .card-label {
        font-size: 0.85rem;
        color: #8a6320;
        text-transform: uppercase;
        letter-spacing: 0.8px;
    }
    .tax-card .card-amount {
        font-size: 1.4rem;
        font-weight: 700;
        color: #6a4a10;
    }

    .final-card {
        background: linear-gradient(135deg, #e8f8ee 0%, #cde9d4 100%);
        border-left: 6px solid #2e7d32;
        border-radius: 10px;
        padding: 20px 24px;
        margin: 18px 0;
        box-shadow: 0 2px 8px rgba(46,125,50,0.15);
    }
    .final-card .card-label {
        font-size: 0.95rem;
        color: #2e7d32;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 700;
    }
    .final-card .card-amount {
        font-size: 2.4rem;
        font-weight: 900;
        color: #1b5e20;
        margin-top: 4px;
    }

    /* ---- summary table ---- */
    .summary-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.9rem;
        margin-top: 10px;
    }
    .summary-table th {
        background: #1b3a5c;
        color: white;
        padding: 8px 12px;
        text-align: left;
    }
    .summary-table td {
        padding: 8px 12px;
        border-bottom: 1px solid #e0e8f0;
    }
    .summary-table tr:nth-child(even) td { background: #f4f8fd; }

    /* ---- disclaimer ---- */
    .disclaimer-box {
        background: #fff8e1;
        border: 1px solid #ffe082;
        border-left: 5px solid #f9a825;
        border-radius: 8px;
        padding: 14px 18px;
        margin-top: 22px;
        font-size: 0.85rem;
        color: #5d4037;
        line-height: 1.6;
    }
    .disclaimer-box strong { color: #e65100; }

    /* ---- button ---- */
    div.stButton > button {
        background: linear-gradient(135deg, #0d1b2a, #1a78c2);
        color: white !important;
        font-size: 1.1rem;
        font-weight: 700;
        padding: 12px 0;
        border-radius: 8px;
        border: none;
        width: 100%;
        letter-spacing: 0.5px;
        transition: opacity 0.2s;
    }
    div.stButton > button:hover { opacity: 0.88; }

    hr { border: none; border-top: 1px solid #dde8f0; margin: 20px 0; }
</style>
""", unsafe_allow_html=True)


# ── Constants & helpers ───────────────────────────────────────────────────────
# Labor days cheat sheet, per Canopy_Instructions.md
LABOR_DAYS_TABLE = {
    2: 3,
    3: 5,
    4: 7,
    5: 10,
    6: 12,
    7: 12,
    8: 14,
    9: 15,
    10: 15,
}
DEFAULT_DEPTH = 24.0  # Spec references depth via cheat sheet but only width formula is given;
                      # use 24 ft as the standard default (matches 4-dispenser baseline).

def get_labor_days(columns: int) -> int:
    """Return labor days for a given column count.
    Spec covers 2–10 columns; clamp at the nearest end for out-of-range values."""
    if columns <= 2:
        return LABOR_DAYS_TABLE[2]
    if columns >= 10:
        return LABOR_DAYS_TABLE[10]
    return LABOR_DAYS_TABLE[columns]

def get_misc_cost(columns: int) -> float:
    """MISC flat-rate pricing per spec:
       1–4 cols → $2,500   5–8 cols → $4,000   9–12 cols → $6,000"""
    if columns <= 4:
        return 2_500
    elif columns <= 8:
        return 4_000
    else:
        return 6_000

def calc_auto_width(dispensers: int, distance: float, overhang: float) -> float:
    """Width = overhang + (n_dispensers − 1) × distance + overhang.
       Per spec example: 4 disp → 10 + 28 + 28 + 28 + 10 = 104."""
    return 2 * overhang + max(dispensers - 1, 0) * distance

def fmt(value: float) -> str:
    return f"${value:,.2f}"


# ── Calculation engine ────────────────────────────────────────────────────────
def calculate(width, depth, columns, dispensers, gp_pct, tax_rate):
    # ACM
    a1 = 2 * width + 2 * depth
    a2 = math.ceil(a1 / 10)
    acm = a2 * 450

    # Steel
    steel = 4_000 * columns

    # Decking
    decking = 5_000 * columns

    # Lights
    lights = dispensers * 4 * 150

    # Misc
    misc = get_misc_cost(columns)

    # Labor (4 workers × 10 hrs × $30/hr × labor days)
    labor_days = get_labor_days(columns)
    labor = 4 * 10 * 30 * labor_days

    # Cost of canopy
    cost = acm + steel + decking + lights + misc + labor

    # Retail = Cost + GP markup
    retail = cost * (1 + gp_pct / 100)

    # Tax — applied to all items EXCEPT labor.
    # Labor portion of retail (proportional) is excluded from tax base.
    labor_share_of_retail = labor * (1 + gp_pct / 100)
    taxable_base = retail - labor_share_of_retail
    tax_amount = taxable_base * (tax_rate / 100)

    final_price = retail + tax_amount

    return {
        "acm": acm,
        "steel": steel,
        "decking": decking,
        "lights": lights,
        "misc": misc,
        "labor": labor,
        "labor_days": labor_days,
        "cost": cost,
        "retail": retail,
        "taxable_base": taxable_base,
        "tax_amount": tax_amount,
        "final_price": final_price,
        "sq_ft": width * depth,
    }


# ── Banner / Headline ─────────────────────────────────────────────────────────
st.markdown("""
<div class="apec-banner">
    <div class="apec-logo-text">APEC</div>
    <div class="apec-sub">Imaging &amp; Canopies</div>
    <div class="apec-tagline">Price Configurator</div>
</div>
""", unsafe_allow_html=True)


# ── Input form ────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">⚙️ Project Configuration</div>', unsafe_allow_html=True)

col_left, col_right = st.columns(2)

with col_left:
    # 1) Branding — Unbranded default
    branded = st.selectbox("Branding", ["Unbranded", "Branded"], index=0)

    # 2) Canopy Type — Dive-in default
    canopy_type = st.selectbox("Canopy Type", ["Dive-in", "Stacked"], index=0)

    # 3) Dispensers — default 4
    num_disp = st.number_input(
        "Number of Dispensers",
        min_value=1, max_value=12, value=4, step=1,
    )

    # 4) Columns = dispensers
    num_cols = int(num_disp)
    st.markdown(
        f"<div class='auto-readout'>🏛️ <strong>Columns:</strong> {num_cols} "
        f"<small>(equals dispensers)</small></div>",
        unsafe_allow_html=True,
    )

    # 5) Distance between dispensers — default 28
    distance = st.number_input(
        "Distance Between Dispensers (ft)",
        min_value=1.0, max_value=200.0, value=28.0, step=1.0,
    )

    # 6) Overhang — default 10
    overhang = st.number_input(
        "Overhang (ft)",
        min_value=0.0, max_value=50.0, value=10.0, step=1.0,
    )

with col_right:
    # 7) Width — auto-calculated, with override
    auto_width = calc_auto_width(int(num_disp), distance, overhang)
    st.markdown(
        f"<div class='auto-readout'>📏 <strong>Auto-calculated width:</strong> "
        f"{auto_width:.0f} ft &nbsp;<small>"
        f"({overhang:.0f} + {max(int(num_disp)-1,0)}×{distance:.0f} + {overhang:.0f})</small></div>",
        unsafe_allow_html=True,
    )
    override_width = st.checkbox("Override width", value=False)
    if override_width:
        width = st.number_input(
            "Width (ft)", min_value=1.0, max_value=1000.0,
            value=float(auto_width), step=1.0,
        )
    else:
        width = float(auto_width)

    # 8) Depth — defaults to 24, with override
    st.markdown(
        f"<div class='auto-readout'>📐 <strong>Default depth:</strong> "
        f"{DEFAULT_DEPTH:.0f} ft</div>",
        unsafe_allow_html=True,
    )
    override_depth = st.checkbox("Override depth", value=False)
    if override_depth:
        depth = st.number_input(
            "Depth (ft)", min_value=1.0, max_value=1000.0,
            value=float(DEFAULT_DEPTH), step=1.0,
        )
    else:
        depth = float(DEFAULT_DEPTH)

    # 11) Tax Rate — default 8.25
    tax_rate = st.number_input(
        "Tax Rate (%)",
        min_value=0.0, max_value=30.0, value=8.25, step=0.25, format="%.2f",
    )

    # 12) GP — default 45
    gp_pct = st.number_input(
        "Desired GP / Markup (%)",
        min_value=0.0, max_value=300.0, value=45.0, step=1.0, format="%.1f",
    )

# Live computed pills
sq_ft_live    = width * depth
labor_days_lv = get_labor_days(int(num_cols))
st.markdown(f"""
<div style="margin:14px 0 8px 0;">
    <span class="info-pill">📐 Dimensions: <strong>{width:.0f} × {depth:.0f} ft</strong></span>
    <span class="info-pill">🟦 Square Footage: <strong>{sq_ft_live:,.0f} sq ft</strong></span>
    <span class="info-pill">🔨 Labor Days: <strong>{labor_days_lv} days</strong></span>
    <span class="info-pill">💡 Lights: <strong>{int(num_disp) * 4} fixtures</strong></span>
</div>
""", unsafe_allow_html=True)

st.markdown("<hr/>", unsafe_allow_html=True)

# ── Calculate button ──────────────────────────────────────────────────────────
if st.button("🧮  Calculate Canopy Price"):

    r = calculate(width, depth, int(num_cols), int(num_disp), gp_pct, tax_rate)

    # ── Pricing output (no calculations/breakdown shown, per spec) ────────────
    st.markdown('<div class="section-title">💰 Pricing</div>', unsafe_allow_html=True)

    st.markdown(f"""
    <div class="cost-card">
        <div class="card-label">Canopy Cost</div>
        <div class="card-amount">{fmt(r['cost'])}</div>
    </div>

    <div class="retail-card">
        <div class="card-label">Retail Price &nbsp;(Cost + {gp_pct:.1f}% GP)</div>
        <div class="card-amount">{fmt(r['retail'])}</div>
    </div>

    <div class="tax-card">
        <div class="card-label">Tax — Bracket {tax_rate:.2f}%</div>
        <div class="card-amount">{fmt(r['tax_amount'])}</div>
        <div style="font-size:0.78rem;color:#8a6320;margin-top:4px;">
            (Tax applied to all items <em>except labor</em>)
        </div>
    </div>

    <div class="final-card">
        <div class="card-label">Final Price for Customer</div>
        <div class="card-amount">{fmt(r['final_price'])}</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Project summary ──────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📋 Project Summary</div>', unsafe_allow_html=True)

    summary_rows = [
        ("Branding",                    branded),
        ("Canopy Type",                 canopy_type),
        ("Dispensers",                  str(int(num_disp))),
        ("Columns",                     str(int(num_cols))),
        ("Distance Between Dispensers", f"{distance:.0f} ft"),
        ("Overhang",                    f"{overhang:.0f} ft"),
        ("Width",                       f"{width:.0f} ft" + ("  (overridden)" if override_width else "  (auto)")),
        ("Depth",                       f"{depth:.0f} ft" + ("  (overridden)" if override_depth else "  (default)")),
        ("Square Footage",              f"{r['sq_ft']:,.0f} sq ft"),
        ("Labor Days",                  f"{r['labor_days']} days"),
        ("GP Markup",                   f"{gp_pct:.1f}%"),
        ("Tax Rate",                    f"{tax_rate:.2f}%"),
    ]

    trs = "".join(
        f"<tr><td><strong>{k}</strong></td><td>{v}</td></tr>"
        for k, v in summary_rows
    )
    st.markdown(f"""
    <table class="summary-table">
        <thead><tr><th>Detail</th><th>Value</th></tr></thead>
        <tbody>{trs}</tbody>
    </table>
    <div style="font-size:0.85rem;color:#555;margin-top:14px;line-height:1.6;">
        <strong>This quote includes:</strong> ACM panels, steel structure, decking, lights,
        miscellaneous items (paint, transportation, insurance, project management,
        shipping &amp; handling), and labor.
    </div>
    """, unsafe_allow_html=True)

    # ── Disclaimer ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="disclaimer-box">
        <strong>⚠️ Disclaimer:</strong> Price does not include Imaging material, price sign material,
        Price Sign Labor, permit fees, electrical work, or canopy piers.<br><br>
        <em>"This is an experimental configurator and not to be used in quoting real jobs at this time."</em>
    </div>
    """, unsafe_allow_html=True)
