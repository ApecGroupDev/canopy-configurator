"""
Canopy Proposal Builder — Phase III (v4.5 — Disclaimers heading, installation warranty, 7pt body)
Public entry: build_proposal(data, output_path) -> Path
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor

APEC_PALETTE = {
    "band_bg":         RGBColor(0x2B, 0x33, 0x40),
    "grand_total_bg":  RGBColor(0x2B, 0x33, 0x40),
    "accent":          RGBColor(0xE0, 0x20, 0x20),
    "text_emphasis":   RGBColor(0x2B, 0x33, 0x40),
    "customer_bg":     RGBColor(0xF5, 0xF5, 0xF7),
    "subtle_bg":       RGBColor(0xF5, 0xF5, 0xF7),
    "subtle_bg_alt":   RGBColor(0xEE, 0xED, 0xF0),
}

GEO_PALETTE = {
    "band_bg":         RGBColor(0x0F, 0x8A, 0x60),
    "grand_total_bg":  RGBColor(0x0A, 0x5C, 0x42),
    "accent":          RGBColor(0x0A, 0x5C, 0x42),
    "text_emphasis":   RGBColor(0x0A, 0x5C, 0x42),
    "customer_bg":     RGBColor(0xEA, 0xF4, 0xF0),
    "subtle_bg":       RGBColor(0xEA, 0xF4, 0xF0),
    "subtle_bg_alt":   RGBColor(0xDF, 0xEE, 0xEC),
}

PALETTE = APEC_PALETTE


def _H(rgb):
    return bytes(rgb).hex().upper()


COMPANIES = {
    "APEC": {
        "name":    "APEC Imaging and Canopies",
        "tagline": "Petroleum Canopy Specialists",
        "address": "4732 N Royal Atlanta Drive, Suite E, Tucker, GA 30084",
        "phone":   "855-444-APEC",
        "email":   "sales@TheAPECgroup.com",
        "website": "www.TheAPECgroup.com",
        "logo":    "Apec Imaging Logo.jpg",
        "signature_label": "APEC Authorized Representative & Date",
        "palette": APEC_PALETTE,
    },
    "GEO": {
        "name":    "GEO Canopies",
        "tagline": "Petroleum Canopy Specialists",
        "address": "40 Lyerly Street, Houston, TX 77022",
        "phone":   "844-GEO-4040",
        "email":   "sales@GeoPetroleum.com",
        "website": "www.GeoPetroleum.com",
        "logo":    "GEO Canopies logo.jpg",
        "signature_label": "GEO Authorized Representative & Date",
        "palette": GEO_PALETTE,
    },
}

USABLE_W = 7.3
GREY  = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED   = RGBColor(0xC6, 0x28, 0x28)

CANOPY_BUILD_SPECS = [
    "Clearance: 16.5 ft",
    "Columns: 10″ × 10″ square",
    "Decking: white embossed, 20-gauge",
    "Perimeter gutter",
    "Scuppers",
    "Anchor bolts",
]

_TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
    "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
    "tcFitText", "vAlign", "hideMark", "headers",
    "cellIns", "cellDel", "cellMerge", "tcPrChange",
]

def _insert_tc_pr_ordered(tc_pr, new_element):
    new_name = new_element.tag.split("}")[-1]
    new_pos = _TCPR_ORDER.index(new_name)
    for child in list(tc_pr):
        cname = child.tag.split("}")[-1]
        if cname == new_name:
            tc_pr.remove(child)
            break
    for i, child in enumerate(tc_pr):
        cname = child.tag.split("}")[-1]
        if cname in _TCPR_ORDER and _TCPR_ORDER.index(cname) > new_pos:
            tc_pr.insert(i, new_element)
            return
    tc_pr.append(new_element)


def _shade_cell(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    _insert_tc_pr_ordered(tc_pr, shd)


def _remove_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:color"), "auto")
        borders.append(b)
    tbl_pr.append(borders)


def _set_fixed_layout(table, col_widths_inches):
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    for ci, w in enumerate(col_widths_inches):
        table.columns[ci].width = Inches(w)
    for row in table.rows:
        for ci, w in enumerate(col_widths_inches):
            row.cells[ci].width = Inches(w)


def _set_cell_margins(cell, top=40, left=80, bottom=40, right=80):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    _insert_tc_pr_ordered(tc_pr, mar)


def _format_run(run, *, size=10, bold=False, italic=False, color=None, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _add_para(cell_or_doc, text="", *, size=10, bold=False, italic=False,
              color=None, align=None, space_after=2, space_before=0):
    p = cell_or_doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if align is not None:
        p.alignment = align
    if text:
        r = p.add_run(text)
        _format_run(r, size=size, bold=bold, italic=italic, color=color)
    return p


def _clear_default_para(container):
    if container.paragraphs and not container.paragraphs[0].text:
        p = container.paragraphs[0]
        p._element.getparent().remove(p._element)


def _add_cell_bottom_border(cell, sz=8, color=None):
    if color is None:
        color = _H(PALETTE["band_bg"])
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(sz))
    b.set(qn("w:color"), color)
    borders.append(b)
    _insert_tc_pr_ordered(tc_pr, borders)


def _money(n):
    return f"${n:,.2f}"


def _section_totals(data):
    p = data["pricing"]
    gp = p["gp_rate"]
    canopy_mat_marked = (p["acm"] + p["steel"] + p["decking"] + p["lights"] + p["misc"]) * (1 + gp)
    canopy_mat = canopy_mat_marked + p.get("steel_secondary", 0)
    canopy_inst_marked = p["base_labor"] * (1 + gp)
    canopy_inst = canopy_inst_marked + p.get("two_disp_labor_add", 0)
    branding_mat  = p.get("brand_imaging", 0) + p.get("shipping", 0)
    branding_inst = p.get("branded_labor_add", 0)
    price_sign_mat  = p.get("mid_material", 0)
    price_sign_inst = p.get("mid_labor", 0)
    material_total    = canopy_mat + branding_mat + price_sign_mat
    installation_total = canopy_inst + branding_inst + price_sign_inst
    tax = material_total * p["tax_rate"]
    grand_total = material_total + installation_total + tax
    return {
        "canopy_mat": canopy_mat, "canopy_inst": canopy_inst,
        "branding_mat": branding_mat, "branding_inst": branding_inst,
        "price_sign_mat": price_sign_mat, "price_sign_inst": price_sign_inst,
        "material_total": material_total, "installation_total": installation_total,
        "tax": tax, "grand_total": grand_total,
    }


def _derive_inclusions(data):
    inc = [
        "Canopy structure (ACM, steel, decking, lighting, miscellaneous)",
        "Engineering drawing (one per canopy)",
        "Canopy installation",
    ]
    if data["canopy"].get("branded"):
        inc.append(f"Brand imaging — {data['canopy']['brand_name']}")
        inc.append("Brand-imaging shipping & handling")
        inc.append("Branding installation")
    else:
        inc.append("Forecourt painting (columns, island forms, bollards)")
    if data["mid"].get("include"):
        inc.append(f"Price sign — {data['mid'].get('brand','')}")
        inc.append("Price sign installation")
    return inc


def _derive_exclusions(data):
    exc = [
        "Permit fees",
        "Electrical work",
        "Canopy piers / foundations",
        "Bricking / decorative masonry",
        "Third-party inspections (provided at additional cost if required)",
    ]
    if not data["canopy"].get("branded"):
        exc.append("Brand imaging material & installation")
    if not data["mid"].get("include"):
        exc.append("Price sign material & installation")
    for extra in data["pricing"].get("items_not_included", []):
        if extra and extra not in exc:
            exc.append(extra)
    return exc


def _add_section_band(doc, label):
    t = doc.add_table(rows=1, cols=1)
    _remove_table_borders(t)
    _set_fixed_layout(t, [USABLE_W])
    c = t.cell(0, 0)
    _shade_cell(c, _H(PALETTE["band_bg"]))
    _set_cell_margins(c, top=70, left=120, bottom=70, right=120)
    _clear_default_para(c)
    _add_para(c, label, size=10, bold=True, color=WHITE, space_after=0)


def _add_section_subtotal(doc, mat, inst, *, section_name="Section"):
    total = mat + inst
    t = doc.add_table(rows=1, cols=2)
    _remove_table_borders(t)
    _set_fixed_layout(t, [USABLE_W * 0.55, USABLE_W * 0.45])
    _shade_cell(t.cell(0, 0), _H(PALETTE["subtle_bg"]))
    _shade_cell(t.cell(0, 1), _H(PALETTE["subtle_bg"]))
    for ci in range(2):
        _set_cell_margins(t.cell(0, ci), top=60, bottom=60, left=160, right=160)
        _clear_default_para(t.cell(0, ci))
    _add_para(t.cell(0, 0),
              f"Material: {_money(mat)}     Installation: {_money(inst)}",
              size=9, bold=True, color=PALETTE["text_emphasis"], space_after=0,
              align=WD_ALIGN_PARAGRAPH.LEFT)
    _add_para(t.cell(0, 1),
              f"{section_name} Total: {_money(total)}",
              size=11, bold=True, color=PALETTE["accent"], space_after=0,
              align=WD_ALIGN_PARAGRAPH.RIGHT)


def _add_section_subtotal_note(doc, note_text):
    t = doc.add_table(rows=1, cols=1)
    _remove_table_borders(t)
    _set_fixed_layout(t, [USABLE_W])
    cell = t.cell(0, 0)
    _shade_cell(cell, _H(PALETTE["subtle_bg"]))
    _set_cell_margins(cell, top=50, bottom=50, left=160, right=160)
    _clear_default_para(cell)
    _add_para(cell, note_text, size=9, bold=True, italic=True,
              color=PALETTE["text_emphasis"], space_after=0,
              align=WD_ALIGN_PARAGRAPH.RIGHT)


def _add_canopy_section(doc, data, totals):
    _add_section_band(doc, "CANOPY")
    body = doc.add_table(rows=1, cols=1)
    _remove_table_borders(body)
    _set_fixed_layout(body, [USABLE_W])
    cell = body.cell(0, 0)
    _set_cell_margins(cell, top=100, left=160, bottom=80, right=160)
    _clear_default_para(cell)
    c = data["canopy"]
    items_provided = [
        "ACM panels & fascia",
        f"Structural steel framing ({c['columns']} columns)",
        "Decking system",
        f"Canopy lighting ({c['dispensers'] * 4} fixtures)",
        "Miscellaneous (paint, transport, project management, insurance)",
    ]
    twocol = cell.add_table(rows=1, cols=2)
    _remove_table_borders(twocol)
    _set_fixed_layout(twocol, [3.45, 3.45])
    left_c, right_c = twocol.cell(0, 0), twocol.cell(0, 1)
    _set_cell_margins(left_c,  top=0, bottom=0, left=0,   right=120)
    _set_cell_margins(right_c, top=0, bottom=0, left=120, right=0)
    _clear_default_para(left_c)
    _clear_default_para(right_c)
    _add_para(left_c, "Items provided:", size=9, bold=True,
              color=PALETTE["text_emphasis"], space_after=2)
    for it in items_provided:
        _add_para(left_c, f"•  {it}", size=9, space_after=1)
    _add_para(right_c, "Build specifications:", size=9, bold=True,
              color=PALETTE["text_emphasis"], space_after=2)
    for spec in CANOPY_BUILD_SPECS:
        _add_para(right_c, f"•  {spec}", size=9, space_after=1)
    _add_para(cell, "", size=2, space_after=2)
    _add_para(cell, "Engineering drawing: One included per canopy.",
              size=9, italic=True, color=GREY, space_after=2)
    _add_para(cell,
              "Third-party inspections: Not included; if required by the "
              "authority having jurisdiction, provided at additional cost.",
              size=9, italic=True, color=GREY, space_after=0)
    _add_section_subtotal(doc, totals["canopy_mat"], totals["canopy_inst"],
                          section_name="Canopy")


def _add_branding_section(doc, data, totals):
    _add_section_band(doc, "CANOPY BRANDING")
    body = doc.add_table(rows=1, cols=1)
    _remove_table_borders(body)
    _set_fixed_layout(body, [USABLE_W])
    cell = body.cell(0, 0)
    _set_cell_margins(cell, top=100, left=160, bottom=80, right=160)
    _clear_default_para(cell)
    c = data["canopy"]
    if c.get("branded"):
        status_text = f"Branded: Yes — {c['brand_name']}"
    else:
        status_text = "Branded: No (Unbranded)"
    _add_para(cell, status_text, size=10, bold=True,
              color=PALETTE["text_emphasis"], space_after=4)
    _add_para(cell, "Items provided:", size=9, bold=True,
              color=PALETTE["text_emphasis"], space_after=2)
    always_items = [
        "Fascia / ACM imaging or paint",
        "Forecourt: paint of columns, island forms, and bollards",
    ]
    for it in always_items:
        _add_para(cell, f"•  {it}", size=9, space_after=1)
    if c.get("branded"):
        for it in ("Brand valances", "Trash bins", "Pump toppers", "Flags"):
            _add_para(cell, f"•  {it}", size=9, space_after=1)
    _add_para(cell, "", size=2, space_after=2)
    concluding = ("Branding retail prices include material (with shipping and "
                  "handling) and Installation.")
    _add_para(cell, concluding, size=9, italic=True, color=GREY, space_after=0)
    if c.get("branded"):
        _add_section_subtotal(doc, totals["branding_mat"], totals["branding_inst"],
                              section_name="Branding")
    else:
        _add_section_subtotal_note(doc, "Included in canopy price")


def _add_price_sign_section(doc, data, totals):
    _add_section_band(doc, "PRICE SIGN")
    body = doc.add_table(rows=1, cols=1)
    _remove_table_borders(body)
    _set_fixed_layout(body, [USABLE_W])
    cell = body.cell(0, 0)
    _set_cell_margins(cell, top=100, left=160, bottom=80, right=160)
    _clear_default_para(cell)
    brand = data["mid"].get("brand", "")
    _add_para(cell, f"Price sign: {brand}", size=10, bold=True,
              color=PALETTE["text_emphasis"], space_after=2)
    _add_para(cell, "Includes price sign material and installation.",
              size=9, italic=True, color=GREY, space_after=0)
    _add_section_subtotal(doc, totals["price_sign_mat"], totals["price_sign_inst"],
                          section_name="Price Sign")


def _add_notes_box(doc):
    _add_section_band(doc, "ADDITIONAL NOTES")
    body = doc.add_table(rows=1, cols=1)
    _remove_table_borders(body)
    _set_fixed_layout(body, [USABLE_W])
    cell = body.cell(0, 0)
    _shade_cell(cell, "FFFDF5")
    _set_cell_margins(cell, top=120, left=160, bottom=600, right=160)
    _clear_default_para(cell)
    for _ in range(3):
        _add_para(cell, "", size=10, space_after=0)


def _add_incl_excl(doc, data):
    incexc = doc.add_table(rows=1, cols=2)
    _remove_table_borders(incexc)
    _set_fixed_layout(incexc, [USABLE_W / 2, USABLE_W / 2])
    inc_cell, exc_cell = incexc.cell(0, 0), incexc.cell(0, 1)
    _shade_cell(inc_cell, "EAF7EE")
    _shade_cell(exc_cell, "FCEEEE")
    _set_cell_margins(inc_cell, top=120, left=140, bottom=120, right=140)
    _set_cell_margins(exc_cell, top=120, left=140, bottom=120, right=140)
    _clear_default_para(inc_cell)
    _clear_default_para(exc_cell)
    _add_para(inc_cell, "✓  INCLUDED", size=10, bold=True, color=GREEN, space_after=4)
    for item in _derive_inclusions(data):
        _add_para(inc_cell, f"•  {item}", size=9, color=PALETTE["text_emphasis"], space_after=1)
    _add_para(exc_cell, "✗  NOT INCLUDED", size=10, bold=True, color=RED, space_after=4)
    for item in _derive_exclusions(data):
        _add_para(exc_cell, f"•  {item}", size=9, color=PALETTE["text_emphasis"], space_after=1)


def _add_grand_total(doc, totals, tax_rate):
    t = doc.add_table(rows=4, cols=2)
    _remove_table_borders(t)
    _set_fixed_layout(t, [USABLE_W - 2.0, 2.0])
    rows = [
        ("Material total", _money(totals["material_total"]), False),
        ("Installation",   _money(totals["installation_total"]), False),
        (f"Sales tax ({tax_rate*100:.2f}% on material)", _money(totals["tax"]), False),
        ("GRAND TOTAL",    _money(totals["grand_total"]), True),
    ]
    for ri, (label, amount, is_grand) in enumerate(rows):
        l, r = t.cell(ri, 0), t.cell(ri, 1)
        _set_cell_margins(l, top=60, bottom=60, left=160)
        _set_cell_margins(r, top=60, bottom=60, right=160)
        _clear_default_para(l)
        _clear_default_para(r)
        if is_grand:
            _shade_cell(l, _H(PALETTE["grand_total_bg"]))
            _shade_cell(r, _H(PALETTE["grand_total_bg"]))
            _add_para(l, label, size=12, bold=True, color=WHITE, space_after=0)
            _add_para(r, amount, size=14, bold=True, color=WHITE, space_after=0,
                      align=WD_ALIGN_PARAGRAPH.RIGHT)
        else:
            _shade_cell(l, _H(PALETTE["subtle_bg"]))
            _shade_cell(r, _H(PALETTE["subtle_bg"]))
            _add_para(l, label, size=10, bold=False, color=PALETTE["text_emphasis"], space_after=0)
            _add_para(r, amount, size=11, bold=True, color=PALETTE["text_emphasis"],
                      space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)


def _add_terms(doc):
    t = doc.add_table(rows=1, cols=1)
    _remove_table_borders(t)
    _set_fixed_layout(t, [USABLE_W])
    cell = t.cell(0, 0)
    _shade_cell(cell, _H(PALETTE["subtle_bg"]))
    _set_cell_margins(cell, top=100, left=160, bottom=100, right=160)
    _clear_default_para(cell)
    _add_para(cell, "Disclaimers:", size=9, bold=True,
              color=PALETTE["text_emphasis"], space_after=3)
    disclosures = (
        "Prices are valid for 30 days from the quote date. "
        "Pricing is contingent on jobsite accessibility and favorable "
        "ground conditions. "
        "Docusign signature and a 50% down payment are required to execute "
        "this agreement. "
        "All material cost has been included on an estimated basis using "
        "certain assumptions. This may be changed based on the prevailing "
        "prices and actual quantities at the time of the order and hence "
        "will be charged to the customer separately. "
        "Material warranty is based on manufacturers’ warranties while "
        "installation warranty is 90 days from the date of installation. "
        "Any requirement such as soil testing, electrical, survey drawings, "
        "third party or Govt. special inspections costs are not included "
        "and if needed for permitting, will be done at owner’s expense. "
        "Carrying charges of 1.5% per month will be charged on all past "
        "due payments."
    )
    _add_para(cell, disclosures, size=7, color=GREY, space_after=0)


def _add_signature_block(doc, company):
    sig = doc.add_table(rows=2, cols=2)
    _remove_table_borders(sig)
    _set_fixed_layout(sig, [USABLE_W / 2, USABLE_W / 2])
    for ci in range(2):
        line_cell = sig.cell(0, ci)
        _set_cell_margins(line_cell, top=280, bottom=20)
        _clear_default_para(line_cell)
        _add_cell_bottom_border(line_cell)
        _add_para(line_cell, "", size=8, space_after=0)
    _clear_default_para(sig.cell(1, 0))
    _clear_default_para(sig.cell(1, 1))
    _add_para(sig.cell(1, 0), "Customer Signature, Name & Date",
              size=8, color=GREY, space_after=0, align=WD_ALIGN_PARAGRAPH.LEFT)
    _add_para(sig.cell(1, 1), company["signature_label"],
              size=8, color=GREY, space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)


def _spec_rows(data):
    c = data["canopy"]
    col_arr = ("Double column" if (c["type"] == "Dive-in" and c.get("double_col"))
               else "Single column")
    canopy_label = (f"{c['type']} ({col_arr})" if c["type"] == "Dive-in"
                    else c["type"])
    branded_label = (f"Yes — {c.get('brand_name','')}" if c.get("branded")
                     else "No")
    price_sign_label = (f"Yes — {data['mid'].get('brand','')}"
                        if data["mid"].get("include") else "No")
    return [
        ("Canopy Type", canopy_label),
        ("Branding", branded_label),
        ("Price Sign", price_sign_label),
        ("Dispensers / Columns", f"{c['dispensers']}  /  {c['columns']}"),
        ("Dimensions (W × D)", f"{int(c['width'])}′  ×  {int(c['depth'])}′"),
        ("Square Footage", f"{int(c['width']*c['depth']):,} sq ft"),
        ("Distance / Overhang", f"{int(c.get('distance',0))}′  /  {int(c.get('overhang',0))}′"),
    ]


def build_proposal(data, output_path):
    global PALETTE
    output_path = Path(output_path)
    company = COMPANIES[data["company_key"]]
    PALETTE = company["palette"]
    totals = _section_totals(data)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    header = doc.add_table(rows=1, cols=2)
    _remove_table_borders(header)
    _set_fixed_layout(header, [2.1, USABLE_W - 2.1])
    logo_cell = header.cell(0, 0)
    info_cell = header.cell(0, 1)
    logo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    info_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _clear_default_para(logo_cell)
    logo_path = Path(__file__).parent / company["logo"]
    if logo_path.exists():
        p = logo_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.add_run().add_picture(str(logo_path), width=Inches(1.8))
    _clear_default_para(info_cell)
    _add_para(info_cell, company["name"], size=14, bold=True, color=PALETTE["accent"],
              align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=1)
    _add_para(info_cell, company["tagline"], size=9, italic=True, color=GREY,
              align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=4)
    _add_para(info_cell, company["address"], size=9, color=GREY,
              align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=1)
    _add_para(info_cell,
              f'{company["phone"]}  •  {company["email"]}  •  {company["website"]}',
              size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)

    title = doc.add_table(rows=1, cols=3)
    _remove_table_borders(title)
    _set_fixed_layout(title, [2.7, 2.4, USABLE_W - 5.1])
    for ci in range(3):
        _shade_cell(title.cell(0, ci), _H(PALETTE["band_bg"]))
        _set_cell_margins(title.cell(0, ci), top=70, bottom=70)
    c0, c1, c2 = title.cell(0, 0), title.cell(0, 1), title.cell(0, 2)
    _clear_default_para(c0); _clear_default_para(c1); _clear_default_para(c2)
    _add_para(c0, "PROPOSAL", size=15, bold=True, color=WHITE, space_after=0)
    p1 = c1.add_paragraph()
    p1.paragraph_format.space_after = Pt(0)
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _format_run(p1.add_run("QUOTE #  "), size=9, color=WHITE)
    _format_run(p1.add_run(data["quote_number"]), size=11, bold=True, color=WHITE)
    p2 = c2.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _format_run(p2.add_run("DATE  "), size=9, color=WHITE)
    _format_run(p2.add_run(data["quote_date"]), size=11, bold=True, color=WHITE)
    _add_para(doc, "", size=2, space_after=0)

    cust_specs = doc.add_table(rows=1, cols=2)
    _remove_table_borders(cust_specs)
    _set_fixed_layout(cust_specs, [3.6, USABLE_W - 3.6])
    left = cust_specs.cell(0, 0)
    _clear_default_para(left)
    _shade_cell(left, _H(PALETTE["customer_bg"]))
    _set_cell_margins(left, top=120, left=140, bottom=120, right=140)
    cust = data["customer"]
    _add_para(left, "PREPARED FOR", size=8, bold=True, color=PALETTE["text_emphasis"], space_after=2)
    _add_para(left, cust.get("company", ""), size=11, bold=True,
              color=PALETTE["text_emphasis"], space_after=1)
    if cust.get("name"):
        _add_para(left, f"Attn: {cust['name']}", size=9, color=GREY, space_after=1)
    contact_bits = [b for b in (cust.get("phone"), cust.get("email")) if b]
    if contact_bits:
        _add_para(left, "  •  ".join(contact_bits), size=9, color=GREY, space_after=8)
    else:
        _add_para(left, "", space_after=8)
    _add_para(left, "PROJECT SITE", size=8, bold=True, color=PALETTE["text_emphasis"], space_after=2)
    _add_para(left, cust.get("street", ""), size=10, space_after=1)
    citystatezip = ", ".join(b for b in
                             (cust.get("city"),
                              f"{cust.get('state','')} {cust.get('zip','')}".strip())
                             if b and b.strip())
    _add_para(left, citystatezip, size=10, space_after=6)
    sp = data.get("sales_person", {})
    if sp.get("name"):
        _add_para(left, "SALES REPRESENTATIVE", size=8, bold=True,
                  color=PALETTE["text_emphasis"], space_after=2)
        _add_para(left, sp.get("name", ""), size=10, bold=True, space_after=1)
        sp_contact = [b for b in (sp.get("phone"), sp.get("email")) if b]
        if sp_contact:
            _add_para(left, "  •  ".join(sp_contact), size=9, color=GREY, space_after=0)

    right = cust_specs.cell(0, 1)
    _clear_default_para(right)
    _set_cell_margins(right, top=120, left=140, bottom=120, right=140)
    _add_para(right, "PROJECT SPECIFICATIONS", size=8, bold=True,
              color=PALETTE["text_emphasis"], space_after=4)
    spec_rows = _spec_rows(data)
    spec_tbl = right.add_table(rows=len(spec_rows), cols=2)
    _remove_table_borders(spec_tbl)
    _set_fixed_layout(spec_tbl, [1.7, 1.7])
    for ri, (k, v) in enumerate(spec_rows):
        kcell, vcell = spec_tbl.cell(ri, 0), spec_tbl.cell(ri, 1)
        if ri % 2 == 0:
            _shade_cell(kcell, _H(PALETTE["subtle_bg_alt"]))
            _shade_cell(vcell, _H(PALETTE["subtle_bg_alt"]))
        _set_cell_margins(kcell, top=30, bottom=30, left=80, right=20)
        _set_cell_margins(vcell, top=30, bottom=30, left=20, right=80)
        _clear_default_para(kcell); _clear_default_para(vcell)
        _add_para(kcell, k, size=9, color=GREY, space_after=0)
        _add_para(vcell, v, size=9, bold=True, color=PALETTE["text_emphasis"],
                  space_after=0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _add_para(doc, "", size=2, space_after=0)

    _add_canopy_section(doc, data, totals)
    _add_para(doc, "", size=2, space_after=0)
    _add_branding_section(doc, data, totals)
    _add_para(doc, "", size=2, space_after=0)
    if data["mid"].get("include"):
        _add_price_sign_section(doc, data, totals)
        _add_para(doc, "", size=2, space_after=0)
    _add_notes_box(doc)
    _add_para(doc, "", size=2, space_after=0)
    _add_incl_excl(doc, data)
    _add_para(doc, "", size=2, space_after=0)
    _add_grand_total(doc, totals, data["pricing"]["tax_rate"])
    _add_para(doc, "", size=2, space_after=0)
    _add_terms(doc)
    _add_para(doc, "", size=2, space_after=0)
    _add_signature_block(doc, company)

    doc.save(str(output_path))
    return output_path


def sample_data():
    return {
        "company_key":  "APEC",
        "quote_number": f"Q-{_dt.date.today().strftime('%Y%m%d')}-001",
        "quote_date":   _dt.date.today().strftime("%B %d, %Y"),
        "customer": {
            "company": "Burleson Fuel Stop", "name": "Test Customer",
            "phone": "(817) 555-0100", "email": "owner@burlesonfuelstop.test",
            "street": "898 NE Alsbury Blvd", "city": "Burleson",
            "state": "TX", "zip": "76028",
        },
        "sales_person": {
            "name": "Walid Husain", "phone": "(555) 555-0123",
            "email": "walid@apecimaging.com",
        },
        "canopy": {
            "type": "Dive-in", "branded": False, "brand_name": "",
            "double_col": False, "dispensers": 4, "columns": 4,
            "width": 104, "depth": 24, "distance": 28, "overhang": 10,
            "labor_days": 7,
        },
        "mid": {"include": False, "brand": None},
        "pricing": {
            "gp_rate": 0.45, "tax_rate": 0.0825,
            "acm": 11_700, "steel": 16_000, "decking": 20_000,
            "lights": 2_400, "misc": 2_500,
            "steel_secondary": 0, "shipping": 0, "brand_imaging": 0,
            "mid_material": 0, "base_labor": 8_400,
            "branded_labor_add": 0, "two_disp_labor_add": 0, "mid_labor": 0,
            "items_not_included": [],
        },
    }


if __name__ == "__main__":
    out = build_proposal(sample_data(),
                         Path(__file__).parent / "Proposal_v4_Sample.docx")
    print(f"Wrote {out}")
