# proposal_doc.py
import io
import numbers
from typing import Optional

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========
# Taiga-teema
# =========
TAIGA_ACCENT      = RGBColor(0x23, 0x4F, 0x3B)   # #234F3B
TAIGA_BORDER_HEX  = "E5E7EB"                     # #E5E7EB (hexa ilman #)
TAIGA_TEXT        = RGBColor(0x11, 0x18, 0x27)   # #111827
TAIGA_CARD_BG_HEX = "F8FAF9"                     # #F8FAF9


# =========
# Tyyli & asettelu
# =========
def _configure_styles(doc: Document) -> None:
    """Määritä dokumentin perustyylit Taigan ilmeeseen."""
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    normal.font.color.rgb = TAIGA_TEXT

    for lvl, size in [(1, 20), (2, 16), (3, 13)]:
        style_name = f"Heading {lvl}"
        if style_name in doc.styles:
            s = doc.styles[style_name]
            s.font.name = "Arial"
            s.font.size = Pt(size)
            s.font.bold = True
            s.font.color.rgb = TAIGA_TEXT
            s.paragraph_format.space_before = Pt(6)
            s.paragraph_format.space_after = Pt(6)

    # Korttityylinen väliotsikko
    if "TaigaCardHeading" not in doc.styles:
        s = doc.styles.add_style("TaigaCardHeading", WD_STYLE_TYPE.PARAGRAPH)
        s.font.name = "Arial"
        s.font.size = Pt(13)
        s.font.bold = True
        s.font.color.rgb = TAIGA_TEXT
        s.paragraph_format.space_before = Pt(2)
        s.paragraph_format.space_after = Pt(2)


def _set_page_margins(doc: Document, left=2.0, right=2.0, top=2.0, bottom=2.0) -> None:
    """Aseta marginaalit senttimetreinä."""
    section = doc.sections[0]
    section.left_margin = Inches(left / 2.54)
    section.right_margin = Inches(right / 2.54)
    section.top_margin = Inches(top / 2.54)
    section.bottom_margin = Inches(bottom / 2.54)


# =========
# Pienet apurit
# =========
def _add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


def _add_kv_paragraph(doc: Document, label: str, value: str) -> None:
    p = doc.add_paragraph()
    run1 = p.add_run(f"{label}: ")
    run1.bold = True
    run1.font.color.rgb = TAIGA_TEXT
    p.add_run(value).font.color.rgb = TAIGA_TEXT
    p.space_after = Pt(2)


# =========
# Muotoilut (FI-tyyli)
# =========
def _fmt_num_int_fi(v) -> str:
    """Kokonaisluvut: '12 345' (välilyönti tuhaterottimena)."""
    try:
        n = float(v)
        s = f"{n:,.0f}".replace(",", " ").replace(".", ",")
        return s
    except Exception:
        return str(v)


def _fmt_num_2dec_fi(v) -> str:
    """Kaksi desimaalia: '1 234,57'."""
    try:
        n = float(v)
        s = f"{n:,.2f}".replace(",", " ").replace(".", ",")
        return s
    except Exception:
        return str(v)


def _fmt_eur(v, decimals=0) -> str:
    if decimals == 0:
        return f"{_fmt_num_int_fi(v)} €"
    return f"{_fmt_num_2dec_fi(v)} €"


# =========
# DOCX-solujen tyylit
# =========
def _shade_cell(cell, hex_fill: str) -> None:
    """Aseta solun taustaväri heksalla (RRGGBB)."""
    tcPr = cell._tc.get_or_add_tcPr()
    # Poista mahdollinen vanha shading, luodaan uusi
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _set_cell_border(cell, color_hex: str = TAIGA_BORDER_HEX, size: int = 6) -> None:
    """
    Aseta solun reunaviivat.
    - color_hex: 'RRGGBB' (ilman '#')
    - size: viivan paksuus 1/8 pt (6 ~ 0,75 pt)
    """
    tcPr = cell._tc.get_or_add_tcPr()

    # Luo/hanki w:tcBorders
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    else:
        # Tyhjennä edelliset reunamääritykset
        for el in list(tcBorders):
            tcBorders.remove(el)

    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color_hex)
        tcBorders.append(el)


def _style_table(table) -> None:
    """Taulukon otsikko vihreäksi, zebra-raidat ja kevyet reunat."""
    # Header-rivi
    hdr = table.rows[0].cells
    for c in hdr:
        _shade_cell(c, "234F3B")  # Taiga-vihreä
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Reunat + zebra
    for i, row in enumerate(table.rows):
        for c in row.cells:
            _set_cell_border(c, color_hex=TAIGA_BORDER_HEX, size=6)
            # Zebra vain datariveihin
            if i % 2 == 0 and i != 0:
                _shade_cell(c, TAIGA_CARD_BG_HEX)


# =========
# Pivotin lisäys
# =========
def _add_pivot_table(doc: Document, title: str, df) -> None:
    """Lisää pivot-taulukko korttimaisena lohkona, jos df on annettu."""
    if df is None or getattr(df, "empty", True):
        return

    t = doc.add_paragraph(title, style="TaigaCardHeading")
    t.paragraph_format.space_before = Pt(8)
    t.paragraph_format.space_after = Pt(4)

    n_cols = len(df.columns) + 1  # +1 indeksisarakkeelle
    table = doc.add_table(rows=1, cols=n_cols)

    # Header
    hdr = table.rows[0].cells
    hdr[0].text = "Cost item"
    for j, col in enumerate(df.columns, start=1):
        hdr[j].text = str(col)

    # Rows
    for idx, row in df.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(idx)
        for j, val in enumerate(row, start=1):
            text = _fmt_eur(val, 0) if isinstance(val, numbers.Number) else str(val)
            cells[j].text = text

    _style_table(table)
    doc.add_paragraph("")  # väli


# =========
# Pääfunktio
# =========
def generate_proposal_doc(
    payload,
    df_pivot_taiga=None,
    df_pivot_trad=None,
    df_pivot_delta=None,
    locale: str = "fi_FI",
    logo_path: Optional[str] = "logo.png",
    **kwargs
) -> bytes:
    """
    Rakenna Taiga-tyylinen Word-yhteenveto (logo, yhteenveto, pivotit).

    Taaksepäinyhteensopivuus:
    Jos funktiota kutsutaan muodossa (payload, "fi_FI"), tulkitaan toinen arg locale-merkiksi.
    """
    # Backward compat: (payload, "fi_FI")
    if isinstance(df_pivot_taiga, str) and df_pivot_trad is None and df_pivot_delta is None:
        locale = df_pivot_taiga
        df_pivot_taiga = None

    d = Document()
    _configure_styles(d)
    _set_page_margins(d, left=2.0, right=2.0, top=2.0, bottom=2.0)

    # --- Taiga Cover Page ---
    if logo_path:
        try:
            # Luo kansilehti omalle sivulleen
            section = d.sections[0]
            section.start_type = 0  # Ensimmäinen sivu
            section.page_height = Inches(11.69)  # A4 pystyasento
            section.page_width = Inches(8.27)

            cover = d.add_paragraph()
            cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_logo = cover.add_run()
            run_logo.add_picture(logo_path, width=Inches(2.4))
            cover.space_after = Pt(70)

            # Otsikko
            title = d.add_paragraph("Total Ownership Cost Summary")
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_title = title.runs[0]
            run_title.font.name = "Inter"
            run_title.font.size = Pt(28)
            run_title.font.bold = True
            run_title.font.color.rgb = TAIGA_TEXT
            title.space_after = Pt(80)

            # Alatiedot (asiakas, projekti, päivämäärä)
            cust = payload.get("customer_name", "")
            proj = payload.get("project_name", "")
            date = payload.get("date_str", "")

            info = d.add_paragraph()
            info.alignment = WD_ALIGN_PARAGRAPH.CENTER
            info.add_run(f"Customer: {cust}\n").font.size = Pt(12)
            info.add_run(f"Project: {proj}\n").font.size = Pt(12)
            info.add_run(f"Date: {date}").font.size = Pt(12)

            # Tyhjä väli + sivunvaihto
            d.add_paragraph("")
            d.add_page_break()
        except Exception as e:
            print(f"[Warning] Failed to create cover page: {e}")

    d.sections[0].page_width, d.sections[0].page_height = Inches(11.69), Inches(8.27)

    # Otsikko
    title = d.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = title.add_run("Total Ownership Cost Summary")
    r.font.name = "Inter"
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = TAIGA_TEXT

# TCO-selostus
    p = d.add_paragraph()
    run = p.add_run(
        "Total Cost of Ownership (TCO) represents the sum of all costs "
        "incurred throughout a product’s entire lifecycle — including acquisition, "
        "operation, maintenance, downtime, and end-of-life. In Taiga’s context, "
        "TCO highlights how modular and circular solutions reduce long-term costs "
        "compared to traditional construction, especially when workspace needs "
        "and layouts change over time."
    )
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(80, 80, 80)
    p.paragraph_format.space_after = Pt(5)
    
    # Kevyt vihreä erotinviiva
    hr = d.add_paragraph()
    run_hr = hr.add_run(" ")
    run_hr.underline = True
    run_hr.font.color.rgb = TAIGA_ACCENT
    hr.paragraph_format.space_after = Pt(5)

    # Header-kentät
    cust = payload.get("customer_name", "")
    proj = payload.get("project_name", "")
    date = payload.get("date_str", "")

    _add_kv_paragraph(d, "Customer", cust)
    _add_kv_paragraph(d, "Project", proj)
    _add_kv_paragraph(d, "Date", date)

    d.add_paragraph("")

    # Tulosten yhteenveto
    _add_heading(d, "Results overview", level=2)
    res = payload.get("results", {}) or {}
    t_trad = res.get("TCO_TRAD_PV", 0.0)
    t_taiga = res.get("TCO_TAIGA_PV", 0.0)
    diff = res.get("DIFF_TRAD_TAIGA", 0.0)

    _add_kv_paragraph(d, "TCO TRAD PV", _fmt_eur(t_trad, 0))
    _add_kv_paragraph(d, "TCO TAIGA PV", _fmt_eur(t_taiga, 0))
    _add_kv_paragraph(d, "Difference (TRAD − TAIGA)", _fmt_eur(diff, 0))

    d.add_paragraph("")

    # Parametrit (tiivistetty)
    p = payload.get("params", {}) or {}
    if p:
        _add_heading(d, "Key parameters", level=2)

        def _get_pct(v):
            try:
                return f"{_fmt_num_2dec_fi(float(v) * 100)} %"
            except Exception:
                return str(v)

        def _fmt_generic(val):
            if isinstance(val, numbers.Number):
                return _fmt_num_2dec_fi(val)
            return str(val)

        for k in ["years", "wacc", "area_m2", "kwh_m2yr", "elec_price", "cycle_year"]:
            if k in p:
                if k == "wacc":
                    _add_kv_paragraph(d, "WACC", _get_pct(p[k]))
                elif k == "elec_price":
                    _add_kv_paragraph(d, "Electricity price (€/kWh)", _fmt_num_2dec_fi(p[k]))
                else:
                    _add_kv_paragraph(d, k.replace("_", " ").title(), _fmt_generic(p[k]))
        d.add_paragraph("")
        d.add_page_break()
    # Pivotit
    _add_pivot_table(d, "Taiga yearly breakdown (PV)", df_pivot_taiga)
    d.add_page_break()
    _add_pivot_table(d, "TRAD yearly breakdown (PV)", df_pivot_trad)
    d.add_page_break()
    _add_pivot_table(d, "Delta (Taiga − TRAD)", df_pivot_delta)

    # Byte-palautus
    bio = io.BytesIO()
    d.save(bio)
    bio.seek(0)
    return bio.getvalue()
