# proposal_doc.py
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generate_proposal_doc(payload: dict, locale: str = "fi_FI") -> bytes:
    """
    Luo ammattimaisesti muotoillun Word-yhteenvedon TCO-laskennasta asiakkaalle.
    payload = {
        'customer_name': str,
        'project_name': str,
        'date_str': str,
        'params': dict,
        'results': dict
    }
    """
    d = Document()

    # ----------- Kansilehti -----------
    d.add.picture("logo_logo.png", width=Inches(2))
    title = d.add_heading("Total Cost of Ownership (TCO) – Comparison Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    d.add_paragraph()
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        f"Customer: {payload.get('customer_name','-')}\n"
        f"Project: {payload.get('project_name','-')}\n"
        f"Date: {payload.get('date_str','-')}"
    )
    run.font.size = Pt(12)

    d.add_page_break()

    # ----------- Parametrit taulukossa -----------
    d.add_heading("Key Parameters", level=1)
    params = payload.get("params", {})

    if params:
        table = d.add_table(rows=1, cols=2)
        table.style = "Light Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Parameter"
        hdr[1].text = "Value"

        for k, v in params.items():
            row = table.add_row().cells
            row[0].text = str(k)
            row[1].text = str(v)

    d.add_page_break()

    # ----------- Tulokset taulukossa -----------
    d.add_heading("Results (Present Value)", level=1)
    results = payload.get("results", {})

    # Tulostetaan TRAD vs TAIGA erot taulukossa
    table = d.add_table(rows=1, cols=3)
    table.style = "Light Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "TRAD (€)"
    hdr[2].text = "TAIGA (€)"

    trad_val = results.get("TCO_TRAD_PV", "-")
    taiga_val = results.get("TCO_TAIGA_PV", "-")
    diff_val = results.get("DIFF_TRAD_TAIGA", "-")

    row1 = table.add_row().cells
    row1[0].text = "Total TCO (PV)"
    row1[1].text = f"{trad_val:,.0f}" if isinstance(trad_val,(int,float)) else str(trad_val)
    row1[2].text = f"{taiga_val:,.0f}" if isinstance(taiga_val,(int,float)) else str(taiga_val)

    d.add_paragraph()
    d.add_heading("Difference", level=2)
    d.add_paragraph(
        f"Difference (TRAD − TAIGA): "
        f"{diff_val:,.0f} €" if isinstance(diff_val,(int,float)) else str(diff_val)
    )

    d.add_page_break()

    # ----------- Yhteenveto -----------
    d.add_heading("Summary & Conclusion", level=1)
    d.add_paragraph(
        "This analysis demonstrates the financial advantages of Taiga’s modular workspace solutions "
        "compared to traditional construction methods. Over the selected time horizon, the Total Cost "
        "of Ownership (TCO) for Taiga is lower, reflecting benefits from reduced maintenance, minimized "
        "downtime, and potential buyback value. \n\n"
        "Beyond cost savings, Taiga offers flexibility, scalability, and faster deployment, supporting "
        "customers’ evolving needs while aligning with sustainability objectives."
    )

    # ----------- Palautus bytes muodossa -----------
    bio = BytesIO()
    d.save(bio)
    bio.seek(0)
    return bio.read()
