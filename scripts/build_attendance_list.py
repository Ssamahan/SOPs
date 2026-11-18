"""
Attendance list — Onsite Refresher Training, Use and Maintenance of the GeneXpert
Machine, Ruli District Hospital Laboratory, 18 November 2026.

Standalone builder: emits a Word document with the official hospital letterhead
(the same 7-line block, logo and address used by every SOP in this repository),
a blank attendance table for up to ten participants, and a facilitator/sign-off block.

Not part of build_sop.py — a form, not a controlled SOP. It is generated from the
same letterhead constants so that a change to the address or the telephone number
is made in one place.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from build_sop import LOGO_PATH, LETTERHEAD  # noqa: E402

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

INK = "1F3864"
GREY = "595959"
RULE = "B4C6E7"

TITLE = "Use and Maintenance of the GeneXpert Machine"
DATE_LONG = "Wednesday, 18 November 2026"
DATE_SHORT = "18 / 11 / 2026"
# No names: the sheet goes to the room blank so each person writes their own
# name and signs against it. Ten rows: the group is never more than ten, and the
# whole sheet, letterhead to sign-off, has to sit on one page.
BLANK_ROWS = 10
def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def rule(paragraph, size=6):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), RULE)
    pbdr.append(bottom)
    pPr.append(pbdr)


def cell_margins(table, top=0, bottom=0):
    """Trim the cell padding so the table rules stay tight to the text."""
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("bottom", bottom)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def heading(doc, text, space_before=9):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.font.size, r.font.bold = Pt(10), True
    r.font.color.rgb = RGBColor.from_string(INK)
    return p


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = Cm(1.0)
    sec.bottom_margin = Cm(1.2)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    # ---- letterhead: logo, then the address block --------------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), height=Cm(1.6))

    for text, style in LETTERHEAD:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        if style == "lab":
            r.font.size, r.font.bold = Pt(15), True
            r.font.color.rgb = RGBColor.from_string(INK)
        elif style == "kind":
            r.font.size, r.font.bold = Pt(9), True
            r.font.color.rgb = RGBColor.from_string(GREY)
        else:
            r.font.size, r.font.italic = Pt(7), True
            r.font.color.rgb = RGBColor.from_string(GREY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run("Laboratory Department \u00b7 (+250) 0787 055 121 \u00b7 "
                  "rulidhlab@gmail.com")
    r.font.size = Pt(7)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    rule(p, 8)

    # ---- document title ----------------------------------------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("ATTENDANCE LIST")
    r.font.size, r.font.bold = Pt(16), True
    r.font.color.rgb = RGBColor.from_string(INK)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Onsite Refresher Training")
    r.font.size, r.font.bold = Pt(11), True
    r.font.color.rgb = RGBColor.from_string(GREY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(TITLE)
    r.font.size, r.font.bold = Pt(12.5), True
    r.font.color.rgb = RGBColor.from_string(INK)

    # ---- training particulars ----------------------------------------------
    particulars = [
        ("Venue", "Ruli District Hospital \u2014 Laboratory Department, TB / Molecular "
                  "Laboratory"),
        ("Date", DATE_LONG + " \u00b7 one day"),
        ("Facilitator", "Kelvin SAMAHANI \u2014 Laboratory Manager"),
        ("Reference documents", "TB-011 (archiving and preventive maintenance); ADM-063 "
                                "(use, care and maintenance); ADM-065 (troubleshooting)"),
        ("Governed by", "ADM-064 (signing of SOPs; training, competence and authorisation "
                        "of staff)"),
    ]
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_margins(t, 5, 5)
    for label, value in particulars:
        row = t.add_row().cells
        row[0].width = Cm(4.1)
        row[1].width = Cm(12.9)
        pa = row[0].paragraphs[0]
        ra = pa.add_run(label)
        ra.font.size, ra.font.bold = Pt(8.5), True
        ra.font.color.rgb = RGBColor.from_string(INK)
        pa.paragraph_format.space_after = Pt(0)
        shade(row[0], "EFF3FA")
        pb = row[1].paragraphs[0]
        rb = pb.add_run(value)
        rb.font.size = Pt(8.5)
        pb.paragraph_format.space_after = Pt(0)

    # ---- attendance table ---------------------------------------------------
    heading(doc, "ATTENDANCE", space_before=6)

    headers = ["No.", "Full name (write in block letters)", "Designation / Department",
               "Date", "Signature"]
    widths = [1.0, 6.6, 3.6, 2.2, 3.6]
    assert abs(sum(widths) - 17.0) < 2e-9, sum(widths)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].width = Cm(widths[i])
        pa = hdr[i].paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ra = pa.add_run(h)
        ra.font.size, ra.font.bold = Pt(8.5), True
        ra.font.color.rgb = RGBColor.from_string("FFFFFF")
        pa.paragraph_format.space_after = Pt(2)
        shade(hdr[i], INK)

    for n in range(1, BLANK_ROWS + 1):
        cells = t.add_row().cells
        for i, value in enumerate((str(n), "", "", "", "")):
            cells[i].width = Cm(widths[i])
            pa = cells[i].paragraphs[0]
            if i == 0:
                pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ra = pa.add_run(value)
            ra.font.size = Pt(9)
            pa.paragraph_format.space_after = Pt(3.2)
            pa.paragraph_format.space_before = Pt(3.2)

    # ---- training content and sign-off, side by side ----------------------
    # The sheet has to be one page. Content runs in the left column and the
    # facilitator block sits beside it in the right, so nothing is stranded in
    # a band of white space below the list.
    outer = doc.add_table(rows=1, cols=2)
    outer.autofit = False
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    left, right = outer.rows[0].cells
    left.width, right.width = Cm(11.6), Cm(5.4)
    for cell in (left, right):
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        for edge in ("top", "bottom", "left", "right"):
            pass

    content = [
        "Principle of the GeneXpert Dx system and the Xpert MTB/RIF Ultra assay; what the "
        "module does and what it does not do.",
        "Daily checks before the first run: power, module status, the self-test, and the "
        "reagent and cartridge storage conditions.",
        "Sample processing and cartridge loading, including the volume, the transfer pipette "
        "and the handling of a specimen that is too thick or too thin.",
        "Reading and releasing a result, and the meaning of MTB DETECTED, MTB NOT DETECTED, "
        "INVALID, ERROR and NO RESULT \u2014 and which of them may be reported to a clinician.",
        "Repeat criteria: when a repeat is permitted, when it is not, and what is recorded "
        "each time.",
        "Troubleshooting to module level, with the error codes in ADM-065 and the correlation "
        "rules for failures across modules.",
        "Preventive maintenance: weekly cleaning, the module cleaning kit, the probe, the "
        "filter and the fan; the schedule and who signs it.",
        "Data archiving and the export of results for the record, per TB-011.",
        "Biosafety and waste handling for cartridges and consumables.",
        "What the operator may fix, and what must be escalated to the Laboratory Manager and "
        "the supplier.",
    ]

    p = left.paragraphs[0]
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2.5)
    r = p.add_run("CONTENT COVERED")
    r.font.size, r.font.bold = Pt(10), True
    r.font.color.rgb = RGBColor.from_string(INK)

    for i, item in enumerate(content, 1):
        p = left.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.space_after = Pt(0.8)
        p.paragraph_format.line_spacing = 0.96
        r = p.add_run(f"{i}.  ")
        r.font.size, r.font.bold = Pt(8.5), True
        r.font.color.rgb = RGBColor.from_string(INK)
        r = p.add_run(item)
        r.font.size = Pt(8.5)

    # sign-off block, right column
    p = right.paragraphs[0]
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2.5)
    r = p.add_run("SIGN-OFF")
    r.font.size, r.font.bold = Pt(10), True
    r.font.color.rgb = RGBColor.from_string(INK)

    p = right.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Facilitator")
    r.font.size, r.font.bold = Pt(8.5), True
    r.font.color.rgb = RGBColor.from_string(GREY)
    p = right.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("Kelvin SAMAHANI\nLaboratory Manager")
    r.font.size = Pt(8.5)

    p = right.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    rule(p, 4)
    p = right.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Signature")
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor.from_string(GREY)

    p = right.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Date")
    r.font.size, r.font.bold = Pt(8.5), True
    r.font.color.rgb = RGBColor.from_string(GREY)
    p = right.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(DATE_SHORT)
    r.font.size = Pt(8.5)

    # ---- footer ---------------------------------------------------------------
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run("Ruli District Hospital \u00b7 Laboratory Department \u00b7 Attendance "
                       "list \u2014 Onsite Refresher Training, Use and Maintenance of the "
                       "GeneXpert Machine \u00b7 18 November 2026")
    r.font.size = Pt(7)
    r.font.color.rgb = RGBColor.from_string(GREY)

    out = ROOT / "forms" / ("Ruli DH Laboratory - Attendance List - GeneXpert Refresher "
                            "Training - 18 November 2026.docx")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print("written:", out)
    return out


if __name__ == "__main__":
    build()
