"""
Attendance list — Onsite Refresher Training, Use and Maintenance of the GeneXpert
Machine, Ruli District Hospital Laboratory, 15 November 2026.

Standalone builder: emits a Word document with the official hospital letterhead
(the same 7-line block, logo and address used by every SOP in this repository),
an attendance table for the eight active staff, and a facilitator/sign-off block.

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
DATE_LONG = "Sunday, 15 November 2026"
ATTENDEES = [
    ("1", "Kelvin SAMAHANI", "Laboratory Manager"),
    ("2", "Jean Paul NSENGIYUMVA", "Laboratory Quality Manager"),
    ("3", "Aurelie YANSONEYE", "HIV Laboratory Quality Coordinator"),
    ("4", "Beatrice NGEZEKURE", "Lab Technologist \u2014 Reception and Phlebotomy"),
    ("5", "Patrick ISHIMWE", "Lab Technician"),
    ("6", "Bernadette MUHAWENIMANA", "Lab Technician"),
    ("7", "Jeanne Claudine UWAYISABA", "Lab Technician"),
    ("8", "Dieudonne NIYIDUHABYOSE", "Laboratory Safety Officer"),
]


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


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = Cm(1.2)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    # ---- letterhead: logo, then the 7-line address block -------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), height=Cm(2.3))

    for text, style in LETTERHEAD:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(text)
        if style == "lab":
            r.font.size, r.font.bold = Pt(16), True
            r.font.color.rgb = RGBColor.from_string(INK)
        elif style == "kind":
            r.font.size, r.font.bold = Pt(9.5), True
            r.font.color.rgb = RGBColor.from_string(GREY)
        else:
            r.font.size, r.font.italic = Pt(7.5), True
            r.font.color.rgb = RGBColor.from_string(GREY)

    # the hospital's own telephone, as printed on the laboratory's letterhead
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Laboratory Department \u00b7 (+250) 0787 055 121 \u00b7 "
                  "rulidhlab@gmail.com")
    r.font.size = Pt(7.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    rule(p, 8)

    # ---- document title ----------------------------------------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("ATTENDANCE LIST")
    r.font.size, r.font.bold = Pt(17), True
    r.font.color.rgb = RGBColor.from_string(INK)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run("Onsite Refresher Training")
    r.font.size, r.font.bold = Pt(12), True
    r.font.color.rgb = RGBColor.from_string(GREY)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(TITLE)
    r.font.size, r.font.bold = Pt(13.5), True
    r.font.color.rgb = RGBColor.from_string(INK)

    # ---- training particulars ---------------------------------------------
    particulars = [
        ("Training", "Onsite Refresher Training \u2014 " + TITLE),
        ("Venue", "Ruli District Hospital \u2014 Laboratory Department, TB / Molecular "
                  "Laboratory"),
        ("Date", DATE_LONG),
        ("Duration", "One day"),
        ("Organised by", "Ruli District Hospital, Laboratory Department"),
        ("Facilitator", "Kelvin SAMAHANI \u2014 Laboratory Manager"),
        ("Reference documents", "TB-011 (Data archiving and preventive maintenance of the "
                                "GeneXpert Dx system); ADM-063 (Use, Care and Preventive "
                                "Maintenance of the GeneXpert Dx System); ADM-065 "
                                "(Troubleshooting of laboratory analysers)"),
        ("Governed by", "ADM-064 \u2014 Signing of SOPs, and the Training, Competence and "
                        "Authorisation of Staff"),
    ]
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for label, value in particulars:
        row = t.add_row().cells
        row[0].width = Cm(4.4)
        row[1].width = Cm(12.6)
        pa = row[0].paragraphs[0]
        ra = pa.add_run(label)
        ra.font.size, ra.font.bold = Pt(9), True
        ra.font.color.rgb = RGBColor.from_string(INK)
        pa.paragraph_format.space_after = Pt(3)
        shade(row[0], "EFF3FA")
        pb = row[1].paragraphs[0]
        rb = pb.add_run(value)
        rb.font.size = Pt(9)
        pb.paragraph_format.space_after = Pt(3)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)

    # ---- attendance table --------------------------------------------------
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("ATTENDANCE \u2014 laboratory staff")
    r.font.size, r.font.bold = Pt(10.5), True
    r.font.color.rgb = RGBColor.from_string(INK)

    headers = ["No.", "Full name", "Designation", "Signature"]
    widths = [1.3, 5.6, 4.6, 5.5]
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].width = Cm(widths[i])
        pa = hdr[i].paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ra = pa.add_run(h)
        ra.font.size, ra.font.bold = Pt(9), True
        ra.font.color.rgb = RGBColor.from_string("FFFFFF")
        pa.paragraph_format.space_after = Pt(2)
        shade(hdr[i], INK)

    for no, name, role in ATTENDEES:
        cells = t.add_row().cells
        for i, value in enumerate((no, name, role, "")):
            cells[i].width = Cm(widths[i])
            pa = cells[i].paragraphs[0]
            if i == 0:
                pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ra = pa.add_run(value)
            ra.font.size = Pt(9)
            pa.paragraph_format.space_after = Pt(5)
            pa.paragraph_format.space_before = Pt(5)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p = doc.add_paragraph()
    r = p.add_run("Eight (8) active laboratory staff. A member of staff who attends and "
                  "cannot sign on the day signs against their name within 24 hours; the "
                  "facilitator initials the late signature.")
    r.font.size = Pt(8)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY)
    p.paragraph_format.space_after = Pt(8)

    # ---- training content --------------------------------------------------
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("CONTENT COVERED")
    r.font.size, r.font.bold = Pt(10.5), True
    r.font.color.rgb = RGBColor.from_string(INK)

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
    for i, item in enumerate(content, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{i}.  ")
        r.font.size, r.font.bold = Pt(9), True
        r.font.color.rgb = RGBColor.from_string(INK)
        r = p.add_run(item)
        r.font.size = Pt(9)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)

    # ---- assessment and sign-off ------------------------------------------
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("ASSESSMENT ON THE DAY")
    r.font.size, r.font.bold = Pt(10.5), True
    r.font.color.rgb = RGBColor.from_string(INK)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("Each participant performs one observed run on the GeneXpert, from sample "
                  "to result, and is assessed against the competency checklist. A participant "
                  "who does not pass is not signed off for independent operation and is "
                  "re-trained. The result of the assessment is recorded on the competency "
                  "record, not on this sheet.")
    r.font.size = Pt(9)

    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    w = [4.25, 4.25, 4.25, 4.25]
    hdr = t.rows[0].cells
    for i, h in enumerate(["Facilitator", "Signature", "Date", "Competency records updated"]):
        hdr[i].width = Cm(w[i])
        pa = hdr[i].paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ra = pa.add_run(h)
        ra.font.size, ra.font.bold = Pt(8.5), True
        ra.font.color.rgb = RGBColor.from_string("FFFFFF")
        pa.paragraph_format.space_after = Pt(2)
        shade(hdr[i], INK)

    cells = t.add_row().cells
    for i in range(4):
        cells[i].width = Cm(w[i])
        pa = cells[i].paragraphs[0]
        pa.paragraph_format.space_before = Pt(14)
        pa.paragraph_format.space_after = Pt(14)
    cells[0].paragraphs[0].add_run("Kelvin SAMAHANI\nLaboratory Manager").font.size = Pt(8.5)
    cells[2].paragraphs[0].add_run("15 / 11 / 2026").font.size = Pt(8.5)

    # ---- footer ------------------------------------------------------------
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run("Ruli District Hospital \u00b7 Laboratory Department \u00b7 Attendance "
                       "list \u2014 Onsite Refresher Training, Use and Maintenance of the "
                       "GeneXpert Machine \u00b7 15 November 2026")
    r.font.size = Pt(7)
    r.font.color.rgb = RGBColor.from_string(GREY)

    out = ROOT / "forms" / ("Ruli DH Laboratory - Attendance List - GeneXpert Refresher "
                            "Training - 15 November 2026.docx")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print("written:", out)
    return out


if __name__ == "__main__":
    build()
