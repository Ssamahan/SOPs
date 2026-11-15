#!/usr/bin/env python3
"""Generate the Master List insert sheet from the document REGISTER.

The Master List in Drive is the controlled register. This script produces the
sheet that carries a change into it, in the Master List's own column order, so
the rows can be typed or pasted into the crosscutting section in one action.

Because the sheet is generated from the same REGISTER as the SOPs themselves,
the insert and the documents cannot disagree — which is the whole point: a code
change must reach every document that references it, including the register.

    .venv/bin/python scripts/build_masterlist_insert.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from build_sop import REGISTER, SERIES_KNOWN_MAX, code, form

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "MASTER LIST insert - code register (Nov 2026).docx"

INK = "1F3864"      # navy, as in the SOP house style
GREY = "595959"
LINE = "BFBFBF"
EFFECTIVE = "14/11/2026"   # proposed, as everywhere else in this block
REVIEW = "14/11/2028"

# --- the defects found in the Drive Master List while cross-checking ----------
# Recorded, not silently corrected: amending the Master List is a controlled act
# and needs its own sign-off.
DEFECTS = [
    ("Serial 001 is used twice",
     "RliDHL-ADM-SOP-001-VERS-004 appears as both “Laboratory Temperature Control” "
     "and “Document and Record Control”. Document and Record Control is 004 in the "
     "numbering actually in force — it is the code every SOP cites for clause 8.3 — "
     "so the row reading 001 for it is the error.",
     "Correct that row to RliDHL-ADM-SOP-004-VERS-004."),
    ("Serial 034 is used twice",
     "RliDHL-ADM-SOP-034-VERS-003 appears as both “Laboratory Waste Management” and "
     "“Handling of biological spills”.",
     "Give the spills SOP its own serial at its next revision; until then, cite it by "
     "title, not by number."),
    ("Crosscutting serials are not contiguous",
     "The first six crosscutting SOPs sit at 001, 002, 004, 020, 024 and 034, so the "
     "series already skips 003 and 005-019.",
     "Adopt the no-gap rule: treat those serials as retired numbers, or renumber the "
     "rows once, at each document's next revision. Not both."),
    ("Document and Record Control is cited under two numbers",
     "Because of the 001/004 duplication, a reader following the 8.3 cross-reference "
     "cannot tell which document to open.",
     "Resolved by the first correction above."),
]


def shade(cell, hex_fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), hex_fill)
    tcPr.append(el)


def h(text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK)


def body(text: str, size=9.5, after=4) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.font.size = Pt(size)


def bullet(text: str, size=9.5) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("•  " + text)
    r.font.size = Pt(size)


def table(headers, rows, widths, body_size=8.5):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    for j, text in enumerate(headers):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(text)
        r.font.bold = True
        r.font.size = Pt(body_size)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, INK)
    for i, vals in enumerate(rows, start=1):
        for j, val in enumerate(vals):
            c = t.rows[i].cells[j]
            c.text = ""
            r = c.paragraphs[0].add_run(str(val))
            r.font.size = Pt(body_size)
    return t


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
for attr in ("left_margin", "right_margin"):
    setattr(sec, attr, Cm(2.0))
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)

# ---- letterhead ----
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
r = p.add_run("RULI DISTRICT HOSPITAL – LABORATORY DEPARTMENT")
r.font.size = Pt(11)
r.font.bold = True
r.font.color.rgb = RGBColor.from_string(INK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(10)
r = p.add_run("P.O. Box 1285 Kigali · Tel: (+250) 0781 316 368 · "
              "ruli.hospital@moh.gov.rw · rulihospital.rw")
r.font.size = Pt(8)
r.font.italic = True
r.font.color.rgb = RGBColor.from_string(GREY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run("MASTER LIST – INSERT SHEET")
r.font.size = Pt(15)
r.font.bold = True
r.font.color.rgb = RGBColor.from_string(INK)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(12)
r = p.add_run("Code corrections, new SOP 060, and the forms bound to it (November 2026)")
r.font.size = Pt(10.5)
r.font.bold = True
r.font.color.rgb = RGBColor.from_string(GREY)

# ---- what this sheet is ----
h("What this sheet is")
body(
    "Reviewing the SOP collection against this Master List found four problems with "
    "document numbers: two serials are each used by two different documents, two SOPs "
    "were carrying a number that already belonged to another document, and the forms in "
    "the SOPs were identified by a code that tied them to no SOP at all. This sheet sets "
    "out the corrections, the one new document that has to exist for the signature pages "
    "to have a parent, and the forms, in the Master List's own column order, so the rows "
    "can be typed or pasted into the crosscutting section in one action. It carries the "
    "same signature block, so the addition is authorised at the same time as the "
    "corrections it makes."
)
body(
    "Nothing here is in force until it is signed. Where a correction and a document "
    "disagree, the signed Master List wins and the document is corrected."
)

# ---- A. corrections ----
h("A.  CORRECTIONS to existing rows in this Master List")
body(
    "These are defects in the register itself. They are listed rather than corrected "
    "silently, because amending the Master List is a controlled act under the Document "
    "and Record Control SOP.", size=9)
table(
    ["#", "Defect", "What is wrong", "Correction proposed"],
    [[str(i + 1), d[0], d[1], d[2]] for i, d in enumerate(DEFECTS)],
    [0.8, 3.6, 7.0, 5.0], body_size=8,
)

# ---- B. SOPs ----
h("B.  STANDARD OPERATING PROCEDURES – rows to add or amend")
rows = []
n = 0
for key, r in REGISTER.items():
    if key in ("doc_control", "pre_exam", "glucose"):
        continue          # already correct in the Master List
    n += 1
    status = ("NEW – serial 060 is the next free number after the 053–059 "
              "equipment block" if r["code"].endswith("060")
              else "RENUMBERED – was RliDHL-IH-SOP-001-VERS-001, which already "
                   "belongs to Vein Blood Collection (HEM, v003)")
    rows.append([n, code(key), r["title"], EFFECTIVE, REVIEW,
                 "RliDHL Office; all bench rooms", status])
table(["N°", "Document Code", "Document Title", "Effective date", "Review date",
       "Locations", "Status"], rows, [0.8, 3.9, 4.4, 1.9, 1.9, 2.6, 3.5], body_size=7.5)

body(
    "The venous SOP appears in the Master List as “Vein Blood Collection and Preparation "
    "of Hematology sample”, RliDHL-HEM-SOP-008-VERS-003. That is a different document "
    "with a different scope. Two documents cannot answer to one number, and the "
    "hematology SOP is the one already in force, so the new pre-examination SOP moves to "
    "IH-SOP-002 and the HEM entry is left alone.", size=9
)
body(
    f"Serials continue from {SERIES_KNOWN_MAX['RliDHL-ADM-SOP']:03d}, the highest in use "
    f"in the ADM series. The next free ADM serial after this insert is 061; check the "
    f"Master List again before allocating it.", size=9
)

# ---- C. forms ----
h("C.  FORMS – recorded against the SOP that controls them")
form_rows = [
    [form("training", "F01"),
     "Training and competence acknowledgement, competence assessment and authorisation "
     "record — the section 21 signature page of every SOP",
     "060", "On training and on authorisation", EFFECTIVE, "RliDHL Office; all rooms"],
    [form("training", "F01-2"),
     "Continuation sheet, for a section with more than 18 staff",
     "060", "As required", EFFECTIVE, "RliDHL Office; all rooms"],
]
table(["Form number", "Form title", "Parent SOP", "Frequency", "Effective date",
       "Locations"], form_rows, [3.6, 6.2, 1.5, 2.5, 1.9, 2.4], body_size=8)

body(
    "The forms were previously cited in the SOPs as F-TRN-001 and F-TRN-001-2. That "
    "breaks two rules at once: TRN is not one of the laboratory's area codes, and the "
    "form carried no parent SOP, so nothing controlled it. The forms now carry the code "
    "of the SOP that controls them, which is the convention already used for "
    "RliDHL-ADM-SOP-058-F01 and the rest of the November block.", size=9
)

# ---- D. why 060 ----
h("D.  Why SOP 060 has to exist")
body(
    "Every SOP in the collection ends with a signature page that records that a member "
    "of staff was trained on it, that their competence was assessed, and that they are "
    "authorised to perform the procedure. ISO 15189:2022 requires those records (6.2.2 "
    "and 6.2.3) and requires them to be controlled (8.4)."
)
body(
    "The Master List contains no SOP that governs training, competence assessment, "
    "authorisation, or the signing of SOPs. The signature page and its forms therefore "
    "had no parent document, which is why they had no control either. "
    "RliDHL-ADM-SOP-060-VERS-001 is that parent."
)
body(
    "Two existing SOPs cover part of the ground and are not duplicates of it: "
    "RliDHL-ADM-SOP-020 covers specimen collection, not people, and RliDHL-ADM-SOP-024 "
    "covers quality control, not competence."
)

# ---- E. status ----
h("E.  Status and what still has to happen")
bullet("All entries are first issues at VERS-001, except glucose, which stays at 005.")
bullet("The effective date of 14/11/2026 is the proposed date. If signature happens "
       "later, the effective date in this insert, in the SOP cover block and in the file "
       "name must be changed to the date of authorisation.")
bullet("SOP 060 has not been written yet. The serial is reserved here so that the "
       "signature pages and their forms have a code to point at; the SOP itself must be "
       "prepared, reviewed and approved before those records can be relied on.")
bullet("The four corrections in section A need the same sign-off as the additions. "
       "Until then the Master List and the SOPs disagree in the four places listed.")

# ---- F. authorisation ----
h("F.  Authorisation of this addition to the Master List")
table(["Role", "Name", "Signature", "Date"],
      [["Prepared by – Head of Laboratory", "", "", ""],
       ["Reviewed by – BMeT", "", "", ""],
       ["Approved by – Director General", "", "", ""]],
      [6.0, 5.0, 4.0, 2.5], body_size=9)
for row in doc.tables[-1].rows[1:]:
    row.height = Cm(1.5)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(16)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Ruli District Hospital – Laboratory Department  •  Master List insert  "
              "•  prepared 15 November 2026")
r.font.size = Pt(7.5)
r.font.italic = True
r.font.color.rgb = RGBColor.from_string(GREY)

OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUT)
print(f"Written: {OUT.relative_to(ROOT)}")
print("\nRows that must reach the Master List:")
for key, r in REGISTER.items():
    print(f"  {code(key):28s} {r['title']}")
