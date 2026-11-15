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
    ("008 on two documents",
     "RliDHL-ADM-SOP-008 is the legacy “Authorization” (v001) and also “Use and "
     "Maintenance of Biomedical Equipment” VERS-005 (30/04/2026). This one has already "
     "been propagated: all seven SOPs in the November 2026 equipment block cite "
     "ADM-SOP-008-VERS-005 in their reference lists, so a reader following that citation "
     "may open the wrong document.",
     "The Laboratory Manager and Director General decide which document keeps 008; the "
     "other is renumbered and every reference list updated in the same act."),
    ("049 used twice",
     "RliDHL-ADM-SOP-049-VERS-001 is both “BIOLOGICAL REFERENCE INTERVALS-1” and "
     "“MEET THE CLINICIANS”.",
     "One of the two takes the next free serial at its next revision."),
    ("035 never used",
     "The series runs 034 then 036. Nothing is registered at 035.",
     "Declare 035 retired, so the gap is not later mistaken for a lost document."),
    ("053–059 cannot be proven free",
     "Two management SOPs above 052 exist on disk but were never registered, notably "
     "“52. Resolution of complaints” and ADM-SOP-008-VERS-005. The November pack therefore "
     "starts at 060 rather than 053.",
     "Recorded as a limit, not a decision. Do not allocate 053–059 without checking the "
     "files first."),
    ("Unregistered file on disk",
     "“52. Resolution of complaints” exists but is not in the Master List.",
     "Register it, or withdraw it."),
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
r = p.add_run("Code corrections, new SOP 064, and the forms bound to it (November 2026)")
r.font.size = Pt(10.5)
r.font.bold = True
r.font.color.rgb = RGBColor.from_string(GREY)

# ---- what this sheet is ----
h("What this sheet is")
body(
    "Cross-checking the SOP collection against this Master List found that the register "
    "itself carries five numbering defects, two of which are already in force: number 008 "
    "identifies two documents and is cited by all seven SOPs in the November 2026 equipment "
    "block, and number 049 is carried twice. None of these can be settled by whoever wrote "
    "the documents — they need the Laboratory Manager and the Director General. This sheet "
    "records them so they do not disappear, and adds the one new document that has to exist "
    "for the signature pages to have a parent, in the Master List's own column order, with "
    "the same signature block."
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
    if key in ("doc_control", "pre_exam"):
        continue          # already correct in the Master List
    n += 1
    serial = r["code"].rsplit("-", 1)[1]
    if key == "venous":
        status = ("AMENDED – serial 001 already registered for this subject; this is the "
                  "revision to VERS-002, not a new number")
    elif key == "glucose":
        status = ("CODE UNCHANGED – 001 kept, version 005. The other holder of 001, the "
                  "Humastar 80 SOP, is WITHDRAWN — the analyser is decommissioned — not "
                  "renumbered")
    elif key == "training":
        status = ("NEW – first serial after the 060–063 equipment block of November 2026; "
                  "053–059 are not shown to be free")
    elif "-ADM-" in r["code"]:
        status = ("NEW – first serial after the 060–063 equipment block of "
                  "November 2026; 053–059 are not shown to be free")
    else:
        status = (f"NEW – serial {serial} is not carried in this Master List "
                  f"for the {r['code'].split('-')[1]} series")
    rows.append([n, code(key), r["title"], EFFECTIVE, REVIEW,
                 "RliDHL Office; all bench rooms", status])
# The legacy Biochemistry and Haematology rows must not be renumbered. Their equipment
# has been decommissioned, so they are superseded in place: the serial stays with the
# document and nothing is re-issued under that number.
for code_ref, title, note in [
    ("RliDHL-BIO-SOP-001-VERS-002", "Use and preventive maintenance of Humastar 80",
     "WITHDRAWN – Humastar 80 decommissioned. Serial 001 also held by the glucose SOP, "
     "which is the live document under that number"),
    ("RliDHL-BIO-SOP-002-VERS-001",
     "Sample analysis for Clinical Biochemistry using Humalyzer 3500",
     "WITHDRAWN – Humalyzer 3500 decommissioned. Superseded in place; nothing replaces it"),
    ("RliDHL-BIO-SOP-003 to 009", "Urea, ALAT, ASAT, creatinine, glucose prep, "
     "bilirubin total and direct — all by use of Humastar 80",
     "WITHDRAWN as a block – every one is a Humastar 80 procedure and the analyser is gone"),
    ("RliDHL-HEM-SOP-002-VERS-001", "Hematology sample analysis using Sysmex Pochi 100i",
     "WITHDRAWN – Pochi decommissioned. No Pochi is in service"),
    ("RliDHL-HEM-SOP-003-VERS-001", "Hematology sample analysis using Humacount 5",
     "WITHDRAWN – Humacount 5 decommissioned"),
    ("RliDHL-HEM-SOP-004-VERS-001", "Hematology sample analysis using Sysmex KX.N21",
     "WITHDRAWN – KX-21N decommissioned"),
]:
    n += 1
    rows.append([n, code_ref, title, "—", "—", "—", note])

table(["N°", "Document Code", "Document Title", "Effective date", "Review date",
       "Locations", "Status"], rows, [0.8, 3.9, 4.4, 1.9, 1.9, 2.6, 3.5], body_size=7.5)

body(
    "The venous SOP is RliDHL-HEM-SOP-001-VERS-002. It is a revision of the registered "
    "“Vein blood collection and preparation of hematology sample”, not a new document, "
    "so it keeps serial 001 and the version goes to 002. No IH series is used: the "
    "register has no such department, and inventing one would have given the laboratory "
    "two SOPs for drawing blood.", size=9
)
body(
    f"The ADM series stands at {SERIES_KNOWN_MAX['RliDHL-ADM-SOP']:03d}: serials 001–049 are "
    f"accounted for (035 was never used) and the November 2026 equipment block takes 050–052 "
    f"and 060–063. That leaves 064 as the next free ADM serial, which is what this SOP takes. "
    f"053–059 are not shown to be free — two management SOPs above 052 exist on disk but were "
    f"never registered — so they must not be allocated without checking the files.", size=9
)

# ---- C. forms ----
h("C.  FORMS – recorded against the SOP that controls them")
form_rows = [
    [form("training", "F01"),
     "Training and competence acknowledgement, competence assessment and authorisation "
     "record — the section 21 signature page of every SOP",
     "064", "On signing, and on authorisation; cites ADM-FORM-014 and ADM-FORM-001", EFFECTIVE, "RliDHL Office; all rooms"],
    [form("training", "F01-2"),
     "Continuation sheet, for a section with more than 18 staff",
     "064", "As required", EFFECTIVE, "RliDHL Office; all rooms"],
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

body(
    "Serial 001 was carried twice in the Biochemistry series: by this glucose SOP at "
    "VERS-005, and by “Use and preventive maintenance of Humastar 80” at VERS-002. "
    "Serial 001 is retained by the glucose SOP. The Humastar SOP is not moved and not "
    "reissued — the Humastar 80 has been decommissioned, so that document is withdrawn "
    "and its serial stays with it. Nothing is skipped and nothing is resurrected.",
    size=9
)
body(
    "The same applies to the rest of the legacy Biochemistry and Haematology rows. "
    "Humalyzer 3500, Humacount 5, Sysmex KX-21N and Sysmex Pochi have all been "
    "decommissioned. Their SOPs are marked superseded in place, never renumbered. The "
    "analysers actually in service are the Sysmex XN-330, XQ-320 and Alere Pima CD4.",
    size=9
)

# ---- D. why 064 ----
h("D.  Why SOP 064 has to exist")
body(
    "Every SOP in the collection ends with a signature page that records that a member "
    "of staff was trained on it, that their competence was assessed, and that they are "
    "authorised to perform the procedure. ISO 15189:2022 requires those records (6.2.2 "
    "and 6.2.3) and requires them to be controlled (8.4)."
)
body(
    "The Master List already carries SOP 033 (Laboratory personnel training) and SOP 002 "
    "(Competency assessment), so the training and assessment duties are not missing. What "
    "is missing is the SOP that governs signing an SOP: who signs, in what order, what the "
    "signature page is evidence of, and how the authorisation is recorded. Every SOP in the "
    "collection ends with such a page, and no document owns it."
)
body(
    "SOP 064 must cross-reference 033 and 002 rather than duplicate them. Writing it as a "
    "general training SOP would create a second source of truth for competence, which is "
    "worse than the gap."
)
body(
    "Two controlled forms already exist for the record itself: ADM-FORM-014 “Signature "
    "page” and ADM-FORM-001 “Staff competency assessment form”. Form 064-F01 cites both "
    "and replaces neither, so no existing form has to be withdrawn and the laboratory is "
    "not given two instruments for one purpose. Both form numbers are printed on the "
    "signature page of every SOP."
)

# ---- E. status ----
h("E.  Status and what still has to happen")
bullet("All entries are first issues at VERS-001, except glucose, which stays at 005.")
bullet("The effective date of 14/11/2026 is the proposed date. If signature happens "
       "later, the effective date in this insert, in the SOP cover block and in the file "
       "name must be changed to the date of authorisation.")
bullet("SOP 064 has not been written yet. The serial is reserved here so that the "
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
