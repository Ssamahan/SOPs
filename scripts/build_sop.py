#!/usr/bin/env python3
"""
Build an ISO 15189:2022-compliant sample SOP in two formats from one
content model:

  sop/SOP-PRE-001-Venous-Blood-Collection.docx   (editable Word template)
  sop/SOP-PRE-001-Venous-Blood-Collection.pdf    (print-exact preview)

Layout rules implemented (see docs/ISO-15189-2022-SOP-Guide.md):
  * Fresh page starts: cover | revision history + table of contents |
    body | Annex A training & signature sheet.
  * Continuous flow inside the body (no wasted pages).
  * Footer on every page: document number | version | "Page X of Y" |
    controlled-copy statement.
  * Approval signatures (prepared / reviewed / approved) on the cover.
  * Annex A provides 15 signature rows so at least 12 staff can sign.

Usage:  .venv/bin/python scripts/build_sop.py
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCX_PATH = ROOT / "sop" / "SOP-PRE-001-Venous-Blood-Collection.docx"
PDF_PATH = ROOT / "sop" / "SOP-PRE-001-Venous-Blood-Collection.pdf"

# ---------------------------------------------------------------------------
# Shared content model
# ---------------------------------------------------------------------------

META = {
    "lab_name": "[LABORATORY NAME]",
    "logo": "[ LABORATORY LOGO ]",
    "doc_no": "SOP-PRE-001",
    "title": "Venous Blood Sample Collection by Venipuncture",
    "version": "1.0",
    "effective": "01 December 2026",
    "review": "01 December 2028",
    "supersedes": "\u2014 (First issue)",
    "owner": "Quality Officer",
    "classification": "Controlled Document \u2014 Master copy retained by the Quality Office",
}

DOC_CONTROL_ROWS = [
    ("Document Number", META["doc_no"]),
    ("Document Title", META["title"]),
    ("Version / Revision", META["version"]),
    ("Effective Date", META["effective"]),
    ("Next Review Date", META["review"] + "  (review at least every 2 years or on change)"),
    ("Supersedes", META["supersedes"]),
    ("Document Owner", META["owner"]),
    ("Classification", META["classification"]),
]

APPROVAL_ROWS = [
    ("Prepared by", "[Name \u2014 Quality Officer]"),
    ("Reviewed by", "[Name \u2014 Laboratory Manager]"),
    ("Approved by", "[Name \u2014 Laboratory Director]"),
]

APPROVAL_STATEMENT = (
    "This document is the property of " + META["lab_name"] +
    ". It is a controlled document and becomes effective on the date approved above. "
    "It must not be modified, copied, or distributed without written authorization of the "
    "Laboratory Director. Uncontrolled when printed unless stamped \u201cCONTROLLED COPY\u201d in red. "
    "Control of documents and records is maintained in accordance with ISO 15189:2022, "
    "Clauses 8.3 and 8.4."
)

REVISION_HEADERS = ["Version", "Date", "Description of change", "Prepared", "Approved"]
REVISION_ROWS = [
    ["1.0", "14 Nov 2026", "First issue \u2014 initial release", "[Name]", "[Name]"],
    ["", "", "", "", ""],
    ["", "", "", "", ""],
]

DISTRIBUTION = [
    "Master copy (signed): Quality Office \u2014 document control register.",
    "Controlled copies: Phlebotomy station; Sample reception; Laboratory Manager\u2019s office; "
    "each collection point.",
    "A read-only copy is published on the laboratory intranet. Printed copies are "
    "uncontrolled unless stamped \u201cCONTROLLED COPY\u201d in red and listed in the distribution register "
    "(ISO 15189:2022, 8.3).",
]

# Body blocks: ('p', text) | ('b', [items]) | ('h2', title) |
#              ('steps', [(label, text), ...]) | ('tbl', headers, rows, widths)
BODY = [
    ("1. Purpose", [
        ("p", "This Standard Operating Procedure (SOP) defines the standard method for collecting "
              "venous blood samples by peripheral venipuncture at " + META["lab_name"] + ". Following "
              "this procedure protects patient safety, guarantees positive patient identification and "
              "sample integrity, and supports reliable examination results in accordance with "
              "ISO 15189:2022, in particular the pre-examination process requirements of Clause 7.2 "
              "and the personnel competence requirements of Clause 6.2."),
    ]),
    ("2. Scope", [
        ("p", "This SOP applies to all personnel (employees, students and contractors) who collect "
              "venous blood samples at the main laboratory and at every collection point operated by "
              "the laboratory, and to all patient populations served."),
        ("p", "Out of scope \u2014 covered by separate procedures:"),
        ("b", [
            "arterial blood collection (SOP-PRE-004);",
            "capillary blood sampling (SOP-PRE-005);",
            "blood culture collection (SOP-PRE-002);",
            "blood donation and donor phlebotomy (SOP-BB-001).",
        ]),
    ]),
    ("3. Definitions and Abbreviations", [
        ("tbl",
         ["Term", "Definition"],
         [
             ["Positive patient identification",
              "Verification of patient identity using at least two unique identifiers (full name and "
              "date of birth or national identification number) matched to the request."],
             ["Order of draw",
              "The standard sequence in which blood collection tubes are filled during a single "
              "venipuncture, preventing additive cross-contamination between tubes."],
             ["Haemolysis",
              "Rupture of red blood cells releasing intracellular constituents into plasma or serum; "
              "may invalidate results and is a cause for sample rejection."],
             ["Mislabelled sample",
              "A sample whose label does not correctly and uniquely identify the patient from whom "
              "it was collected; always treated as nonconforming work (ISO 15189:2022, 7.5)."],
             ["EDTA / LIS / PPE / TAT / NC / IQC",
              "Ethylenediaminetetraacetic acid (tube anticoagulant) / laboratory information system / "
              "personal protective equipment / turnaround time / nonconforming (work) / internal "
              "quality control."],
         ],
         [4.5, 12.5]),
    ]),
    ("4. Responsibilities", [
        ("tbl",
         ["Role", "Responsibility"],
         [
             ["Phlebotomist / sample collector",
              "Perform venipuncture exactly as written in this SOP; identify every patient with two "
              "identifiers; label samples at the point of collection; report and record incidents and "
              "nonconformities."],
             ["Reception / clerical staff",
              "Verify completeness of the request form; register samples in the LIS; protect patient "
              "confidentiality (ISO 15189:2022, 4.2)."],
             ["Laboratory Manager",
              "Ensure sufficient trained and authorised personnel; formally authorise staff to perform "
              "venipuncture after competence assessment (6.2); review quality indicators."],
             ["Quality Officer",
              "Maintain this SOP under document control (8.3); monitor sample rejection and haemolysis "
              "indicators; coordinate corrective actions (8.7); schedule periodic review."],
             ["Health & Safety Officer",
              "Maintain sharps disposal infrastructure and PPE availability; investigate exposure "
              "incidents; oversee hepatitis B vaccination of collectors."],
         ],
         [4.5, 12.5]),
    ]),
    ("5. Safety and Precautions", [
        ("b", [
            "Apply standard precautions for every patient regardless of diagnosis: hand hygiene before "
            "and after every collection, gloves for every procedure, gloves changed between patients.",
            "Wear a protective gown and eye protection when splashing of blood is anticipated.",
            "Never recap needles. Activate the needle safety device immediately after withdrawal and "
            "dispose of the complete assembly in a puncture-resistant sharps container at the point of use.",
            "Hepatitis B vaccination is mandatory for all personnel who perform venipuncture.",
            "Needlestick or mucous-membrane exposure: wash or irrigate immediately, report at once to "
            "the Health & Safety Officer and follow SOP-SAF-003. Record the event; raise a "
            "nonconformity if procedural failure contributed (ISO 15189:2022, 7.5).",
            "Keep all patient information confidential at all times (ISO 15189:2022, 4.2).",
        ]),
    ]),
    ("6. Materials and Equipment", [
        ("b", [
            "Disposable vacuum blood collection tubes with approved additives: serum (plain red or "
            "gold top), sodium citrate 3.2% (light blue), EDTA (lavender), lithium heparin (green), "
            "fluoride/oxalate (grey).",
            "Safety-engineered needles 21G or 22G; butterfly sets for difficult veins and for "
            "paediatric patients.",
            "70% isopropyl alcohol swabs, dry sterile gauze, adhesive plaster.",
            "Disposable single-patient tourniquets.",
            "Pre-printed barcode labels and/or LIS request forms; ballpoint pen for manual labels.",
            "Biohazard transport bags, rigid transport containers and cooled carriers where required.",
            "Puncture-resistant sharps containers and clinical waste bags.",
        ]),
    ]),
    ("7. Procedure", [
        ("h2", "7.1 Patient preparation and identification"),
        ("b", [
            "Confirm the request: tests requested, patient details, fasting status and any special "
            "timing requirements.",
            "Identify the patient using two unique identifiers \u2014 full name plus date of birth or "
            "national identification number \u2014 matched against the request form and, where available, "
            "the patient\u2019s wristband. Ask inpatients to state their own name.",
            "Explain the procedure, confirm consent and position the patient comfortably, seated or "
            "lying down. Never leave a patient unattended in a collection chair.",
            "Perform hand hygiene and put on gloves.",
        ]),
        ("h2", "7.2 Order of draw"),
        ("p", "When more than one tube is collected during a single venipuncture, fill the tubes in "
              "the sequence below (consistent with CLSI GP41). Invert each filled tube gently 8 to 10 "
              "times; never shake. Fill citrate and EDTA tubes to the marked fill volume."),
        ("tbl",
         ["Order", "Tube cap colour / additive", "Typical examinations"],
         [
             ["1", "Blood culture bottles (yellow)", "Blood cultures"],
             ["2", "Light blue \u2014 sodium citrate 3.2%", "Coagulation (PT, aPTT, INR)"],
             ["3", "Red / gold \u2014 plain or clot activator (with gel)", "Clinical chemistry, serology, immunology"],
             ["4", "Green \u2014 lithium heparin", "Plasma chemistry, emergency tests"],
             ["5", "Lavender \u2014 EDTA", "Haematology (CBC), HbA1c"],
             ["6", "Grey \u2014 fluoride / oxalate", "Glucose, lactate"],
         ],
         [1.8, 7.2, 8.0]),
        ("h2", "7.3 Venipuncture"),
        ("steps", [
            ("1.", "Assemble all required materials within reach before approaching the patient."),
            ("2.", "Select the site: prefer the median cubital vein of the forearm. Avoid the arm on "
                   "the side of a mastectomy, arteriovenous fistula, active infusion, burn or scar."),
            ("3.", "Apply the tourniquet 3 to 4 cm above the intended puncture site, tight enough to "
                   "impede venous return but not arterial flow. Do not leave the tourniquet on for "
                   "more than 1 minute."),
            ("4.", "Palpate the vein with the index finger. Clean the site with a 70% isopropyl "
                   "alcohol swab using a circular motion from the centre outwards over at least 2 cm "
                   "and allow to air-dry for at least 30 seconds. Do not touch, blow on or fan the site."),
            ("5.", "Anchor the vein with the thumb of the non-dominant hand, holding the skin taut "
                   "below the puncture point."),
            ("6.", "Insert the needle with the bevel upwards at an angle of 15 to 30 degrees, in a "
                   "single smooth motion, in the direction of the vein."),
            ("7.", "When blood flows, fill the tubes in the order of draw given in 7.2. Under-filling "
                   "or over-filling citrate and EDTA tubes is a cause for rejection."),
            ("8.", "Release the tourniquet before withdrawing the needle. Withdraw the needle and "
                   "apply dry gauze with firm pressure for 2 to 3 minutes; apply plaster if needed."),
            ("9.", "Activate the needle safety device immediately and drop the entire unit into the "
                   "sharps container at the point of use \u2014 never carry an unsheathed needle and never recap."),
            ("10.", "Keep the patient\u2019s arm extended and observe for haematoma or dizziness before "
                    "the patient leaves the collection area."),
            ("11.", "If blood does not flow or flows slowly: do not probe more than twice. Release the "
                    "tourniquet, withdraw, apply pressure, and use a second site or the other arm. "
                    "Record every difficult draw."),
        ]),
        ("h2", "7.4 Labelling and documentation"),
        ("b", [
            "Label every tube at the bedside, at the point of collection, immediately after filling. "
            "Never pre-label tubes and never label tubes away from the patient.",
            "Minimum label data: patient full name; unique identifier (date of birth or national ID / "
            "LIS number); collection date and exact time; collector\u2019s initials.",
            "Verify that every label matches the request form before leaving the patient. Place the "
            "label lengthwise on the tube without covering the cap.",
            "For manual labels, write legibly with a ballpoint pen. Correct errors with a single line "
            "through the error, the correct data, an initial and the date (ISO 15189:2022, 8.4).",
        ]),
        ("h2", "7.5 Transport to the laboratory"),
        ("b", [
            "Place tubes upright in a biohazard bag; keep the request form from direct contact with "
            "the tubes.",
            "Deliver to the laboratory within 2 hours of collection, at 18 to 25 degrees C unless "
            "otherwise specified (for example cooled carriers at 2 to 8 degrees C for glucose and "
            "lactate).",
            "Do not use the pneumatic tube system for coagulation samples or blood culture bottles "
            "unless the system has been validated for them.",
        ]),
        ("h2", "7.6 Sample acceptance and rejection criteria"),
        ("p", "Samples are checked at reception against the criteria below. Rejected samples are "
              "recorded in the Sample Rejection Register and, where the cause is a systemic failure, "
              "handled as nonconforming work (ISO 15189:2022, 7.5)."),
        ("tbl",
         ["Observation", "Action"],
         [
             ["Patient identification missing, incomplete, or not matching the request form",
              "Reject; notify the requester; arrange recollection; record as mislabelling NC"],
             ["Visible haemolysis in plasma/serum, or analyser haemolysis flag",
              "Reject the affected test(s); notify; recollect"],
             ["Incorrect tube or additive for the test requested",
              "Reject the affected test(s); recollect"],
             ["Insufficient volume (especially citrate and EDTA tubes)",
              "Reject the affected test(s); recollect"],
             ["Clots present in an EDTA tube (haematology)",
              "Reject; recollect immediately"],
             ["Transport time or temperature outside defined limits",
              "Assess analyte stability; if stability is not demonstrated, reject and recollect"],
             ["Damaged or leaking container",
              "Reject; recollect; decontaminate the transport container"],
         ],
         [7.0, 10.0]),
    ]),
    ("8. Quality Control and Monitoring", [
        ("b", [
            "Quality indicators monitored monthly by the Quality Officer: sample rejection rate "
            "(target at most 2% of all samples); haemolysis rate (target at most 1%); mislabelled "
            "samples (target zero); samples delivered within the defined turnaround time.",
            "Indicator results are reported at the monthly quality meeting and reviewed in management "
            "review (ISO 15189:2022, 8.8 and 8.9).",
            "Where available, the laboratory participates in external quality assessment schemes for "
            "pre-analytical phases and acts on unsatisfactory performance (8.7).",
            "Phlebotomy technique is observed directly at least annually for each collector as part of "
            "competence assessment (6.2).",
        ]),
    ]),
    ("9. Risk Considerations", [
        ("p", "Risks relevant to this procedure are recorded below in line with the risk-based approach "
              "of ISO 15189:2022 (Clauses 5.6 and 8.5). Residual risks are reviewed at each periodic "
              "review of this SOP."),
        ("tbl",
         ["Risk", "Control measure"],
         [
             ["Wrong-patient identification",
              "Two identifiers confirmed at the bedside; labelling at the point of collection; "
              "labels verified against the request before leaving the patient"],
             ["Needlestick injury",
              "Safety-engineered devices; no recapping; sharps container at the point of use; "
              "mandatory hepatitis B vaccination; SOP-SAF-003 for exposure management"],
             ["Haemolysed or clotted samples",
              "Correct tube fill volumes; gentle inversion; order of draw; avoid prolonged "
              "tourniquet time; prompt transport"],
             ["Delayed transport affecting analyte stability",
              "Two-hour delivery rule; scheduled courier runs; temperature-controlled carriers"],
         ],
         [6.0, 11.0]),
    ]),
    ("10. Documentation and Records", [
        ("b", [
            "Records generated by this procedure: blood collection register / LIS collection log; "
            "sample rejection register (F-PRE-002); exposure incident reports (F-SAF-001); the "
            "training acknowledgement in Annex A (F-TRN-001).",
            "Records are created at the time the activity is performed, and remain legible, traceable "
            "to the collector and retrievable (ISO 15189:2022, 8.4).",
            "Corrections: single line through the error, the correct value, an initial and the date. "
            "The original entry must remain legible.",
            "Retention: at least 5 years, or as required by national regulation where longer; disposal "
            "by approved confidential destruction.",
        ]),
    ]),
    ("11. References", [
        ("b", [
            "ISO 15189:2022 \u2014 Medical laboratories: requirements for quality and competence "
            "(in particular Clauses 4.2, 4.3, 5.6, 6.2, 7.2, 7.5, 8.3, 8.4, 8.5).",
            "CLSI GP41 \u2014 Collection of Diagnostic Venous Blood Specimens.",
            "CLSI H3 \u2014 Procedures for the Handling and Processing of Blood Specimens.",
            "WHO \u2014 Laboratory Quality Management System (LQMS) Handbook.",
            "[National phlebotomy, sharps-safety and infection-prevention guidelines \u2014 adapt to "
            "the applicable national authority].",
        ]),
    ]),
    ("12. Related Documents and Forms", [
        ("tbl",
         ["Reference", "Title"],
         [
             ["SOP-PRE-002", "Sample Reception, Accessioning and Rejection"],
             ["SOP-SAF-003", "Occupational Exposure and Post-Exposure Management"],
             ["SOP-QM-004", "Control of Documents and Records"],
             ["F-PRE-001", "Blood Collection Register"],
             ["F-PRE-002", "Sample Rejection Register / Nonconforming Work Form"],
             ["F-TRN-001", "Training Acknowledgement and Competence Record (Annex A)"],
         ],
         [3.5, 13.5]),
    ]),
]

TRAINING_TITLE = "Annex A \u2014 Training and Competence Acknowledgement (Form F-TRN-001)"
TRAINING_STATEMENT = (
    "I confirm that I have received training on this Standard Operating Procedure, that I have read "
    "and understood its contents including the safety precautions, and that I agree to perform the "
    "procedure as written. I understand that I must not perform this procedure until this form has "
    "been signed and filed in my training record. Training, competence assessment and authorisation "
    "of personnel are maintained in accordance with ISO 15189:2022, Clause 6.2; the completed form "
    "is retained as a record in accordance with Clause 8.4."
)
TRAINING_HEADERS = ["S/No", "Full Name (print)", "Staff / Employee ID", "Designation", "Signature", "Date"]
TRAINING_ROWS = [[f"{i}", "", "", "", "", ""] for i in range(1, 16)]
TRAINING_WIDTHS = [1.2, 5.0, 2.6, 3.2, 3.0, 2.0]
TRAINING_FOOTER = (
    "Trainer / Supervisor: ______________________    Name (print): ______________    Date: ____________"
)
TRAINING_NOTE = (
    "Complete one row per staff member (15 rows provided \u2014 at least 12 anticipated). "
    "If more rows are needed, attach continuation sheet F-TRN-001-2 and staple it to this page. "
    "File the completed form in the personnel training record; reference it in the SOP training "
    "register maintained by the Quality Officer."
)

FOOTER_LEFT = (META["doc_no"] + "  |  Version " + META["version"] +
               "  |  Uncontrolled when printed unless stamped \u201cCONTROLLED COPY\u201d in red")
HEADER_LEFT = META["doc_no"] + " \u2014 " + META["title"]
HEADER_RIGHT = "Version " + META["version"] + "  |  Effective " + META["effective"]

INK = "#1F4E79"      # headings
INK2 = "#2E74B5"     # subheadings
GREY = "#595959"     # running text

# ---------------------------------------------------------------------------
# DOCX builder
# ---------------------------------------------------------------------------

def build_docx():
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()

    # -- page setup ---------------------------------------------------------
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.header_distance = Cm(0.9)
    sec.footer_distance = Cm(0.9)
    sec.different_first_page_header_footer = True

    # -- base styles --------------------------------------------------------
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.06

    for name, size, color, before, after in (
        ("Heading 1", 13.5, INK, 10, 5),
        ("Heading 2", 11.5, INK2, 8, 3),
        ("Title", 20, INK, 0, 4),
    ):
        st = doc.styles[name]
        st.font.name = "Calibri"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(color.lstrip("#"))
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)

    # -- helpers ------------------------------------------------------------
    def shade(cell, hexfill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), hexfill.lstrip("#"))
        tcPr.append(shd)

    def fixed_layout(table):
        tblPr = table._tbl.tblPr
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        tblPr.append(layout)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def set_cell(cell, text, *, bold=False, size=9.5, color=None, align=None):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        if align is not None:
            p.alignment = align
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))

    def make_table(headers, rows, widths_cm, *, header_fill=INK, body_size=9.5,
                   row_height=None, first_col_bold=False):
        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = "Table Grid"
        fixed_layout(table)
        for j, w in enumerate(widths_cm):
            for row in table.rows:
                row.cells[j].width = Cm(w)
        for j, h in enumerate(headers):
            set_cell(table.rows[0].cells[j], h, bold=True, size=9.5,
                     color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
            shade(table.rows[0].cells[j], header_fill)
        for i, rowvals in enumerate(rows, start=1):
            for j, val in enumerate(rowvals):
                set_cell(table.rows[i].cells[j], val, size=body_size,
                         bold=(first_col_bold and j == 0))
            if row_height:
                table.rows[i].height = Cm(row_height)
        return table

    def add_field(paragraph, instr, placeholder="1", size=8):
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), instr)
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rPr.append(sz)
        r.append(rPr)
        t = OxmlElement("w:t")
        t.text = placeholder
        r.append(t)
        fld.append(r)
        paragraph._p.append(fld)

    def tab_stop(paragraph, pos_cm):
        pPr = paragraph._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:pos"), str(int(pos_cm * 567)))  # twips (1 cm = 567 twips)
        tabs.append(tab)
        pPr.append(tabs)

    def bottom_border(paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "BFBFBF")
        pBdr.append(bottom)
        pPr.append(pBdr)

    def fill_footer(footer):
        p = footer.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(0)
        top = p.add_run(FOOTER_LEFT)
        top.font.size = Pt(7.5)
        top.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
        p.add_run("\t")
        pg = p.add_run("Page ")
        pg.font.size = Pt(7.5)
        pg.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
        add_field(p, "PAGE", size=7.5)
        mid = p.add_run(" of ")
        mid.font.size = Pt(7.5)
        mid.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
        add_field(p, "NUMPAGES", placeholder="1", size=7.5)
        tab_stop(p, 17.0)
        bottom_border(p)

    # -- running header (all pages except the first) ------------------------
    hp = sec.header.paragraphs[0]
    hr = hp.add_run(HEADER_LEFT)
    hr.font.size = Pt(8)
    hr.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
    hp.add_run("\t")
    hr2 = hp.add_run(HEADER_RIGHT)
    hr2.font.size = Pt(8)
    hr2.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
    tab_stop(hp, 17.0)
    bottom_border(hp)

    fill_footer(sec.footer)
    fill_footer(sec.first_page_footer)
    # first page: no header text
    fph = sec.first_page_header.paragraphs[0]
    fph.text = ""

    # -- page 1: cover ------------------------------------------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(META["logo"])
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
    r.font.italic = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(META["lab_name"])
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("STANDARD OPERATING PROCEDURE")
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(META["title"])
    r.font.size = Pt(19)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    make_table(["Field", "Detail"], [[k, v] for k, v in DOC_CONTROL_ROWS], [5.0, 12.0],
               body_size=9.5, first_col_bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(110)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("Approval and Authorisation")
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    ap = make_table(["Role", "Name (print)", "Signature", "Date"],
                    [[role, name, "", ""] for role, name in APPROVAL_ROWS],
                    [4.2, 5.3, 4.5, 3.0], body_size=9.5, first_col_bold=True,
                    row_height=1.05)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run(APPROVAL_STATEMENT)
    r.font.size = Pt(8.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    doc.add_page_break()

    # -- page 2: revision history + TOC ------------------------------------
    doc.add_heading("Revision History", level=1)
    make_table(REVISION_HEADERS, REVISION_ROWS, [2.0, 2.6, 7.4, 2.5, 2.5], body_size=9.5)

    doc.add_heading("Distribution and Control of Copies", level=1)
    for item in DISTRIBUTION:
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.6)
        bp.paragraph_format.first_line_indent = Cm(-0.4)
        bp.paragraph_format.space_after = Pt(2)
        r = bp.add_run("\u2022  " + item)
        r.font.size = Pt(9.5)

    tp_head = doc.add_paragraph()
    tp_head.paragraph_format.space_before = Pt(10)
    tp_head.paragraph_format.space_after = Pt(5)
    r = tp_head.add_run("Table of Contents")
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))
    tp = doc.add_paragraph()
    add_field(tp, 'TOC \\o "1-1" \\h \\z \\u',
              placeholder="Right-click here and choose \u201cUpdate Field\u201d to build the "
                          "table of contents with page numbers.", size=9.5)

    doc.add_page_break()

    # -- body pages (flow continuously) ------------------------------------
    for title, blocks in BODY:
        doc.add_heading(title, level=1)
        for block in blocks:
            kind = block[0]
            if kind == "p":
                bp = doc.add_paragraph(block[1])
                bp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif kind == "b":
                for item in block[1]:
                    bp = doc.add_paragraph()
                    bp.paragraph_format.left_indent = Cm(0.6)
                    bp.paragraph_format.first_line_indent = Cm(-0.4)
                    bp.paragraph_format.space_after = Pt(2)
                    bp.add_run("\u2022  " + item)
            elif kind == "h2":
                doc.add_heading(block[1], level=2)
            elif kind == "steps":
                for label, text in block[1]:
                    bp = doc.add_paragraph()
                    bp.paragraph_format.left_indent = Cm(0.85)
                    bp.paragraph_format.first_line_indent = Cm(-0.85)
                    bp.paragraph_format.space_after = Pt(2)
                    r = bp.add_run(label + "  ")
                    r.font.bold = True
                    bp.add_run(text)
            elif kind == "tbl":
                _, headers, rows, widths = block
                make_table(headers, rows, widths, body_size=9.0)
                doc.add_paragraph().paragraph_format.space_after = Pt(0)

    doc.add_page_break()

    # -- Annex A: training & signature sheet -------------------------------
    doc.add_heading(TRAINING_TITLE, level=1)
    sp = doc.add_paragraph(TRAINING_STATEMENT)
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    make_table(TRAINING_HEADERS, TRAINING_ROWS, TRAINING_WIDTHS,
               body_size=9.5, row_height=0.82)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    r = p.add_run(TRAINING_FOOTER)
    r.font.size = Pt(9.5)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(TRAINING_NOTE)
    r.font.size = Pt(8.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    # -- update fields (TOC, page numbers) on open -------------------------
    settings = doc.settings.element
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    settings.append(upd)

    DOCX_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc.save(DOCX_PATH)
    print(f"DOCX written: {DOCX_PATH.relative_to(ROOT)}")


# ---------------------------------------------------------------------------
# PDF builder
# ---------------------------------------------------------------------------

def build_pdf():
    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate,
                                    Paragraph, Spacer, Table, TableStyle)
    from reportlab.platypus.tableofcontents import TableOfContents

    ink = colors.HexColor(INK)
    ink2 = colors.HexColor(INK2)
    grey = colors.HexColor(GREY)
    line_grey = colors.HexColor("#BFBFBF")
    head_fill = colors.HexColor(INK)

    styles = {
        "logo": ParagraphStyle("logo", fontName="Helvetica-Oblique", fontSize=10,
                               textColor=grey, alignment=1, spaceAfter=2),
        "lab": ParagraphStyle("lab", fontName="Helvetica-Bold", fontSize=15,
                              textColor=ink, alignment=1, spaceAfter=4),
        "kind": ParagraphStyle("kind", fontName="Helvetica-Bold", fontSize=11.5,
                               textColor=grey, alignment=1, spaceAfter=2),
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=18,
                                textColor=ink, alignment=1, spaceAfter=12),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=12,
                             textColor=ink, spaceBefore=10, spaceAfter=5),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10.5,
                             textColor=ink2, spaceBefore=7, spaceAfter=3),
        "p": ParagraphStyle("p", fontName="Helvetica", fontSize=9.5, leading=12.4,
                            textColor=colors.black, spaceAfter=4,
                            alignment=4),
        "b": ParagraphStyle("b", fontName="Helvetica", fontSize=9.5, leading=12.2,
                            textColor=colors.black, spaceAfter=2.5,
                            leftIndent=14, firstLineIndent=-10),
        "s": ParagraphStyle("s", fontName="Helvetica", fontSize=9.5, leading=12.2,
                            textColor=colors.black, spaceAfter=2.5,
                            leftIndent=20, firstLineIndent=-20),
        "cell": ParagraphStyle("cell", fontName="Helvetica", fontSize=8.5,
                               leading=10.4, textColor=colors.black),
        "cellb": ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=8.5,
                                leading=10.4, textColor=colors.black),
        "cellhead": ParagraphStyle("cellhead", fontName="Helvetica-Bold", fontSize=8.5,
                                   leading=10.4, textColor=colors.white),
        "small": ParagraphStyle("small", fontName="Helvetica-Oblique", fontSize=8,
                                textColor=grey, leading=10),
        "smallplain": ParagraphStyle("smallplain", fontName="Helvetica", fontSize=8.5,
                                     textColor=grey, leading=11),
        "coverlabel": ParagraphStyle("coverlabel", fontName="Helvetica-Bold", fontSize=8.5,
                                     leading=10.5, textColor=colors.black),
        "coverval": ParagraphStyle("coverval", fontName="Helvetica", fontSize=8.5,
                                   leading=10.5, textColor=colors.black),
    }

    PAGE_W, PAGE_H = A4
    FRAME_W = 17.0 * cm

    class NumberedCanvas(pdfcanvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self._draw_footer(total)
                super().showPage()
            super().save()

        def _draw_footer(self, total):
            self.saveState()
            self.setStrokeColor(line_grey)
            self.setLineWidth(0.5)
            self.line(2.0 * cm, 1.65 * cm, PAGE_W - 2.0 * cm, 1.65 * cm)
            self.setFont("Helvetica", 7)
            self.setFillColor(grey)
            self.drawString(2.0 * cm, 1.30 * cm, FOOTER_LEFT)
            self.drawRightString(PAGE_W - 2.0 * cm, 1.30 * cm,
                                 f"Page {self._pageNumber} of {total}")
            self.restoreState()

    def on_page(canv, doc_):
        if doc_.page <= 1:
            return
        canv.saveState()
        canv.setFont("Helvetica", 7.5)
        canv.setFillColor(grey)
        canv.drawString(2.0 * cm, PAGE_H - 1.35 * cm, HEADER_LEFT)
        canv.drawRightString(PAGE_W - 2.0 * cm, PAGE_H - 1.35 * cm, HEADER_RIGHT)
        canv.setStrokeColor(line_grey)
        canv.setLineWidth(0.5)
        canv.line(2.0 * cm, PAGE_H - 1.55 * cm, PAGE_W - 2.0 * cm, PAGE_H - 1.55 * cm)
        canv.restoreState()

    class SOPDoc(BaseDocTemplate):
        def afterFlowable(self, flowable):
            entry = getattr(flowable, "toc_entry", None)
            if entry is not None:
                self.notify("TOCEntry", (0, entry, self.page))

    def P(text, style="p"):
        return Paragraph(escape(text), styles[style])

    def H1(text):
        para = Paragraph(escape(text), styles["h1"])
        para.toc_entry = text
        return para

    def make_pdf_table(headers, rows, widths_cm, *, first_col_bold=False,
                       repeat_header=True, row_heights=None):
        data = [[Paragraph(escape(str(h)), styles["cellhead"]) for h in headers]]
        for row in rows:
            cells = []
            for j, val in enumerate(row):
                st = "cellb" if (first_col_bold and j == 0) else "cell"
                cells.append(Paragraph(escape(str(val)), styles[st]))
            data.append(cells)
        t = Table(data, colWidths=[w * cm for w in widths_cm],
                  rowHeights=row_heights, repeatRows=1 if repeat_header else 0)
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), head_fill),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, line_grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]
        t.setStyle(TableStyle(style_cmds))
        return t

    story = []

    # -- cover --------------------------------------------------------------
    story += [
        Spacer(1, 0.6 * cm),
        P(META["logo"], "logo"),
        P(META["lab_name"], "lab"),
        P("STANDARD OPERATING PROCEDURE", "kind"),
        P(META["title"], "title"),
        make_pdf_table(["Field", "Detail"],
                       [[k, v] for k, v in DOC_CONTROL_ROWS],
                       [5.0, 12.0], first_col_bold=True),
        Spacer(1, 1.9 * cm),
        P("Approval and Authorisation", "h2"),
        make_pdf_table(["Role", "Name (print)", "Signature", "Date"],
                       [[role, name, "", ""] for role, name in APPROVAL_ROWS],
                       [4.2, 5.3, 4.5, 3.0], first_col_bold=True,
                       row_heights=[None] + [0.95 * cm] * len(APPROVAL_ROWS)),
        Spacer(1, 0.35 * cm),
        P(APPROVAL_STATEMENT, "small"),
        PageBreak(),
    ]

    # -- revision history + TOC --------------------------------------------
    story.append(H1("Revision History"))
    story.append(make_pdf_table(REVISION_HEADERS, REVISION_ROWS,
                                [2.0, 2.6, 7.4, 2.5, 2.5],
                                row_heights=[None] + [0.55 * cm] * len(REVISION_ROWS)))
    story.append(H1("Distribution and Control of Copies"))
    for item in DISTRIBUTION:
        story.append(P("\u2022  " + item, "b"))
    story.append(Paragraph(escape("Table of Contents"), styles["h1"]))  # not itself a TOC entry
    toc = TableOfContents()
    toc.levelStyles = [ParagraphStyle("toc0", fontName="Helvetica", fontSize=9.5,
                                      leading=13.5, textColor=colors.black)]
    story.append(toc)
    story.append(PageBreak())

    # -- body ---------------------------------------------------------------
    for title, blocks in BODY:
        story.append(H1(title))
        for block in blocks:
            kind = block[0]
            if kind == "p":
                story.append(P(block[1]))
            elif kind == "b":
                for item in block[1]:
                    story.append(P("\u2022  " + item, "b"))
            elif kind == "h2":
                story.append(Paragraph(escape(block[1]), styles["h2"]))
            elif kind == "steps":
                for label, text in block[1]:
                    story.append(P(label + "   " + text, "s"))
            elif kind == "tbl":
                _, headers, rows, widths = block
                story.append(make_pdf_table(headers, rows, widths))
                story.append(Spacer(1, 0.15 * cm))

    story.append(PageBreak())

    # -- Annex A ------------------------------------------------------------
    story.append(H1(TRAINING_TITLE))
    story.append(P(TRAINING_STATEMENT))
    story.append(Spacer(1, 0.15 * cm))
    story.append(make_pdf_table(TRAINING_HEADERS, TRAINING_ROWS, TRAINING_WIDTHS,
                                row_heights=[None] + [0.82 * cm] * len(TRAINING_ROWS)))
    story.append(Spacer(1, 0.45 * cm))
    story.append(P(TRAINING_FOOTER, "p"))
    story.append(Spacer(1, 0.15 * cm))
    story.append(P(TRAINING_NOTE, "small"))

    PDF_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc = SOPDoc(str(PDF_PATH), pagesize=A4,
                 leftMargin=2.0 * cm, rightMargin=2.0 * cm,
                 topMargin=2.0 * cm, bottomMargin=2.0 * cm,
                 title=META["doc_no"] + " \u2014 " + META["title"],
                 author=META["lab_name"],
                 subject="ISO 15189:2022-compliant SOP sample")
    frame = Frame(doc.leftMargin, doc.bottomMargin, FRAME_W, PAGE_H - 4.0 * cm,
                  id="main")
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])
    doc.multiBuild(story, canvasmaker=NumberedCanvas)
    print(f"PDF written:  {PDF_PATH.relative_to(ROOT)}  ({doc.page} pages)")


if __name__ == "__main__":
    build_docx()
    build_pdf()
