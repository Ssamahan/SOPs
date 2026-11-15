#!/usr/bin/env python3
"""
Build Ruli District Hospital Laboratory SOP documents (DOCX + PDF) from one
content model, in the 2026 standard layout mapped to ISO 15189:2022.

Documents produced (names come from REGISTER, not from this docstring):
  Venous Blood Sample Collection  — IH series, first free serial
  Glucose Testing using Cobas c 111 — BIO-SOP-001, version 005

Rules applied:
  * Every document code is taken from REGISTER below, never typed in the text:
    RliDHL-<AREA>-<TYPE>-<NNN>-VERS-<NNN>. Serials run with NO GAPS within a
    series; a new document takes the next free serial; a form carries its
    parent SOP's code with an -F<NN> suffix.
  * scripts/check_codes.py proves the rules hold after every build.
  * Dates are filled in (DD/MM/YYYY).
  * Approval block: designation below the name; rows tall enough for the
    Director General's hospital stamp.
  * Explanatory prose appears only where ISO 15189:2022 requires the content.
  * Rwanda district-hospital roles: Environmental Health Officer (EHO) for
    health & safety; no "specialist" posts.

Usage:  .venv/bin/python scripts/build_sop.py
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = ROOT / "assets" / "Ruli-District-Hospital-logo.png"
SOP_DIR = ROOT / "sop"

INK = "#1F4E79"
INK2 = "#2E74B5"
GREY = "#595959"

# ---------------------------------------------------------------------------
# Shared content
# ---------------------------------------------------------------------------

LAB = "Ruli District Hospital Laboratory Department"

# Official letterhead (from the Laboratory Normal Ranges / Critical Results / TAT lists)
LETTERHEAD = [
    ("REPUBLIC OF RWANDA · MINISTRY OF HEALTH · GAKENKE DISTRICT", "kind"),
    ("Ruli District Hospital", "lab"),
    ("Laboratory Department", "kind"),
    ("P.O. Box 1285 Kigali · ruli.hospital@moh.gov.rw · (+250) 0781 316 368 · rulihospital.rw", "small"),
]

# (role, name, designation) — designation is printed below the name
APPROVAL_ROWS = [
    ("Prepared by (Author)", "Jean Paul NSENGIYUMVA", "Lab Quality Manager"),
    ("Reviewed by", "Kelvin SAMAHANI", "Laboratory Manager"),
    ("Authorized by", "Dr. Déogratias KANEZA", "Hospital’s Director General"),
]

REVISION_HEADERS = ["Prepared/Reviewed by", "Version", "Date", "Description of Change"]

# ---------------------------------------------------------------------------
# DOCUMENT REGISTER — the single source of truth for every document code
# ---------------------------------------------------------------------------
# Numbering rule for Ruli District Hospital Laboratory (ISO 15189:2022, 8.3):
#
#   RliDHL-<AREA>-<TYPE>-<NNN>-VERS-<NNN>
#
#   AREA  3-letter section: ADM BAC BIO HEM IH IMM MOL MYC PAR REC SER STC SAF
#   TYPE  MAN (manual) GUID (guideline) SOP
#   NNN   sequential serial, 3 digits, runs 001, 002, 003 ... with NO GAPS within
#         one AREA+TYPE series. A number is allocated once and belongs to that
#         document for life: it never changes at revision, and a number released
#         by a withdrawn document is not re-used.
#   VERS  version, 3 digits, 001 at first issue.
#
# Forms take their parent SOP's code and an -FNN suffix, so a form can never
# drift away from the SOP that controls it:
#
#   RliDHL-ADM-SOP-042-F01
#
# Every code used anywhere in the generated documents is derived from this
# register — never typed into the document text. Change a code here and every
# reference, the cross-reference lists and the file names follow.
#
# Serials below are the ones actually allocated, read from MASTER LIST 2025.docx
# (Drive 1cl3zHlZ2-SDW-1gUrrwrYVeRA7-YwPuI) and from SOP-PROGRAMME-HANDOFF.md §2.
#
# The ADM series is 001-049 in the register, NOT 001-052. The November 2026
# equipment block then took 050, 051, 052 and 060-063 — 060 onwards because files
# exist above 52 that were never registered, so 053-059 cannot be proven free.
# The next free ADM serial is therefore 064.
#
# "new": True marks a serial this library is allocating for the first time. Such a
# serial must not already be in SERIES_ALLOCATED. An entry without the flag is
# pre-existing and is cited, not allocated.
SERIES_KNOWN_MAX = {
    # Highest SOP serial registered per series, read line by line from
    # MASTER LIST 2025.docx on 15/11/2026. These are observed maxima, not
    # permission to allocate -- always re-read the register first.
    #
    #   ADM 001-049, with 016/017 unreadable and 035 never used
    #   BIO 001-009   HEM 001-004   IMM 001-004
    #
    # NOTE for anyone tempted to use this list as an inventory: several of those
    # serials govern equipment that has since been decommissioned. See WITHDRAWN.
    # A serial existing in the register says nothing about the machine existing in
    # the laboratory.
    #
    # There is NO IH section in the register: the areas it uses are ADM, BAC,
    # BIO, HEM, IMM, MYC, PAR, SER and STC. "IH" was introduced by this library
    # and is not an established department code.
    "RliDHL-ADM-SOP": 49,
    "RliDHL-BAC-SOP": 6,
    "RliDHL-BIO-SOP": 9,
    "RliDHL-HEM-SOP": 4,
    "RliDHL-IMM-SOP": 4,
    "RliDHL-MYC-SOP": 12,
    "RliDHL-PAR-SOP": 10,
    "RliDHL-SER-SOP": 7,
    "RliDHL-STC-SOP": 7,
}

# Series where the register has no section at all. A number here is neither free
# nor taken -- the series would be new, which is a decision, not a gap.
SERIES_UNVERIFIED: set[str] = set()   # no invented series remain; IH is gone

# Open questions this library cannot settle, and must not paper over with a green
# check. Each needs the Laboratory Manager and the Director General. The checker
# prints them after the passes, separately, because consistent-with-itself is not
# the same as correct-against-the-register.
DECISIONS = {
    # Resolved 15/11/2026 on the Laboratory Manager's instruction. Kept here, empty of
    # open items, because the history is why the numbers are what they are:
    #
    #   venous    -- IH is gone. Venous collection is HEM-SOP-001-VERS-002, a revision
    #                of the registered "Vein blood collection..." rather than a new
    #                series. No department was invented.
    #   glucose   -- 001 is retained (the instruction: keep 001, move the other document
    #                to the number next to it). "Use and preventive maintenance of
    #                Humastar 80" moves to BIO-SOP-002. Nothing is skipped.
    #   training  -- 064-F01 CITES the existing ADM-FORM-014 signature page and
    #                ADM-FORM-001 competency form instead of replacing them.
    #
    # A new open question belongs in this dict so the checker prints it.
}

# Documents whose equipment has been DECOMMISSIONED. The document is superseded: it is
# not renumbered, not reissued and not replaced. Its serial stays with it forever --
# that is the whole point of a serial being held for life.
#
# Numbering these was a real error in this repository: the Humastar 80 and Humalyzer
# 3500 SOPs were treated as live documents and shuffled between serials, when both
# analysers have been withdrawn from the laboratory. A number existing in the register
# does not mean the machine exists at the bench.
#
# Verified against SOP-PROGRAMME-HANDOFF.md, "Decommissioned" section.
WITHDRAWN = {
    # key: (code, what it governed, why it is withdrawn)
    "adm_biomed": (
        "RliDHL-ADM-SOP-008", "Use and Maintenance of Biomedical Equipment",
        "VERS-005 is a coding defect: serial 008 also identifies the legacy "
        "Authorization document. Not an equipment withdrawal -- listed here so the "
        "008 collision stays visible."),
    "bio_001": (
        "RliDHL-BIO-SOP-001", "Use and preventive maintenance of Humastar 80",
        "Humastar 80 decommissioned. Retained as superseded; serial 001 is also held "
        "by the glucose SOP, which is the live document under that number."),
    "bio_002": (
        "RliDHL-BIO-SOP-002", "Sample analysis for Clinical Biochemistry using "
        "Humalyzer 3500",
        "Humalyzer 3500 decommissioned. Superseded; nothing replaces it."),
    "bio_003_009": (
        "RliDHL-BIO-SOP-003 to 009", "Urea, ALAT, ASAT, creatinine, glucose prep, "
        "bilirubin total and direct, all 'by use of Humastar 80'",
        "Every one of these is a Humastar 80 procedure. The analyser is gone, so the "
        "whole block is superseded and none of it should be cited as current."),
    "hem_002": (
        "RliDHL-HEM-SOP-002", "Hematology sample analysis using Sysmex Pochi 100i",
        "Pochi decommissioned. Superseded; the live analysers are the XN-330 and "
        "XQ-320."),
    "hem_003": (
        "RliDHL-HEM-SOP-003", "Hematology sample analysis using Humacount 5",
        "Humacount 5 decommissioned. Superseded."),
    "hem_004": (
        "RliDHL-HEM-SOP-004", "Hematology sample analysis using Sysmex KX-21N",
        "KX-21N decommissioned. Superseded."),
}

# Analysers actually in service, from the same verified inventory. Cited when a
# document needs to name the equipment; never claim more than this.
IN_SERVICE = [
    ("Sysmex XN-330", "17156", "RULIDH/LAB/ME/0501", "Hematology bench"),
    ("Sysmex XQ-320", "12357", "RULIDH/LAB/ME/0499", "Hematology bench"),
    ("Alere Pima CD4", "PIMA-D-005378", "RULIDH/LAB/ME/0485", "Hematology room"),
]

# Every serial already spoken for beyond the registered set, so a duplicate is
# caught before it ships. This is the list that would have caught ADM-060.
SERIES_ALLOCATED = {
    "RliDHL-ADM-SOP": {
        # the 49 registered rows (035 is the one known gap)
        **{n: "Master List 2025" for n in range(1, 50) if n != 35},
        # the November 2026 equipment block: built, uploaded, unsigned
        50: "ADM SOP 050 Microscope (Nov 2026 block)",
        51: "ADM SOP 051 Centrifuge (Nov 2026 block)",
        52: "ADM SOP 052 Water Bath (Nov 2026 block)",
        60: "ADM SOP 060 Non-Medical Equipment (Nov 2026 block)",
        61: "ADM SOP 061 ICT and LIS (Nov 2026 block)",
        62: "ADM SOP 062 Hematology Analysers (Nov 2026 block)",
        63: "ADM SOP 063 GeneXpert Dx (Nov 2026 block)",
    },
    "RliDHL-BIO-SOP": {
        1: "BIO SOP 001 Glucose Testing (this library, VERS-005)",
    },
    # BIO serial 002 is NOT allocated. Its registered holder is the Humalyzer 3500
    # SOP, and the Humalyzer has been decommissioned. The document is superseded,
    # not renumbered and not replaced -- see WITHDRAWN. Moving the Humastar SOP
    # onto 002 (an earlier attempt) would have re-issued a withdrawn SOP under
    # another analyser's number.
}
REGISTER = {
    # --- management (crosscutting) SOPs ---
    "doc_control": {
        "code": "RliDHL-ADM-SOP-004", "vers": "004",
        "title": "Document and Record Control",
        "master_list": "Master List, Crosscutting SOPs",
    },
    "pre_exam": {
        "code": "RliDHL-ADM-SOP-020", "vers": "004",
        "title": "Pre-Examination Procedure",
        "master_list": "Master List, Crosscutting SOPs",
    },
    # Written to close the gap found when auditing this set: the Master List has
    # no SOP for the SIGNING of SOPs (033 training and 002 competency assessment
    # exist, so those must be cross-referenced rather than duplicated), and no SOP
    # owns the section 21 signature page. Serial 064 is the next free ADM number
    # after the 060-063 equipment block.
    #
    # 060 was allocated here on 15/11/2026 by mistake -- 060 is the Non-Medical
    # Equipment SOP -- and moved to 064 the same day. Its form moved from
    # 060-F01 to 064-F01, which also released 060-F01 back to the temperature
    # monitoring form that owns it.
    "training": {
        "code": "RliDHL-ADM-SOP-064", "vers": "001",
        "title": "Signing of SOPs, and the Training, Competence and Authorisation Record",
        "master_list": "Master List, Crosscutting SOPs",
        "new": True,
    },
    # --- individual SOPs in this repository ---
    "venous": {
        # No IH series exists. The register already holds HEM-SOP-001, "Vein blood
        # collection and preparation of hematology sample", so venous collection is
        # a REVISION of that document, not a new one. Revisions keep the serial and
        # the version goes up (Base rule: serial held for life, never reused).
        "code": "RliDHL-HEM-SOP-001", "vers": "002",
        "title": "Venous Blood Sample Collection by Venipuncture",
        "master_list": "Master List, Haematology",
    },
    "glucose": {
        # 001 was carried by two documents: this SOP at VERS-005, and "Use and
        # preventive maintenance of Humastar 80" at VERS-002. 001 is retained here.
        # The Humastar SOP does NOT move to 002: it is withdrawn, because the
        # analyser has been decommissioned. BIO serial 002 belongs to the Humalyzer
        # SOP, which is withdrawn for the same reason. So 001 is the only live BIO
        # serial below 010, and nothing had to be skipped.
        "code": "RliDHL-BIO-SOP-001", "vers": "005",
        "title": "Glucose Testing using Cobas c 111",
        "master_list": "Master List, Clinical Chemistry",
    },
}

# Forms are addressed as REGISTER[<parent key>] + this suffix.
FORM_SIG = "F01"          # Training and competence acknowledgement + authorisation
FORM_SIG_2 = "F01-2"      # continuation sheet for larger sections


def code(key, vers=None):
    """Document code, e.g. RliDHL-ADM-SOP-004."""
    r = REGISTER[key]
    return f"{r['code']}-VERS-{vers or r['vers']}"


def form(key, suffix):
    """Form code bound to its parent SOP, e.g. RliDHL-ADM-SOP-064-F01."""
    return f"{REGISTER[key]['code']}-{suffix}"


def full_title(key):
    return f"{code(key)} — {REGISTER[key]['title']}"


# References are built from the register so a renumbered document cannot leave a
# stale reference behind in the text.
ADM_DOC_CONTROL = code("doc_control")
ADM_PRE_EXAM = code("pre_exam")
ADM_TRAINING = code("training")
FORM_SIG_CODE = form("training", FORM_SIG)
FORM_SIG2_CODE = form("training", FORM_SIG_2)

# printed on EVERY page footer (ISO 15189:2022, 8.3)
COPY_NOTICE = "Uncontrolled when printed unless stamped CONTROLLED COPY"

CONTROLLED_COPIES = [
    "Controlled copies — Original: document control system (Laboratory Manager’s office). "
    "Read-only copies: laboratory network and EMR document module.",
    "A printed copy is uncontrolled. Distribution, supersession and withdrawal of obsolete "
    f"copies are managed under {ADM_DOC_CONTROL} (ISO 15189:2022, 8.3).",
]

# Two controlled forms already exist for this purpose in MASTER LIST 2025.docx:
#   RliDHL-ADM-FORM-001-VERS 002  "Staff competency assessment form"
#   RliDHL-ADM-FORM-014-VERS 001  "Signature page"
# The instruction is to cite them rather than duplicate or withdraw them, so 064-F01
# binds to them by name and the SOP says so on the page a reader actually signs.
CITED_COMPETENCY_FORM = "RliDHL-ADM-FORM-001"
CITED_SIGNATURE_FORM = "RliDHL-ADM-FORM-014"

SIG_TITLE = ("21. Signature Page — Training and Competence Acknowledgement  "
             f"[{FORM_SIG_CODE} with {CITED_COMPETENCY_FORM}, {CITED_SIGNATURE_FORM}]")
SIG_STATEMENT = (
    "I confirm that I have been trained on this SOP, that I have read and understood it, "
    "including the safety precautions, and that I agree to perform it as written. I will not "
    "perform it without supervision until this record is complete. This page is used with "
    f"{CITED_COMPETENCY_FORM} (staff competency assessment) and {CITED_SIGNATURE_FORM} "
    "(signature page); it does not replace either, and neither needs to be withdrawn."
)
SIG_HEADERS = ["S/No", "Full Name (print)", "IPPIS No.", "Designation", "Signature", "Date"]
# 18 rows so that a section of at least 12 staff signs on ONE page with room to spare
# (ISO 15189:2022, 6.2.2-6.2.3: training, competence and authorisation records).
SIG_ROW_COUNT = 18
SIG_ROWS = [[str(i), "", "", "", "", ""] for i in range(1, SIG_ROW_COUNT + 1)]
SIG_WIDTHS = [1.2, 4.6, 2.6, 3.4, 3.0, 2.2]
SIG_TRAINER = ("Trainer / Supervisor: ______________________    Name (print): ______________    "
               "Date: ____________")
# Competence is ASSESSED, not merely acknowledged (6.2.2) — the block is part of the record.
SIG_COMPETENCE = (
    "Competence was assessed and recorded before authorisation (ISO 15189:2022, 6.2.2):"
)
# one option per line, so no check box is ever split across a line break
SIG_COMPETENCE_ITEMS = [
    "[  ] direct observation of routine work",
    "[  ] review of work records",
    "[  ] assessment of problem-solving",
    "[  ] examination of provided samples",
]
SIG_AUTHORISED = ("Authorised to perform this procedure without supervision — "
                  "authorising officer: ____________________   Signature: ______________   "
                  "Date: ____________")
SIG_NOTE = ("Enter the IPPIS number (staff cards are not numbered). If more rows are needed, "
            f"attach continuation sheet {FORM_SIG2_CODE}.")

AMEND_TITLE = "22. Amendment Page"
AMEND_STATEMENT = (
    "Handwritten amendments are authorised by the Laboratory Manager and are entered here "
    "word for word before the next revision (ISO 15189:2022, 8.3)."
)
AMEND_HEADERS = ["Date", "Section", "Amendment (word for word)", "Reason", "Authorised by"]
AMEND_ROWS = [["", "", "", "", ""] for _ in range(14)]
AMEND_WIDTHS = [2.4, 2.0, 6.4, 3.4, 2.8]
# Reuses the dead space at the foot of the page for a record that is required anyway.
AMEND_FOOTNOTE = (
    "Superseded versions: withdrawn on the effective date of this version, marked OBSOLETE and "
    f"retained for the period in the Document and Record Control SOP ({ADM_DOC_CONTROL}) so "
    "that results issued under them remain traceable (ISO 15189:2022, 8.3)."
)

# ---------------------------------------------------------------------------
# Document 1 — Venous Blood Sample Collection
# ---------------------------------------------------------------------------

VENOUS = {
    # IH-SOP-001 is already "Vein Blood Collection and Preparation of Hematology
    # sample" in the Master List. Taking that serial would have made two different
    # documents answer to one code, so this SOP takes the next free IH serial.
    "out_stem": f"{code('venous')} - Venous Blood Sample Collection",
    "meta": {
        "doc_no": code("venous"),
        "title": "Venous Blood Sample Collection by Venipuncture",
        "version": REGISTER["venous"]["vers"],
    },
    "control_rows": [
        ("SOP Number", code("venous")),
        ("SOP Title", "Venous Blood Sample Collection by Venipuncture"),
        ("Version", "001"),
        ("Effective Date", "14/11/2026"),
        ("Next Review Date", "14/11/2028"),
        ("Supersedes", "— (first issue to the 2026 standard layout)"),
        ("Document Owner", "Kelvin SAMAHANI — Laboratory Manager"),
        ("Classification", "Controlled document — master copy: document control system "
                           f"({ADM_DOC_CONTROL})"),
    ],
    "revision_rows": [
        ["J.P. NSENGIYUMVA", "001", "14/11/2026",
         "First issue to the 2026 standard layout (ISO 15189:2022)."],
        ["", "", "", ""],
    ],
    "body": [
        ("1. Introduction", [
            ("p", "Venous blood collection is the most frequently performed pre-examination "
                  "activity of the laboratory, and pre-examination errors — wrong patient, wrong "
                  "tube, insufficient volume, haemolysis — are the largest source of unusable "
                  "specimens. Requests, patient identification, specimen tracking and results "
                  "are kept in the EMR: no paper request forms accompany specimens inside the "
                  "hospital. Referred specimens travel with a printed Sample Referral Manifest."),
        ]),
        ("2. Principle of the Test", [
            ("p", "Blood is drawn from a peripheral vein through a hollow needle into evacuated "
                  "tubes. The vacuum in each tube draws a fixed volume of blood and mixes it with "
                  "the additive (anticoagulant, clot activator or preservative) required by the "
                  "requested examination. Filling the tubes in a fixed order of draw prevents "
                  "additive carry-over; gentle inversion mixes blood and additive without "
                  "damaging the cells."),
        ]),
        ("3. Purpose", [
            ("p", "This SOP standardises the pre-examination steps for the collection of venous "
                  "blood by peripheral venipuncture so that specimens are collected safely, "
                  "identified without error, and remain fit for the examinations requested."),
        ]),
        ("4. Scope", [
            ("p", "This SOP applies to all laboratory personnel — biotechnologists, technicians, "
                  "interns and students under supervision — who collect venous blood at the main "
                  "laboratory and at every collection point of Ruli District Hospital. It covers "
                  "adult and paediatric venipuncture and all tube types listed in 6.3."),
            ("p", "It does not cover arterial blood collection, capillary blood sampling, blood "
                  "culture collection, or blood donation; these are covered by separate SOPs "
                  "listed in the document control master list."),
        ]),
        ("5. Definitions and Abbreviations", [
            ("h2", "5.1 Definitions"),
            ("tbl",
             ["No.", "Term", "Definition"],
             [
                 ["1", "Positive patient identification",
                  "Verification of identity with at least two identifiers (full name plus date "
                  "of birth or patient number) matched to the EMR request before collection."],
                 ["2", "Order of draw",
                  "The fixed sequence in which tubes are filled during one venipuncture, "
                  "preventing additive carry-over between tubes."],
                 ["3", "Haemolysis",
                  "Rupture of red blood cells releasing intracellular constituents into serum "
                  "or plasma; a common cause of specimen rejection."],
                 ["4", "Mislabelled specimen",
                  "A specimen whose label does not correctly and uniquely identify the "
                  "patient; handled as nonconforming work (ISO 15189:2022, 7.5)."],
                 ["5", "Referred specimen",
                  "A specimen sent to another laboratory for examination; travels with the "
                  "printed Sample Referral Manifest."],
             ],
             [1.2, 4.3, 11.5]),
            ("h2", "5.2 Abbreviations"),
            ("tbl",
             ["No.", "Acronym", "Full Meaning"],
             [
                 ["1", "EHO", "Environmental Health Officer"],
                 ["2", "EMR", "Electronic Medical Record"],
                 ["3", "EQA", "External Quality Assessment"],
                 ["4", "EDTA", "Ethylenediaminetetraacetic acid (tube anticoagulant)"],
                 ["5", "IPPIS", "Integrated Personnel and Payroll Information System"],
                 ["6", "LIS", "Laboratory Information System"],
                 ["7", "NC", "Nonconforming (work)"],
                 ["8", "NRL", "National Reference Laboratory (RBC)"],
                 ["9", "PPE", "Personal Protective Equipment"],
                 ["10", "RBC", "Rwanda Biomedical Centre"],
                 ["11", "RliDHL", "Ruli District Hospital Laboratory"],
                 ["12", "SOP", "Standard Operating Procedure"],
                 ["13", "TAT", "Turnaround Time"],
             ],
             [1.2, 4.3, 11.5]),
        ]),
        ("6. Specimen Requirements", [
            ("h2", "6.1 Request and patient identification"),
            ("b", [
                "Accept only requests entered in the EMR by an authorised requester. An oral "
                "request is accepted only in an emergency: record who made it and when, and "
                "have it entered in the EMR within 24 hours.",
                "Identify the patient with at least two identifiers — full name plus date of "
                "birth or patient number — matched to the EMR request. Scan the wristband where "
                "one is worn.",
                "Each specimen receives its unique laboratory number at reception; the number "
                "is scanned onto the tube label and follows the specimen, aliquots and report "
                "so that every result traces to one patient.",
            ]),
            ("h2", "6.2 Patient preparation and collection"),
            ("b", [
                "Confirm the request (examinations, tube types, fasting or timing) before "
                "approaching the patient. Explain the procedure, confirm consent, and seat or "
                "lie the patient comfortably.",
                "For fasting specimens, confirm at least 8 hours without food or drink other "
                "than water and record the fasting state in the EMR.",
                "Do not collect from an arm with a running infusion, a fistula, a mastectomy "
                "side, or burn/scarred skin unless no alternative exists; if used, note it in "
                "the EMR.",
            ]),
            ("h2", "6.3 Specimen types, volume, and storage"),
            ("tbl",
             ["Tube (cap colour / additive)", "Minimum volume", "Handling and storage"],
             [
                 ["Blood culture bottles (yellow) — if requested", "Per fill mark",
                  "Incubate promptly; do not refrigerate"],
                 ["Light blue — sodium citrate 3.2%", "Fill to the mark (9:1)",
                  "Invert gently 8–10 times; analyse within 24 h at 15–25 °C"],
                 ["Red / gold — plain or clot activator", "3 mL",
                  "Clot 30 min; centrifuge; serum 8 h at 15–25 °C, 72 h at 2–8 °C"],
                 ["Green — lithium heparin", "3 mL",
                  "Invert gently; centrifuge promptly; as for serum"],
                 ["Lavender — EDTA", "1 mL (2 mL paediatric)",
                  "Invert gently 8–10 times; do not shake; examine for clots"],
                 ["Grey — fluoride/oxalate", "2 mL",
                  "Invert gently; glucose preserved in the unseparated tube (24 h)"],
             ],
             [5.6, 4.1, 7.3]),
            ("h2", "6.4 Specimen acceptance and rejection criteria"),
            ("tbl",
             ["Criterion", "Acceptable", "Reject / request recollection"],
             [
                 ["Identification", "Label complete; matches the EMR request",
                  "Unlabelled, mislabelled, or mismatch with the request"],
                 ["Container", "Correct tube type; intact and closed",
                  "Wrong tube or additive; leaking or broken tube"],
                 ["Volume", "At or above the minimum in 6.3",
                  "Below minimum; citrate or EDTA not filled to the mark"],
                 ["Quality", "Clear serum or plasma; slight haemolysis noted",
                  "Gross haemolysis; clots in an EDTA tube; infusion arm"],
                 ["Timing", "Within the limits in 6.3; collection time recorded",
                  "Limits exceeded; collection time unknown"],
                 ["Transport", "Within 2 hours internally; manifest for referrals",
                  "Delayed beyond stability; referral without manifest"],
             ],
             [3.2, 6.9, 6.9]),
            ("p", "Record every rejected specimen in the specimen rejection register and inform "
                  "the requester at once. If a specimen that fails these criteria is "
                  "irreplaceable or clinically critical, the Laboratory Manager or designee may "
                  "authorise testing; the report then states the problem and the caution needed "
                  "to interpret the result."),
        ]),
        ("7. Responsibilities and Authorisation", [
            ("tbl",
             ["Task", "Responsible Person(s)"],
             [
                 ["Authorise this SOP and provide the resources needed to apply it",
                  "Hospital Director General — Dr. Déogratias KANEZA"],
                 ["Review this SOP at least every two years; ensure users are trained, assessed "
                  "and authorised; approve recollection exceptions",
                  "Laboratory Manager — Kelvin SAMAHANI"],
                 ["Maintain this SOP under document control; monitor indicators, nonconformities "
                  "and corrective actions; maintain training and authorisation records",
                  "Lab Quality Manager — Jean Paul NSENGIYUMVA"],
                 ["Perform venipuncture; identify the patient; label at the bedside; record the "
                  "collection in the EMR",
                  "Authorised Laboratory Biotechnologists"],
                 ["Receive specimens; decide acceptance and rejection (6.4); maintain the "
                  "specimen rejection register",
                  "Laboratory Receptionist / Laboratory Biotechnologist on duty"],
                 ["Manage exposure incidents, sharps and waste; keep PPE and vaccination "
                  "current",
                  "Environmental Health Officer (EHO); all laboratory staff"],
             ],
             [9.0, 8.0]),
            ("p", "Authorisation. Only staff trained on this SOP, assessed as competent and "
                  "authorised by the Laboratory Manager perform venipuncture without "
                  "supervision. Students and staff in training work under the direct supervision "
                  "of an authorised person, who checks and countersigns their work. Training, "
                  "competence assessment and authorisation are recorded in the personnel "
                  "training file; acknowledgement of this SOP is recorded on the Signature Page "
                  "(section 21). Competence is reassessed at the interval set in the competence "
                  "assessment procedure and whenever this SOP changes."),
        ]),
        ("8. Equipment, Reagents, and Consumables", [
            ("h2", "8.1 Equipment"),
            ("b", [
                "Vacuum tube holders and safety-engineered needles 21G or 22G; butterfly sets "
                "for difficult veins and paediatric patients.",
                "Disposable single-patient tourniquets.",
                "Mobile label printer and barcode scanner linked to the EMR.",
                "Cooled carriers (2–8 °C) and ambient transport boxes where required.",
            ]),
            ("h2", "8.2 Reagents, calibrators and controls"),
            ("p", "Not applicable — this procedure uses no reagents, calibrators or controls. "
                  "Alcohol swabs and tubes are consumables (8.3)."),
            ("h2", "8.3 Consumables"),
            ("b", [
                "Evacuated tubes listed in 6.3, within expiry and stored as the manufacturer "
                "directs.",
                "70% isopropyl alcohol swabs, dry sterile gauze, adhesive plaster.",
                "EMR-generated barcode tube labels; biohazard transport bags.",
                "Puncture-resistant sharps containers; clinical waste bags; 10% sodium "
                "hypochlorite; 70% ethanol.",
            ]),
            ("h2", "8.4 Calibration and metrological traceability"),
            ("p", "Not applicable — no measuring equipment is used in this procedure. "
                  "Centrifuges and analysers used downstream are controlled under the equipment "
                  "management procedure."),
        ]),
        ("9. Environment and Biosafety Measures", [
            ("b", [
                "Treat every specimen as potentially infectious and apply standard precautions.",
                "Wear a laboratory coat and gloves; change gloves between patients; add eye "
                "protection where splashing is possible. Perform hand hygiene before and after "
                "every collection.",
                "Never recap needles. Activate the safety device and discard the assembly in "
                "the sharps container at the point of use.",
                "Hepatitis B vaccination is mandatory for all staff who collect blood.",
                "After any exposure: wash the area immediately, inform the Laboratory Manager "
                "and follow the post-exposure procedure in the Laboratory Safety Manual. Record "
                "the event; raise a nonconformity where procedural failure contributed (7.5).",
                "Clean spills with 10% sodium hypochlorite; decontaminate benches with 70% "
                "ethanol. Discard used tubes and materials under the Laboratory Waste "
                "Management procedure (waste segregation is supervised by the EHO).",
                "Keep all patient information confidential (ISO 15189:2022, 4.2).",
            ]),
        ]),
        ("10. Procedural Steps", [
            ("h2", "10.1 Preparation and patient identification"),
            ("steps", [
                ("1.", "Open the collection worklist on the EMR station or tablet; verify the "
                       "pending requests."),
                ("2.", "Assemble the tubes in the order of draw (10.2), needle, swabs, gauze, "
                       "labels and sharps container within reach."),
                ("3.", "Identify the patient with two identifiers matched to the EMR request; "
                       "scan the wristband where present. Perform hand hygiene and put on "
                       "gloves."),
            ]),
            ("h2", "10.2 Order of draw"),
            ("p", "When more than one tube is collected during one venipuncture, fill the tubes "
                  "in the sequence below (consistent with CLSI GP41). Invert each filled tube "
                  "gently 8 to 10 times; never shake. Fill citrate and EDTA tubes to the mark."),
            ("tbl",
             ["Order", "Tube cap colour / additive", "Typical examinations"],
             [
                 ["1", "Blood culture bottles (yellow)", "Blood cultures"],
                 ["2", "Light blue — sodium citrate 3.2%", "Coagulation (PT, aPTT, INR)"],
                 ["3", "Red / gold — plain or clot activator (with gel)", "Clinical chemistry, serology, immunology"],
                 ["4", "Green — lithium heparin", "Plasma chemistry, emergency tests"],
                 ["5", "Lavender — EDTA", "Haematology (CBC), HbA1c"],
                 ["6", "Grey — fluoride / oxalate", "Glucose, lactate"],
             ],
             [1.8, 7.2, 8.0]),
            ("h2", "10.3 Venipuncture"),
            ("steps", [
                ("1.", "Select the site: prefer the median cubital vein of the forearm. Avoid "
                       "the areas listed in 6.2."),
                ("2.", "Apply the tourniquet 3 to 4 cm above the site. Do not leave it on for "
                       "more than 1 minute."),
                ("3.", "Palpate the vein. Clean the site with a 70% isopropyl swab in a "
                       "widening circle over at least 2 cm and allow to air-dry for at least 30 "
                       "seconds. Do not touch, blow on or fan the site."),
                ("4.", "Anchor the vein, holding the skin taut. Insert the needle bevel upwards "
                       "at 15 to 30 degrees in one smooth motion."),
                ("5.", "When blood flows, fill the tubes in the order of draw (10.2). "
                       "Under-filling or over-filling citrate and EDTA tubes is a cause for "
                       "rejection."),
                ("6.", "Release the tourniquet before withdrawing the needle. Withdraw, apply "
                       "dry gauze with firm pressure for 2 to 3 minutes."),
                ("7.", "Activate the needle safety device immediately and discard the assembly "
                       "in the sharps container at the point of use."),
                ("8.", "If blood does not flow: do not probe more than twice. Release the "
                       "tourniquet, withdraw, apply pressure, and use a second site or the "
                       "other arm. Record every difficult draw in the EMR."),
            ]),
            ("h2", "10.4 Labelling and identification"),
            ("b", [
                "Label every tube at the bedside, in the patient’s presence, immediately after "
                "filling: scan the wristband, then scan/print the tube label from the EMR. "
                "Never pre-label tubes and never label away from the patient.",
                "Verify each label against the request before leaving the patient. The label "
                "carries the patient name, patient number, laboratory number, collection date "
                "and time, and the collector’s identifier.",
                "For referred specimens, record the Sample Referral Manifest number on the tube "
                "label and in the EMR.",
                "Corrections are made in the EMR only, traceable to the user account and time; "
                "on a printed manifest: single line through, correct entry, initial and date.",
            ]),
            ("h2", "10.5 Transport to the laboratory"),
            ("b", [
                "Internal transport: no paper accompanies the specimen. Deliver within 2 hours "
                "at 15–25 °C, or at 2–8 °C in a cooled carrier where required.",
                "Do not use the pneumatic tube system for coagulation samples or blood culture "
                "bottles unless validated for them.",
                "Referred specimens: print the Sample Referral Manifest as required by the "
                "national referral system; sender and receiving laboratory sign it and a copy "
                "is filed with the EMR episode.",
            ]),
            ("h2", "10.6 Post-collection"),
            ("b", [
                "Record in the EMR: collection time, collector, tube types, fasting state where "
                "relevant, and any note (difficult draw, haematoma, infusion arm).",
                "Dispose of waste and decontaminate the collection area as in section 9.",
            ]),
            ("h2", "10.7 Additional examinations on the same specimen"),
            ("p", "A clinician may ask for a further examination on a specimen already taken. "
                  "The time limits for each analyte and tube type are listed in the examination "
                  "SOPs; requests arriving after the limit are rejected and a new specimen is "
                  "requested. The laboratory number is never reused, so the additional "
                  "examination is unequivocally traceable to the original specimen "
                  "(ISO 15189:2022, 7.2.7.2)."),
        ]),
        ("11. Quality Control and External Quality Assessment", [
            ("h2", "11.1 Quality indicators (internal monitoring)"),
            ("b", [
                "Specimen rejection rate: at most 2% per month.",
                "Haemolysis rate: at most 1% per month.",
                "Mislabelled specimens: zero — every case is a nonconformity.",
                "Specimens delivered within 2 hours: at least 95%.",
                "The Lab Quality Manager extracts these indicators monthly from the EMR/LIS and "
                "reviews them with the Laboratory Manager.",
            ]),
            ("h2", "11.2 External quality assessment"),
            ("p", "The laboratory takes part in the pre-analytical or relevant EQA schemes "
                  "organised or designated by RBC/NRL and acts on unsatisfactory performance."),
            ("h2", "11.3 Actions when monitoring fails"),
            ("p", "When an indicator misses its target or an EQA result is unsatisfactory: "
                  "investigate the cause, define corrective action and evaluate its "
                  "effectiveness (ISO 15189:2022, 8.7). Where wrong-patient or mislabelled "
                  "specimens were involved, perform a look-back to the patients affected and "
                  "inform the Laboratory Manager immediately."),
        ]),
        ("12. Result Interpretation and Validation", [
            ("p", "Not applicable — this procedure produces no examination result. The decision "
                  "on specimen fitness is made at reception (6.4). Interpretation and validation "
                  "of results are covered by the examination SOP concerned."),
        ]),
        ("13. Performance Characteristics and Limitations", [
            ("b", [
                "Tourniquet time beyond 1 minute causes haemoconcentration and raises protein "
                "and lipid results.",
                "Squeezing the puncture site causes haemolysis and tissue-fluid contamination.",
                "Difficult veins and paediatric collection limit volume and quality; record the "
                "limitation in the EMR where it may affect a result.",
                "This procedure controls the pre-examination phase only; method performance, "
                "method verification and validation, and measurement uncertainty are stated in "
                "the examination SOP concerned and in the method verification file "
                "(ISO 15189:2022, 7.3.2-7.3.4).",
            ]),
        ]),
        ("14. Reference Intervals, Clinical Decision Limits, and Critical Results", [
            ("p", "Not applicable — no values are produced by this procedure. Reference "
                  "intervals, decision limits and critical results are defined in the "
                  "examination SOPs."),
        ]),
        ("15. Reporting and Release of Results", [
            ("p", "Not applicable for results — reporting is covered by the examination SOPs. "
                  "The collection event, specimen status (accepted/rejected) and any "
                  "recollection request are recorded in the EMR and are traceable to the "
                  "collector’s user account (ISO 15189:2022, 4.2, 7.4)."),
        ]),
        ("16. Specimen Retention and Disposal", [
            ("b", [
                "Collected specimens are held and transported under the conditions in 6.3 and "
                "10.5; retention after examination follows the examination SOP concerned.",
                "Used tubes, swabs and dressings are discarded as infectious waste; sharps go "
                "into the puncture-resistant container at the point of use.",
            ]),
        ]),
        ("17. Risk Management, Nonconforming Work, and Contingency", [
            ("tbl",
             ["Risk", "Control measure"],
             [
                 ["Wrong-patient identification",
                  "Two identifiers at the bedside; wristband scan; label printed and scanned in "
                  "the patient’s presence; never pre-label"],
                 ["Needlestick injury",
                  "Safety-engineered devices; no recapping; sharps container at the point of "
                  "use; hepatitis B vaccination; post-exposure procedure"],
                 ["Haemolysed or clotted specimens",
                  "Correct fill volumes; gentle inversion; order of draw; tourniquet under 1 "
                  "minute; prompt transport"],
                 ["Delayed transport affecting stability",
                  "Two-hour delivery rule; scheduled collection runs; cooled carriers"],
                 ["Loss of EMR availability",
                  "Downtime procedure below: manual labels double-checked by a second person; "
                  "back-entry within 24 hours"],
             ],
             [6.0, 11.0]),
            ("p", "Nonconforming work (ISO 15189:2022, 7.5): if any step was not followed — "
                  "wrong tube, missed identification, unlabelled tube — stop, report to the "
                  "Laboratory Manager, and record the event. The Laboratory Manager decides on "
                  "the affected specimens, arranges recollection, and determines whether a "
                  "look-back to released results is needed."),
            ("p", "Contingency (ISO 15189:2022, 7.8): when the EMR/LIS is unavailable, use the "
                  "approved downtime request form and manual tube labels; a second person "
                  "double-checks every label; keep a written collection log; enter all data "
                  "into the EMR within 24 hours of restoration. Referral manifests continue to "
                  "be printed as required by the national referral system."),
        ]),
        ("18. Records", [
            ("tbl",
             ["Record", "Where kept", "Retention"],
             [
                 ["Collection events, labels, collector identity", "EMR / LIS (electronic)",
                  "As set in the laboratory records policy"],
                 ["Specimen rejection register", "EMR / LIS (electronic)", "As above"],
                 ["Nonconformity and corrective action records", "Quality module (electronic)",
                  "As above"],
                 ["Training, competence and authorisation (section 21)",
                  "Personnel training file", "Duration of employment plus the period required "
                  "by national rules"],
                 ["Exposure incident reports", "Health & Safety file", "As above"],
                 ["Sample Referral Manifest (printed)", "Referral file; copy with the EMR "
                  "episode", "Per national referral policy"],
             ],
             [6.5, 5.5, 5.0]),
            ("p", "Records are created at the time the activity is performed, are legible, "
                  "traceable to the person or user account, and are retrievable (ISO 15189:2022, "
                  f"8.4; {ADM_DOC_CONTROL})."),
            ("p", "Control of data and information management (ISO 15189:2022, 7.6): the EMR "
                  "and the LIS that carry requests, labels and collection records are validated "
                  "and verified before use, are protected against unauthorised access, "
                  f"tampering and loss, and are backed up under {ADM_DOC_CONTROL}. Every access "
                  "and correction carries the user account and a time stamp, and no record is "
                  "deleted. Access to patient data is limited to the duty in hand "
                  "(confidentiality, ISO 15189:2022, 4.2)."),
        ]),
        ("19. Supporting Documents", [
            ("b", [
                f"Document and Record Control ({ADM_DOC_CONTROL}).",
                f"Pre-Examination Procedure ({ADM_PRE_EXAM}).",
                "Laboratory Safety Manual (post-exposure procedure).",
                "Laboratory Waste Management procedure.",
                "Result Reporting and Release procedure.",
                "Sample Referral procedure and the national Sample Referral Manifest.",
                "Competence Assessment and Training procedure.",
            ]),
        ]),
        ("20. References", [
            ("b", [
                "ISO 15189:2022 — Medical laboratories: requirements for quality and competence.",
                "CLSI GP41 — Collection of Diagnostic Venous Blood Specimens.",
                "CLSI H3 — Procedures for the Handling and Processing of Blood Specimens.",
                "WHO — Laboratory Quality Management System (LQMS) Handbook.",
                "Government of Rwanda paperless operations policy (01 August 2026).",
            ]),
        ]),
    ],
}

# ---------------------------------------------------------------------------
# Document 2 — Glucose Testing using Cobas c 111 (sample)
# ---------------------------------------------------------------------------

GLUCOSE = {
    "out_stem": f"{code('glucose')} - Glucose Testing using Cobas c 111",
    "meta": {
        "doc_no": code("glucose"),
        "title": "Glucose Testing using Cobas c 111",
        "version": REGISTER["glucose"]["vers"],
    },
    "control_rows": [
        ("SOP Number", code("glucose")),
        ("SOP Title", "Glucose Testing using Cobas c 111"),
        ("Version", "005"),
        ("Effective Date", "14/11/2026"),
        ("Next Review Date", "14/11/2028"),
        ("Supersedes", f"{code('glucose', '004')} (effective 16/03/2025, review 16/03/2027)"),
        ("Document Owner", "Kelvin SAMAHANI — Laboratory Manager"),
        ("Classification", "Controlled document — master copy: document control system "
                           f"({ADM_DOC_CONTROL})"),
    ],
    "revision_rows": [
        ["N/A", "001", "15/05/2019", "First / original version"],
        ["Elise NIMUKUZE", "002", "15/04/2021", "Version and effective date update"],
        ["Aurélie YANSONEYE", "003", "15/03/2023", "Version and effective date update"],
        ["Kelvin SAMAHANI", "004", "16/03/2025", "Logo, version, effective date"],
        ["Kelvin SAMAHANI", "005", "14/11/2026",
         "Full revision to the 2026 standard layout aligned with ISO 15189:2022: specimen "
         "handling, calibration and traceability, two-level control with Westgard rules, "
         "measuring range, interferences, reference intervals in both units, clinical decision "
         "limits, critical values, interpretive notes, risks and records."],
    ],
    "body": [
        ("1. Introduction", [
            ("p", "Glucose is the body’s main source of energy. Its concentration in blood is "
                  "kept within narrow limits by uptake from the gut, storage and release by the "
                  "liver, and use by muscle and other tissues. Glucose is measured to diagnose "
                  "and monitor diabetes mellitus, to detect hypoglycaemia, and in pancreatic, "
                  "metabolic, endocrine and critical illness."),
            ("p", "Ruli District Hospital Laboratory measures glucose in serum and plasma on the "
                  "Roche cobas c 111 analyser with the Glucose HK (GLUC2) reagent. Requests and "
                  "results are kept in the EMR."),
        ]),
        ("2. Principle of the Test", [
            ("p", "The GLUC2 assay is an enzymatic ultraviolet test based on the hexokinase "
                  "reference method. Hexokinase phosphorylates glucose to "
                  "glucose-6-phosphate with ATP; glucose-6-phosphate dehydrogenase then "
                  "oxidises it with NADP+ to gluconate-6-phosphate. The rate of NADPH "
                  "formation is directly proportional to the glucose concentration and is "
                  "measured photometrically at 340 nm. The method is standardised by the "
                  "manufacturer against isotope dilution mass spectrometry (ID-MS)."),
        ]),
        ("3. Purpose", [
            ("p", "This SOP standardises the pre-examination, examination and post-examination "
                  "steps for the quantitative measurement of glucose in serum and plasma on the "
                  "cobas c 111, so that results are accurate, reproducible and clinically "
                  "reliable."),
        ]),
        ("4. Scope", [
            ("p", "This SOP applies to all laboratory biotechnologists who operate the cobas "
                  "c 111 and perform glucose testing in the biochemistry section. It covers "
                  "specimen receipt, calibration, quality control, analysis, validation and "
                  "reporting. It does not cover glucose measured with a glucometer at the "
                  "bedside (point-of-care), or glucose in urine or cerebrospinal fluid."),
        ]),
        ("5. Definitions and Abbreviations", [
            ("h2", "5.1 Definitions"),
            ("tbl",
             ["No.", "Term", "Definition"],
             [
                 ["1", "Fasting glucose",
                  "Glucose measured after at least 8 hours without food or drink other than "
                  "water."],
                 ["2", "Random glucose",
                  "Glucose measured at any time, without regard to the last meal; interpreted "
                  "with the collection time and the patient’s state."],
                 ["3", "Glycolysis",
                  "Use of glucose by blood cells in the unseparated tube, which lowers the "
                  "measured value (about 7% per hour at room temperature)."],
                 ["4", "Calibration",
                  "Setting the relationship between the instrument signal and the glucose "
                  "concentration with a calibrator of known value."],
                 ["5", "Measuring range",
                  "The interval of concentrations the method reports without dilution."],
                 ["6", "Westgard rules",
                  "Statistical rules used to decide whether a control result is acceptable."],
             ],
             [1.2, 4.3, 11.5]),
            ("h2", "5.2 Abbreviations"),
            ("tbl",
             ["No.", "Acronym", "Full Meaning"],
             [
                 ["1", "C.f.a.s.", "Calibrator for automated systems (Roche)"],
                 ["2", "CV", "Coefficient of Variation"],
                 ["3", "EHO", "Environmental Health Officer"],
                 ["4", "EMR", "Electronic Medical Record"],
                 ["5", "EQA", "External Quality Assessment"],
                 ["6", "GLUC2", "Roche Glucose HK reagent, second generation"],
                 ["7", "IPPIS", "Integrated Personnel and Payroll Information System"],
                 ["8", "IQC", "Internal Quality Control"],
                 ["9", "LIS", "Laboratory Information System"],
                 ["10", "NRL", "National Reference Laboratory (RBC)"],
                 ["11", "PPE", "Personal Protective Equipment"],
                 ["12", "RBC", "Rwanda Biomedical Centre"],
                 ["13", "RliDHL", "Ruli District Hospital Laboratory"],
                 ["14", "SD", "Standard Deviation"],
                 ["15", "SOP", "Standard Operating Procedure"],
                 ["16", "TAT", "Turnaround Time"],
             ],
             [1.2, 4.3, 11.5]),
        ]),
        ("6. Specimen Requirements", [
            ("h2", "6.1 Request and patient identification"),
            ("b", [
                "Accept requests entered in the EMR by an authorised requester. The request "
                "states the patient’s identifiers, the examination, and whether the specimen "
                "is fasting, random or timed — the result cannot be interpreted without this.",
                "Identify the patient with at least two identifiers matched to the EMR request; "
                "label the specimen in the patient’s presence (Pre-Examination Procedure, "
                f"{ADM_PRE_EXAM}).",
            ]),
            ("h2", "6.2 Patient preparation and collection"),
            ("b", [
                "Fasting glucose: no food or drink except water for at least 8 hours; record "
                "the fasting state and the exact collection time in the EMR.",
                "Do not collect from an arm into which a glucose infusion is running.",
                "Collect venous blood according to Venous Blood Collection "
                f"({code('venous')}): a plain or clot-activator tube for serum, or a "
                "fluoride/oxalate tube when the specimen cannot reach the laboratory and be "
                "centrifuged within 30 minutes.",
            ]),
            ("h2", "6.3 Specimen types, volume, and storage"),
            ("tbl",
             ["Specimen", "Minimum volume", "Handling and storage"],
             [
                 ["Serum (plain or clot-activator tube)", "300 µL in the sample cup",
                  "Centrifuge and separate within 30 min; separated serum: 8 h at 15–25 °C, "
                  "72 h at 2–8 °C"],
                 ["Plasma (lithium heparin or EDTA)", "300 µL in the sample cup",
                  "As for serum"],
                 ["Fluoride/oxalate plasma", "300 µL in the sample cup",
                  "Glucose preserved in the unseparated tube: at least 24 h at 15–25 °C"],
             ],
             [5.6, 4.1, 7.3]),
            ("p", "NOTE: in blood left unseparated at room temperature, glucose falls by about "
                  "7% in one hour. A specimen in a tube without fluoride that reaches the "
                  "laboratory late gives a falsely low result."),
            ("h2", "6.4 Specimen acceptance and rejection criteria"),
            ("tbl",
             ["Criterion", "Acceptable", "Reject / request recollection"],
             [
                 ["Labelling", "Fully labelled; matches the request; collection time and "
                  "fasting state given", "Unlabelled, mislabelled, or mismatch with the request"],
                 ["Container", "Plain, clot-activator, heparin, EDTA or fluoride/oxalate tube",
                  "Wrong tube; leaking or broken tube"],
                 ["Volume", "At least 300 µL of serum or plasma", "Insufficient volume"],
                 ["Quality", "Clear serum or plasma; slight haemolysis, icterus or lipaemia "
                  "accepted and noted on the report",
                  "Specimen drawn from an infusion arm; grossly contaminated specimen"],
                 ["Age", "Separated within 30 min, or collected in fluoride; within the "
                  "storage limits above",
                  "Unseparated, non-fluoride blood older than 30 min; beyond the limits"],
             ],
             [3.2, 6.9, 6.9]),
            ("p", "Record every rejected specimen in the specimen rejection register and inform "
                  "the requester at once. If a specimen that fails these criteria is "
                  "irreplaceable or clinically critical, the Laboratory Manager or designee may "
                  "authorise testing; the report then states the caution needed to interpret "
                  "the result."),
        ]),
        ("7. Responsibilities and Authorisation", [
            ("tbl",
             ["Task", "Responsible Person(s)"],
             [
                 ["Authorise this SOP and provide the resources needed to apply it",
                  "Hospital Director General — Dr. Déogratias KANEZA"],
                 ["Train, assess and authorise staff on this SOP; review it at least every two "
                  "years; validate and release results",
                  "Laboratory Manager — Kelvin SAMAHANI"],
                 ["Monitor IQC, EQA, nonconformities and corrective actions for this procedure; "
                  "keep the master list and controlled copies up to date",
                  "Lab Quality Manager — Jean Paul NSENGIYUMVA"],
                 ["Ensure reagents, controls and consumables are in stock, within expiry and "
                  "stored correctly",
                  "Laboratory Manager / Hospital Pharmacist"],
                 ["Perform start-up checks, calibration, quality control and glucose testing "
                  "on the cobas c 111; record results",
                  "Authorised Laboratory Biotechnologists"],
                 ["Perform daily, weekly and monthly maintenance of the cobas c 111 and record "
                  "it", "Laboratory Biotechnologist on duty"],
                 ["Dispose of biological and chemical waste safely; supervise waste segregation",
                  "Environmental Health Officer (EHO); all laboratory staff"],
             ],
             [9.0, 8.0]),
            ("p", "Authorisation. Only staff trained on this SOP, assessed as competent and "
                  "authorised by the Laboratory Manager perform this procedure and release its "
                  "results without supervision. Staff in training work under the direct "
                  "supervision of an authorised person, who checks and countersigns their work. "
                  "Training, competence assessment and authorisation are recorded in the "
                  "personnel training file; acknowledgement of this SOP is recorded on the "
                  "Signature Page (section 21)."),
        ]),
        ("8. Equipment, Reagents, and Consumables", [
            ("h2", "8.1 Equipment"),
            ("b", [
                "Roche cobas c 111 analyser (in service: serial 65053, hospital tag "
                "RULIDH/LAB/ME/0486, Biochemistry Room — Laboratory Medical Equipment "
                "Presence List, 16/06/2025); the serial number and asset tag are also "
                "recorded in the instrument logbook.",
                "Technical service and installation (electrodes, interface, corrective "
                "maintenance): MediSell Rwanda Ltd.",
                "Uninterruptible power supply (UPS).",
                "Centrifuge; refrigerator at 2–8 °C for reagents, calibrator and controls "
                "(temperature recorded twice daily).",
                "Micropipettes (20–200 µL and 100–1000 µL).",
            ]),
            ("h2", "8.2 Reagents, calibrators and controls"),
            ("b", [
                "Glucose HK (GLUC2) reagent for the cobas c 111; store at 2–8 °C. Record the "
                "lot number, expiry date and loading date; respect the on-board stability in "
                "the method sheet.",
                "Calibrator: C.f.a.s., reconstituted and stored as its insert directs; "
                "deionised water is the zero calibrator.",
                "Controls: PreciControl ClinChem Multi 1 (PCC1, normal) and Multi 2 (PCC2, "
                "pathological), or equivalent with assigned values. Follow the insert for "
                "reconstitution and stability.",
                "On receipt, check each consignment for damage, transport conditions, lot "
                "number and expiry; enter it in the stock records.",
                "When a new reagent lot is loaded, calibrate and then run both control levels "
                "before patient specimens. When a new control lot is started, run it alongside "
                "the lot in use to confirm its target values.",
                "Do not use any reagent, control or device past its expiry date or if storage "
                "conditions were not met. Keep the manufacturer’s current instructions for use "
                "at the bench.",
                "Report any defect or adverse incident linked to a reagent, kit or instrument "
                "to the Laboratory Manager, who records it and informs the supplier and, where "
                "required, RBC/NRL and the Rwanda Food and Drugs Authority.",
            ]),
            ("p", "Procurement traceability. Reagent packs are ordered through Rwanda "
                  "Medical Supply (RMS) local purchase orders — PO 2100000988 (SO 2096), "
                  "PO 2100001448 (SO 2700), PO 2100001746 (SO 3011) and PO 2100002205 "
                  "(SO 3660, delivered 17/03/2026) — fulfilled by MediSell Rwanda Ltd "
                  "(TIN 101897275); the GLUC2 400-test pack is Roche item code "
                  "04657527190. Delivery notes showing lot numbers and expiry dates are "
                  "retained with the stock records."),
            ("h2", "8.3 Consumables"),
            ("b", [
                "Sample cups, cuvette segments, pipette tips, tube racks.",
                "Gloves, laboratory coat, biohazard and liquid waste containers, 10% sodium "
                "hypochlorite, 70% ethanol.",
            ]),
            ("h2", "8.4 Calibration and metrological traceability"),
            ("b", [
                "Calibration mode: linear, two points: deionised water and C.f.a.s.",
                "Calibrate after each change of reagent lot, after major maintenance, and "
                "whenever quality control shows a shift or failure traced to calibration.",
                "Check that the instrument accepts the calibration, then run both control "
                "levels before any patient specimen. Record the calibrator lot, date and "
                "acceptance in the instrument logbook.",
                "Traceability: standardised by the manufacturer against isotope dilution mass "
                "spectrometry (ID-MS), the reference method for glucose.",
            ]),
        ]),
        ("9. Environment and Biosafety Measures", [
            ("b", [
                "Keep the analyser on a stable surface, away from direct sunlight, vibration, "
                "dust, air-conditioning outlets and heat sources. Record the room temperature "
                "each day.",
                "Reagents contain hazardous chemicals: read the safety data sheet; avoid "
                "contact with skin and eyes.",
                "Treat every specimen as potentially infectious: laboratory coat and gloves; "
                "eye protection where splashing is possible; hand hygiene after removing "
                "gloves.",
                "Clean spills with 10% sodium hypochlorite; decontaminate the bench with 70% "
                "ethanol at the end of each session.",
                "After any exposure, wash immediately, inform the Laboratory Manager and follow "
                "the post-exposure procedure in the Laboratory Safety Manual.",
                "Discard used devices, tips, tubes and specimens as infectious waste; empty the "
                "liquid waste container of the analyser as infectious liquid waste (Laboratory "
                "Waste Management procedure; supervised by the EHO).",
            ]),
        ]),
        ("10. Procedural Steps", [
            ("h2", "10.1 Instrument start-up and daily checks"),
            ("steps", [
                ("1.", "Check that the analyser and the UPS are connected. Switch on the cobas "
                       "c 111 and wait for the self-check."),
                ("2.", "Check the system water and the waste container; refill or empty as "
                       "needed."),
                ("3.", "Check the GLUC2 reagent: tests remaining, on-board expiry and lot "
                       "number. Load a new bottle if needed and record it in the reagent log."),
                ("4.", "Confirm the reaction temperature shown by the analyser is 37 °C "
                       "(± 0.5 °C). Perform the daily maintenance the instrument asks for and "
                       "record it."),
            ]),
            ("h2", "10.2 Calibration"),
            ("p", "Calibrate when section 8.4 requires it. Bring the reconstituted calibrator "
                  "to room temperature, run the calibration from the analyser menu, and "
                  "continue only when the calibration is accepted."),
            ("h2", "10.3 Internal quality control"),
            ("b", [
                "Run PCC1 and PCC2 like patient specimens at the start of each day of testing "
                "before any patient specimen, and after every calibration.",
                "Plot each result on the Levey-Jennings chart and apply the rules in section 11.",
                "Proceed to patient specimens only when both levels are acceptable.",
            ]),
            ("h2", "10.4 Specimen preparation and loading"),
            ("steps", [
                ("1.", "Check the identity of the specimen against the request (names, number, "
                       "test, date and time of collection)."),
                ("2.", "Centrifuge the tube at 3000–5000 rpm for 5 minutes. Inspect the serum "
                       "or plasma for haemolysis, lipaemia and icterus and note what you see."),
                ("3.", "Label a sample cup with the laboratory code and pipette at least 300 µL "
                       "of serum or plasma into it."),
                ("4.", "Enter or check the glucose (GLUC2) order in the worklist, place the cup "
                       "in the indicated position, confirm the position matches the laboratory "
                       "code, and start the run."),
            ]),
            ("h2", "10.5 Analysis and result review"),
            ("b", [
                "The analyser performs the reaction and calculates the glucose concentration.",
                "Review each result on the screen for flags and alarms.",
                "If the result is above the measuring range (above 41.6 mmol/L / 750 mg/dL), "
                "dilute with 0.9% sodium chloride as the method sheet directs, repeat, and "
                "multiply by the dilution factor.",
                "If a result is implausible or not matching earlier results, check the specimen "
                "for late separation or contamination and repeat before reporting.",
            ]),
            ("h2", "10.6 Post-procedure steps"),
            ("b", [
                "Cap and keep the specimens as described in section 16.",
                "Discard used cups and tips as infectious waste; decontaminate the bench.",
                "Perform the end-of-day maintenance and shutdown steps and record them.",
            ]),
        ]),
        ("11. Quality Control and External Quality Assessment", [
            ("h2", "11.1 Internal quality control"),
            ("b", [
                "Two levels (PCC1 and PCC2) at the start of each day of testing, after each "
                "calibration, after a reagent lot change, and after maintenance or repair.",
                "Rules: one result outside ± 2 SD is a warning — inspect the chart and repeat "
                "the control. One result outside ± 3 SD, two consecutive results outside "
                "± 2 SD on the same side of the mean, or a range of 4 SD between levels is a "
                "rejection.",
                "Plot every result on the Levey-Jennings chart. The Laboratory Manager or Lab "
                "Quality Manager reviews the charts each week for shifts and trends and each "
                "month calculates the mean, SD and CV for each level.",
                "Record control results, lot numbers and any action in the quality control "
                "register.",
            ]),
            ("h2", "11.2 External quality assessment"),
            ("p", "The laboratory takes part in the clinical chemistry EQA scheme organised or "
                  "designated by RBC/NRL and tests the samples in the same way as patient "
                  "specimens. Unsatisfactory performance is investigated as in 11.3."),
            ("h2", "11.3 Actions when a control fails"),
            ("p", "Stop patient testing, identify the cause (reagent, calibration, instrument "
                  "or specimen), correct it, and repeat both control levels before resuming. "
                  "Assess whether patient results reported since the last acceptable control "
                  "are affected; repeat affected specimens and issue corrected reports where "
                  "needed (ISO 15189:2022, 7.5, 8.7). Record everything in the quality control "
                  "register and the nonconformity log."),
            ("h2", "11.4 Comparability of results"),
            ("p", "Where the same examination is performed by more than one method, on more "
                  "than one analyser, or at more than one site (including a backup c 111 "
                  "during an instrument breakdown), comparability is established with patient "
                  "specimens across the clinically significant interval — not with control "
                  "material alone — whenever a method, analyser or site is added or changed, "
                  "and is reviewed at least once a year. The results and the acceptability "
                  "decision are recorded; where a clinically significant difference is found, "
                  "its effect on the reference intervals and clinical decision limits in "
                  "section 14 is evaluated and users are informed (ISO 15189:2022, 7.3.7.4)."),
            ("h2", "11.5 Control of data and information management"),
            ("p", "Results, control results, calibrations and reagent lot numbers are held in "
                  "the analyser and in the EMR/LIS, which are validated and verified before "
                  "use, protected against tampering and loss, and supported by a documented "
                  "downtime procedure (section 17). Access is by individual user account; "
                  "results are validated and released under the account of the person who "
                  "did so, and no result is altered without an audit trail. Data transferred "
                  "between the analyser and the EMR is checked against the worklist "
                  "(ISO 15189:2022, 7.6)."),
        ]),
        ("12. Result Interpretation and Validation", [
            ("p", "The analyser reports glucose in mmol/L (conversion: mg/dL ÷ 18.02 = mmol/L). "
                  "Before release, the Laboratory Manager or authorised biotechnologist checks: "
                  "the result against the previous result (delta check) and the clinical "
                  "information; the control status of the run; the specimen notes (haemolysis, "
                  "icterus, lipaemia, infusion arm, delayed separation); and the flags raised "
                  "by the analyser. Results are validated in the EMR under the user’s account "
                  "(ISO 15189:2022, 7.4.1)."),
            ("tbl",
             ["Observation", "Interpretive note (common causes)"],
             [
                 ["High glucose (hyperglycaemia)",
                  "Diabetes mellitus; stress response in critical illness; corticosteroids or "
                  "thiazides; Cushing syndrome; pancreatitis; parenteral nutrition; specimen "
                  "from a glucose infusion arm (falsely high)."],
                 ["Low glucose (hypoglycaemia)",
                  "Insulin or oral hypoglycaemic excess; missed meals or prolonged fasting; "
                  "sepsis; liver failure; adrenal insufficiency; insulinoma; neonatal "
                  "hypoglycaemia."],
                 ["Falsely low result",
                  "Delayed separation of a non-fluoride tube (glycolysis, about 7% per hour); "
                  "excess fluoride dilution."],
                 ["Result not matching the patient",
                  "Repeat the specimen; check identification and collection notes before "
                  "release; treat mislabelling as nonconforming work (7.5)."],
             ],
             [5.0, 12.0]),
        ]),
        ("13. Performance Characteristics and Limitations", [
            ("b", [
                "Measuring range: 0.11–41.6 mmol/L (2–750 mg/dL); dilute and repeat above the "
                "range (10.5). Verify against the method sheet of the lot in use.",
                "Within-laboratory precision: CV at most 2.5% at concentrations around 5 and "
                "15 mmol/L (see the method verification file).",
                "Interferences: severe haemolysis, icterus and lipaemia affect the result "
                "according to the indices given in the instructions for use; note any accepted "
                "interference on the report.",
                "Measurement uncertainty: see the method verification file (ISO 15189:2022, "
                "7.3.4); where relevant it is stated with the result.",
            ]),
        ]),
        ("14. Reference Intervals, Clinical Decision Limits, and Critical Results", [
            ("p", "Normal range in use at this laboratory (Laboratory Normal Ranges list, "
                  "Ruli District Hospital, approved 08/07/2025 by the Laboratory Manager and "
                  "the Hospital Director General):"),
            ("tbl",
             ["Test / specimen", "Normal range (in use)", "Normal range (mmol/L)"],
             [
                 ["Serum glucose", "70 – 110 mg/dL", "3.9 – 6.1"],
             ],
             [6.0, 5.5, 5.5]),
            ("p", "The range is verified before first use and after any method or reagent "
                  "change (ISO 15189:2022, 7.3.5). Interpret paediatric and neonatal results "
                  "in the clinical context."),
            ("p", "In the Laboratory Normal Ranges list, an entry written as ‘>x’ "
                  "means that any result under x is normal: the stated figure is the upper "
                  "limit of the normal range."),
            ("p", "Clinical decision limits for the diagnosis of diabetes (WHO/ADA criteria):"),
            ("tbl",
             ["Test / state", "Decision limit", "Meaning"],
             [
                 ["Fasting plasma glucose", ">= 7.0 mmol/L (>= 126 mg/dL)", "Diabetes mellitus"],
                 ["Fasting plasma glucose", "5.6 – 6.9 mmol/L (100 – 125 mg/dL)",
                  "Impaired fasting glucose"],
                 ["2-hour oral glucose tolerance test", ">= 11.1 mmol/L (>= 200 mg/dL)",
                  "Diabetes mellitus"],
                 ["2-hour oral glucose tolerance test", "7.8 – 11.0 mmol/L (140 – 199 mg/dL)",
                  "Impaired glucose tolerance"],
                 ["Random glucose with classic symptoms",
                  ">= 11.1 mmol/L (>= 200 mg/dL)", "Diabetes mellitus"],
                 ["Hypoglycaemia alert", "< 3.9 mmol/L (< 70 mg/dL)", "Alert value — inform the ward"],
                 ["Clinically significant hypoglycaemia", "< 3.0 mmol/L (< 54 mg/dL)",
                  "Treat without delay"],
             ],
             [5.6, 5.9, 5.5]),
            ("p", "Critical results (Laboratory Critical Results list, Ruli District Hospital): "
                  "glucose below 40 mg/dL (2.2 mmol/L) or at 399.6 mg/dL (22.2 mmol/L) and "
                  "above. Report every critical result immediately after verification, even if "
                  "other tests in the panel are still in progress. Notification follows the "
                  "mandatory ‘Write Down and Read Back’ process; record the name of the "
                  "person informed and the date and time (ISO 15189:2022, 7.4.1.3)."),
            ("p", "Escalation when a responsible person cannot be contacted (ISO 15189:2022, "
                  "7.4.1.3 c): telephone the requesting clinician, then the clinical head of "
                  "the ward or department, then the Hospital Director General’s duty officer, "
                  "within 30 minutes of verifying the result. Record each attempt, with the "
                  "time and the number called. If no clinician has been reached within "
                  "30 minutes, telephone the Laboratory Manager, who takes the result to the "
                  "ward in person. The result is never left unreported because a call was not "
                  "answered."),
        ]),
        ("15. Reporting and Release of Results", [
            ("p", "Turnaround time (Laboratory Tests TAT list, Ruli District Hospital): 45 "
                  "minutes from specimen receipt to result communication. Results are "
                  "communicated immediately after verification, even if other tests in the "
                  "panel remain in progress."),
            ("b", [
                "The report shows: patient identifiers, examination, result with units (mg/dL "
                "and mmol/L), the normal range in use for this laboratory, specimen notes, "
                "collection time and the reporting laboratory.",
                "Critical results (14) are reported immediately using the mandatory ‘Write "
                "Down and Read Back’ process, and the communication is recorded.",
                "Corrected reports: issue an amended report identifying the change, the reason "
                "and the date; notify the requester; retain the original entry.",
                "Patient information is confidential (ISO 15189:2022, 4.2).",
            ]),
        ]),
        ("16. Specimen Retention and Disposal", [
            ("b", [
                "After reporting, keep specimens for 7 days at 2–8 °C for repeats or "
                "investigation, then dispose of them under the Laboratory Waste Management "
                "procedure.",
                "Specimens under investigation for a nonconformity or complaint are retained "
                "until the case is closed.",
            ]),
        ]),
        ("17. Risk Management, Nonconforming Work, and Contingency", [
            ("tbl",
             ["Risk", "Control measure"],
             [
                 ["Wrong-patient result released",
                  f"Identification at collection ({code('venous')}); barcode check at loading; "
                  "delta check before release"],
                 ["Control failure not detected",
                  "Two-level IQC before patient testing; Westgard rules; weekly chart review"],
                 ["Reagent lot shift", "Calibration and both control levels on every new lot "
                  "before patient specimens"],
                 ["Power interruption", "UPS; record the event; repeat controls after power "
                  "restoration before testing"],
                 ["Critical result not communicated",
                  "Telephone with read-back; record; escalate to the Laboratory Manager if "
                  "unreachable"],
                 ["EMR/LIS downtime",
                  "Downtime procedure: report forms completed manually; results telephoned "
                  "with read-back; back-entry within 24 hours of restoration"],
             ],
             [6.0, 11.0]),
            ("p", "Nonconforming work (ISO 15189:2022, 7.5): when a result, control or step "
                  "does not conform — stop, report to the Laboratory Manager, assess the "
                  "clinical impact on results already released, repeat or recollect, and issue "
                  "corrected reports where needed. Record the event and the corrective action; "
                  "evaluate its effectiveness (8.7)."),
        ]),
        ("18. Records", [
            ("tbl",
             ["Record", "Where kept", "Retention"],
             [
                 ["Control results and Levey-Jennings charts", "Quality control register", "As set in the laboratory records policy"],
                 ["Calibration and maintenance logs", "Instrument logbook", "As above"],
                 ["Reagent, calibrator and control lot records", "Stock / reagent records", "As above"],
                 ["Temperature records (2–8 °C refrigerator)", "Temperature log", "As above"],
                 ["EQA reports and actions", "Quality file", "As above"],
                 ["Patient results and reports", "EMR / LIS (electronic)", "As above"],
                 ["Training, competence and authorisation (section 21)", "Personnel training file",
                  "Duration of employment plus the period required by national rules"],
             ],
             [6.5, 5.5, 5.0]),
            ("p", "Records are created at the time the activity is performed, are legible, "
                  "traceable to the person or user account, and are retrievable (ISO 15189:2022, "
                  f"8.4; {ADM_DOC_CONTROL})."),
        ]),
        ("19. Supporting Documents", [
            ("b", [
                f"Document and Record Control ({ADM_DOC_CONTROL}).",
                f"Pre-Examination Procedure ({ADM_PRE_EXAM}).",
                f"Venous Blood Sample Collection ({code('venous')}).",
                "Equipment Management and Maintenance procedure.",
                "Internal Quality Control and EQA procedure.",
                "Laboratory Normal Ranges list (approved 08/07/2025).",
                "Laboratory Critical Results list.",
                "Laboratory Tests TAT list (2026).",
                "Laboratory Medical Equipment Presence List (16/06/2025).",
                "RMS local purchase orders and MediSell Rwanda Ltd delivery notes for "
                "cobas c 111 reagents (PO 2100000988, 2100001448, 2100001746, 2100002205).",
                "Result Reporting and Release procedure.",
                "Laboratory Safety Manual; Laboratory Waste Management procedure.",
            ]),
        ]),
        ("20. References", [
            ("b", [
                "ISO 15189:2022 — Medical laboratories: requirements for quality and competence "
                "(in particular 7.3.2 to 7.3.5, 7.3.7, 7.4.1).",
                "Roche cobas c 111 GLUC2 method sheet and instructions for use (of the lot in "
                "use).",
                "WHO/ADA diagnostic criteria for diabetes mellitus and hypoglycaemia.",
                "CLSI GP41 — Collection of Diagnostic Venous Blood Specimens.",
                "WHO — Laboratory Quality Management System (LQMS) Handbook.",
            ]),
        ]),
    ],
}

DOCUMENTS = [VENOUS, GLUCOSE]

# ---------------------------------------------------------------------------
# DOCX builder
# ---------------------------------------------------------------------------

def build_docx(cfg):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    meta = cfg["meta"]
    doc = Document()

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.header_distance = Cm(0.9)
    sec.footer_distance = Cm(0.9)
    sec.different_first_page_header_footer = True

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.06

    for name, size, color, before, after in (
        ("Heading 1", 13.5, INK, 10, 5),
        ("Heading 2", 11.5, INK2, 8, 3),
    ):
        st = doc.styles[name]
        st.font.name = "Calibri"
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(color.lstrip("#"))
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)

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

    def make_table(headers, rows, widths_cm, *, body_size=9.5, row_height=None,
                   first_col_bold=False):
        table = doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = "Table Grid"
        fixed_layout(table)
        for j, w in enumerate(widths_cm):
            for row in table.rows:
                row.cells[j].width = Cm(w)
        for j, h in enumerate(headers):
            set_cell(table.rows[0].cells[j], h, bold=True, size=9.5,
                     color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
            shade(table.rows[0].cells[j], INK)
        for i, rowvals in enumerate(rows, start=1):
            for j, val in enumerate(rowvals):
                set_cell(table.rows[i].cells[j], val, size=body_size,
                         bold=(first_col_bold and j == 0))
            if row_height:
                table.rows[i].height = Cm(row_height)
        return table

    def make_approval_table():
        table = doc.add_table(rows=1 + len(APPROVAL_ROWS), cols=4)
        table.style = "Table Grid"
        fixed_layout(table)
        for j, w in enumerate([3.0, 5.0, 6.0, 3.0]):
            for row in table.rows:
                row.cells[j].width = Cm(w)
        for j, h in enumerate(["Role", "Name / Designation", "Signature", "Date"]):
            set_cell(table.rows[0].cells[j], h, bold=True, size=9.5,
                     color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
            shade(table.rows[0].cells[j], INK)
        for i, (role, name, desig) in enumerate(APPROVAL_ROWS, start=1):
            row = table.rows[i]
            row.height = Cm(1.7)
            set_cell(row.cells[0], role, bold=True, size=9.5)
            c = row.cells[1]
            c.text = ""
            p1 = c.paragraphs[0]
            p1.paragraph_format.space_after = Pt(1)
            r1 = p1.add_run(name)
            r1.font.size = Pt(9.5)
            r1.font.bold = True
            p2 = c.add_paragraph()
            p2.paragraph_format.space_after = Pt(1)
            r2 = p2.add_run(desig)
            r2.font.size = Pt(8.5)
            r2.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
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

    def tab_stop(paragraph, pos_cm, val="right"):
        pPr = paragraph._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), val)
        tab.set(qn("w:pos"), str(int(pos_cm * 567)))
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

    footer_left = meta["doc_no"] + "  |  Version " + meta["version"]

    def fill_footer(footer):
        p = footer.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(0)
        grey_rgb = RGBColor.from_string(GREY.lstrip("#"))
        top = p.add_run(footer_left)
        top.font.size = Pt(7.5)
        top.font.color.rgb = grey_rgb
        p.add_run("\t")
        pg = p.add_run("Page ")
        pg.font.size = Pt(7.5)
        pg.font.color.rgb = grey_rgb
        add_field(p, "PAGE", size=7.5)
        mid = p.add_run(" of ")
        mid.font.size = Pt(7.5)
        mid.font.color.rgb = grey_rgb
        add_field(p, "NUMPAGES", size=7.5)
        tab_stop(p, 17.0)
        bottom_border(p)
        # second footer line: document control (ISO 15189:2022, 8.3) — the status
        # statement is on every page, so a detached photocopy shows it is uncontrolled
        p2 = footer.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        cn = p2.add_run(COPY_NOTICE)
        cn.font.size = Pt(7)
        cn.font.italic = True
        cn.font.color.rgb = grey_rgb

    hp = sec.header.paragraphs[0]
    hr = hp.add_run(meta["doc_no"] + " — " + meta["title"])
    hr.font.size = Pt(8)
    hr.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
    hp.add_run("\t")
    hr2 = hp.add_run("Version " + meta["version"])
    hr2.font.size = Pt(8)
    hr2.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
    tab_stop(hp, 17.0)
    bottom_border(hp)

    fill_footer(sec.footer)
    fill_footer(sec.first_page_footer)
    sec.first_page_header.paragraphs[0].text = ""

    # cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), height=Cm(2.4))

    for text, style in LETTERHEAD:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(text)
        if style == "lab":
            r.font.size = Pt(15)
            r.font.bold = True
            r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))
        elif style == "kind":
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))
        else:
            r.font.size = Pt(7.5)
            r.font.italic = True
            r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("STANDARD OPERATING PROCEDURE")
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(meta["title"])
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    make_table(["Field", "Detail"], [[k, v] for k, v in cfg["control_rows"]], [5.0, 12.0],
               body_size=9.5, first_col_bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("Approval and Authorisation")
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))
    make_approval_table()

    doc.add_page_break()

    # review history + controlled copies + TOC
    doc.add_heading("Review History", level=1)
    make_table(REVISION_HEADERS, cfg["revision_rows"], [3.0, 1.8, 3.2, 9.0], body_size=8.5)

    doc.add_heading("Controlled Copies", level=1)
    for item in CONTROLLED_COPIES:
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.6)
        bp.paragraph_format.first_line_indent = Cm(-0.4)
        bp.paragraph_format.space_after = Pt(2)
        bp.add_run("•  " + item)

    tp_head = doc.add_paragraph()
    tp_head.paragraph_format.space_before = Pt(10)
    tp_head.paragraph_format.space_after = Pt(5)
    tp_head.paragraph_format.keep_with_next = True
    r = tp_head.add_run("Table of Contents")
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))
    tp = doc.add_paragraph()
    # Level 1 only: a two-level contents does not fit on the front-matter page and would
    # cost a whole extra page. Sub-sections stay reachable from the Word navigation pane
    # and from the PDF bookmarks (see H2 in the PDF builder).
    add_field(tp, 'TOC \\o "1-1" \\h \\z \\u',
              placeholder="Right-click here and choose “Update Field” to build the table "
                          "of contents with page numbers.", size=9.5)

    doc.add_page_break()

    # body 1-20
    for title, blocks in cfg["body"]:
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
                    bp.add_run("•  " + item)
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

    # 21 signature page — holds SIG_ROW_COUNT rows; a section of 12+ staff signs one page
    doc.add_heading(SIG_TITLE, level=1)
    sp = doc.add_paragraph(SIG_STATEMENT)
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    make_table(SIG_HEADERS, SIG_ROWS, SIG_WIDTHS, body_size=9, row_height=0.86)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.add_run(SIG_TRAINER).font.size = Pt(9)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    p.add_run(SIG_COMPETENCE).font.size = Pt(8.5)
    for item in SIG_COMPETENCE_ITEMS:
        ip = doc.add_paragraph()
        ip.paragraph_format.left_indent = Cm(0.6)
        ip.paragraph_format.space_after = Pt(0)
        ip.add_run(item).font.size = Pt(8.5)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.add_run(SIG_AUTHORISED).font.size = Pt(9)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(SIG_NOTE)
    r.font.size = Pt(8)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    doc.add_page_break()

    # 22 amendment page
    doc.add_heading(AMEND_TITLE, level=1)
    sp = doc.add_paragraph(AMEND_STATEMENT)
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    make_table(AMEND_HEADERS, AMEND_ROWS, AMEND_WIDTHS, body_size=9, row_height=1.05)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    r = p.add_run(AMEND_FOOTNOTE)
    r.font.size = Pt(8)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    settings = doc.settings.element
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    settings.append(upd)

    out = SOP_DIR / (cfg["out_stem"] + ".docx")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"DOCX written: {out.relative_to(ROOT)}")

# ---------------------------------------------------------------------------
# PDF builder
# ---------------------------------------------------------------------------

def build_pdf(cfg):
    import struct
    from xml.sax.saxutils import escape

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.platypus import (BaseDocTemplate, Frame, Image, PageBreak,
                                    PageTemplate, Paragraph, Spacer, Table, TableStyle)
    from reportlab.platypus.tableofcontents import TableOfContents

    meta = cfg["meta"]
    ink = colors.HexColor(INK)
    ink2 = colors.HexColor(INK2)
    grey = colors.HexColor(GREY)
    line_grey = colors.HexColor("#BFBFBF")

    styles = {
        "lab": ParagraphStyle("lab", fontName="Helvetica-Bold", fontSize=14,
                              textColor=ink, alignment=1, spaceAfter=4),
        "kind": ParagraphStyle("kind", fontName="Helvetica-Bold", fontSize=11.5,
                               textColor=grey, alignment=1, spaceAfter=2),
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=17,
                                textColor=ink, alignment=1, spaceAfter=10),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=12,
                             textColor=ink, spaceBefore=8, spaceAfter=4),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10.5,
                             textColor=ink2, spaceBefore=7, spaceAfter=3),
        "p": ParagraphStyle("p", fontName="Helvetica", fontSize=9.5, leading=11.9,
                            textColor=colors.black, spaceAfter=4, alignment=4),
        "b": ParagraphStyle("b", fontName="Helvetica", fontSize=9.5, leading=11.9,
                            textColor=colors.black, spaceAfter=2.5,
                            leftIndent=14, firstLineIndent=-10),
        "s": ParagraphStyle("s", fontName="Helvetica", fontSize=9.5, leading=11.9,
                            textColor=colors.black, spaceAfter=2.5,
                            leftIndent=20, firstLineIndent=-20),
        "cell": ParagraphStyle("cell", fontName="Helvetica", fontSize=8.5,
                               leading=10.0, textColor=colors.black),
        "cellb": ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=8.5,
                                leading=10.0, textColor=colors.black),
        "cellhead": ParagraphStyle("cellhead", fontName="Helvetica-Bold", fontSize=8.5,
                                   leading=10.4, textColor=colors.white),
        "small": ParagraphStyle("small", fontName="Helvetica-Oblique", fontSize=8,
                                textColor=grey, leading=10),
        "small9": ParagraphStyle("small9", fontName="Helvetica-Oblique", fontSize=8.5,
                                 textColor=grey, leading=11),
        "form": ParagraphStyle("form", fontName="Helvetica", fontSize=9.5,
                               leading=12.4, textColor=colors.black,
                               alignment=0, spaceAfter=4),
        "sigitem": ParagraphStyle("sigitem", fontName="Helvetica", fontSize=8.5,
                                  leading=11, leftIndent=16, textColor=colors.black),
        "appname": ParagraphStyle("appname", fontName="Helvetica-Bold", fontSize=9.5,
                                  leading=11.5, textColor=colors.black),
        "appdesig": ParagraphStyle("appdesig", fontName="Helvetica", fontSize=8.5,
                                   leading=10.5, textColor=grey),
    }

    PAGE_W, PAGE_H = A4
    FRAME_W = 17.0 * cm
    footer_left = meta["doc_no"] + "  |  Version " + meta["version"]

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
            self.line(2.0 * cm, 1.85 * cm, PAGE_W - 2.0 * cm, 1.85 * cm)
            self.setFont("Helvetica", 7)
            self.setFillColor(grey)
            self.drawString(2.0 * cm, 1.52 * cm, footer_left)
            self.drawRightString(PAGE_W - 2.0 * cm, 1.52 * cm,
                                 f"Page {self._pageNumber} of {total}")
            # document control (ISO 15189:2022, 8.3): every page states its own status,
            # so a detached photocopy cannot be mistaken for the master copy
            self.setFont("Helvetica-Oblique", 6.5)
            self.drawCentredString(PAGE_W / 2.0, 1.02 * cm, COPY_NOTICE)
            self.restoreState()

    def on_page(canv, doc_):
        if doc_.page <= 1:
            return
        canv.saveState()
        canv.setFont("Helvetica", 7.5)
        canv.setFillColor(grey)
        canv.drawString(2.0 * cm, PAGE_H - 1.35 * cm, meta["doc_no"] + " — " + meta["title"])
        canv.drawRightString(PAGE_W - 2.0 * cm, PAGE_H - 1.35 * cm, "Version " + meta["version"])
        canv.setStrokeColor(line_grey)
        canv.setLineWidth(0.5)
        canv.line(2.0 * cm, PAGE_H - 1.55 * cm, PAGE_W - 2.0 * cm, PAGE_H - 1.55 * cm)
        canv.restoreState()

    class SOPDoc(BaseDocTemplate):
        def afterFlowable(self, flowable):
            entry = getattr(flowable, "toc_entry", None)
            if entry is not None:
                # PDF bookmarks make the SOP navigable on screen (paperless operation)
                # without adding a line to the printed table of contents.
                self.canv.bookmarkPage(str(id(flowable)))
                self.canv.addOutlineEntry(entry, str(id(flowable)),
                                          level=getattr(flowable, "toc_level", 0),
                                          closed=True)
                # only level-0 headings go in the printed TOC: the front matter is
                # designed to fill exactly one page
                if getattr(flowable, "toc_level", 0) == 0:
                    self.notify("TOCEntry", (0, entry, self.page))

    def P(text, style="p"):
        return Paragraph(escape(text), styles[style])

    def H1(text):
        para = Paragraph(escape(text), styles["h1"])
        para.toc_entry = text
        para.toc_level = 0
        return para

    def H2(text):
        para = Paragraph(escape(text), styles["h2"])
        para.toc_entry = text
        para.toc_level = 1
        return para

    def make_pdf_table(headers, rows, widths_cm, *, first_col_bold=False, row_heights=None):
        data = [[Paragraph(escape(str(h)), styles["cellhead"]) for h in headers]]
        for row in rows:
            cells = []
            for j, val in enumerate(row):
                st = "cellb" if (first_col_bold and j == 0) else "cell"
                cells.append(Paragraph(escape(str(val)), styles[st]))
            data.append(cells)
        t = Table(data, colWidths=[w * cm for w in widths_cm],
                  rowHeights=row_heights, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(INK)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, line_grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    def approval_pdf_table():
        data = [[Paragraph(escape(h), styles["cellhead"])
                 for h in ["Role", "Name / Designation", "Signature", "Date"]]]
        for role, name, desig in APPROVAL_ROWS:
            cell = Paragraph(escape(name) + "<br/>" + escape(desig), styles["appname"])
            data.append([Paragraph(escape(role), styles["cellb"]), cell, "", ""])
        t = Table(data, colWidths=[3.0 * cm, 5.0 * cm, 6.0 * cm, 3.0 * cm],
                  rowHeights=[None] + [1.7 * cm] * len(APPROVAL_ROWS))
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(INK)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, line_grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    def logo_flowable(height_cm=2.2):
        if not LOGO_PATH.exists():
            return Spacer(1, 0.1 * cm)
        raw = LOGO_PATH.read_bytes()
        w, h = struct.unpack(">II", raw[16:24])
        img = Image(str(LOGO_PATH), width=height_cm * (w / h) * cm, height=height_cm * cm)
        img.hAlign = "CENTER"
        return img

    story = [
        Spacer(1, 0.4 * cm),
        logo_flowable(2.2),
        Spacer(1, 0.25 * cm),
        *[P(t, st) for t, st in LETTERHEAD],
        P("STANDARD OPERATING PROCEDURE", "kind"),
        P(meta["title"], "title"),
        make_pdf_table(["Field", "Detail"],
                       [[k, v] for k, v in cfg["control_rows"]],
                       [5.0, 12.0], first_col_bold=True),
        Spacer(1, 0.8 * cm),
        P("Approval and Authorisation", "h2"),
        approval_pdf_table(),
        PageBreak(),
    ]

    story.append(H1("Review History"))
    story.append(make_pdf_table(REVISION_HEADERS, cfg["revision_rows"],
                                [3.0, 1.8, 3.2, 9.0]))
    story.append(H1("Controlled Copies"))
    for item in CONTROLLED_COPIES:
        story.append(P("•  " + item, "b"))
    story.append(Paragraph(escape("Table of Contents"), styles["h1"]))
    toc = TableOfContents()
    # Single level (sections 1-22) with tight leading: the front matter is designed to
    # fill exactly one page. Sub-sections are in the PDF outline/bookmarks instead.
    toc.levelStyles = [ParagraphStyle("toc0", fontName="Helvetica", fontSize=9.5,
                                      leading=10.9, textColor=colors.black)]
    story.append(toc)
    story.append(PageBreak())

    for title, blocks in cfg["body"]:
        story.append(H1(title))
        for block in blocks:
            kind = block[0]
            if kind == "p":
                story.append(P(block[1]))
            elif kind == "b":
                for item in block[1]:
                    story.append(P("•  " + item, "b"))
            elif kind == "h2":
                story.append(H2(block[1]))
            elif kind == "steps":
                for label, text in block[1]:
                    story.append(P(label + "   " + text, "s"))
            elif kind == "tbl":
                _, headers, rows, widths = block
                story.append(make_pdf_table(headers, rows, widths))
                story.append(Spacer(1, 0.15 * cm))

    story.append(PageBreak())
    story.append(H1(SIG_TITLE))
    story.append(P(SIG_STATEMENT))
    story.append(Spacer(1, 0.12 * cm))
    story.append(make_pdf_table(SIG_HEADERS, SIG_ROWS, SIG_WIDTHS,
                                row_heights=[None] + [0.86 * cm] * len(SIG_ROWS)))
    story.append(Spacer(1, 0.34 * cm))
    story.append(P(SIG_TRAINER, "form"))
    story.append(Spacer(1, 0.10 * cm))
    story.append(Paragraph(escape(SIG_COMPETENCE), styles["small9"]))
    for item in SIG_COMPETENCE_ITEMS:
        story.append(Paragraph(escape(item), styles["sigitem"]))
    story.append(Spacer(1, 0.14 * cm))
    story.append(P(SIG_AUTHORISED, "form"))
    story.append(Spacer(1, 0.12 * cm))
    story.append(P(SIG_NOTE, "b"))

    story.append(PageBreak())
    story.append(H1(AMEND_TITLE))
    story.append(P(AMEND_STATEMENT))
    story.append(Spacer(1, 0.12 * cm))
    story.append(make_pdf_table(AMEND_HEADERS, AMEND_ROWS, AMEND_WIDTHS,
                                row_heights=[None] + [1.05 * cm] * len(AMEND_ROWS)))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(escape(AMEND_FOOTNOTE), styles["small9"]))

    out = SOP_DIR / (cfg["out_stem"] + ".pdf")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = SOPDoc(str(out), pagesize=A4,
                 leftMargin=2.0 * cm, rightMargin=2.0 * cm,
                 topMargin=2.0 * cm, bottomMargin=2.0 * cm,
                 title=meta["doc_no"] + " — " + meta["title"],
                 author=LAB, subject="ISO 15189:2022 — 2026 standard layout")
    frame = Frame(doc.leftMargin, doc.bottomMargin, FRAME_W, PAGE_H - 4.0 * cm, id="main")
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])
    doc.multiBuild(story, canvasmaker=NumberedCanvas)
    print(f"PDF written:  {out.relative_to(ROOT)}  ({doc.page} pages)")


def sweep_stale_outputs():
    """Remove built files that no longer correspond to a registered document.

    A renumbered SOP leaves its old PDF and DOCX behind under the old code, and
    a file named with a retired code is exactly the "document still referencing
    the old number" that document control has to prevent. Any file in sop/ whose
    name does not start with a code in REGISTER is deleted.
    """
    live = {code(k) for k in REGISTER}
    removed = []
    for path in sorted(SOP_DIR.iterdir()):
        if not path.is_file() or path.suffix.lower() not in {".pdf", ".docx"}:
            continue
        stem_code = path.name.split(" - ")[0]
        if stem_code not in live:
            path.unlink()
            removed.append(path.name)
    for name in removed:
        print(f"Removed stale output: {name}")


if __name__ == "__main__":
    sweep_stale_outputs()
    for cfg in DOCUMENTS:
        build_docx(cfg)
        build_pdf(cfg)
    print()
    print("Register in force:")
    for key, r in REGISTER.items():
        flag = "  (new serial)" if r.get("new") else ""
        print(f"  {code(key):28s} {r['title']}{flag}")
