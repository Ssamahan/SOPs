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
    # Corrected against the 2026 register (MASTER LIST - 2024 - 2025.docx,
    # modified 17/02/2026). The ADM-SOP serials it actually lists are 001, 002,
    # 009, 020, 024 and 034. The older MASTER LIST 2025.docx ran to 049, but it
    # is not the current register. See docs/register/tb-ept-register-transcript.md.
    "RliDHL-ADM-SOP": 34,
    "RliDHL-ADM-MAN": 2,    # 001 Quality Control Manual, 002 Safety Manual
    "RliDHL-ADM-GUID": 3,   # 001 Scope of Test, 002 Critical Results, 003 Normal Ranges
    "RliDHL-BAC-SOP": 6,
    "RliDHL-BIO-SOP": 9,
    "RliDHL-HEM-SOP": 6,   # 005 XN-330, 006 XQ-320
    "RliDHL-IMM-SOP": 5,   # 005 Pima CD4
    # MYC was renamed TB on 15/11/2026 by the Laboratory Manager: the section is
    # the tuberculosis section and "MYC" was a mycobacteriology label the register
    # used for cells that are TB work. The serials do NOT move with the name --
    # MYC 001-012 become TB 001-012, same numbers, same documents, same dates.
    # A rename is not a renumbering: no serial is freed, skipped or reused.
    "RliDHL-TB-SOP": 12,
    # MOL is new. It exists because HIV early infant diagnosis and the hepatitis
    # B and C viral loads are molecular examinations and were being filed in a
    # mycobacteriology series that is now named for TB. MOL starts at 001.
    # These three documents are MOVED, not renumbered: MOL-002/003/004 keep their
    # serials, because a document that changes series has not changed identity.
    "RliDHL-MOL-SOP": 4,
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
    # These were open questions. They are now settled by fact, and the facts are kept
    # here because they are the reason several documents are withdrawn or marked
    # suspended, and because an auditor will ask for them.
    #
    # NOT open any more: whether the legacy analysers are in service. They were
    # decommissioned in August 2026 and the Decommissioned section of the programme
    # handoff is authoritative. The equipment spreadsheet "RDH EQUIPMENTS CODE25 JANV
    # 2026.xlsx" is stale -- it still shows the Humastar 80 as Operational -- and must
    # not be used to decide equipment status.
    #
    # Equipment status per the maintenance record and the Laboratory Manager:
    #   IN SERVICE   Sysmex XN-330 (17156), Sysmex XQ-320 (12357), Alere Pima CD4
    #   AWAITING REPAIR   cobas c 111 serial 65053, RULIDH/LAB/ME/0486 -- the hospital's
    #                own instrument, out of service since April 2026 (Peltier element,
    #                Error 5040), shipped to MediSell Rwanda Ltd and expected back.
    #   ON LOAN      a cobas c 111 supplied by MediSell while the above is repaired.
    #                Currently DOWN itself since 11/11/2026 (Ready-mode failure).
    #   DECOMMISSIONED (August 2026)   Sysmex KX-21N, XS-500i, Pochi, Humacount 5,
    #                Humastar 80, Humalyzer 3500 -- see WITHDRAWN.
    #
    # Consequence for this repository: BIO-SOP-001 to BIO-SOP-009 are all Humastar 80
    # procedures and are withdrawn; the glucose SOP is valid in content but its
    # examination cannot currently be performed on either cobas c 111.
    # The cobas c 111 outage is recorded here, in the register and in the equipment record,
    # and NOT inside the glucose SOP. A loan analyser is a temporary arrangement; an SOP is
    # a controlled document that outlives it. Naming a borrowed instrument in the procedure
    # guarantees the procedure is wrong the week the unit goes back, so the SOP describes
    # the laboratory's own analyser and section 17 handles the outage.
    "EQUIPMENT_NOTE_cobas_c111": {
        "issue": "The cobas c 111 is out of service — an equipment fault, not a document "
                 "question. No SOP decision is waiting on it.",
        "detail": "The hospital's own cobas c 111 (serial 65053, asset RULIDH/LAB/ME/0486) "
                  "failed in April 2026 with a Peltier element fault (Error 5040) and was "
                  "shipped to MediSell Rwanda Ltd for repair; it is expected back. The unit "
                  "supplied in the meantime is on loan and failed on 11/11/2026: it will not "
                  "enter Ready mode, reporting a Transfer Head Fluid Sensor Calibration Error "
                  "and a Level Detection Frequency Error, persistent after every corrective "
                  "step in the service manual. Biochemistry testing is at a standstill.",
        "options": "This needs the service visit, the loan agreement and the return of the "
                   "hospital's instrument — none of it is document control, and none of it "
                   "belongs in an SOP. The glucose SOP states the contingency (suspension, "
                   "backup method named on the report, referral) without naming any borrowed "
                   "equipment. When the repaired instrument returns, verify it under 7.3.2 "
                   "before use; no SOP revision is needed for that, because the SOP never "
                   "referred to the loan unit.",
    },
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

# Equipment that is not decommissioned but is not working either. This is a MANAGEMENT
# record: it belongs in the equipment register and the nonconformity file, not in an SOP.
# The checker enforces that no SOP names borrowed equipment at all.
AWAITING_REPAIR = [
    ("Roche cobas c 111", "65053", "RULIDH/LAB/ME/0486",
     "hospital's own unit — Peltier element failure, Error 5040, out of service since "
     "April 2026, shipped to MediSell Rwanda Ltd"),
    ("Roche cobas c 111", "not recorded", "no asset tag issued",
     "MediSell loan unit in use while the above is repaired — itself down since "
     "11/11/2026, will not enter Ready mode"),
]

# Every serial already spoken for beyond the registered set, so a duplicate is
# caught before it ships. This is the list that would have caught ADM-060.
SERIES_ALLOCATED = {
    "RliDHL-ADM-SOP": {
        # ---- ADM SOP serials, corrected against the 2026 register --------------
        #
        # The previous version of this table claimed every serial 1-49 except 035
        # was registered. That was read from MASTER LIST 2025.docx (15/04/2025),
        # which is NOT the current register. The register the laboratory actually
        # works to is MASTER LIST - 2024 - 2025.docx, modified 17/02/2026, and it
        # lists only these ADM-SOP serials:
        #
        #     001  002  009  020  024  034
        #
        # Everything else between 001 and 034 is ABSENT from the register -- not
        # proven free, but not proven taken either. The full transcription and the
        # list of defects are in docs/register/tb-ept-register-transcript.md.
        #
        # Serial 035 is retired: the older register skipped from 034 to 036 and no
        # document ever held it (SOP-PROGRAMME-HANDOFF). It stays retired.
        1: "ADM SOP 001 Laboratory Temperature Control (2026 register; DOCUMENT AND "
           "RECORD CONTROL is also printed on 001 -- a register defect)",
        2: "ADM SOP 002 Documentation of Examination Procedure (2026 register)",
        9: "ADM SOP 009 Disposal of sharps and needles (2026 register)",
        20: "ADM SOP 020 Pre-Examination Procedure (2026 register)",
        24: "ADM SOP 024 Quality Control and Quality Assurance (2026 register)",
        34: "ADM SOP 034 Laboratory Waste Management, and Handling of biological spills "
            "(2026 register; TWO documents share this serial -- a register defect)",
        **{n: "ADM SOP 035 retired -- never used, no document ever held it"
           for n in (35,)},
        # the November 2026 equipment block: built, uploaded, unsigned
        # 050 is NOT listed: it was released from the unsigned Microscope draft and
        # re-allocated to the health-centre re-reading SOP, which is in REGISTER as a
        # new allocation. Listing it here as well would be a double allocation.
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
    "RliDHL-HEM-SOP": {
        1: "HEM SOP 001 Venous Blood Sample Collection (this library, VERS-002)",
        2: "HEM SOP 002 Sysmex Pochi 100i — WITHDRAWN, analyser decommissioned",
        3: "HEM SOP 003 Humacount 5 — WITHDRAWN, analyser decommissioned",
        4: "HEM SOP 004 Sysmex KX-21N — WITHDRAWN, analyser decommissioned",
    },
    "RliDHL-IMM-SOP": {
        1: "IMM SOP 001 FACSount usage and maintenance (legacy, register)",
        2: "IMM SOP 002 CD4 sample rejection (legacy, register)",
        3: "IMM SOP 003 CD4 sample preparation on FACSCount (legacy, register)",
        4: "IMM SOP 004 Enumeration of CD4+ T lymphocytes, FACSCount (legacy, register)",
    },
    # MYC 001-012 -> TB 001-012, same serials. The register's Mycobacteriology
    # section is the TB section; the prefix is being renamed to say so.
    # The TB section holds ALL tuberculosis testing, whatever the method: the
    # bacilloscopy (ZN and auramine), the culture and DST work, and the GeneXpert
    # TB assays. TB LAM belongs here too, for the same reason -- it is a TB test
    # performed in the TB section. A test's METHOD does not decide its series;
    # the section that performs it does. That is why GeneXpert work is split
    # between TB (the TB assays) and MOL (the viral loads) even though it is the
    # same instrument, and why HIV *qualitative* serology stays in SER and does
    # not follow the instrument into MOL.
    "RliDHL-TB-SOP": {
        1: "TB SOP 001 Sputum collection, preservation and transport using CPC "
           "(was MYC SOP 001 — register, renamed 15/11/2026)",
        2: "TB SOP 002 TB specimen rejection criteria for microscopy, culture and DST "
           "(was MYC SOP 002)",
        3: "TB SOP 003 TB microscopy by hot Ziehl-Neelsen staining for quality control "
           "(was MYC SOP 003)",
        4: "TB SOP 004 Fluorescence TB microscopy by auramine staining (was MYC SOP 004)",
        5: "TB SOP 005 TB sample transportation (was MYC SOP 005)",
        6: "TB SOP 006 Internal quality control for reagents and techniques in the TB "
           "laboratory (was MYC SOP 006)",
        7: "TB SOP 007 Turnaround time for TB tests at RliDH and the NRL (was MYC SOP 007)",
        8: "TB SOP 008 Quality control for TB microscopy slide re-testing (was MYC SOP 008)",
        9: "TB SOP 009 — register row has NO TITLE recorded (was MYC SOP 009). The "
           "serial is allocated; the subject is unrecorded. This is a register defect, "
           "recorded not fixed.",
        10: "TB SOP 010 Sample processing using GeneXpert (was MYC SOP 010)",
        11: "TB SOP 011 Data archiving and preventive maintenance of the GeneXpert Dx "
            "system (was MYC SOP 011)",
        12: "TB SOP 012 Sample processing using GeneXpert (was MYC SOP 012 — duplicate "
            "subject with 010; a register defect, recorded not fixed)",
    },
    # MOL takes the MOLECULAR examinations: HIV early infant diagnosis and the
    # hepatitis B and C viral loads. The three documents below already carry
    # these serials; the move out of MYC does not renumber them.
    #
    # Serial 001 is HELD, not free. RliDHL-MOL-SOP-001-VERS-003, "HIV-1
    # Qualitative Testing Using Xpert HIV-1 Assays", held it; MOL-002-VERS-002
    # records that it "replaces" that document. A withdrawn document keeps its
    # serial for life, so 001 cannot be allocated.
    #
    # HIV SEROLOGY IS NOT IN MOL AND IS NOT TOUCHED BY THIS MOVE. Rapid
    # Determine, Unigold, capillary collection for HIV testing and HIV QC
    # transport are SEROLOGY, and the register files them under SER (SER-001,
    # -002, -004, -005). Those SOPs are unaffected by anything here. The
    # distinction is method, not analyte: an antibody rapid test is serology; an
    # assay that amplifies nucleic acid is molecular. The one HIV document that
    # WAS molecular was the Xpert Qual assay, and it is withdrawn.
    "RliDHL-MOL-SOP": {
        1: "MOL SOP 001 HIV-1 Qualitative Testing using Xpert HIV-1 Assays — "
           "WITHDRAWN, absorbed into MOL-002-VERS-002. Serial held for life.",
        2: "MOL SOP 002 HIV Early Infant Diagnosis using GeneXpert (this library)",
        3: "MOL SOP 003 Hepatitis B Virus Viral Load using GeneXpert",
        4: "MOL SOP 004 Hepatitis C Virus Viral Load using GeneXpert",
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
    "xn330": {
        # HEM serials: 001 = venipuncture revision, 002/003/004 = withdrawn analysers.
        # 005 is the first genuinely free one.
        "code": "RliDHL-HEM-SOP-005", "vers": "001",
        "title": "Examination of Blood by the Sysmex XN-330 Haematology Analyser",
        "master_list": "Master List, Haematology",
        "new": True,
    },
    "xq320": {
        "code": "RliDHL-HEM-SOP-006", "vers": "001",
        "title": "Examination of Blood by the Sysmex XQ-320 Haematology Analyser",
        "master_list": "Master List, Haematology",
        "new": True,
    },
    "pima": {
        # CD4 counting is immunology. IMM 001-004 are the legacy FACSCount SOPs;
        # 005 is the first free serial.
        "code": "RliDHL-IMM-SOP-005", "vers": "001",
        "title": "CD4 Enumeration by the Alere Pima CD4 Analyser",
        "master_list": "Master List, Immunology",
        "new": True,
    },
    "exam_docs": {
        "code": "RliDHL-ADM-SOP-041", "vers": "001",
        "new": True,
        "title": "Documentation of examination procedures",
        "master_list": "Master List, Crosscutting SOPs",
    },
    "equip_maint": {
        "code": "RliDHL-ADM-SOP-062", "vers": "001",
        "title": "Use, Care and Preventive Maintenance of Hematology Analysers",
        "master_list": "Master List, crosscutting (Nov 2026 equipment block)",
    },
    # The GeneXpert maintenance SOP is already in the Nov 2026 equipment block
    # (ADM-SOP-063) and is cited by the troubleshooting SOP. Registering it here
    # lets it be referenced by code() rather than typed.
    "genexpert_maint": {
        "code": "RliDHL-ADM-SOP-063", "vers": "001",
        "title": "Use, Care and Preventive Maintenance of the GeneXpert Dx System",
        "master_list": "Master List, crosscutting (Nov 2026 equipment block)",
    },
    # --- documents cited by the EQA/PT SOP, registered so they are cited by
    # code rather than typed. All are in the 2026 register.
    "qca": {
        "code": "RliDHL-ADM-SOP-024", "vers": "004",
        "title": "Quality Control and Quality Assurance",
        "master_list": "Master List, Crosscutting SOPs",
    },
    "scope_test": {
        "code": "RliDHL-ADM-GUID-001", "vers": "004",
        "title": "Laboratory Scope of Test",
        "master_list": "Master List, Guidelines",
    },
    "qc_manual": {
        "code": "RliDHL-ADM-MAN-001", "vers": "006",
        "title": "Laboratory Quality Control Manual",
        "master_list": "Master List, Manuals",
    },
    "safety_manual": {
        "code": "RliDHL-ADM-MAN-002", "vers": "006",
        "title": "Laboratory Safety Manual",
        "master_list": "Master List, Manuals",
    },
    # Serial 050. The November 2026 equipment block opened 050 (Microscope), 051
    # (Centrifuge) and 052 (Water Bath); those three are built and uploaded but
    # UNSIGNED, so they are not in force. 050 has therefore been released back to
    # the queue and re-used here -- see SERIES_ALLOCATED. Nothing was ever issued
    # under 050, so no document loses its number.
    "hc_slide_review": {
        "code": "RliDHL-ADM-SOP-050", "vers": "001",
        "title": "Re-reading of Health-Centre Slides and Re-Testing of Health-Centre Samples",
        "master_list": "Master List, Crosscutting SOPs",
        "new": True,
    },
    # Serial 025. The 2026 register (MASTER LIST - 2024 - 2025.docx, modified
    # 17/02/2026) holds ADM-024 (Quality Control and Quality Assurance) and
    # ADM-020 (Pre-Examination Procedure) and nothing between them. The
    # Laboratory Manager's instruction was to fill the gap rather than skip past
    # it, so this SOP takes the free serial 025 instead of a number past 065.
    "eqa_pt": {
        "code": "RliDHL-ADM-SOP-025", "vers": "001",
        "title": "External Quality Assessment and Proficiency Testing",
        "master_list": "Master List, Crosscutting SOPs",
        "new": True,
    },
    # Serial 065. The gap 064 -> 065 is not a gap: 064 is taken by the SOP on the
    # signing of SOPs, which this library issues. 065 is the next free serial.
    #
    # One document for every analyser, not one per analyser. The failure modes the
    # manufacturer names are mostly common to a family (the Sysmex pair share a
    # software generation), and four copies of the same table would drift apart
    # within a year -- which is the defect this library was built to stop.
    "trouble": {
        "code": "RliDHL-ADM-SOP-065", "vers": "001",
        "title": "Troubleshooting of laboratory analysers: error codes, causes and "
                 "corrective action",
        "master_list": "Master List, crosscutting (Nov 2026 equipment block)",
        "new": True,
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
            ("p", "Ruli District Hospital Laboratory measures glucose in serum and plasma on its "
                  "Roche cobas c 111 analyser, serial 65053, asset RULIDH/LAB/ME/0486, using the "
                  "Glucose HK (GLUC2) reagent. Requests and results are kept in the EMR. When "
                  "that analyser is not available, biochemistry follows the contingency in "
                  "section 17; this SOP describes the examination on the laboratory's own "
                  "instrument and on nothing else."),
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
            ("p", "This SOP applies to all laboratory biotechnologists who operate the "
                  "laboratory's cobas c 111 and perform glucose testing in the biochemistry "
                  "section. It covers "
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
                "Roche cobas c 111 analyser, Biochemistry Room, serial 65053, hospital tag "
                "RULIDH/LAB/ME/0486 — Laboratory Medical Equipment Presence List and "
                "instrument logbook. This SOP is written for that instrument.",                "Technical service and installation (electrodes, interface, corrective "
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
                ("1.", "Confirm that the analyser is the instrument identified in 8.1 and "
                       "that it is not under a withdrawal from service (section 17). Check that "
                       "it and the UPS are connected, switch on, and wait for the self-check. "
                       "An analyser that will not reach Ready mode is out of service: stop and "
                       "report it."),
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
                 ["Analyser withdrawn from service, or not in Ready mode",
                  "Biochemistry is suspended and the suspension is reported to the Laboratory "
                  "Manager and the Director General. Critical glucose is measured with the "
                  "backup method recorded in the contingency plan (glucometer with a "
                  "documented, traceable method) and the method used is stated on every "
                  "report. Requests that cannot be met are referred to the National "
                  "Reference Laboratory or the nearest district hospital laboratory, with "
                  "the referral recorded. Do not report a result from an instrument that is "
                  "not in Ready mode (ISO 15189:2022, 7.5, 7.8)"],
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

# ---------------------------------------------------------------------------
# Documents 3-5 — examination SOPs for the analysers actually in service
#
# ISO 15189:2022, 7.3.6 a) requires the laboratory to document its examination
# procedures "to the extent necessary to ensure the consistent application of its
# activities and the validity of its results". 7.3.6 d) allows the manufacturer's
# instructions for use to be incorporated BY REFERENCE where they carry sufficient
# information. That is what is done here: the analyser's own parameter list,
# measuring ranges, flagging rules and interference limits are cited to the IFU for
# the model in use rather than retyped from memory into an uncontrolled SOP. What is
# written out is what the IFU cannot know: this laboratory's specimen acceptance,
# workflow, QC rules, critical limits and escalation.
#
# Nothing in these documents states a performance figure (CV, measuring range,
# reference interval) that has not been verified for the instrument on the bench.
# 7.3.2 requires those to be verified and recorded in the method verification file,
# and 7.3.5 requires the reference intervals to be established and reviewed. The SOP
# points at those records; it does not invent them.
# ---------------------------------------------------------------------------

ADM_EXAM_DOCS = code("exam_docs")        # ADM SOP 041, Documentation of examination procedures
# 041 is a NEW allocation by this library; the 2026 register does not carry it.
ADM_EQUIP_MAINT = code("equip_maint")    # ADM SOP 062, Hematology analyser maintenance


def analyser_sop(key, *, model, manufacturer, serial, asset, bench, discipline,
                 clinical, principle, specimen, specimen_notes, startup,
                 run_steps, results, interp_intro, interp_rows, critical,
                 reports, risks, limitations, ifu, extra_support, extra_refs,
                 replacement_note):
    """Build one examination SOP, sharing everything that does not depend on the model.

    The instrument-specific sections are passed in; the sections that are the same for
    any analyser — responsibilities, environment, records, waste, contingency — are
    written once here so five documents cannot drift apart.
    """
    c = code(key)
    body = []

    body.append(("1. Introduction", [
        ("p", f"{clinical}"),
        ("p", f"At Ruli District Hospital Laboratory this examination is performed on the "
              f"{manufacturer} {model}, {bench}, asset code {asset}, serial number {serial}. "
              f"Requests and results are kept in the EMR. {replacement_note}"),
    ]))

    body.append(("2. Principle of the Test", [("p", principle)]))

    body.append(("3. Purpose", [
        ("p", f"This SOP standardises the examination of {discipline.lower()} specimens on the "
              f"{model} so that results are correct, comparable over time, and traceable to the "
              f"person who produced and released them."),
    ]))

    body.append(("4. Scope", [
        ("p", f"This SOP applies to every member of laboratory staff authorised under section 7 "
              f"who operates the {model} at the main laboratory of Ruli District Hospital. It "
              f"covers patient specimens and, where stated, quality control and EQA material. It "
              f"does not cover the maintenance of the analyser, which is controlled by a separate "
              f"procedure (section 19), nor the collection of the specimen, which is covered by "
              f"the pre-examination procedure."),
    ]))

    body.append(("5. Definitions and Abbreviations", [
        ("b", [
            "CBC — complete blood count: the panel of cell counts and derived indices reported "
            "by a haematology analyser.",
            "IQC — internal quality control; EQA — external quality assessment.",
            "MU — measurement uncertainty (ISO 15189:2022, 7.3.4). IFU — the manufacturer's "
            "instructions for use for the analyser and the reagent lot in use.",
            "Delta check — comparison of a result with the patient's previous result.",
            f"Authorised person — a member of staff recorded as competent and authorised "
            f"under {ADM_TRAINING}.",
        ]),
    ]))

    body.append(("6. Specimen Requirements", specimen))

    body.append(("7. Responsibilities and Authorisation", [
        ("b", [
            "The Laboratory Manager (Kelvin SAMAHANI) owns this SOP, ensures the analyser is "
            "verified and maintained, and authorises staff to report results unsupervised.",
            "The Laboratory Quality Manager (Jean Paul NSENGIYUMVA) reviews control performance, "
            "EQA results and any nonconformity arising from this examination.",
            "The BMeT (Pacifique ISHIMWE) performs and records the maintenance and repair of the "
            "analyser under the equipment maintenance procedure.",
            "Only staff recorded as competent and authorised on the signature page of this SOP "
            "may operate the analyser unsupervised. Staff in training operate it under direct "
            "supervision and do not release results (ISO 15189:2022, 7.2).",
            "The person named against a result in the EMR is accountable for its verification and "
            "release (ISO 15189:2022, 7.3.1 d).",
        ]),
    ]))

    body.append(("8. Equipment, Reagents, and Consumables", [
        ("tbl",
         ["Item", "Detail"],
         [
             ["Analyser", f"{manufacturer} {model}"],
             ["Asset code", asset],
             ["Serial number", serial],
             ["Location", bench],
             ["Reagents and controls", "As specified by the manufacturer for the model; the lot "
              "in use, its expiry and its on-board stability are those stated in the IFU and "
              "recorded in the reagent log."],
             ["Consumables", "Sample tubes, cuvettes, cleaning and rinse solutions specified for "
              "the model, and printer/interface consumables. Use only consumables specified for "
              "the instrument."],
             ["Power protection", "The analyser is connected through the laboratory UPS. A power "
              "event is recorded and controls are repeated before testing resumes."],
         ],
         [5.2, 11.8]),
        ("p", "Before use, confirm the analyser carries a valid maintenance status and that no "
              f"withdrawal is in force ({ADM_EQUIP_MAINT})."),
    ]))

    body.append(("9. Environment and Biosafety Measures", [
        ("b", [
            "Work at the bench with gloves, a laboratory coat and eye protection. Treat all "
            "specimens as potentially infectious.",
            "Keep the analyser on a level bench away from direct sunlight, heat sources and strong "
            "air currents, with the clearance the IFU specifies for ventilation.",
            "Have the spill kit available. A spill is cleaned with the laboratory disinfectant "
            "and recorded; a spill into the analyser is handled as a fault in section 17.",
            "Sharps and contaminated consumables go into the sharps container and the "
            "infectious waste stream, never into general waste. Decontaminate the bench at the "
            "end of the shift.",
        ]),
    ]))

    body.append(("10. Procedural Steps", [
        ("h2", f"10.1 Start-up before the first patient specimen"),
        ("steps", startup),
        ("h2", "10.2 Running a specimen"),
        ("steps", run_steps),
        ("h2", "10.3 Reading and accepting the result"),
        ("b", results),
    ]))

    body.append(("11. Quality Control and External Quality Assessment", [
        ("h2", "11.1 Internal quality control"),
        ("b", [
            "Run the control material specified for this analyser at the start of every day of "
            "testing, after calibration, after a reagent lot change, and after maintenance or "
            "repair.",
            "Accept the run only when each control lies within its assigned limits. Apply the "
            "laboratory's rejection rules: a single result outside ±3 SD, two consecutive "
            "results beyond ±2 SD on the same side of the mean, or a range exceeding 4 SD between "
            "levels is a rejection.",
            "Plot every control result so that shifts and trends are visible, and review the "
            "charts at defined intervals (ISO 15189:2022, 7.3.7.2 e).",
            "Record control results, lot numbers and every action taken in the quality control "
            "register.",
        ]),
        ("h2", "11.2 External quality assessment"),
        ("p", "The laboratory takes part in the haematology EQA scheme organised or designated by "
              "RBC/NRL and tests the samples exactly as patient specimens, by the staff who "
              "normally perform the examination. Unsatisfactory performance is investigated as in "
              "11.3 and reported to the Laboratory Quality Manager."),
        ("h2", "11.3 When a control fails"),
        ("p", "Stop patient testing. Identify the cause — control material, reagent, calibration, "
              "analyser or specimen — and correct it before repeating both control levels. Then "
              "assess whether any result released since the last acceptable control is affected; "
              "repeat the affected specimens and issue corrected reports where needed "
              "(ISO 15189:2022, 7.5, 7.4.1). Record the event, the action and the assessment of "
              "clinical impact."),
        ("h2", "11.4 Comparability"),
        ("p", f"The laboratory operates more than one haematology analyser. Comparability of "
              f"results between them is established with patient specimens across the clinically "
              f"significant interval — not with control material alone — when an analyser is "
              f"added or returned to service, and is reviewed at least once a year. Results and "
              f"the acceptability decision are recorded. Where a clinically significant difference "
              f"is found, its effect on the reference intervals in section 14 is evaluated and "
              f"users are informed (ISO 15189:2022, 7.3.7.4)."),
    ]))

    body.append(("12. Result Interpretation and Validation", [
        ("p", interp_intro),
        ("tbl", ["Observation", "Interpretive note (common causes)"], interp_rows, [5.0, 12.0]),
        ("p", "Interpretation supports the clinician; it does not replace the clinical picture. "
              "Where the result and the clinical information disagree, telephone the requester "
              "before release and record the conversation."),
    ]))

    body.append(("13. Performance Characteristics and Limitations", [
        ("p", "The performance specifications of the method — measuring interval, precision, "
              "trueness and specificity including interfering substances — are those confirmed "
              "for this instrument during verification and recorded in the method verification "
              "file (ISO 15189:2022, 7.3.2). The manufacturer's stated specifications, and the "
              "parameter list with the analyser's flagging and interpretive rules, are held in "
              "the IFU for the model in use and are incorporated by reference (7.3.6 d)."),
        ("b", limitations),
        ("p", "Measurement uncertainty is estimated and recorded for this examination under "
              "ISO 15189:2022, 7.3.4, and is made available to users on request. Where the "
              "laboratory cannot estimate MU for a parameter, the rationale is recorded in the "
              "same file."),
    ]))

    body.append(("14. Reference Intervals, Clinical Decision Limits, and Critical Results", [
        ("p", "The biological reference intervals in use are those of the Laboratory Normal "
              "Ranges list approved for Ruli District Hospital Laboratory, established and "
              "reviewed under ISO 15189:2022, 7.3.5. They are not restated here, so that a "
              "revision of the ranges cannot leave this SOP contradicting them; the list in "
              "force hangs at the bench and is available to users."),
        ("p", "Critical results are those on the Laboratory Critical Results list. Any result "
              "within a critical limit is telephoned to the requester as soon as it is verified, "
              "using the mandatory Write Down and Read Back process, and the communication is "
              "recorded with date, time, the person notified and the result conveyed "
              "(ISO 15189:2022, 7.4.1.3)."),
        ("b", critical),
    ]))

    body.append(("15. Reporting and Release of Results", [
        ("p", reports),
        ("b", [
            "Report patient identifiers, the examination, the result with its unit, the reference "
            "interval in force for this laboratory, and any comment or flag the analyser raised.",
            "Corrected reports: issue an amended report that identifies the change, the reason and "
            "the date; notify the requester; retain the original entry (ISO 15189:2022, 7.4.1.7).",
            "Patient information is confidential (ISO 15189:2022, 4.2).",
        ]),
    ]))

    body.append(("16. Specimen Retention and Disposal", [
        ("b", [
            "After reporting, keep the specimen for the period set in the laboratory records "
            "policy, under the storage conditions the IFU specifies, in case a repeat or an "
            "investigation is needed.",
            "Specimens under investigation for a nonconformity or complaint are retained until the "
            "case is closed.",
            "Dispose of specimens and contaminated consumables under the Laboratory Waste "
            "Management procedure.",
        ]),
    ]))

    body.append(("17. Risk Management, Nonconforming Work, and Contingency", [
        ("tbl", ["Risk", "Control measure"], risks, [6.0, 11.0]),
        ("p", "Nonconforming work (ISO 15189:2022, 7.5): when a result, control or step does not "
              "conform, stop, report to the Laboratory Manager, assess the clinical impact on "
              "results already released, repeat or recollect, and issue corrected reports. "
              "Record the event and the corrective action and evaluate its effectiveness (8.7)."),
        ("p", "If the analyser is withdrawn from service, the laboratory reverts to the backup "
              "analyser or to the manual method recorded in the contingency plan, states the "
              "method used on the report, and restores the analyser only after maintenance and "
              "control are acceptable again."),
    ]))

    body.append(("18. Records", [
        ("tbl", ["Record", "Where kept", "Retention"],
         [
             ["Patient results and reports", "EMR / LIS and analyser", "Per the laboratory records policy"],
             ["Control results and Levey-Jennings charts", "Quality control register", "As above"],
             ["Calibration and maintenance records", "Instrument logbook", "As above"],
             ["Reagent, control and calibrator lot records", "Reagent records", "As above"],
             ["EQA reports and actions", "Quality file", "As above"],
             ["Training, competence and authorisation (section 21)", "Personnel training file",
              "Duration of employment plus the period required by national rules"],
         ],
         [6.5, 5.5, 5.0]),
        ("p", "Records are created at the time the activity is performed, remain legible and "
              "traceable to the person or user account, and are retrievable (ISO 15189:2022, 8.4; "
              f"{ADM_DOC_CONTROL})."),
    ]))

    # Sections 19 and 20 are deliberately one section. They were two, and the split left
    # a page holding three orphaned bullets and nothing else -- the half-empty page this
    # layout exists to avoid. They serve the same purpose: what a reader consults while
    # using this SOP, whether it is a controlled internal document or an external standard.
    body.append(("19. Supporting Documents and References", [
        ("h2", "19.1 Controlled documents of this laboratory"),
        ("b", [
            f"Documentation of examination procedures ({ADM_EXAM_DOCS}).",
            f"Document and Record Control ({ADM_DOC_CONTROL}).",
            f"Pre-Examination Procedure ({ADM_PRE_EXAM}).",
            f"Venous Blood Sample Collection ({code('venous')}).",
            f"Maintenance of haematology analysers ({ADM_EQUIP_MAINT}).",
            f"Troubleshooting of laboratory analysers — error codes, causes and corrective "
            f"action ({code('trouble')}).",
            "Laboratory Normal Ranges list; Critical Results list; Tests TAT list; Safety "
            "Manual; Waste Management procedure; Result Reporting and Release procedure.",
        ] + extra_support),
        ("h2", "19.2 External references"),
        ("b", [
            "ISO 15189:2022 — Medical laboratories (in particular 7.3.1, 7.3.2, 7.3.4 to "
            "7.3.7).",
            f"{ifu}.",
            "CLSI GP41 — Collection of Diagnostic Venous Blood Specimens; WHO — Laboratory "
            "Quality Management System (LQMS) Handbook.",
        ] + extra_refs),
    ]))

    return {
        "out_stem": f"{c} - {model} examination",
        "meta": {"doc_no": c, "title": REGISTER[key]["title"],
                 "version": REGISTER[key]["vers"]},
        "control_rows": [
            ("SOP Number", c),
            ("SOP Title", REGISTER[key]["title"]),
            ("Version", REGISTER[key]["vers"]),
            ("Effective Date", "14/11/2026"),
            ("Next Review Date", "14/11/2028"),
            ("Supersedes", "— (first issue; no SOP previously covered this examination)"),
            ("Document Owner", "Kelvin SAMAHANI — Laboratory Manager"),
            ("Classification", "Controlled document — master copy: document control system "
                               f"({ADM_DOC_CONTROL})"),
        ],
        "revision_rows": [
            ["J.P. NSENGIYUMVA", REGISTER[key]["vers"], "14/11/2026",
             "First issue. Written because the analyser is in service and no examination SOP "
             "covered it."],
            ["", "", "", ""],
        ],
        "body": body,
    }


XN330 = analyser_sop(
    "xn330", model="Sysmex XN-330", manufacturer="Sysmex", serial="17156",
    asset="RULIDH/LAB/ME/0501", bench="Haematology bench", discipline="Haematology",
    clinical="The complete blood count is the most frequently requested haematology "
             "examination. It counts and characterises the cells of peripheral blood — red "
             "cells, white cells and platelets — and derives the indices used to classify "
             "anaemia, to detect and follow infection and inflammation, to recognise "
             "leukaemia and other haematological disorders, and to monitor the effect of "
             "treatment including transfusion and chemotherapy.",
    principle="A measured volume of anticoagulated whole blood is drawn into the analyser, "
              "divided into aliquots and diluted. Cells are counted and sized by electrical "
              "impedance and/or optical (flow cytometric) measurement as the model provides; "
              "haemoglobin is measured photometrically after lysis of the red cells. The "
              "analyser derives the indices (MCV, MCH, MCHC, RDW) from the measured counts "
              "and the haemoglobin, and produces the white cell differential from the "
              "channels fitted to the instrument. The exact measurement channels, reagents "
              "and calculation rules are those of the IFU for this model and are incorporated "
              "by reference (ISO 15189:2022, 7.3.6 d).",
    specimen=[
        ("p", "Venous blood collected by the procedure in the Pre-Examination Procedure and in "
              "the venipuncture SOP, into the tube specified for a CBC on this analyser, filled "
              "to the marked volume."),
        ("h2", "6.1 Request and patient identification"),
        ("b", [
            "Accept requests entered in the EMR by an authorised requester. Clinical details "
            "relevant to interpretation (fever, bleeding, fatigue, transfusion, chemotherapy, "
            "recent travel) must be visible: a blood count cannot be interpreted without them.",
            "Identify the patient with at least two identifiers matched to the EMR request and "
            "label the tube in the patient's presence.",
        ]),
        ("h2", "6.2 Collection and transport"),
        ("b", [
            "Fill the tube to the mark. An under-filled tube changes the dilution and gives "
            "falsely low counts.",
            "Mix by gentle inversion immediately after collection; do not shake. Transport "
            "promptly, do not refrigerate before analysis, and state the collection time on "
            "the request. A clotted specimen is rejected.",
            "Reject and request a new specimen when: the tube is unlabelled or the label "
            "disagrees with the request; the tube is clotted, haemolysed or under-filled; the "
            "specimen is beyond the stability limit in the IFU for the parameters requested.",
            "Record every rejection and inform the requester (ISO 15189:2022, 7.3.4).",
        ]),
    ],
    specimen_notes=None,
    startup=[
        ("1.", "Check that the analyser and the UPS are connected, and that the waste container "
               "has room and the reagents are loaded."),
        ("2.", "Switch on and allow the warm-up and background checks the IFU requires."),
        ("3.", "Confirm the reagent status: volume, on-board expiry and lot number. Record a new "
               "lot in the reagent log."),
        ("4.", "Run the controls specified for this analyser and confirm they are acceptable "
               "before any patient specimen is analysed (section 11)."),
        ("5.", "Record the start-up, the control results and any remark in the instrument log."),
    ],
    run_steps=[
        ("1.", "Check the specimen against the request: identity, tube, filling, clots, "
               "haemolysis."),
        ("2.", "Load the tube in the position the analyser requires and start the run. Do not "
               "re-load a tube that has already been aspirated unless the analyser asks for it."),
        ("3.", "Watch the aspiration and the flags. A short sample, a clot or an obstruction is a "
               "fault, not a result: clear it and repeat the specimen (section 17). When the "
               "analyser shows an error message or code, act on it from the code table for this "
               "analyser and record the code — "
               f"{code('trouble')}."),
        ("4.", "Repeat any specimen that the analyser flags as needing review, or that gives a "
               "result outside the clinically plausible range."),
        ("5.", "For a flagged differential or an unexpected cell population, prepare and examine "
               "a stained film before releasing the result (10.3)."),
    ],
    results=[
        "Accept a result only when the controls for that run are acceptable.",
        "Check the result against the patient's previous result where one exists (delta check) "
        "and against the clinical information on the request.",
        "Examine the flags: a flag is a prompt to look, not a diagnosis and not a reason to "
        "suppress a result.",
        "Release the result in the EMR under your own account; the person named against a result "
        "is accountable for it.",
        "Any result within a critical limit is telephoned immediately (section 14).",
    ],
    interp_intro="The analyser reports the counts, indices, differential and flags. Before "
                 "release, check the result against the previous result, the clinical "
                 "information, the control status of the run and the specimen quality. The "
                 "common patterns are:",
    interp_rows=[
        ["Low haemoglobin — classify by MCV", "Low MCV: iron deficiency, thalassaemia, anaemia "
         "of chronic disease. High MCV: B12 or folate deficiency, alcohol, liver disease, "
         "hypothyroidism, myelodysplasia. Normal MCV: acute blood loss, haemolysis, chronic "
         "kidney disease. Confirm with ferritin, reticulocyte count and the film."],
        ["High white cell count", "Bacterial infection; inflammation; stress; corticosteroid "
         "therapy; leukaemia. A differential with blasts or unexplained immature forms needs a "
         "film and urgent review."],
        ["Low white cell count or low platelets", "Viral infection; drug- or "
         "chemotherapy-induced suppression; hypersplenism; marrow failure; immune "
         "thrombocytopenia; sepsis; dengue; platelet clumping in the tube (repeat in citrate "
         "or examine the film)."],
        ["High platelets", "Reactive thrombocytosis after infection, inflammation or bleeding; "
         "iron deficiency; myeloproliferative disease."],
        ["Result not matching the patient", "Repeat the specimen and check identification before "
         "release; treat mislabelling as nonconforming work (7.5)."],
    ],
    critical=[
        "Haemoglobin at or below the transfusion threshold on the Critical Results list.",
        "Platelet count at or below the critical limit: risk of spontaneous bleeding.",
        "White cell count above or below the critical limits, or blasts on the film.",
        "Any result the authorised person judges immediately life-threatening.",
    ],
    reports="Turnaround time is that of the Laboratory Tests TAT list for a full blood count. "
            "Results are communicated immediately after verification, even where other "
            "examinations on the same request are still in progress.",
    risks=[
        ["Wrong-patient result released",
         f"Identification at collection ({code('venous')}); label checked against the EMR at "
         f"loading; delta check before release"],
        ["Clotted or under-filled tube counted as valid",
         "Filling and clot check at loading; rejection criteria in 6.2; every rejection recorded"],
        ["Control failure, or an analyser fault, mistaken for a result",
         "Controls before patient testing with defined rejection rules and chart review; "
         "aspiration watched and short-sample or obstruction flags treated as faults, with the "
         "specimen repeated after clearance"],
        ["Power interruption", "UPS; event recorded; controls repeated before testing resumes"],
        ["Critical result not communicated",
         "Telephone with read-back; recorded; escalated to the Laboratory Manager if the "
         "requester cannot be reached"],
        ["EMR/LIS downtime",
         "Results recorded on the downtime form, telephoned with read-back, and back-entered "
         "within 24 hours of restoration"],
    ],
    limitations=[
        "An analyser result is not a diagnosis. Abnormal results, and results that disagree with "
        "the clinical picture, require examination of a stained film.",
        "Cold agglutinins, lipaemia, icterus, very high white cell counts, platelet clumping and "
        "red cell fragments all interfere with one or more parameters; the interfering "
        "concentrations are those stated in the IFU.",
        "Specimens outside the stability limit in the IFU give unreliable indices, so the "
        "elapsed time from collection must be known. Nucleated red cells and giant platelets "
        "are counted with the wrong population unless the analyser's correction is applied; "
        "confirm on the film.",
    ],
    ifu="Sysmex XN-330 instructions for use and the package inserts of the reagents and controls "
        "in use (EXT-RliDHL-HEM series, Master List)",
    extra_support=[],
    extra_refs=[],
    replacement_note="It replaces no earlier SOP: the analysers previously covered were withdrawn "
                     "from service.",
)

XQ320 = analyser_sop(
    "xq320", model="Sysmex XQ-320", manufacturer="Sysmex", serial="12357",
    asset="RULIDH/LAB/ME/0499", bench="Haematology bench", discipline="Haematology",
    clinical="The complete blood count characterises the cells of peripheral blood and is "
             "used to classify anaemia, to detect and follow infection and inflammation, to "
             "recognise haematological disease and to monitor treatment. The laboratory "
             "operates two haematology analysers so that the service continues when one is "
             "unavailable, and so that the workload can be shared.",
    principle="A measured volume of anticoagulated whole blood is aspirated, diluted and "
              "measured. Cells are counted and sized by the measurement channels the model "
              "provides, and haemoglobin is measured photometrically after lysis of the red "
              "cells; the indices are derived from the measured values. The measurement "
              "channels, reagents and calculation rules are those of the IFU for this model "
              "and are incorporated by reference (ISO 15189:2022, 7.3.6 d).",
    specimen=[
        ("p", "Venous blood collected by the procedure in the Pre-Examination Procedure and in "
              "the venipuncture SOP, into the tube specified for a CBC on this analyser, filled "
              "to the marked volume."),
        ("h2", "6.1 Request and patient identification"),
        ("b", [
            "Accept requests entered in the EMR by an authorised requester, with the clinical "
            "details needed to interpret the result.",
            "Identify the patient with at least two identifiers matched to the EMR request and "
            "label the tube in the patient's presence.",
        ]),
        ("h2", "6.2 Collection and transport"),
        ("b", [
            "Fill the tube to the mark and mix by gentle inversion immediately; a clotted or "
            "under-filled tube is rejected.",
            "Transport promptly; do not refrigerate before analysis; record the collection time.",
            "Reject and request a new specimen when the label disagrees with the request, or the "
            "tube is clotted, haemolysed, under-filled or beyond the stability limit in the IFU.",
            "Record every rejection and inform the requester (ISO 15189:2022, 7.3.4).",
        ]),
    ],
    specimen_notes=None,
    startup=[
        ("1.", "Check the analyser and UPS connections, the waste container and the reagents."),
        ("2.", "Switch on and complete the warm-up and background checks the IFU requires."),
        ("3.", "Confirm reagent volume, on-board expiry and lot number; record a new lot."),
        ("4.", "Run the controls specified for this analyser and confirm they are acceptable."),
        ("5.", "Record the start-up and control results in the instrument log."),
    ],
    run_steps=[
        ("1.", "Check the specimen against the request: identity, tube, filling, clots, "
               "haemolysis."),
        ("2.", "Load and run the tube as the analyser requires."),
        ("3.", "Treat a short sample, clot or obstruction as a fault: clear it and repeat."),
        ("4.", "Repeat any specimen flagged for review or giving an implausible result."),
        ("5.", "Examine a stained film before releasing a flagged differential or an unexpected "
               "population (10.3)."),
    ],
    results=[
        "Accept a result only when the controls for that run are acceptable.",
        "Check against the previous result (delta check) and the clinical information.",
        "A flag is a prompt to look, not a diagnosis and not a reason to suppress a result.",
        "Release in the EMR under your own account.",
        "Any result within a critical limit is telephoned immediately (section 14).",
    ],
    interp_intro="The analyser reports the counts, indices, differential and flags. The same "
                 "interpretive principles apply as on the other haematology analyser, and the "
                 "two are kept comparable under 11.4. The common patterns are:",
    interp_rows=[
        ["Low haemoglobin with low MCV", "Iron deficiency; thalassaemia; anaemia of chronic "
         "disease. Confirm with ferritin and the film."],
        ["Low haemoglobin with high MCV", "B12 or folate deficiency; alcohol; liver disease; "
         "hypothyroidism; myelodysplasia."],
        ["High white cell count", "Bacterial infection; inflammation; stress; corticosteroid "
         "therapy; leukaemia. Blasts or unexplained immature forms need a film and urgent review."],
        ["Low white cell count", "Viral infection; drug-induced suppression; hypersplenism; "
         "marrow failure."],
        ["Low platelets", "Immune thrombocytopenia; sepsis; dengue and other infections; drugs; "
         "clumping in the tube."],
        ["High platelets", "Reactive thrombocytosis; iron deficiency; myeloproliferative disease."],
        ["Result differing from the XN-330", "Repeat on both analysers and apply 11.4 before "
         "release; report the method used if one analyser is out of service."],
    ],
    critical=[
        "Haemoglobin at or below the transfusion threshold on the Critical Results list.",
        "Platelet count at or below the critical limit.",
        "White cell count above or below the critical limits, or blasts on the film.",
        "Any result the authorised person judges immediately life-threatening.",
    ],
    reports="Turnaround time is that of the Laboratory Tests TAT list. Where the result is "
            "produced on this analyser while the other is out of service, the report states the "
            "method used.",
    risks=[
        ["Wrong-patient result released",
         f"Identification at collection ({code('venous')}); label checked at loading; delta check"],
        ["Results not comparable between the two analysers",
         "Comparability study under 11.4 whenever an analyser is added or returned to service, "
         "and at least yearly"],
        ["Clotted or under-filled tube", "Filling and clot check; rejection criteria in 6.2"],
        ["Control failure not detected", "Controls before testing; rejection rules; chart review"],
        ["Analyser fault mistaken for a result", "Aspiration watched; faults cleared and the "
         "specimen repeated"],
        ["Critical result not communicated", "Telephone with read-back; recorded; escalated"],
        ["EMR/LIS downtime", "Downtime form; telephone with read-back; back-entry within 24 hours"],
    ],
    limitations=[
        "An analyser result is not a diagnosis; abnormal or discordant results require a film.",
        "Lipaemia, icterus, cold agglutinins, platelet clumping and very high white cell counts "
        "interfere with one or more parameters; interfering concentrations are those in the IFU.",
        "Specimens beyond the stability limit in the IFU give unreliable indices.",
        "This analyser is not used for a parameter, or a specimen type, that the IFU does not "
        "claim for it; such requests are performed on the other analyser or by the manual method.",
    ],
    ifu="Sysmex XQ-320 instructions for use and the package inserts of the reagents and controls "
        "in use (EXT-RliDHL-HEM series, Master List)",
    extra_support=[],
    extra_refs=[],
    replacement_note="It replaces no earlier SOP. It exists so that the laboratory is not "
                     "dependent on a single analyser.",
)

PIMA = analyser_sop(
    "pima", model="Alere Pima CD4", manufacturer="Alere", serial="PIMA-D-005378",
    asset="RULIDH/LAB/ME/0485", bench="Haematology room", discipline="Immunology",
    clinical="The CD4 count measures the number of CD4-positive T lymphocytes per microlitre "
             "of blood and is the laboratory marker used to stage HIV infection, to decide "
             "when to start cotrimoxazole prophylaxis and antiretroviral therapy, and to "
             "monitor immune recovery on treatment. The absolute count and the percentage of "
             "lymphocytes that are CD4-positive are both reported; the percentage is less "
             "affected by specimen ageing and by total lymphocyte changes.",
    principle="A sample of whole blood is mixed with a reagent that lyses the red cells and "
              "labels CD4-bearing lymphocytes with a fluorophore-labelled antibody. The "
              "labelled cells are counted in a disposable cartridge by the instrument's optical "
              "system, and the result is expressed as an absolute CD4 count and as a percentage "
              "of lymphocytes. The cartridges, reagents and the counting rules are those of the "
              "IFU for this instrument, incorporated by reference (ISO 15189:2022, 7.3.6 d).",
    specimen=[
        ("p", "Venous blood collected by the procedure in the Pre-Examination Procedure and in "
              "the venipuncture SOP, into the tube specified for CD4 counting on this "
              "instrument, filled to the marked volume."),
        ("h2", "6.1 Request and patient identification"),
        ("b", [
            "Accept requests entered in the EMR by an authorised requester, with the clinical "
            "details needed to interpret the result.",
            "Identify the patient with at least two identifiers matched to the EMR request and "
            "label the tube in the patient's presence. Misidentification here attaches a CD4 "
            "count to the wrong person's HIV record, so the check is made twice.",
        ]),
        ("h2", "6.2 Collection, transport and stability"),
        ("b", [
            "Fill the tube to the mark and mix by gentle inversion. A clotted or under-filled "
            "tube is rejected.",
            "Analyse within the time limit stated in the IFU for the cartridge and the specimen; "
            "the count falls as the specimen ages. Record the collection time and the analysis "
            "time.",
            "Transport at ambient temperature; do not refrigerate or heat the specimen.",
            "Reject and request a new specimen when the label disagrees with the request, or the "
            "tube is clotted, haemolysed, under-filled or beyond the stability limit.",
            "Record every rejection and inform the requester (ISO 15189:2022, 7.3.4).",
        ]),
    ],
    specimen_notes=None,
    startup=[
        ("1.", "Check the instrument, the battery or mains supply and the printer."),
        ("2.", "Switch on; complete the self-test the IFU requires."),
        ("3.", "Check the reagent and cartridge lot numbers and expiry dates; record a new lot."),
        ("4.", "Run the control material specified for this instrument and confirm it is "
               "acceptable before patient specimens."),
        ("5.", "Record the start-up and the control result in the instrument log."),
    ],
    run_steps=[
        ("1.", "Check the specimen against the request: identity, tube, filling, clots, "
               "haemolysis, elapsed time since collection."),
        ("2.", "Label the cartridge with the specimen identifiers before adding the sample."),
        ("3.", "Prepare and load the specimen as the IFU directs for the cartridge in use."),
        ("4.", "Run the cartridge within the time the IFU allows after preparation."),
        ("5.", "Read and record the absolute count and the percentage; retain the printout with "
               "the specimen identifiers on it."),
    ],
    results=[
        "Accept a result only when the control for that run is acceptable.",
        "Check the absolute count against the percentage for internal consistency.",
        "Check the result against the patient's previous count (delta check) and the clinical "
        "information.",
        "Release in the EMR under your own account; the printout is attached to the record.",
        "Results that fall within a critical limit are telephoned immediately (section 14).",
    ],
    interp_intro="CD4 counts are interpreted against the clinical stage and the treatment "
                 "history. The thresholds in use at this laboratory are those of the national "
                 "HIV programme and the Critical Results list; this table gives the common "
                 "patterns only.",
    interp_rows=[
        ["CD4 count falling over successive visits",
         "HIV disease progression; assess adherence, viral load and prophylaxis; repeat to "
         "confirm before acting on a single value."],
        ["CD4 count rising on treatment",
         "Immune recovery on antiretroviral therapy; interpret with the viral load."],
        ["CD4 count much lower than expected",
         "Check specimen age and handling, repeat the count, and check identification. A count "
         "that does not fit the clinical picture is repeated before release."],
        ["Percentage and absolute count discordant",
         "Repeat the cartridge; consider a change in total lymphocyte count; do not release until "
         "the two agree with the previous record."],
        ["Control outside limits", "Do not report. Repeat the control; if still unacceptable, "
         "stop and treat as a fault (section 17)."],
    ],
    critical=[
        "A CD4 count below the threshold that triggers prophylaxis on the Critical Results list.",
        "A fall in CD4 count between consecutive visits beyond the limit set by the HIV "
        "programme.",
        "Any result the authorised person judges immediately clinically significant.",
    ],
    reports="The report states the absolute CD4 count (cells/µL) and the percentage, the "
            "collection time and the analysis time. Where a result is produced while the "
            "laboratory's other CD4 method is unavailable, the report states the method used.",
    risks=[
        ["Wrong-patient result attached to an HIV record",
         "Two identifier checks, at collection and again at the instrument; label on the "
         "cartridge before the sample is added; retained printout"],
        ["Specimen aged beyond the stability limit", "Collection time and analysis time both "
         "recorded; elapsed time checked at loading"],
        ["Control failure not detected", "Control before patient specimens; stop if unacceptable"],
        ["Cartridge or reagent lot shift", "Lot number recorded; control before use of a new lot"],
        ["Power or battery failure", "Mains or charged battery; control repeated after any "
         "interruption"],
        ["Critical result not communicated", "Telephone with read-back; recorded; escalated"],
        ["EMR/LIS downtime", "Result recorded on the downtime form, telephoned with read-back, "
         "back-entered within 24 hours"],
    ],
    limitations=[
        "The count is affected by the age of the specimen; the limit is that stated in the IFU.",
        "The instrument measures CD4 on a sample of whole blood; it does not measure CD8, and "
        "does not substitute for a viral load.",
        "A result that does not fit the clinical picture, or disagrees with the previous record, "
        "is repeated before release.",
        "The instrument is not used outside the specimen types and conditions the IFU claims.",
    ],
    ifu="Alere Pima CD4 instructions for use and the package inserts of the cartridges and "
        "controls in use (EXT-RliDHL-IMM series, Master List)",
    extra_support=[],
    extra_refs=["National HIV programme guidance on CD4 thresholds in use in Rwanda."],
    replacement_note="The laboratory's earlier CD4 SOPs govern a different instrument and are "
                     "not applicable to this analyser.",
)

# ---------------------------------------------------------------------------
# ADM SOP 065 -- Troubleshooting of laboratory analysers
#
# One document, every analyser. This replaces the phrase "troubleshoot the
# error" with a procedure: a code, what it means, what caused it, what the
# operator does, what the supervisor does, and who is called.
#
# EVERY code, message and corrective action below is taken from the
# manufacturer's own documentation for the instrument (listed in section 20).
# Nothing is an operator's guess and nothing is remembered. Where a
# manufacturer does not publish a numeric code for a failure, that is stated
# rather than invented -- the Pima and the Xpert failures without codes are the
# clearest example: they are real failures with no number, and a table that
# pretended otherwise would send staff hunting for a code that does not exist.
#
# The last column of each table is the part the laboratory was missing: WHO is
# responsible. An error code that anyone may act on is an error code nobody
# acts on.
# ---------------------------------------------------------------------------

TRB = REGISTER["trouble"]
TRB_CODE = code("trouble")

TRB_ESC = [
    "Stop testing on the affected module or instrument, and do not release the result. "
    "An error is not a result and must never be reported as one, not even with a comment.",
    "Record the numeric code, the approximate time, the module, the assay and your initials "
    "in the instrument log at the moment it appears. Record the code from the screen, not "
    "from memory at the end of the shift: a code read later is a guess.",
    "If the same code appears twice in a row on the same module, treat the module as suspect "
    "even if the third attempt works. Intermittent is not fixed.",
    "Tell the Laboratory Manager or the senior biotechnologist on duty the same day, and tell "
    "the requester that the result will be delayed and why.",
    "The Laboratory Manager reviews the failure log monthly against the error rate, looks for "
    "a code that has become common, a module that has become worse and an operator whose rate "
    "is high, and records the review.",
]

TROUBLE = {
    "meta": {"doc_no": TRB_CODE, "title": TRB["title"], "version": TRB["vers"]},
    "out_stem": f"{TRB_CODE} - Troubleshooting of laboratory analysers",
    "revision_rows": [
        ["Kelvin SAMAHANI", TRB["vers"], "14/11/2026",
         "First issue. Written because none of the examination SOPs carried the error codes "
         "of the analyser they govern: they told staff to record a code and to consult the "
         "manufacturer's manual, and supplied neither. Every code and corrective action here "
         "is taken from the manufacturer's own documentation for the instrument."],
    ],
    "control_rows": [
        ("SOP Number", TRB_CODE),
        ("SOP Title", TRB["title"]),
        ("Version", TRB["vers"]),
        ("Effective Date", "14/11/2026"),
        ("Next Review Date", "14/11/2028"),
        ("Supersedes", "— (first issue; no previous SOP covered troubleshooting)"),
        ("Document Owner", "Kelvin SAMAHANI — Laboratory Manager"),
        ("Applies to", "Every analyser, instrument and device in the laboratory that displays "
                       "an error code or reports an error, no result or invalid result"),
        ("Classification", "Controlled document — master copy: document control system "
                           f"({ADM_DOC_CONTROL})"),
    ],
    "body": [
        ("1. Introduction", [
            ("p", "The laboratory holds four analysers in service and one awaiting repair. Each of "
                  "them fails: the GeneXpert returned 938 failed tests in 9,460 runs between "
                  "February 2022 and April 2025, and the Alere Pima CD4 is reported in the "
                  "literature to return invalid results on about one test in eight. A failure is "
                  "not an accident to be apologised for; it is a defined event with a code, a "
                  "cause and a response."),
            ("p", "Until this SOP was issued, the examination SOPs said what to do in principle — "
                  "\"note the error code\", \"troubleshoot and repeat\" — and then referred the "
                  "reader to a manufacturer's manual held somewhere else, in a form the SOP never "
                  "converted. The codes were therefore never read, never recorded and never "
                  "analysed. This SOP puts that content where the operator can reach it and makes "
                  "the recording of a code mandatory."),
            ("p", "It does not replace the manufacturer's manual, and it does not licence anyone to "
                  "go beyond it. Where this SOP and the current manual differ, the manual is "
                  "authoritative and this SOP is corrected: that is an amendment, not a decision "
                  "for the bench."),
        ]),
        ("2. Purpose", [
            ("p", "To give, for each analyser in the laboratory, the error codes and messages the "
                  "instrument actually displays, what each one means, the cause the manufacturer "
                  "attributes to it, the corrective action the manufacturer specifies, and the "
                  "point at which the operator stops and escalates — so that a failure is handled "
                  "the same way by every member of staff and is recorded well enough to be "
                  "investigated."),
        ]),
        ("3. Scope", [
            ("p", "This SOP applies to every analyser, instrument and device in the laboratory that "
                  "can display an error code or can return an error, no-result or invalid result:"),
            ("b", [
                "Sysmex XN-330 haematology analyser — section 5.",
                "Sysmex XQ-320 haematology analyser — section 6.",
                "Roche cobas c 111 clinical chemistry analyser — section 7.",
                "GeneXpert Dx system, all assays in use (MTB/RIF, MTB/RIF Ultra, Xpert HIV-1 "
                "Qual XC, Xpert HBV viral load, Xpert HCV viral load) — section 8.",
                "Alere Pima CD4 analyser — section 9.",
            ]),
            ("p", "A device not listed here is added by amendment before it is used to produce a "
                  "result. Devices that cannot fail in a reportable way (a centrifuge, a "
                  "thermometer, a water bath) are covered by their own maintenance SOPs and by the "
                  "laboratory's procedure for nonconforming work, not by this document."),
        ]),
        ("4. How to use this SOP, and where responsibility stops", [
            ("p", "Every table in this SOP has the same seven columns: the code or message as the "
                  "instrument displays it; what it means; the cause the manufacturer gives; what "
                  "the operator does first; what is done when the first action does not clear it; "
                  "whether the module is taken out of service; and who does it. Read the row for "
                  "the code shown on the screen, and do what the row says."),
            ("h2", "4.1 The rules that apply to every analyser"),
            ("b", TRB_ESC),
            ("h2", "4.2 What the operator may and may not do"),
            ("p", "An operator may: clear the error and restart the instrument; repeat the test "
                  "with a new cartridge, a new specimen aliquot or a new reagent; clean what the "
                  "manufacturer's instructions for that code tell them to clean; take a module out "
                  "of service; and call the service provider."),
            ("p", "An operator may not: open the instrument beyond the access the manufacturer's "
                  "maintenance instructions allow; remove a stuck cartridge by force or by "
                  "dismantling a module (this needs the supervisor, the service provider's "
                  "instructions and, in practice, the service provider); defeat, disable or ignore "
                  "an alarm; re-run a test to get the answer that was expected; or report a result "
                  "that came with an error, no matter how plausible it looks."),
            ("h2", "4.3 Recording — the part that was missing"),
            ("p", "The single reason the laboratory cannot today say whether its failures are "
                  "cartridges, specimens, modules or people is that the numeric code was never "
                  "written down. The raw GeneXpert export for 2022 to 2025 records a failure as "
                  "\"Error\" 772 times, \"Load Error\" 38 times and \"Warning\" twice, and does not "
                  "record one code. A failure with no code cannot be investigated, cannot be "
                  "counted by cause, and cannot be shown to have been fixed."),
            ("p", "From the effective date of this SOP, the instrument log carries the code. The "
                  "log is a controlled record: code, date and time, module (where the instrument "
                  "has modules), assay, specimen or control, the action taken, whether the repeat "
                  "passed, and the initials. Entries are made in ink at the time; a correction is a "
                  "single line through the entry, the correct entry beside it, and initials and "
                  "date."),
            ("h2", "4.4 Monthly review"),
            ("p", "The Laboratory Manager reviews, each month: the number of tests and the number of "
                  "failures by analyser; the ten most frequent codes; the failure rate by module "
                  "and by operator; and whether any code has risen against the previous three "
                  "months. A code that recurs on one module, a rate above the laboratory's "
                  "objective, or a rate that has doubled is investigated and recorded as "
                  "nonconforming work with a corrective action. The review is a record and is "
                  "kept."),
        ]),
        ("5. Sysmex XN-330 haematology analyser", [
            ("p", "The XN-330 does not display a numeric error code. It displays an error message in "
                  "words, colour-coded by severity — a red background is a warning that stops "
                  "analysis, an orange background is a caution — and it sorts each failure into "
                  "[Result], [Func.] or a barcode read error. The message is the identifier, and it "
                  "is the message that is recorded."),
            ("tbl",
             ["Message or indication", "Meaning", "Cause (Sysmex)", "Operator action", "If it "
              "persists", "Module out of service", "Responsible"],
             [
                 ["Blood cannot be aspirated.",
                  "The aspiration sensor did not detect blood. No result is produced.",
                  "Short sample, clot at the tube bottom, tube not presented correctly, or a "
                  "blocked probe.",
                  "Check the tube volume against the minimum on the analyser; remix by inversion; "
                  "check for a clot with a applicator stick; rerun.",
                  "Clean the aspiration probe as the maintenance SOP directs; rerun once.",
                  "No — the analyser, not a module.",
                  "Operator; supervisor if the second rerun fails."],
                 ["Insufficient blood volume (short sample).",
                  "The sample was aspirated but the volume was below the analyser's requirement.",
                  "Under-filled tube, or the tube was pulled from the aspirator too early.",
                  "Check the fill line; recollect if the tube is short; rerun.",
                  "If it recurs on correctly filled tubes, check the aspiration volume setting and "
                  "call service.",
                  "No.",
                  "Operator."],
                 ["Low count error / Blank data.",
                  "A parameter counted abnormally low against the blank.",
                  "Clotted or very dilute specimen, aspiration fault, or a dirty flow cell.",
                  "Remix, inspect for clots and haemolysis, rerun once.",
                  "Run a background check; if it fails three times in a row, stop and call service.",
                  "No.",
                  "Operator; supervisor for the background check."],
                 ["0.06 MPa pressure error / 0.25 MPa pressure error / Pressure sensor error / "
                  "Vacuum sensor error.",
                  "The pneumatic pressure or vacuum is outside the range the analyser requires.",
                  "Blocked or leaking tubing, a stuck valve, a waste container that is full, or a "
                  "failed pressure sensor.",
                  "Empty the waste; check that all reagent and waste connections are seated; "
                  "restart the analyser.",
                  "Do not open the pneumatic path beyond the maintenance SOP. Call service.",
                  "No — instrument-level.",
                  "Operator to restart; service provider beyond that."],
                 ["Temperature or thermistor error (reagent heater 34 °C or 41 °C, ambient "
                  "temperature, FCM reaction chamber, FCM sheath, detector).",
                  "A temperature sensor reads outside its range.",
                  "Room temperature above the analyser's operating range, a blocked or dirty fan "
                  "filter, insufficient clearance around the analyser, or a failed heater.",
                  "Check the room temperature against the range in the maintenance SOP; check the "
                  "10 cm clearance; check that the rear fan is running and the filter is clean; "
                  "restart.",
                  "Call service. A heater or thermistor is not an operator-replaceable part. Record "
                  "the module temperature from the Maintenance menu with the call.",
                  "No.",
                  "Operator to check the room; service provider for the heater."],
                 ["Laser output error / APD thermistor error / FCM detector thermistor error.",
                  "A failure in the optical or detector subsystem.",
                  "Component failure. The manufacturer suppresses this message on screen and "
                  "performs an automatic repeat; when it is visible, the repeat did not clear it.",
                  "Do not attempt repair. Stop using the analyser and call service the same day.",
                  "—",
                  "Yes, the analyser.",
                  "Service provider; supervisor informs the Laboratory Manager."],
                 ["ID read error (sample or rack barcode).",
                  "A barcode label could not be read.",
                  "Damaged, wet, wrinkled or badly placed label; dirty read window; a label outside "
                  "the specification.",
                  "Wipe the label and the reader window; re-present the tube; if it still fails, "
                  "enter the sample number manually and record that it was entered manually.",
                  "Report a label that cannot be read to the collection site; the specimen is not "
                  "rejected for this alone.",
                  "No.",
                  "Operator."],
                 ["QC alarm / X-barM limit error / L-J limit error / Control expired error / "
                  "Unregistered control.",
                  "Quality control has failed, or the control in use is expired or not registered.",
                  "Control out of limits, expired control, wrong control file, or a real shift in "
                  "the analyser.",
                  "Do not test patient specimens. Check the control's expiry and storage, repeat "
                  "the control once.",
                  "Follow the quality control procedure: stop, record the nonconformity, review the "
                  "patient results released since the last acceptable control.",
                  "The analyser is stopped for patient testing until the control passes.",
                  "Operator stops; supervisor authorises resumption."],
                 ["Reagent expired error.",
                  "An expired reagent is loaded.",
                  "Expiry passed, or the on-board stability has been exceeded.",
                  "Replace the reagent; record the lot and the date opened.",
                  "Check whether results were produced with the expired reagent and review them.",
                  "No.",
                  "Operator; supervisor reviews affected results."],
                 ["HC off-line / HC buffer full / LAN no response.",
                  "Communication with the laboratory information system or host computer has "
                  "failed.",
                  "Network or interface fault, host down, or a full transmission buffer.",
                  "Check the cable and the network indicator; continue in offline mode and record "
                  "results on the paper worksheet.",
                  "Inform whoever supports the LIS; results are entered when the link is restored, "
                  "and the report states that they were entered retrospectively.",
                  "No.",
                  "Operator; LIS support."],
                 ["Red status light, alarm sounding.",
                  "An error has occurred and the analyser has stopped.",
                  "Read the message on the screen — the light alone says nothing.",
                  "Touch the screen to silence the alarm, read the error message, and follow that "
                  "message's row in this table.",
                  "—",
                  "Only if the message says so.",
                  "Operator."],
             ],
             [2.5, 2.6, 2.9, 2.9, 2.9, 1.6, 1.6]),
            ("p", "IP messages (Blasts/Abn Lympho?, Left Shift?, Atypical Lympho?, NRBC?, RBC "
                  "Agglutination?, Turbidity/HGB Interf?, Fragments?, PLT Clumps?, and the rest) "
                  "are not errors. They are the analyser reporting a suspicion about the specimen. "
                  "They do not stop the run and they are not cleared: they are acted on by making a "
                  "film and by a manual differential, as the examination SOP requires. A result "
                  "carrying an IP message is released with the review the SOP defines, never "
                  "silently."),
        ]),
        ("6. Sysmex XQ-320 haematology analyser", [
            ("p", "The XQ-320 belongs to the same Sysmex generation and reports through the same "
                  "software, so its messages, meanings and corrective actions are the ones in "
                  "section 5. What differs is the hardware: three-part differential only, no "
                  "reticulocyte channel, a single aspiration path and no sampler. Two consequences "
                  "matter here, and both are recorded in the maintenance SOP rather than repeated:"),
            ("b", [
                "A failure that stops the analyser stops the whole laboratory's XQ-320 work, "
                "because there is no second module to fall back on. The XN-330 is then the backup "
                "and the report states which analyser produced the result.",
                "The part numbers, the fluid volumes and the maintenance intervals differ from the "
                "XN-330. When a code in section 5 sends the operator to service, the call is made "
                "quoting the XQ-320 serial number, not the XN-330's.",
            ]),
            ("p", "A message on the XQ-320 that is not in the section 5 table is treated as a "
                  "failure with no known corrective action: stop, record the exact wording, and "
                  "call service. Do not try the corrective action for a similar-looking message on "
                  "another analyser."),
        ]),
        ("7. Roche cobas c 111 clinical chemistry analyser", [
            ("p", "The cobas c 111 is the only analyser in the laboratory whose manufacturer "
                  "publishes a complete numeric alarm list, and the laboratory has already been "
                  "stopped by one of them: the instrument was withdrawn from service in April 2026 "
                  "with error 5040, which is a failed Peltier element in the cooling system. That "
                  "is a service intervention, not a bench fix, and the table says so."),
            ("p", "Codes are four digits. Alarms in the 5xxx range are hardware and fluidics; 6xxx "
                  "are photometer; 7xxx are transfer, position and level detection. Every row below "
                  "is from the cobas c 111 Operator's Manual, chapter 9, \"Messages and alarms\"."),
            ("tbl",
             ["Code", "Message", "Meaning and cause (Roche)", "Operator action", "If it persists",
              "Out of service", "Responsible"],
             [
                 ["5040", "Error cooling system. One of the Peltier elements does not react "
                  "anymore.",
                  "A Peltier element in the cooling system has failed. This is the fault that "
                  "withdrew this analyser from service in April 2026.",
                  "Restart the system once.",
                  "Call service. Do not open the instrument: the cooling system is not a "
                  "user-serviceable assembly.",
                  "Yes — the analyser.",
                  "Service provider; the Laboratory Manager records the withdrawal."],
                 ["5038", "Error analyzer heating system. The heating system does not react "
                  "anymore.",
                  "The heating system has failed.",
                  "Restart once.",
                  "Call service.", "Yes.", "Service provider."],
                 ["5034", "Error wash valve V1. The valve does not react anymore.",
                  "A fluidic valve has failed.", "Restart once.",
                  "Call service.", "Yes.", "Service provider."],
                 ["5028 / 5030 / 5032 / 5036", "Error temperature sensor (analyzer heating, "
                  "analyzer cooling, cooling-system heat sink, or instrument ambient).",
                  "A temperature sensor does not respond.",
                  "Restart once. Check the room temperature and that the vents are clear.",
                  "Call service.", "Yes.", "Service provider."],
                 ["5026", "Error wash pump P1 or mixer motor. The pump or motor could not be "
                  "started.",
                  "A drive has failed or is jammed.",
                  "Switch off; check the mixing mechanism and the fan for an obstruction as the "
                  "manual directs; restart.",
                  "Call service.", "Yes.", "Service provider."],
                 ["5052 / 5055", "Analyzer temperature is outside the range. / Analyzer heating "
                  "system overheated.",
                  "The instrument is above or below its operating temperature, or the heating "
                  "system has overheated.",
                  "Check the room temperature and the air circulation (unrestricted, vents not "
                  "blocked); restart.",
                  "Call service if it recurs with the room in range.", "Yes.", "Operator, then "
                  "service."],
                 ["5056", "The calibration of the transfer head fluid sensor failed.",
                  "The fluid sensor in the transfer head cannot be calibrated.",
                  "Check that the sample probe tube and the syringe are correctly installed; "
                  "restart.",
                  "Call service.", "Yes.", "Operator, then service."],
                 ["5013 / 5019", "Error fluid sensor transfer head. / Error fluid sensor "
                  "degasser.",
                  "A fluid sensor does not respond, or air is in the fluid system.",
                  "For 5019, perform Initialise Degasser Fluid Sensor as the screen directs. "
                  "Restart.",
                  "Call service.", "Yes.", "Operator, then service."],
                 ["6024 / 6025 / 6026 / 6027", "Photometer error (lamp fan speed, dark shutter "
                  "motor or shutter jamming, internal circuitry, signals out of range during a "
                  "dark measurement).",
                  "A fault in the photometer or its shutter.",
                  "Restart once.",
                  "Call service. Do not adjust the photometer.", "Yes.", "Service provider."],
                 ["6001 / 7071", "Photometer error. Signal is invalid or unstable. / Photometer "
                  "lamp error.",
                  "The photometer signal is unusable: dirty rotor, ageing lamp, or blocked light "
                  "path.",
                  "Switch off. Remove and clean the rotor as the maintenance SOP directs; restart.",
                  "Perform Replace Photometer Lamp and then Abs Air/Water Calibration — supervisor "
                  "level — and if it persists call service.",
                  "Yes, until the lamp and calibration are acceptable.", "Supervisor; service for "
                  "the lamp."],
                 ["7013 / 7130 / 7131 / 7132", "Position error X-transfer. / Rotor position error. "
                  "/ Position error X-transfer. / Initialization error Z-transfer.",
                  "A mechanical axis is blocked or jammed, or a home position is invalid.",
                  "Switch off immediately. Check the rotor, the transfer head and the probe for an "
                  "obstruction; check that the transfer head cover is seated; restart.",
                  "Call service. Never force a jammed axis.", "Yes.", "Supervisor; service "
                  "provider."],
                 ["7072", "The reagent mixing failed due to a level detection problem.",
                  "The probe could not detect the reagent level.",
                  "Check the probe for an obstruction and check the reagent bottle's fill level.",
                  "Call service if the level is adequate.", "Yes.", "Operator, then service."],
                 ["7111 / 7122 / 7123 / 7124", "Transfer movement failed due to open cover. / "
                  "Prepare stopped due to open cover. / Extra wash cycle after reagent pipetting "
                  "failed. / Analyzer temperature outside the range.",
                  "A cover was opened while the instrument was operating, or a wash cycle failed.",
                  "Ensure every cover is closed during initialisation, operation and maintenance; "
                  "for 7123 perform Deproteinise Probe before the next run; restart.",
                  "Call service if the temperature alarm recurs with the room in range.", "No, "
                  "unless it recurs.", "Operator."],
                 ["5002 / 5003 / 5004 / 5010", "Error motor X-transfer, Z-transfer, syringe "
                  "initialisation, or mixer motor.",
                  "A motor could not be started, or an initialisation failed.",
                  "Restart once. Check the relevant mechanism for jamming or blockage and that the "
                  "syringe is correctly installed.",
                  "Call service.", "Yes.", "Service provider."],
                 ["7001 / 7002 / 13002 / 13005 / 13009 / 13010", "Instrument Control software "
                  "exception / software error.",
                  "A software fault.",
                  "Restart the system.",
                  "Export the log files with the system still in the error condition and call "
                  "service. Note the code: it identifies the fault.",
                  "Yes, until service has the log.", "Service provider; supervisor exports the "
                  "log."],
                 ["Any code not in this table", "—", "—",
                  "Stop. Record the exact code and message text from the screen.",
                  "Call service quoting the code and send the system log.",
                  "Yes.", "Supervisor."],
             ],
             [1.7, 2.7, 3.0, 3.2, 2.8, 1.4, 1.6]),
            ("p", "The 5040 entry is the reason the laboratory's own analyser is away for repair and "
                  "the reason a report cannot be produced on it today. It is recorded here so that "
                  "the next occurrence is recognised in seconds rather than after a service visit, "
                  "and so that nobody at the bench spends a shift attempting a repair the "
                  "manufacturer reserves for a service engineer."),
        ]),
        ("8. GeneXpert Dx system", [
            ("p", "The GeneXpert is the analyser the laboratory uses for everything molecular: "
                  "tuberculosis, HIV early infant diagnosis, and the hepatitis B and C viral "
                  "loads. It is also, by the laboratory's own record, the analyser that fails "
                  "most — 938 failed tests out of 9,460 between February 2022 and April 2025."),
            ("p", "Two things about its errors are not obvious and both matter. First, the result "
                  "codes INVALID and NO RESULT are failures with no number, so a table that listed "
                  "only numeric codes would leave the commonest failures unexplained. Second, a "
                  "test result of ERROR is not one code: the numeric code is in the Errors tab of "
                  "the result, and it must be read from there and recorded. An entry that says only "
                  "\"ERROR\" destroys the information needed to fix it."),
            ("h2", "8.1 Failures without a numeric code"),
            ("tbl",
             ["Result", "Meaning", "Cause (Cepheid)", "Operator action", "If it persists", "Module "
              "out of service", "Responsible"],
             [
                 ["INVALID",
                  "An internal control failed: the Sample Processing Control, the Sample Adequacy "
                  "Control or the internal quantification standard, depending on the assay.",
                  "The reaction was inhibited — blood, food particles or pus in the specimen — or "
                  "too little specimen was added, or the kit was stored outside its conditions.",
                  "Check the specimen for blood or particles before repeating; repeat once with a "
                  "new cartridge from the same specimen.",
                  "If INVALID repeats, the specimen is inhibiting the assay: do not test it a third "
                  "time. Request a new specimen and record the event.",
                  "No, unless it repeats on one module.",
                  "Operator; supervisor for the repeat specimen request."],
                 ["NO RESULT",
                  "The run did not complete and insufficient data were collected.",
                  "Power interruption during the run, the Stop Test function was used, or the "
                  "computer froze.",
                  "Check the power supply and the UPS; repeat with a new cartridge. Do not use Stop "
                  "Test unless the run must be abandoned, and never leave other applications "
                  "running on the instrument computer.",
                  "If NO RESULT follows a power event, record the event and inform the Laboratory "
                  "Manager — the UPS is at fault, not the cartridge.",
                  "No.", "Operator."],
             ],
             [1.8, 3.0, 3.2, 3.2, 3.0, 1.4, 1.8]),
            ("h2", "8.2 Failures with a numeric code"),
            ("p", "Read the code from the Errors tab of the result before removing the cartridge. "
                  "The Xpert assays in use differ: the TB assays load a liquefied specimen with "
                  "sample reagent, while the whole-blood and dried-blood-spot assays load a small "
                  "volume of blood directly, so the viscosity and particle causes below apply "
                  "mainly to the TB and stool work, and the inhibition causes to the blood work."),
            ("tbl",
             ["Code", "Message", "Meaning and cause (Cepheid)", "Operator action", "If it persists",
              "Module out of service", "Responsible"],
             [
                 ["5006 / 5007 / 5008",
                  "Probe Check failed.",
                  "The Probe Check Control failed and the run was stopped before amplification. "
                  "Cause: specimen too viscous or the wrong volume added; the reaction tube "
                  "improperly filled or containing bubbles; damaged reagents from incorrect "
                  "cartridge storage; overdue maintenance or Xpert Check.",
                  "Check that the specimen is completely liquefied and that the volume is the one "
                  "the assay's instructions require; load without introducing bubbles; use a new "
                  "cartridge; confirm the cartridges have been stored within their range and are "
                  "in date.",
                  "Repeat once more. If 5006 or 5007 recurs on the same module, suspect the module "
                  "and exclude it; if it recurs across modules, suspect the cartridge lot and "
                  "quarantine it.",
                  "Yes if it recurs on one module.", "Operator; supervisor to exclude the module."],
                 ["5011", "Signal loss detected in the amplification curve.",
                  "Loss of tube pressure: the cartridge tube is not airtight or the cartridge valve "
                  "is not working. Can be a module fluidics fault.",
                  "Repeat with a new cartridge.",
                  "If 5011 recurs, the module's fluidics are suspect: exclude the module and call "
                  "service.", "Yes if it recurs.", "Supervisor; service provider."],
                 ["2008 / 2005", "Abnormal pressure detected. / Pressure exceeds the maximum.",
                  "The pressure reading exceeded the maximum: the cartridge filter is clogged by a "
                  "viscous specimen or by particles, or the pressure sensor has failed.",
                  "Ensure the specimen is fully liquefied and free of solid particles; repeat with "
                  "a new cartridge.",
                  "Run a new cartridge with sample reagent only: if the error persists with no "
                  "specimen, the module is at fault. Exclude it and call service.",
                  "Yes if it persists without specimen.", "Operator, then supervisor."],
                 ["2037", "Cartridge integrity test failed.",
                  "At the start of the run the instrument checks the pressure in the tube; the "
                  "cartridge is not airtight.",
                  "Discard the cartridge — do not use it — and repeat with a new one from a "
                  "different lot if possible.",
                  "If several cartridges from one lot fail, quarantine the lot, record the lot "
                  "number and inform the supplier and the Laboratory Manager.",
                  "No.", "Operator; supervisor for the lot."],
                 ["2120 / 2122 / 2123 / 2124 / 2126 / 2127",
                  "Module communication loss detected.",
                  "Communication between the module and the software was lost: unstable mains or "
                  "UPS supply, a loose or damaged Ethernet cable, a poor connection at the gateway "
                  "board, or excessive room temperature.",
                  "Check and reseat the Ethernet cable between the computer and the instrument; "
                  "restart the system; check the room temperature.",
                  "Check the UPS and fit a surge protector. If the code returns, the gateway board "
                  "or a cable is at fault: call service.",
                  "Yes if it recurs.", "Operator; supervisor for the UPS; service provider."],
                 ["1001 / 1002 / 2014 / 4009 / 4010 / 4017 / 3074 / 3075",
                  "Temperature or heater failure.",
                  "A heater component in a module has failed, the ambient temperature is outside "
                  "the range, or a fan has failed or its filter is blocked.",
                  "Check the room temperature (the instrument requires 15–30 °C), that there is "
                  "10–15 cm clearance on every side, that the rear fan runs and that the filter is "
                  "clean. Check the module temperature in the Maintenance menu.",
                  "If the error follows one module, that module's heater has failed: exclude it and "
                  "call service.", "Yes, the affected module.", "Operator, then service."],
                 ["Any other code", "—",
                  "Not in this SOP.",
                  "Stop. Record the exact code, the module and the assay.",
                  "Look the code up in the current GeneXpert Dx Operator Manual, chapter 9, and in "
                  "the manufacturer's troubleshooting guide; if it is not there, call service "
                  "quoting it. Report it to the Laboratory Manager so it is added here by "
                  "amendment.",
                  "Yes until it is understood.", "Supervisor."],
             ],
             [2.2, 2.6, 3.2, 3.0, 3.0, 1.5, 1.7]),
            ("h2", "8.3 Module out of service, and the stuck cartridge"),
            ("p", "A module is taken out of service by excluding it in the software — Setup, then "
                  "System configuration, then Select Modules — and the exclusion is recorded in the "
                  "instrument log with the reason and the code. An excluded module that has been "
                  "repaired is not returned to use until it has completed the manufacturer's check "
                  "and a control has passed; the exclusion is then removed and the removal "
                  "recorded."),
            ("p", "A cartridge stuck in a module is a supervisor-level event. The door is first "
                  "opened from the software (Maintenance, Open Module Door). Only if that fails, and "
                  "only with the service provider's instruction and in consultation with the "
                  "Laboratory Manager, is the manual removal the manufacturer describes performed. "
                  "The instrument must be disconnected from power and the computer first, and it "
                  "must never be turned upside down or on its side: spilled cartridge contents "
                  "destroy the module."),
            ("h2", "8.4 Error rate"),
            ("p", "The laboratory's objective is an error rate at or below 5% per assay per month, "
                  "which is the figure the World Health Organization gives for molecular "
                  "diagnostics. The rate is calculated from the instrument log and reviewed "
                  "monthly. An assay above the objective for two consecutive months is a "
                  "nonconformity and is investigated by cause, which is only possible because the "
                  "numeric code is now recorded."),
        ]),
        ("9. Alere Pima CD4 analyser", [
            ("p", "The Pima reports a failure as an invalid test with a numeric code, and it "
                  "initialises with a self-test that reports its own codes. Published evaluations of "
                  "this analyser in routine use found 10% to 13% of tests invalid, most of them "
                  "attributable to how the cartridge was filled rather than to the instrument — "
                  "which is why the operator actions below concentrate on filling and on which "
                  "specimen type is being used."),
            ("tbl",
             ["Code", "Meaning", "Cause", "Operator action", "If it persists", "Out of service",
              "Responsible"],
             [
                 ["810", "Invalid test — channel filling error.",
                  "The cartridge channel was not filled correctly with blood.",
                  "Repeat with a new cartridge, filling slowly to the mark and without bubbles; do "
                  "not squeeze the finger repeatedly and do not use a clotted specimen.",
                  "If it recurs on venous specimens, suspect the cartridge lot; if on capillary "
                  "specimens, retrain the collection technique.",
                  "No.", "Operator; supervisor reviews the technique."],
                 ["850", "Invalid test — the most frequently reported Pima code, most often seen "
                  "with blood applied directly from a finger-stick.",
                  "Improper loading and filling of the cartridge, or a specimen affected by the "
                  "collection — haemolysis, interstitial fluid, a clotted or short sample. The "
                  "cartridge and the analyser cannot be excluded as causes.",
                  "Repeat with a new cartridge. Prefer blood collected into an EDTA microtube or a "
                  "venous EDTA tube over direct finger-stick application: the published data show "
                  "fewer invalid results and the possibility of a repeat without a second prick.",
                  "If it recurs on one analyser with venous specimens, stop using the analyser and "
                  "call the distributor; keep a sample of the failing lot.",
                  "Yes if venous specimens fail repeatedly.", "Operator; supervisor for the "
                  "analyser."],
                 ["Other invalid test codes", "An invalid test with a code not in this table.",
                  "Not established.",
                  "Record the exact number, repeat once with a new cartridge.",
                  "Report it to the Laboratory Manager, who obtains the meaning from the "
                  "distributor and adds it here by amendment.",
                  "No.", "Supervisor."],
                 ["Self-test failure code at switch-on", "The analyser failed its initialisation "
                  "and self-test.",
                  "A hardware fault, or the analyser was moved or the temperature changed.",
                  "Switch off, wait, and restart. Allow 30 minutes for temperature equilibration if "
                  "the analyser has been moved.",
                  "Record the code and call the distributor. The analyser is maintenance-free and "
                  "has no user-serviceable parts; it is not opened.",
                  "Yes.", "Supervisor; distributor."],
             ],
             [2.2, 2.8, 3.2, 3.4, 2.9, 1.4, 1.7]),
            ("p", "A Pima result is reported only from a test the analyser completed. An invalid "
                  "test is not a low CD4 count, and a repeat that gives a result does not mean the "
                  "first attempt may be forgotten: both attempts are recorded."),
        ]),
        ("10. Failures of the specimen rather than the instrument", [
            ("p", "Some of what arrives at the bench as an instrument error is a specimen problem "
                  "that the instrument has correctly detected. These are handled here only far "
                  "enough to tell them apart; the rejection criteria themselves belong to the "
                  "examination SOP of each section and are not repeated."),
            ("tbl",
             ["What is seen", "Instrument or specimen?", "How to tell", "What happens next"],
             [
                 ["Clotted specimen; short sample",
                  "Specimen.",
                  "Visible clot, or a fill volume below the tube's minimum.",
                  "Reject, record in the specimen rejection log, and request a new specimen. Do not "
                  "repeat on the analyser hoping for a result."],
                 ["Repeated INVALID on the same specimen, or a control passed on the same run",
                  "Specimen (inhibition), not the instrument.",
                  "The control is acceptable and the specimen fails repeatedly.",
                  "Do not test a third time; request a new specimen."],
                 ["The same code on one module, with the other modules working",
                  "Module.",
                  "The other modules run the same assay correctly.",
                  "Exclude that module, record it, and call service."],
                 ["The same code on every module or the whole analyser",
                  "Instrument, environment or consumable lot.",
                  "Check the room temperature, the power supply and the cartridge or reagent lot.",
                  "Stop, record, and follow sections 5 to 9 and the maintenance SOP."],
                 ["One operator's failures much higher than the others",
                  "Neither: technique.",
                  "The monthly review in 4.4 shows it.",
                  "Retraining and supervised re-assessment, recorded in the training file."],
             ],
             [3.4, 2.6, 4.0, 6.0]),
        ]),
        ("11. Reagents, controls and consumables", [
            ("b", [
                "An error is not treated as a reagent fault until the obvious has been excluded: "
                "expiry, storage temperature, lot number, and whether the kit was opened and "
                "handled as the manufacturer directs.",
                "A lot that produces the same code repeatedly across more than one module is "
                "quarantined, the lot number is recorded, and the supplier is informed. The "
                "Laboratory Manager decides whether the remaining stock is used and records the "
                "decision.",
                "A defect in a reagent, kit or device that could have affected a patient result is "
                "reported to the Laboratory Manager, who records it and informs the supplier and, "
                "where required, the Rwanda Food and Drugs Authority and the national reference "
                "programme.",
            ]),
        ]),
        ("12. Quality control, and what a passed control does not prove", [
            ("p", "A passed control shows that the analyser and the reagents were working when the "
                  "control ran. It does not show that a failed test was caused by the specimen, and "
                  "it does not authorise the release of a result that came with an error. The two "
                  "questions are separate: the control is evidence about the run, the error code is "
                  "evidence about the test."),
            ("p", "When quality control fails, the preparation is stopped and the failure is "
                  "handled as nonconforming work under the quality control procedure of the "
                  "section. Nothing produced since the last acceptable control is released until it "
                  "has been reviewed."),
        ]),
        ("13. Nonconforming work and the corrective action that follows", [
            ("p", "Every error that is not cleared by the manufacturer's stated operator action is a "
                  "nonconformity, and so is every error that recurs, every module taken out of "
                  "service and every error rate above the objective. The procedure is the "
                  "laboratory's: stop the affected work; do not release what it produced; record "
                  "the event; assess whether results already released are affected and act on them; "
                  "identify the cause; take the corrective action; check that the action worked, "
                  "using the same error-rate review that raised it."),
            ("p", "The cause is recorded as one of four, because these four need four different "
                  "responses: specimen, operator technique, consumable or reagent, instrument. "
                  "Until this SOP, the laboratory could not assign a cause at all, because the code "
                  "was not recorded."),
        ]),
        ("14. Records", [
            ("tbl", ["Record", "Where kept", "Retention"],
             [
                 ["Instrument error log — code, date and time, module, assay, action, outcome, "
                  "initials", "Instrument logbook of each analyser",
                  "At least 10 years"],
                 ["Exclusion and return of a module", "Instrument logbook and maintenance log",
                  "At least 10 years"],
                 ["Service call, intervention report and engineer's report",
                  "Equipment file of each analyser", "Life of the instrument"],
                 ["System log export (cobas, GeneXpert) sent to the manufacturer",
                  "Equipment file", "Life of the instrument"],
                 ["Monthly error-rate review and the corrective actions raised from it",
                  "Quality file", "As the laboratory records policy requires"],
                 ["Nonconformity and corrective action records", "Quality file", "As above"],
                 ["Training and re-assessment after a technique-related failure",
                  "Personnel file", "Duration of employment plus the period national rules "
                  "require"],
             ],
             [6.0, 5.0, 5.0]),
            ("p", "Records are made at the time of the activity, in ink, and corrected as described "
                  "in section 4.3. Records are legible, traceable to the person or the user account, "
                  "and retrievable (ISO 15189:2022, 8.4; "
                  f"{ADM_DOC_CONTROL})."),
        ]),
        ("15. Responsibilities and Authorisation", [
            ("tbl", ["Task", "Responsible Person(s)"],
             [
                 ["Authorise this SOP and provide the resources to apply it",
                  "Hospital Director General"],
                 ["Ensure every user is trained, assessed as competent and authorised, and that "
                  "this is recorded; review this SOP at least every two years",
                  "Laboratory Manager"],
                 ["Read the code, record it, act on it within the limits of section 4.2, and "
                  "escalate at the point the table says",
                  "Authorised Laboratory Biotechnologists"],
                 ["Exclude a module; authorise resumption of testing; decide the fate of a "
                  "quarantined lot; review the monthly error rate",
                  "Laboratory Manager or Senior Biotechnologist"],
                 ["Call the service provider and keep the equipment file complete",
                  "Laboratory Manager and Quality Officer"],
                 ["Report a defect in a reagent, kit or device to the supplier and the authority",
                  "Laboratory Manager"],
                 ["Add a newly observed code to this SOP by amendment", "Laboratory Manager"],
             ],
             [8.5, 8.5]),
            ("p", "Authorisation. Only staff trained on this SOP, assessed as competent and "
                  "authorised by the Laboratory Manager act on an error code. Staff in training act "
                  "under direct supervision, and the supervisor countersigns. Training, competence "
                  "and authorisation are recorded in the personnel file "
                  f"({code('training')})."),
        ]),
        ("16. Training and competence for this SOP", [
            ("b", [
                "A new member of staff is trained on the analysers they will use, not on all of "
                "them, and on this SOP once.",
                "Competence includes: finding the code on the instrument; recording it correctly; "
                "choosing the right row of the table; knowing the point at which they must stop; "
                "and knowing who to call.",
                "A competence assessment is repeated when a new analyser is introduced, when this "
                "SOP is amended, and after a failure attributed to technique.",
            ]),
        ]),
        ("17. Risk Management, Nonconforming Work and Contingency", [
            ("tbl", ["Risk", "Control measure"],
             [
                 ["A result released from a failed test",
                  "An error is never a result; the release check in each examination SOP; the "
                  "recording rule in 4.3"],
                 ["The code is not recorded, so the failure cannot be investigated",
                  "Section 4.3 and the monthly review in 4.4 — the specific gap this SOP closes"],
                 ["The same fault is fixed repeatedly without the cause being found",
                  "The four-way cause classification in section 13 and the monthly review"],
                 ["A module deteriorates while the others mask it",
                  "Error rate by module, reviewed monthly; exclusion at the second recurrence"],
                 ["An operator works outside the manufacturer's limits and damages the instrument",
                  "Section 4.2; the service call instead of the improvised repair"],
                 ["A quarantined lot is put back into use",
                  "Quarantine is recorded and the Laboratory Manager alone releases it"],
                 ["Testing stops and the requester is not told",
                  "Every escalation includes informing the requester of the delay"],
                 ["The instrument is unavailable and specimens expire",
                  "Refer specimens as the examination SOP's contingency requires; storage limits "
                  "are those in that SOP"],
             ],
             [6.0, 11.0]),
        ]),
        ("18. Supporting Documents and References", [
            ("h2", "18.1 Controlled documents of this laboratory"),
            ("b", [
                f"Document and Record Control ({ADM_DOC_CONTROL}).",
                f"Signing of SOPs, and the Training, Competence and Authorisation Record "
                f"({code('training')}).",
                f"Use, Care and Preventive Maintenance of Hematology Analysers ({ADM_EQUIP_MAINT}).",
                f"Use, Care and Preventive Maintenance of the GeneXpert Dx System "
                f"({code('genexpert_maint')}).",
                f"Examination of Blood by the Sysmex XN-330 ({code('xn330')}); by the Sysmex XQ-320 "
                f"({code('xq320')}).",
                f"CD4 Enumeration by the Alere Pima CD4 Analyser ({code('pima')}).",
                f"Glucose Testing using Cobas c 111 ({code('glucose')}).",
                f"HIV Early Infant Diagnosis, and the hepatitis B and C viral loads, on the "
                f"GeneXpert: MOL-SOP-002, -003 and -004, in the new MOL series "
                f"(RliDHL-{'MOL'}-SOP).",
                "Laboratory Waste Management procedure; procedure for nonconformities and "
                "corrective action; quality control procedure; Laboratory Safety Manual.",
            ]),
            ("h2", "18.2 Manufacturer documentation consulted for this SOP"),
            ("p", "The codes, causes and corrective actions in sections 5 to 9 are taken from these "
                  "documents. A copy of each is held at the bench and in the equipment file, and the "
                  "revision in force is the one cited by the examination SOP of the section."),
            ("b", [
                "Sysmex. Automated Hematology Analyzer XN-L Series XN-530/XN-430/XN-330, Basic "
                "Operation; and the XN Series Instruction for Use, chapter \"Troubleshooting\", "
                "error message list. Sysmex Corporation, Kobe.",
                "Sysmex. XN-L Series Flagging Guide, document 1399-LSS — for the IP messages that "
                "section 5 states are not errors.",
                "Sysmex. Automated Hematology Analyzer XQ-320 Instructions for Use — the message "
                "list and maintenance intervals.",
                "Roche Diagnostics. cobas c 111 Operator's Manual, version 3.0, chapter 9, "
                "\"Messages and alarms\", and the list of alarm messages from which every code in "
                "section 7 is taken — including 5040, the Peltier element fault.",
                "Cepheid. GeneXpert Dx System Operator Manual, chapter 9, troubleshooting; and the "
                "Genexpert Dx troubleshooting guidance for INVALID, NO RESULT, 5006/5007, 2008, "
                "2037, 2127, 5011 and the temperature codes.",
                "Global Laboratory Initiative. Xpert MTB/RIF Training Package, Module 9: "
                "Troubleshooting — the source for the corrective actions on the code table and for "
                "the NO RESULT and INVALID definitions.",
                "Abbott (Alere). Pima Analyser User Guide, sections 8 and 9 — error codes and "
                "troubleshooting; WHO Prequalification public report PQDx 0099-032-00, Pima CD4 "
                "Test.",
                "Evidence for the invalid-test rates quoted in sections 8 and 9: published "
                "evaluations of the Pima CD4 analyser in routine use, including the multi-country "
                "study reporting 13.1% invalid results and the specimen-type evaluation reporting "
                "error 850 as the most frequent code.",
            ]),
            ("h2", "18.3 Standards"),
            ("b", [
                "ISO 15189:2022 — Medical laboratories: requirements for quality and competence, in "
                "particular 7.3.1, 7.3.6, 7.3.7, 7.5 and 8.4 to 8.7.",
                "WHO. Laboratory Quality Management System (LQMS) Handbook — equipment "
                "troubleshooting and maintenance.",
            ]),
        ]),
    ],
}

# ---------------------------------------------------------------------------
# ADM SOP 025 -- External Quality Assessment and Proficiency Testing
#
# Serial 025, taken on the Laboratory Manager's instruction to fill the gap
# between ADM-024 and ADM-026 rather than skip past it.
#
# Written against what the laboratory actually does, which is recorded in
# docs/register/tb-ept-register-transcript.md: Dried Tube Specimen HIV serology
# PT from RBC/NRL, MTB PT on GeneXpert, and the quarterly NRL biochemistry and
# haematology PT that covers creatinine, glucose, ALT and the full blood count.
#
# The content follows ISO 15189:2022 7.3.7 (and 7.3.7.3 for EQA specifically),
# which is not a suggestion: 7.3.7.3 requires participation in EQA appropriate
# to the examinations performed, review of the data at regular intervals against
# pre-defined acceptability criteria, and defined action where performance falls
# outside them, including an assessment of whether the nonconformance is
# clinically significant. It also follows the SLIPTA checklist items 8.20 and
# 8.21, which are what the accreditation assessor will score.
# ---------------------------------------------------------------------------

EQ = REGISTER["eqa_pt"]
EQ_CODE = code("eqa_pt")
EQ_EFF = "14/11/2026"
EQ_REV = "14/11/2028"

EQA = {
    "meta": {"doc_no": EQ_CODE, "title": EQ["title"], "version": EQ["vers"]},
    "out_stem": f"{EQ_CODE} - External Quality Assessment and Proficiency Testing",
    "revision_rows": [
        ["Kelvin SAMAHANI", EQ["vers"], EQ_EFF,
         "First issue. The laboratory takes part in EQA for HIV serology, TB on GeneXpert, "
         "biochemistry and the full blood count, and no controlled document governed it: who "
         "enrols, who tests the panels, who submits the results, what happens when a result is "
         "unsatisfactory, and how the annual plan is made were all held in one person's head. "
         "Written to ISO 15189:2022, 7.3.7 and 7.3.7.3."],
    ],
    "control_rows": [
        ("SOP Number", EQ_CODE),
        ("SOP Title", EQ["title"]),
        ("Version", EQ["vers"]),
        ("Effective Date", EQ_EFF),
        ("Next Review Date", EQ_REV),
        ("Supersedes", "— (first issue; no previous SOP covered EQA or PT)"),
        ("Document Owner", "Kelvin SAMAHANI — Laboratory Manager"),
        ("Applies to", "Every examination the laboratory performs, including point-of-care "
                       "testing, and every examination referred to a laboratory whose EQA covers it"),
        ("Classification", "Controlled document — master copy: document control system "
                           f"({ADM_DOC_CONTROL})"),
    ],
    "body": [
        ("1. Introduction", [
            ("p", "Internal quality control tells the laboratory whether a run was under control. It "
                  "cannot tell the laboratory whether its results are right in the way another "
                  "laboratory's results are right. Two laboratories can each be perfectly in control "
                  "of an assay that both have calibrated wrongly, and both will report the same wrong "
                  "answer with complete confidence. Only comparison with an outside body detects "
                  "that, and that comparison is external quality assessment."),
            ("p", "This laboratory already participates. The records show Dried Tube Specimen "
                  "proficiency testing for HIV serology from the National Reference Laboratory of the "
                  "Rwanda Biomedical Centre, MTB proficiency testing on the GeneXpert, and the "
                  "quarterly NRL biochemistry and haematology programme covering creatinine, glucose, "
                  "ALT and the full blood count. In the January–March 2026 event the laboratory's "
                  "peers included 56 public hospital laboratories and 36 private laboratories."),
            ("p", "What was missing was not participation but governance. No document said who "
                  "enrols for what, who tests the panel, who signs the submission, what must happen "
                  "when a result is unsatisfactory, or how the laboratory knows that it is enrolled "
                  "in EQA for everything it reports. This SOP supplies that. It does not replace the "
                  "provider's instructions; where this SOP and the provider's instructions differ, "
                  "the provider's instructions govern that event and the difference is recorded."),
            ("p", "The reason this matters beyond the certificate: an unsatisfactory EQA result is "
                  "the clearest early warning a laboratory gets that patients may be receiving wrong "
                  "results. Handled as a formality it becomes a tick in a file. Handled as this SOP "
                  "requires, it becomes the trigger for finding out what went wrong and whether a "
                  "patient was affected."),
        ]),
        ("2. Purpose", [
            ("p", "To set out how this laboratory plans, performs, records, reviews and acts on "
                  "external quality assessment and proficiency testing, so that participation is "
                  "complete and demonstrable, the panels are tested exactly as patient specimens are "
                  "tested, every unsatisfactory result is investigated to its cause, and corrective "
                  "action is verified as effective rather than merely declared."),
        ]),
        ("3. Scope", [
            ("p", "This SOP applies to every examination on the laboratory's Scope of Test list "
                  "that is performed at Ruli District Hospital, including point-of-care examinations, "
                  "and to examinations referred to another laboratory where the referring laboratory "
                  "relies on that laboratory's EQA. It covers the full EQA cycle: the annual plan, "
                  "enrolment, receipt and handling of panels, testing, submission, review of the "
                  "report, investigation and corrective action, verification of effectiveness, and "
                  "the reports to management."),
            ("p", "**It governs one direction of quality assessment only: the assessment of this "
                  "laboratory.** The laboratory takes part in programmes run by the National "
                  "Reference Laboratory and other providers, and this SOP is about that "
                  "participation — the panel arrives, the laboratory tests it, the result goes back, "
                  "the report is reviewed."),
            ("p", "**It does not govern the assessment the laboratory performs on others.** The "
                  "laboratory also acts as a **referral laboratory** for the eight health centres in "
                  "its network: it re-reads their stained slides, compares its reading with theirs, "
                  "scores the discordance and reports back. That is a different activity, running in "
                  "the opposite direction, with different criteria, different records and a different "
                  "regulator's interest. It is governed by the relevant **examination SOP of the "
                  "section performing it** — for malaria films, `PAR-SOP-008`; for TB smear, "
                  "`TB-SOP-002` and `TB-SOP-003` — with the records kept as the records of those "
                  "examinations."),
            ("p", "**The two are never combined in one figure.** A health-centre discordance rate is "
                  "not the laboratory's EQA performance, and the laboratory's PT z-scores are not a "
                  "statement about the health centres. Where both appear in the same management "
                  "report they appear in separate sections, separately titled, with the subject of "
                  "each stated: **\"External quality assessment of this laboratory\"** and "
                  "**\"Re-reading of health-centre slides performed by this laboratory\"**. A single "
                  "composite score mixing them is not permitted, because it would describe neither."),
            ("p", "It does not govern internal quality control, which is required by the quality "
                  "control procedure and by the examination SOPs of each section. Nor does it govern "
                  "the laboratory's participation in a provider's **on-site assessment or mentorship "
                  "visit**, which is handled under the nonconformity and improvement procedure even "
                  "where the same provider is involved."),
        ]),
        ("4. Definitions", [
            ("tbl", ["Term", "Meaning in this laboratory"],
             [
                 ["External quality assessment (EQA)",
                  "Evaluation of the laboratory's performance against pre-established criteria by "
                  "comparison with other laboratories — ISO 15189:2022, 3.10. Encompasses "
                  "proficiency testing, re-testing and onsite assessment."],
                 ["Proficiency testing (PT)",
                  "EQA in which samples of known but undisclosed composition are circulated and "
                  "participants' results are scored against the expected values or against the "
                  "consensus of participants."],
                 ["Panel / survey",
                  "One distribution of EQA material, with its provider code, shipment code and "
                  "closing date."],
                 ["Target value",
                  "The value against which the participant's result is scored: a reference value, "
                  "the consensus value of the peer group, or the value assigned by the scheme "
                  "organiser."],
                 ["SDI / z-score",
                  "The standard deviation index or the standard score used by the scheme to express "
                  "how far the result lies from the target, in units of the acceptable standard "
                  "deviation. Different schemes use different names for the same idea."],
                 ["Satisfactory / unsatisfactory",
                  "Satisfactory means inside the acceptability criteria the scheme publishes and "
                  "this SOP adopts. Unsatisfactory means outside them. The laboratory does not "
                  "redefine a provider's criteria to make a result satisfactory."],
                 ["Re-testing / re-reading (as performed by this laboratory)",
                  "The laboratory re-reads slides or re-tests samples **of the health centres it "
                  "serves**, and reports the discordance back. This is quality assessment performed "
                  "**for others**. It is not this laboratory's own EQA and is governed by the "
                  "examination SOP concerned."],
                 ["Slide re-reading (as performed on this laboratory)",
                  "Where a provider re-reads **this laboratory's** slides, the laboratory is the "
                  "participant and the activity is EQA under this SOP, even though no panel of "
                  "manufactured samples is involved."],
                 ["Alternative approach",
                  "Where no EQA programme exists for an examination: comparison with another "
                  "laboratory, re-testing, analysis of commutable reference material, clinical "
                  "correlation, or a blind re-test, with the rationale recorded — "
                  "ISO 15189:2022, 7.3.7.3, note."],
             ],
             [4.5, 12.0]),
        ]),
        ("5. Responsibilities and Authorisation", [
            ("tbl", ["Task", "Responsible Person(s)"],
             [
                 ["Authorise this SOP and provide the funds for enrolment and participation",
                  "Hospital Director General"],
                 ["Own the EQA programme: the annual plan, the enrolment, the register of panels, "
                  "the reports to management, and this SOP", "Laboratory Manager"],
                 ["Prepare the annual EQA plan against the Scope of Test list and check that every "
                  "examination is covered or has a recorded alternative", "Laboratory Manager and "
                  "Quality Officer"],
                 ["Receive, log and store the panel; test it; submit the results by the closing date",
                  "Section head or designated authorised biotechnologist"],
                 ["Approve the submission before it is sent", "Laboratory Manager or Senior "
                  "Biotechnologist"],
                 ["Review each provider report and assess it against 8.4", "Laboratory Manager"],
                 ["Investigate an unsatisfactory result, identify the cause, propose and implement "
                  "the corrective action", "Section head, with the Laboratory Manager"],
                 ["Verify that the corrective action was effective", "Quality Officer and "
                  "Laboratory Manager"],
                 ["Assess whether any patient result may have been affected, and act if so",
                  "Laboratory Manager"],
                 ["Keep the EQA file complete and retrievable", "Quality Officer"],
             ],
             [8.0, 9.0]),
            ("p", "Authorisation. Only staff trained on this SOP and authorised by the Laboratory "
                  "Manager plan or review EQA. Any authorised biotechnologist of the section "
                  "performs the panel, because the whole value of the exercise is that the panel is "
                  "handled by the person who handles patients' specimens "
                  f"({code('training')})."),
        ]),
        ("6. The annual EQA plan", [
            ("p", "Before the end of each calendar year the Laboratory Manager prepares an EQA plan "
                  "for the coming year. It is a controlled record, approved by the Hospital Director "
                  "General, and it is the document an assessor asks for first."),
            ("p", "The plan is built from the Scope of Test list, not from memory. Every examination "
                  "on that list is entered against one of three columns: **covered by an EQA "
                  "programme** (naming the programme and the provider), **covered by a recorded "
                  "alternative approach** (naming it, with the rationale), or **not covered** "
                  "(naming the reason and the action being taken to close it). An examination in the "
                  "third column is a finding, and it is reported to management review."),
            ("tbl", ["Examination", "Programme and provider", "Frequency", "Status"],
             [
                 ["HIV serology — rapid testing per the national algorithm",
                  "Dried Tube Specimen PT for HIV antibody diagnostics, RBC / National Reference "
                  "Laboratory", "Per NRL survey schedule",
                  "In place. Panels on file for 2022, 2023 and 2024."],
                 ["TB — Xpert MTB/RIF Ultra on GeneXpert",
                  "MTB proficiency testing, RBC / National Reference Laboratory",
                  "Per NRL survey schedule",
                  "In place. Panels on file for 2023 and 2024."],
                 ["TB — smear microscopy (ZN and auramine bacilloscopy)",
                  "Where the provider re-reads **this laboratory's own** slides: the laboratory is "
                  "the participant and the activity is EQA under this SOP",
                  "As the provider schedules it",
                  "**Not evidenced.** No report was found of a provider re-reading this laboratory's "
                  "TB slides. The reports on file are the laboratory re-reading the health centres' "
                  "slides — the opposite direction, governed by `TB-SOP-002`/`TB-SOP-003`, not by "
                  "this SOP. Recorded as a gap in section 6.4."],
                 ["Malaria — thick and thin film microscopy",
                  "Where the provider re-reads **this laboratory's own** slides",
                  "As the provider schedules it",
                  "**Not evidenced**, for the same reason. The malaria re-reading reports on file "
                  "assess the health centres, not this laboratory; they belong to `PAR-SOP-008`."],
                 ["Biochemistry — creatinine, glucose, ALT",
                  "NRL biochemistry and haematology EQA, quarterly",
                  "Quarterly", "In place. January–March 2026 report on file."],
                 ["Haematology — full blood count: RBC, Hb, Hct, WBC, platelet count",
                  "NRL biochemistry and haematology EQA, quarterly",
                  "Quarterly", "In place. January–March 2026 report on file."],
                 ["HIV early infant diagnosis (Xpert HIV-1 Qual XC)",
                  "No PT report held by the laboratory", "—",
                  "**GAP.** Either enrolled with RBC/NRL and the reports are not filed, or not "
                  "enrolled. The Laboratory Manager establishes which, within one month of this SOP "
                  "taking effect, and records the answer."],
                 ["HBV viral load, HCV viral load, TB LAM", "No PT report held", "—",
                  "**GAP.** Same action. If no programme exists, an alternative approach under "
                  "section 9 is designed and recorded."],
                 ["CD4 enumeration (Alere Pima)", "PT results on file from 2023", "—",
                  "Confirm the programme is still active and file the intervening reports."],
             ],
             [4.2, 4.6, 2.4, 5.3]),
        ]),
        ("6.4 Gaps in this laboratory's own EQA coverage", [
            ("p", "The following examinations are on the Scope of Test list but **no assessment of "
                  "this laboratory** was found on file for them. Each is either an enrolment to be "
                  "confirmed or a participation to be arranged, and none may be closed by pointing "
                  "at the re-reading the laboratory performs for its health centres — that is a "
                  "different activity and gives no evidence about this laboratory's own performance:"),
            ("tbl",
             ["Examination", "What was searched", "Position at 15 November 2026", "Action and owner"],
             [
                 ["TB smear microscopy (ZN and auramine)",
                  "NRL re-reading and PT reports held by the laboratory",
                  "**No report of this laboratory's own slides being re-read.** The four NRL TB "
                  "microscopy reports on file assess Nyange HC and Rwankuba HC, not Ruli DH",
                  "Confirm with NRL whether the laboratory is enrolled for its own slides; record "
                  "the answer either way. Quality Manager, within one month"],
                 ["Malaria film microscopy",
                  "NRL re-reading and PT reports held by the laboratory",
                  "**No report of this laboratory's own slides being re-read.** The four NRL "
                  "malaria reports on file assess the eight health centres",
                  "As above. Quality Manager, within one month"],
                 ["HBV viral load, HCV viral load, HIV early infant diagnosis, TB LAM",
                  "PT reports and provider correspondence",
                  "No PT report held for any of the four",
                  "Establish whether the laboratory is enrolled and the reports are unfiled, or not "
                  "enrolled at all. Laboratory Manager, within one month"],
             ],
             [3.8, 3.4, 5.2, 4.6]),
            ("p", "**Why this section exists.** It is easy to read a quarterly NRL report with the "
                  "laboratory's name in the header and treat it as the laboratory's EQA result. It is "
                  "not. On both malaria and TB microscopy the subject of those reports is the health "
                  "centre, the reader is this laboratory, and the beneficiary is the health centre's "
                  "patient. If this laboratory is never itself assessed on smear microscopy, that is "
                  "a genuine gap in its own quality assurance and it is recorded here as one."),
        ]),
        ("7. Enrolment and participation", [
            ("b", [
                "The laboratory is enrolled in a programme for every examination in the first column "
                "of the plan. Enrolment is renewed before it lapses; the renewal date is entered in "
                "the EQA register so that a lapse is visible before it happens.",
                "Where a provider allocates a participant code and a site code, both are recorded and "
                "used on every submission. The laboratory's participant code on the MTB programme "
                "the laboratory holds is **20044**; a submission under any other code is traced "
                "before it is sent.",
                "Where two laboratories share a programme, or an examination is performed on more "
                "than one analyser, the panel is tested on the analyser in routine use for patients "
                "and the submission states which one was used.",
                "Where the laboratory starts a new examination, the plan is amended and an EQA route "
                "identified before results are reported — not after the first accreditation visit.",
                "The laboratory does not share its panel, its result or its expected values with "
                "another laboratory, and does not accept another laboratory's results as its own.",
            ]),
        ]),
        ("8. Handling a panel", [
            ("h2", "8.1 On receipt"),
            ("p", "The panel is opened by the section head, who records on the EQA register: the "
                  "provider, the programme, the survey or shipment code, the date received, the "
                  "storage conditions the provider requires, the closing date for results, and who "
                  "is assigned to test it. Any damage, delay in transit or breach of the transport "
                  "conditions is recorded at once and reported to the provider before testing."),
            ("p", "NRL's own recommendation is that PT panels are tested **immediately upon "
                  "receipt**, within one day. A panel left in a refrigerator while routine work is "
                  "done is being assessed as a different sample from the one the provider intended."),
            ("h2", "8.2 Testing — the rule that gives the exercise its value"),
            ("p", "The panel is tested exactly as a patient specimen is tested, by the person who "
                  "normally performs that examination, using the routine method, the routine "
                  "reagents, the routine calibrators and the routine analyser, in the routine run, "
                  "with the same quality control as a patient run. It is **not** given to the most "
                  "senior member of staff, **not** repeated until it agrees with an expectation, "
                  "**not** run on a different analyser chosen for the occasion, and **not** discussed "
                  "with another laboratory before submission."),
            ("p", "The panel's results are entered in the same worksheet or record the examination "
                  "uses for patients, marked as EQA, so that both the panel and the patient run "
                  "carry the same traceable record."),
            ("h2", "8.3 Review and submission"),
            ("b", [
                "Before submission, the Laboratory Manager or a Senior Biotechnologist reviews the "
                "result sheet for completeness: every sample answered, every required field "
                "completed — including the kit name, the lot number and the expiry date of the "
                "reagent used, which are among the commonest causes of a submission being rejected.",
                "The submission is made through the channel the provider requires, before the "
                "closing date, and the date sent and the person who sent it are recorded. NRL asks "
                "for feedback within **two weeks** of receipt of the panel.",
                "A copy of the submitted result sheet is kept in the EQA file, whether it was "
                "submitted on paper, by e-mail or in an electronic system.",
                "A submission that is late or incomplete is a nonconformity and is recorded as one, "
                "even if the result itself would have been satisfactory.",
            ]),
        ]),
        ("9. Where no EQA programme exists", [
            ("p", "ISO 15189:2022, 7.3.7.3 permits an alternative approach where no EQA programme is "
                  "available, provided the rationale is justified and the effectiveness of the "
                  "alternative is shown. The laboratory may use, in this order of preference:"),
            ("b", [
                "A comparison with another laboratory performing the same examination on "
                "commutable material — the National Reference Laboratory by preference.",
                "Re-testing of the same material by a second person, on a second analyser or by a "
                "second method.",
                "Analysis of certified reference material that is commutable with patient samples, "
                "where such material exists for the examination.",
                "A blind re-test of previously examined patient material with known results, "
                "organised so that the person performing it does not know which specimens they are.",
                "Clinical correlation, where no material can be compared and the examination's "
                "interpretation is the thing being assessed.",
            ]),
            ("p", "The choice, the reason for it and the evidence that it worked are recorded in the "
                  "EQA plan. An alternative approach that is claimed but never performed is worse "
                  "than an acknowledged gap, because it looks like coverage."),
        ]),
        ("10. Reviewing the provider's report", [
            ("p", "Each report is reviewed by the Laboratory Manager within **one month** of receipt, "
                  "and the review is recorded with a date and a signature. The review answers four "
                  "questions for every examination in the report."),
            ("tbl", ["Question", "What the reviewer does"],
             [
                 ["Did the laboratory take part?",
                  "Checks that the report shows the laboratory per programme and per examination, "
                  "and that the participation rate is 100%. A nil return or an absence from the "
                  "report is followed up with the provider on the same day."],
                 ["Was the performance satisfactory?",
                  "Applies the acceptability criteria in 8.4. Every result outside them is listed, "
                  "not just the overall score: a satisfactory overall score can hide one badly wrong "
                  "result among four correct ones."],
                 ["Is there a trend?",
                  "Reads the current result against the previous reports for the same examination. A "
                  "**satisfactory** result that is consistently positive or consistently negative of "
                  "the target is a bias and is investigated as an adverse trend, even though it "
                  "passes."],
                 ["Could a patient have been affected?",
                  "Considers every unsatisfactory or biased result against the results the laboratory "
                  "reported in the same period, under 8.5."],
             ],
             [5.0, 12.0]),
            ("p", "The review is not a formality performed for the file. Where the answer to any of "
                  "the four questions is not a clean yes, an investigation is opened."),
        ]),
        ("11. Acceptability criteria, and what counts as unsatisfactory", [
            ("h2", "11.1 The criteria the laboratory uses"),
            ("b", [
                "The scheme's own criteria govern, where it publishes them. The NRL biochemistry "
                "and haematology programme scores against **± 2 SDI**: a result outside ± 2 SDI from "
                "the target is outside the acceptable range.",
                "The Dried Tube Specimen HIV serology programme scores on the **number of correct "
                "reports** across the five-sample panel, and on whether the reported result matches "
                "the expected result for each sample. A single wrong sample on a qualitative panel "
                "is a serious error, not a rounding error.",
                "Where a scheme publishes no numeric criterion, the laboratory adopts the criterion "
                "the provider states in its instructions and records it in the plan.",
                "The laboratory may set a stricter internal target; it may never set a laxer one.",
            ]),
            ("h2", "11.2 What makes a result unsatisfactory"),
            ("p", "Any of the following is an unsatisfactory EQA result and is handled as a "
                  "nonconformity:"),
            ("b", [
                "A result outside the scheme's acceptability criteria.",
                "A qualitative result reported wrongly, on any sample of the panel.",
                "A result reported for the wrong sample, or two samples transposed.",
                "A result that could not be scored because a required field — kit name, lot, expiry "
                "date — was missing.",
                "A panel not tested, or tested but not submitted by the closing date.",
                "A satisfactory result showing a bias or a trend against previous reports.",
                "A report showing that the laboratory did not in fact participate for an examination "
                "on its scope.",
            ]),
        ]),
        ("12. Investigation and corrective action", [
            ("p", "An unsatisfactory result is investigated to its cause, not explained away. The "
                  "investigation is recorded on the nonconformity record, and it follows "
                  "ISO 15189:2022, 8.7: contain the effect, find the cause, decide whether corrective "
                  "action is needed, take it, and **verify that it worked**."),
            ("h2", "12.1 First: contain"),
            ("b", [
                "Establish whether the affected examination is still producing results and whether "
                "those results can be relied on. If not, stop reporting for that examination.",
                "Inform the Laboratory Manager the same day and, where patients may be affected, the "
                "Hospital Director General.",
                "Do not wait for the root cause before containing the risk.",
            ]),
            ("h2", "12.2 Then: find the cause"),
            ("p", "The investigation works through the whole examination process, in this order, "
                  "because that is the order in which the causes are commonest:"),
            ("tbl", ["Stage", "What is examined", "Examples of a finding"],
             [
                 ["Pre-examination",
                  "The panel's own handling: receipt date, storage, whether it was tested "
                  "immediately, whether the panel was reconstituted as the instructions require, "
                  "whether it was tested within the stated time after reconstitution",
                  "A rehydrated Dried Tube Specimen tested outside the hours the instructions "
                  "allow — one of the corrective actions NRL's own report required"],
                 ["Examination — the method",
                  "Whether the result was produced by the routine method; whether the national "
                  "testing algorithm was followed for a serology panel; whether the correct assay "
                  "definition file or programme was selected",
                  "A rapid test read against the wrong algorithm; the wrong assay file loaded for a "
                  "whole-blood versus dried-blood-spot EID panel"],
                 ["Examination — the reagent",
                  "Kit and lot in use, expiry, storage temperature, whether the same lot was in "
                  "routine use for patients, whether a new lot was introduced in the period",
                  "An expired kit in routine use — a finding about patients, not only about the "
                  "panel"],
                 ["Examination — the equipment",
                  "Calibration status, the analyser's recent internal quality control, the error "
                  "rate and error codes for the period, module performance where the analyser has "
                  "modules",
                  "A module that has been failing and was masked by the others; a calibration that "
                  "was overdue"],
                 ["Examination — the person",
                  "Whether the panel was tested by the person who normally performs the "
                  "examination; their training, competence and authorisation record; whether the "
                  "failure mode has appeared before for that person",
                  "A tester whose failure rate is high — the same signal the GeneXpert error review "
                  "produces"],
                 ["Post-examination",
                  "Transcription of the result, transposition of samples, the units reported, "
                  "rounding, the fields submitted",
                  "Two samples transposed on the submission sheet; a result reported in the wrong "
                  "unit"],
             ],
             [2.8, 7.0, 7.2]),
            ("p", "The cause is recorded in one of the four categories the laboratory uses "
                  "everywhere — **specimen**, **operator technique**, **consumable or reagent**, "
                  "**instrument** — plus a fifth for EQA, **the submission itself**, which covers "
                  "transcription, transposition and missing fields. A cause that cannot be "
                  "established is recorded as not established, with what was excluded; it is not "
                  "invented to close the record."),
            ("h2", "12.3 Then: corrective action"),
            ("b", [
                "The action addresses the cause, not the symptom. Re-testing the panel is not a "
                "corrective action; it is a re-check, and on a PT panel it is not even permissible "
                "after submission.",
                "Where the cause is technique, the action is retraining and re-assessment, recorded "
                "in the personnel file, with the effectiveness checked by supervised work or a "
                "further blind comparison.",
                "Where the cause is a reagent or a lot, the action includes reviewing the patient "
                "results produced with the same lot, informing the supplier, and, where required, "
                "the Rwanda Food and Drugs Authority.",
                "Where the cause is the instrument, the action is the service intervention and the "
                "verification the maintenance and troubleshooting SOPs require.",
                "Where the cause is the submission, the action is a change to the review step that "
                "let it through.",
            ]),
            ("h2", "12.4 Finally: verify effectiveness, and report"),
            ("p", "Corrective action is not closed when it is implemented; it is closed when it is "
                  "shown to have worked. Effectiveness is verified by the **next EQA event for the "
                  "same examination** being satisfactory, or, where the next event is too far away, "
                  "by a defined internal check — a blind comparison, a supervised batch, or a review "
                  "of internal quality control over the intervening period. The verification is "
                  "recorded against the original nonconformity."),
            ("p", "Where the provider conducts an onsite corrective-action visit — NRL does so within "
                  "three weeks of issuing a report — the visit's findings and the laboratory's "
                  "response are filed with the nonconformity."),
        ]),
        ("13. Patient results, and what an EQA failure means for them", [
            ("p", "This is the section that makes EQA worth doing. An unsatisfactory EQA result may "
                  "mean that patients examined in the same period received wrong results, and the "
                  "laboratory is required to assess that, not to assume it is not so "
                  "(ISO 15189:2022, 7.3.7.3 h and 7.5)."),
            ("p", "The assessment starts from the cause. Where the cause is confined to the panel — a "
                  "storage error affecting only the EQA material, a submission error — the "
                  "assessment is short and states why patients are not affected. Where the cause "
                  "could have affected patient work — a reagent lot, a calibration, an instrument "
                  "fault, a person's technique — the assessment covers every result the laboratory "
                  "reported in the period the cause was active, and answers:"),
            ("b", [
                "Which patients' results could have been wrong, and how many?",
                "How wrong could they have been, in the units the clinician acts on?",
                "Is the error clinically significant — would it have changed a decision about this "
                "patient?",
                "Do the results need to be repeated, and can the specimen be obtained again?",
                "Does the requester, and the patient's clinician, need to be told, and a corrected "
                "report issued?",
            ]),
            ("p", "Where a patient result is found to have been affected, a corrected report is issued "
                  "under the laboratory's procedure for reporting, the requester is informed, and the "
                  "event is recorded. A corrected report states plainly that it replaces an earlier "
                  "report and why. The laboratory does not quietly withdraw a result."),
        ]),
        ("14. Records", [
            ("tbl", ["Record", "Where kept", "Retention"],
             [
                 ["EQA annual plan, approved", "Quality file (EQA section)", "At least 6 years"],
                 ["EQA register — every panel: provider, code, dates received and submitted, "
                  "assignee", "Quality file", "At least 6 years"],
                 ["Panels as received, and the result sheet as submitted", "Quality file",
                  "At least 6 years"],
                 ["Provider reports and the record of the review, with date and signature",
                  "Quality file", "At least 6 years"],
                 ["Investigation, cause, corrective action and verification of effectiveness",
                  "Nonconformity file, cross-referenced", "At least 6 years"],
                 ["Assessment of the effect on patient results", "Nonconformity file",
                  "At least 6 years"],
                 ["Reports to management review", "Quality file", "At least 6 years"],
                 ["Records of training and re-assessment arising from an EQA failure",
                  "Personnel file", "Duration of employment plus the period national rules "
                  "require"],
             ],
             [6.4, 5.0, 4.6]),
            ("p", "Records are made at the time of the activity, in ink where paper, corrected by a "
                  "single line through the entry with the correction beside it and initials and date "
                  f"({ADM_DOC_CONTROL}). Electronic records are held with a backup and are "
                  "attributable to a named user account."),
        ]),
        ("15. Reporting to management and to the authority", [
            ("b", [
                "Every EQA report and the outcome of its review is an input to the management "
                "review (ISO 15189:2022, 8.9.2 c and e): participation, performance by examination, "
                "trends, unsatisfactory results, the corrective actions and whether they worked.",
                "An annual summary states, for every examination on the scope, whether EQA was "
                "available, whether the laboratory participated, the results obtained and the status "
                "of any open corrective action. An examination with no EQA and no recorded "
                "alternative is shown as such — that is the finding, not a paragraph to be softened.",
                "Where the provider or the regulator requires the laboratory to submit its results "
                "or its corrective action, the submission is made within the stated period and "
                "recorded.",
                "Where a scheme is withdrawn, changes its scoring, or ceases to be offered, the plan "
                "is amended and the change recorded.",
            ]),
        ]),
        ("16. Risk Management, Nonconforming Work and Contingency", [
            ("tbl", ["Risk", "Control measure"],
             [
                 ["The laboratory believes it is enrolled and is not",
                  "The annual plan is built from the Scope of Test list and the renewal date is "
                  "held in the EQA register"],
                 ["An examination on the scope has no EQA at all",
                  "The three-column plan makes the gap visible; an alternative approach under "
                  "section 9 is designed and recorded"],
                 ["The panel is tested as a special exercise, so the result says nothing about "
                  "patient work", "Section 8.2: routine person, routine method, routine analyser"],
                 ["The panel is repeated until it agrees, or discussed with another laboratory",
                  "Section 8.2 prohibits both; the submission is reviewed before it is sent"],
                 ["A submission is late or incomplete",
                  "The closing date and the review step in 8.3; a late submission is recorded as a "
                  "nonconformity"],
                 ["An unsatisfactory result is explained rather than investigated",
                  "Section 12: cause, corrective action and verification, on the nonconformity "
                  "record"],
                 ["Corrective action is declared effective without evidence",
                  "Section 12.4: effectiveness is verified by the next event or a defined internal "
                  "check"],
                 ["Patients were affected and nobody looked",
                  "Section 13: the effect on patient results is assessed every time, and a corrected "
                  "report issued where needed"],
                 ["A satisfactory result with a bias is ignored",
                  "Section 10 and 11.2: a trend against previous reports is investigated even when "
                  "it passes"],
                 ["The EQA file cannot be produced at an assessment",
                  "Section 14: the file is a defined record set, held in one place, retrievable"],
                 ["The provider's report is received and filed unread",
                  "Section 10: review within one month, recorded with a date and a signature"],
             ],
             [6.0, 11.0]),
        ]),
        ("17. Supporting Documents and References", [
            ("h2", "17.1 Controlled documents of this laboratory"),
            ("b", [
                f"Document and Record Control ({ADM_DOC_CONTROL}).",
                f"Signing of SOPs, and the Training, Competence and Authorisation Record "
                f"({code('training')}).",
                f"Troubleshooting of laboratory analysers — error codes, causes and corrective "
                f"action ({code('trouble')}).",
                "Quality Control and Quality Assurance ({code('qca')}) — internal quality "
                "control, with which this SOP is used and which it does not replace.",
                f"Pre-Examination Procedure ({ADM_PRE_EXAM}).",
                "Laboratory Scope of Test ({code('scope_test')}) — the list the annual plan is built "
                "from.",
                "Laboratory Critical Results (ADM-GUID-002); Normal Ranges for Laboratory "
                "Tests (RliDHL-ADM-GUID-003).",
                f"Laboratory Quality Control Manual ({code('qc_manual')}); Laboratory Safety "
                f"Manual ({code('safety_manual')}).",
                "The examination SOPs of every section, each of which states the EQA the section "
                "takes part in.",
                "Procedure for nonconformities and corrective action; procedure for reporting and "
                "release of results; Laboratory Waste Management procedure.",
            ]),
            ("h2", "17.2 Standards and external references"),
            ("b", [
                "ISO 15189:2022 — Medical laboratories: requirements for quality and competence, "
                "in particular 7.3.7 (and 7.3.7.3, external quality assessment), 7.5, 8.4, 8.7, "
                "8.8 and 8.9.",
                "ISO/IEC 17043 — Conformity assessment: general requirements for proficiency "
                "testing, which governs what a provider of PT must do.",
                "ASLM / SLIPTA Checklist version 3:2023, items 8.20 (procedure for EQA) and 8.21 "
                "(participation in EQA), which are scored during accreditation preparation.",
                "WHO. Laboratory Quality Management System (LQMS) Handbook — the EQA chapter.",
                "Rwanda Biomedical Centre / National Reference Laboratory. Biochemistry and "
                "Haematology External Quality Assessment Proficiency Testing performance report, "
                "January–March 2026 (RBC/AG. DIVISION MANAGER OF NRL/17302/2026), including its "
                "recommendations on immediate testing of panels, feedback within two weeks, and "
                "onsite corrective action within three weeks.",
                "RBC / National Reference Laboratory. Proficiency Testing Program for HIV Antibody "
                "Diagnostics using Dried Tube Specimen — all participants summary reports.",
                "RBC / National Reference Laboratory. MTB proficiency testing reports, participant "
                "code 20044.",
                "The current instructions, scheme description and acceptability criteria of each "
                "EQA provider the laboratory uses, held with the EQA file.",
            ]),
        ]),
    ],
}


# ---------------------------------------------------------------------------
# ADM-SOP-050 -- Re-reading of health-centre slides and re-testing of
# health-centre samples.
#
# This is the OTHER direction of quality assessment. ADM-025 governs the
# laboratory being assessed. This document governs the laboratory assessing the
# eight health centres in its network: it re-reads their stained slides,
# compares its reading with theirs, scores the discordance and reports back.
#
# The two must never be merged into one figure. A health-centre discordance rate
# says nothing about this laboratory's own accuracy.
# ---------------------------------------------------------------------------

HCR = REGISTER["hc_slide_review"]
HCR_CODE = code("hc_slide_review")
HCR_EFF = "14/11/2026"
HCR_REV = "14/11/2028"

HC_REVIEW = {
    "meta": {"doc_no": HCR_CODE, "title": HCR["title"], "version": HCR["vers"]},
    "out_stem": f"{HCR_CODE} - Re-reading of Health-Centre Slides and Re-Testing of "
                "Health-Centre Samples",
    "revision_rows": [
        ["Kelvin SAMAHANI", HCR["vers"], HCR_EFF,
         "First issue. The laboratory re-reads the stained slides of the eight health centres "
         "in its network and reports the discordance back to the National Reference "
         "Laboratory, but no controlled document governed the activity: who selects the "
         "slides, who reads them, how the discordant rate is calculated, what happens when a "
         "site fails and what the laboratory may conclude from a sample of two slides were all "
         "unwritten. Written to ISO 15189:2022, 7.3.7, and kept deliberately separate from "
         "ADM-SOP-025, which governs the assessment of this laboratory."],
    ],
    "control_rows": [
        ("SOP Number", HCR_CODE),
        ("SOP Title", HCR["title"]),
        ("Version", HCR["vers"]),
        ("Effective Date", HCR_EFF),
        ("Next Review Date", HCR_REV),
        ("Supersedes", "\u2014 (first issue; the activity was performed without a controlled "
                       "document)"),
        ("Document Owner", "Kelvin SAMAHANI \u2014 Laboratory Manager"),
        ("Applies to", "The re-reading and re-testing the laboratory performs FOR its health "
                       "centres. NOT to the assessment of this laboratory, which is "
                       "ADM-SOP-025"),
        ("Classification", "Controlled document \u2014 master copy: document control system "
                           f"({ADM_DOC_CONTROL})"),
    ],
    "body": [
        ("1. Introduction", [
            ("p", "There are two directions of quality assessment in this laboratory's work and "
                  "they are easily confused. In one, an outside body assesses **this laboratory**: "
                  "a panel arrives, the laboratory tests it, the result goes back and a report "
                  f"follows. That is governed by the EQA and proficiency testing SOP "
                  f"({code('eqa_pt')}), and it is about this laboratory's own accuracy."),
            ("p", "In the other, **this laboratory assesses its health centres**. The health centres "
                  "stain slides and read them; the slides come to Ruli; the laboratory re-reads them "
                  "blind; the two readings are compared slide by slide; the discordance is scored "
                  "and reported back to the National Reference Laboratory. The subject of that "
                  "assessment is the health centre. **The laboratory is the reader, not the "
                  "participant.**"),
            ("p", "This document governs the second direction. It exists because the activity was "
                  "being performed to a real standard without a controlled procedure behind it: the "
                  "reports on file show it working, quarter after quarter, and show four reporting "
                  "defects and two coverage problems that no procedure caught."),
            ("p", "**The two must never be merged.** A health-centre discordance rate is not this "
                  "laboratory's EQA performance. A PT z-score is not a statement about the health "
                  "centres. In any management report they appear under separate headings with the "
                  "subject stated, and no composite score mixes them \u2014 a blended figure would "
                  "describe neither and would mislead an assessor in the direction the laboratory "
                  "would least like."),
        ]),
        ("2. Purpose", [
            ("p", "To define how the laboratory selects, re-reads, scores and reports back the "
                  "slides and samples of its health centres, so that the comparison is blind, "
                  "representative, arithmetically sound and actionable \u2014 and so that a failing "
                  "site is identified early rather than after a patient is misdiagnosed."),
        ]),
        ("3. Scope", [
            ("p", "**In scope:** malaria thick and thin films from the eight health centres in the "
                  "network \u2014 Ruli, Coko, Nyange, Minazi, Rushashi, Muyongwe, Rwankuba and "
                  "Muhondo; TB smear microscopy from the centres that send slides; and any other "
                  "slide-based or sample-based re-testing the laboratory is asked to perform for a "
                  "health centre, including chemistry and haematology re-testing where a site's "
                  "result is queried."),
            ("p", "**Out of scope:** the assessment of this laboratory, which is the EQA and "
                  "proficiency testing SOP; internal quality control within the laboratory, which "
                  "is the quality control SOP; and the laboratory's own examination procedures for "
                  "malaria and TB, which are `PAR-SOP-008` for Giemsa malaria and `TB-SOP-002` and "
                  "`TB-SOP-003` for auramine and Ziehl-Neelsen microscopy. This SOP governs the "
                  "**programme**, not the staining or the reading technique."),
        ]),
        ("4. Definitions", [
            ("tbl", ["Term", "Meaning in this laboratory"],
             [
                 ["Re-reading", "Reading a slide a second time to compare two independent readings. "
                                "The second reader must not know the first result."],
                 ["Blind re-reading", "Re-reading without sight of the original result or the "
                                      "patient's identity. Without this the comparison measures "
                                      "nothing."],
                 ["Discordant rate", "The number of slides on which the two readings disagree, "
                                     "divided by the number of slides read, expressed as a "
                                     "percentage. It is calculated, never estimated."],
                 ["Concordant", "The two readings agree. For TB this means the same category \u2014 "
                                "positive, scanty or negative \u2014 not merely both non-negative."],
                 ["Scanty", "The TB smear grade between negative and 1+ (1\u20139 acid-fast bacilli "
                            "per 100 fields). **A category in its own right**, scored as such and "
                            "never folded into positive or negative."],
                 ["Passing rate", "The tolerance printed on the provider report. On TB microscopy "
                                  "it is 90 per cent, equivalently a discordant rate at or below "
                                  "10 per cent. On malaria no tolerance is printed and a discordant "
                                  "rate of zero is recorded."],
                 ["False negative (TB)", "A slide the health centre called negative that the "
                                         "laboratory reads as positive or scanty. **The most "
                                         "dangerous error in the programme**, because the patient "
                                         "was sent home."],
                 ["False positive (TB)", "A slide the health centre called positive that the "
                                         "laboratory reads as negative. Costs unnecessary "
                                         "treatment and a notification that should not have been "
                                         "made."],
                 ["Quantification error", "A discrepancy in the grade rather than in the presence of "
                                          "the organism \u2014 printed on the NRL TB report as QE."],
                 ["Site", "A health centre in the network. The unit at which results are scored and "
                          "reported back."],
             ],
             [4.6, 11.9]),
        ]),
        ("5. Responsibilities and Authorisation", [
            ("tbl", ["Task", "Responsible Person(s)"],
             [
                 ["Authorise this SOP and resource the programme", "Laboratory Manager"],
                 ["Own the programme, hold the schedule, keep the reports", "Quality Manager"],
                 ["Select slides for re-reading", "Section head, Parasitology / TB"],
                 ["Perform the blind re-reading",
                  "A competent reader who did not report the original result"],
                 ["Second-read a discordant slide before it is reported",
                  "Quality Manager or a second competent reader"],
                 ["Compile the quarterly return and check its arithmetic", "Quality Manager"],
                 ["Sign the return", "Laboratory Manager"],
                 ["Send the return and the queries to the provider", "Quality Manager"],
                 ["Notify a health centre of a clinically significant discordance",
                  "Laboratory Manager, within one working day of the discrepancy being confirmed"],
                 ["Feed a confirmed discrepancy into the health centre's corrective action",
                  "Quality Manager with the district laboratory focal person"],
             ],
             [8.8, 8.2]),
        ]),
        ("6. The slide sample", [
            ("p", "The comparison is only as good as the sample, and this is where the programme on "
                  "file is weakest. **A site re-read on two slides cannot be judged.** The procedure:"),
            ("steps", [
                ("1.", "Slides are taken from the site's routine work for the quarter as they come, "
                       "in the order they were reported. They are **never** chosen for appearance, "
                       "for being clearly positive, or for being the ones the site got right."),
                ("2.", "A **minimum of fifteen slides per site per quarter** is the target, and no "
                       "site is scored on fewer than **ten**. Where a site reported fewer than ten "
                       "slides in the quarter, **every slide it reported** is re-read and the return "
                       "states that the sample was the whole of the site's work, not a sample of it."),
                ("3.", "The sample must contain both positive and negative slides **in roughly the "
                       "proportion the site reported**, so that neither a false negative nor a false "
                       "positive can hide. A quarter that is entirely negative slides is recorded as "
                       "a **sampling limitation**, not as a clean quarter."),
                ("4.", "At least the **scanty** slides reported by the site are included, because "
                       "scanty is where readings diverge most."),
                ("5.", "Every slide is identified by the site, the quarter and a re-reading number, "
                       "with the patient identifier removed from the re-reading sheet. The reader "
                       "sees neither the identity nor the original result."),
                ("6.", "The numbers sent, the numbers received and the numbers readable are "
                       "recorded. A slide that has deteriorated in transit or storage is recorded as "
                       "**unreadable**, not as concordant."),
            ]),
            ("p", "**Coverage.** All eight health centres are in scope on malaria. On TB the "
                  "programme on file covers only Nyange and Rwankuba; widening it to every site that "
                  "sends TB smears is an action of this SOP, and until it is widened every return "
                  "states the coverage limitation in words."),
        ]),
        ("7. Performing the re-reading", [
            ("steps", [
                ("1.", "The re-reader examines the slide independently, using the same method the "
                       "health centre used, and records a result **before** the original is looked at."),
                ("2.", "The result is recorded in the same categories the site used: for TB, "
                       "negative, scanty or positive with a grade; for malaria, negative or positive "
                       "with the species, and the parasite density where the site reported one."),
                ("3.", "Staining quality is recorded on a three-point scale and commented on where "
                       "the slide is difficult to read. **A finding about staining is a finding**, "
                       "even where the reading agrees, because the next slide from that site may be "
                       "unreadable."),
                ("4.", "The original result is then revealed and the pair recorded as concordant or "
                       "discordant, with the direction of the discordance."),
                ("5.", "Every **discordant** slide, and every slide on which staining quality was "
                       "poor, is read a second time by a different reader before it is reported. Two "
                       "readers agreeing that a site was wrong is evidence; one reader alone is an "
                       "opinion."),
                ("6.", "The second reading and its reader are recorded. Where the two laboratory "
                       "readers disagree, the slide is referred to the National Reference Laboratory "
                       "and is not counted against the site."),
            ]),
            ("p", "**The person who reported the original must not re-read it.** Where the only "
                  "competent reader for a site is the person who read it first, the slide is sent "
                  "to NRL instead. This is the single rule most likely to be broken in a laboratory "
                  "of eight staff, and it is the rule that decides whether the programme means "
                  "anything."),
        ]),
        ("8. Calculating and scoring", [
            ("p", "The arithmetic is done in full on a worksheet that travels with the return, and "
                  "it is recomputed by the Quality Manager. **A return whose rows do not balance is "
                  "not sent.**"),
            ("tbl",
             ["Quantity", "How it is calculated", "Where it is recorded"],
             [
                 ["Slides received", "Count of slides that arrived from the site", "Worksheet 1"],
                 ["Slides readable", "Received minus unreadable, with each unreadable slide listed "
                                     "and the reason", "Worksheet 2"],
                 ["Positives and negatives, site side",
                  "As the site reported them; they must sum to the slides read", "Worksheet 3"],
                 ["Positives and negatives, laboratory side",
                  "As the laboratory read them; they must sum to the same total", "Worksheet 4"],
                 ["Discordant count",
                  "Slides on which the two readings differ, listed by slide number", "Worksheet 5"],
                 ["Discordant rate",
                  "Discordant count \u00f7 slides readable \u00d7 100, to one decimal place",
                  "Worksheet 6"],
                 ["Score", "100 minus the discordant rate, on the provider's convention",
                  "Worksheet 7"],
                 ["Verdict", "Pass where the discordant rate is at or below the tolerance printed "
                             "on the report; otherwise Fail", "Worksheet 8"],
             ],
             [4.3, 8.0, 4.7]),
            ("p", "**Tolerances.** TB microscopy: a discordant rate at or below **10 per cent**, "
                  "which the NRL report states as a passing rate of 90 per cent. Malaria: the NRL "
                  "report prints **no tolerance**; the discordant rate is recorded and a rate of "
                  "zero passes. Where a provider prints a criterion, the **provider's** criterion "
                  "governs and is quoted on the return. Where none is printed, the laboratory "
                  "applies the TB tolerance and says on the return that it has done so \u2014 it "
                  "does not invent a silent one."),
            ("p", "**Consistency checks before the return is signed.** The following are checked "
                  "every quarter and the check is initialled on the return: the rows add up on both "
                  "sides; the slide counts match on both sides; the discordant rate reproduces from "
                  "the numbers in the row; the verdict matches the criterion printed on the report; "
                  "the quarter on the tab matches the quarter in the text; and every site that "
                  "should appear does appear. **A site that has dropped out of a quarter is a "
                  "finding, not a blank.**"),
        ]),
        ("9. Reporting back", [
            ("p", "The quarterly return goes to the National Reference Laboratory and to each health "
                  "centre, and it contains for every site: the slides read, the two readings, the "
                  "discordant count with the slides listed, the discordant rate, the verdict against "
                  "the printed criterion, and a comment. A verdict is never sent without the numbers "
                  "that produced it."),
            ("p", "**Clinically significant discrepancies are notified before the return.** Where the "
                  "laboratory finds a **false negative** on a TB smear, or a missed malaria positive, "
                  "the health centre is telephoned the same working day and the patient is traced. "
                  "The return then records what was done. A quarterly cycle is far too slow to "
                  "correct a patient who is being sent home with a wrong answer."),
            ("p", "**What the laboratory may and may not conclude.** A pass at a site means the "
                  "reading agreed on the slides re-read in that quarter, at that sample size, and "
                  "nothing more. It is not a statement that the site's microscopy is accurate, it "
                  "is not a competence assessment of an individual, and it is not a licence to "
                  "reduce supervision. A fail means the sample contained discordance that needs a "
                  "cause, not that the site is bad."),
        ]),
        ("10. When a site fails, or a discrepancy is confirmed", [
            ("p", "A discordance is investigated to its cause in the same way an unsatisfactory PT "
                  "result is, and the cause is sought at the site rather than assumed:"),
            ("b", [
                "**Pre-analytical** \u2014 was the slide made from a good sample, at the right "
                "thickness, and fixed properly?",
                "**Staining** \u2014 was the stain fresh, correctly filtered, correctly timed, "
                "correctly decolourised? Is the site's stain control slide in place and read?",
                "**Reading** \u2014 did the reader examine enough fields, at the right "
                "magnification, for long enough?",
                "**Materials and equipment** \u2014 is the microscope clean and correctly adjusted, "
                "is the objective intact, is the light right?",
                "**Person** \u2014 is the reader competent for this examination, and when was that "
                "last verified?",
            ]),
            ("p", "Corrective action is agreed **with the site**, has an owner and a date, and its "
                  "**effectiveness is verified on the next quarter's re-reading** \u2014 not on a "
                  "declaration that it was done. A site that fails the same way twice goes to the "
                  "Laboratory Manager and into the district performance review."),
        ]),
        ("11. Records", [
            ("tbl",
             ["Record", "Who keeps it", "How long"],
             [
                 ["The re-reading worksheet, with both readings and the arithmetic",
                  "Quality Manager", "6 years"],
                 ["The signed quarterly return, as sent and as returned by the provider",
                  "Quality Manager", "6 years"],
                 ["The list of slides, their identifiers and their fate", "Section head", "6 years"],
                 ["Second-reading records for every discordant slide", "Quality Manager", "6 years"],
                 ["Notifications of clinically significant discrepancies, with the time and the "
                  "person told", "Quality Manager", "6 years"],
                 ["Corrective action records for each failing site", "Quality Manager",
                  "6 years, and until effectiveness is verified"],
             ],
             [7.6, 4.4, 5.0]),
            ("p", "Records are held so that any quarter can be reconstructed: which slides, read by "
                  "whom, with what result, against what original, and with what arithmetic. A return "
                  "that cannot be reconstructed from the file is treated as not having been made."),
        ]),
        ("12. Defects found in the programme as it stands", [
            ("p", "The review of 15 November 2026 read every re-reading report on file. Five matters "
                  "were found. They are recorded here rather than corrected on the provider's "
                  "behalf, and each is raised with NRL in writing:"),
            ("tbl",
             ["Finding", "Why it matters", "Action and owner"],
             [
                 ["Malaria Q3 2025/26 does not add up: Coko HC is reported 8 positive and 7 "
                  "negative against 8 and 7 yet carries a discordant rate of 0 and a pass; "
                  "Muyongwe reads 3 and 12 against 3 and 10; Rushashi 5 and 10 against 5 and 12",
                  "A quarter whose rows do not balance cannot be signed as a pass, and a zero "
                  "computed from unbalanced rows is not a measured zero",
                  "Query to NRL with the rows reproduced; hold the quarter **unverified** until "
                  "answered. Quality Manager, within one month"],
                 ["Two TB sheets are both headed \"Quarter II\" \u2014 Q4 2024/25, whose tab and "
                  "dates make it Quarter IV, and Q3 2025/26",
                  "Two quarters carrying one label cannot be trended, and a report filed under the "
                  "wrong quarter misstates the year",
                  "Query the labelling; file by tab and date range and annotate. Quality Manager"],
                 ["TB Q4 2024/25 marks both sites **False** with a discordant rate of **0**, which "
                  "on the printed 90 per cent passing rate should read Pass",
                  "Either a transcription error or a criterion other than the printed one; the "
                  "difference decides whether a site was wrongly failed",
                  "Query with NRL; if the fail stands, obtain and record the criterion applied. "
                  "Quality Manager"],
                 ["TB re-reading covers two of eight centres, and Nyange's sample ran 11, then "
                  "**2**, then 5, then 8 slides",
                  "A two-slide quarter supports no conclusion about that site, and six sites are "
                  "not covered at all",
                  "Widen coverage and enforce the ten-slide floor from section 6. Laboratory "
                  "Manager, from the next quarter"],
                 ["No tolerance is printed on the malaria report",
                  "A pass with no stated criterion cannot be defended to an assessor or to a site "
                  "that disputes it",
                  "Apply the TB tolerance and state it on every malaria return until NRL prints "
                  "one. Quality Manager"],
             ],
             [5.6, 5.4, 6.0]),
        ]),
        ("13. Supporting documents and references", [
            ("b", [
                "ISO 15189:2022, 7.3.7 \u2014 external quality assessment, and the requirement that "
                "where results fall outside pre-defined criteria the laboratory acts under 8.7 "
                "including an assessment of clinical significance.",
                "ISO 15189:2022, 7.4 \u2014 reporting of results, which governs how a discrepancy "
                "is communicated to a health centre.",
                "ISO 15189:2022, 8.7 \u2014 nonconforming work, under which a confirmed discrepancy "
                "in a reported patient result is handled.",
                "ASLM SLIPTA Checklist v3:2023, items 8.20 and 8.21 \u2014 the procedure for EQA "
                "and participation in it, which the assessor scores.",
                "NRL quarterly evaluation reports for malaria and TB microscopy re-testing of "
                "health centres \u2014 the provider format this SOP works to.",
            ]),
        ]),
    ],
}

DOCUMENTS = [VENOUS, GLUCOSE, XN330, XQ320, PIMA, TROUBLE, EQA, HC_REVIEW]

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
    from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether, PageBreak,
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
                items = block[1]
                rows_ = [P("•  " + item, "b") for item in items]
                # No orphans: a short list is kept together so a few lines cannot be
                # stranded on a page of their own. Long lists are split in blocks.
                if len(items) <= 8:
                    story.append(KeepTogether(rows_))
                else:
                    for i in range(0, len(rows_), 4):
                        story.append(KeepTogether(rows_[i:i + 4]))
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
