#!/usr/bin/env python3
"""
Build the Ruli District Hospital Laboratory SOP sample/template in two formats:

  sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.docx
  sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.pdf

Conventions implemented:
  * Facility acronym in every document code: RliDHL-XXX-SOP-NNN-VERS-NNN.
  * 22-section technical-SOP layout (2026 standard layout mapped to ISO 15189:2022).
  * Paperless operation (Rwanda, since 01 August 2026): electronic controlled
    document, e-signatures, EMR records; only referred specimens carry a printed
    Sample Referral Manifest; printed SOP copies are contingency-only.
  * Signature page: 15 rows (IPPIS No. column) so at least 12 staff fit on one page.
  * Fresh page starts: cover | review history + controlled copies + TOC |
    body (1-20) | 21 Signature Page | 22 Amendment Page.

Usage:  .venv/bin/python scripts/build_sop.py
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = ROOT / "assets" / "Ruli-District-Hospital-logo.png"
OUT_STEM = "RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection"
DOCX_PATH = ROOT / "sop" / (OUT_STEM + ".docx")
PDF_PATH = ROOT / "sop" / (OUT_STEM + ".pdf")

# ---------------------------------------------------------------------------
# Shared content model
# ---------------------------------------------------------------------------

META = {
    "lab": "Ruli District Hospital Laboratory Department",
    "doc_no": "RliDHL-IH-SOP-001-VERS-001",
    "title": "Venous Blood Sample Collection by Venipuncture",
    "version": "001",
    "effective": "____/____/________  (entered at authorisation)",
    "review": "____/____/________  (two years after the effective date)",
    "supersedes": "\u2014 (first issue to the 2026 standard layout)",
    "owner": "Kelvin SAMAHANI \u2014 Laboratory Manager",
}

DOC_CONTROL_ROWS = [
    ("SOP Number", META["doc_no"]),
    ("SOP Title", META["title"]),
    ("Version", META["version"]),
    ("Effective Date", META["effective"]),
    ("Next Review Date", META["review"]),
    ("Supersedes", META["supersedes"]),
    ("Document Owner", META["owner"]),
    ("Classification", "Controlled electronic document \u2014 master copy in the laboratory "
                       "document control system (RliDHL-ADM-SOP-004)"),
]

APPROVAL_ROWS = [
    ("Prepared by (Author)", "Jean Paul NSENGIYUMVA \u2014 Lab Quality Manager"),
    ("Reviewed by", "Kelvin SAMAHANI \u2014 Laboratory Manager"),
    ("Authorized by", "Dr. D\u00e9ogratias KANEZA \u2014 Hospital\u2019s Director General"),
]

APPROVAL_STATEMENT = (
    "This is a controlled electronic document of the Ruli District Hospital Laboratory "
    "Department. It becomes effective when authorised above. Authorisation is applied "
    "electronically in the document control system and is traceable to the signatory\u2019s "
    "user account. Paperless operation applies since 01 August 2026: a printed copy is "
    "uncontrolled unless stamped \u201cCONTINGENCY COPY\u201d and verified against the document "
    "control system. Control of documents and records: ISO 15189:2022, Clauses 8.3 and 8.4; "
    "RliDHL-ADM-SOP-004."
)

REVISION_HEADERS = ["Prepared/Reviewed by", "Version", "Date", "Description of Change"]
REVISION_ROWS = [
    ["J.P. NSENGIYUMVA", "001", "Pending authorisation",
     "First issue to the 2026 standard layout (ISO 15189:2022). Rewritten for paperless "
     "operation: electronic requests, labels and records in the EMR; printed Sample Referral "
     "Manifest only for referred specimens."],
    ["", "", "", ""],
    ["", "", "", ""],
]

CONTROLLED_COPIES = [
    "Controlled copies \u2014 Original (electronic master): document control system, "
    "Laboratory Manager\u2019s office. Controlled electronic copies (read-only): laboratory "
    "network and EMR document module; collection stations.",
    "Printed copies: paperless operation applies since 01 August 2026. A printout is made "
    "only for contingency use during information-system downtime, is stamped "
    "\u201cCONTINGENCY COPY\u201d, and is destroyed when the electronic version is restored.",
    "Supersedes: \u2014 (first issue). When this SOP replaces an earlier document, the "
    "superseded number, version, effective date and review date are recorded here and in "
    "the review history.",
]

# Body blocks: ('p', text) | ('b', [items]) | ('h2', title) |
#              ('steps', [(label, text), ...]) | ('tbl', headers, rows, widths)
BODY = [
    ("1. Introduction", [
        ("p", "Venous blood collection is the most frequently performed pre-examination "
              "activity of the laboratory, and pre-examination errors \u2014 wrong patient, wrong "
              "tube, insufficient volume, haemolysis \u2014 are the largest source of unusable "
              "specimens. This procedure standardises how venous blood is collected at Ruli "
              "District Hospital so that every specimen is correctly identified, safely "
              "collected, and fit for examination."),
        ("p", "The hospital operates on an electronic medical record (EMR). Requests, patient "
              "identification, specimen tracking and results are electronic: no paper request "
              "forms accompany specimens inside the hospital. Referred specimens travel with a "
              "printed Sample Referral Manifest, as required by the national referral system."),
    ]),
    ("2. Principle of the Test", [
        ("p", "Blood is drawn from a peripheral vein through a hollow needle into evacuated "
              "tubes. The vacuum in each tube draws a fixed volume of blood and mixes it with "
              "the additive (anticoagulant, clot activator or preservative) that the requested "
              "examination requires. Filling the tubes in a fixed order of draw prevents "
              "additive carry-over between tubes; gentle inversion mixes blood and additive "
              "without damaging the cells."),
    ]),
    ("3. Purpose", [
        ("p", "This SOP standardises the pre-examination steps for the collection of venous "
              "blood by peripheral venipuncture, so that specimens are collected safely, "
              "identified without error, and remain fit for the examinations requested."),
    ]),
    ("4. Scope", [
        ("p", "This SOP applies to all laboratory personnel \u2014 biotechnologists, technicians, "
              "interns and students under supervision \u2014 who collect venous blood at the main "
              "laboratory and at every collection point of Ruli District Hospital. It covers "
              "adult and paediatric venipuncture and all evacuated tube types listed in 6.3."),
        ("p", "It does not cover: arterial blood collection; capillary blood sampling; blood "
              "culture collection; blood donation and donor phlebotomy; or venous access "
              "devices (cannulas, ports). These are covered by separate SOPs listed in the "
              "document control master list."),
    ]),
    ("5. Definitions and Abbreviations", [
        ("h2", "5.1 Definitions"),
        ("tbl",
         ["No.", "Term", "Definition"],
         [
             ["1", "Positive patient identification",
              "Verification of identity with at least two identifiers (full name plus date of "
              "birth or patient number) matched to the EMR request before collection."],
             ["2", "Order of draw",
              "The fixed sequence in which tubes are filled during one venipuncture, preventing "
              "additive carry-over between tubes."],
             ["3", "Haemolysis",
              "Rupture of red blood cells releasing intracellular constituents into serum or "
              "plasma; a common cause of specimen rejection."],
             ["4", "Mislabelled specimen",
              "A specimen whose label does not correctly and uniquely identify the patient; "
              "always handled as nonconforming work (ISO 15189:2022, 7.5)."],
             ["5", "Contingency copy",
              "A printout of this SOP made only during information-system downtime, stamped "
              "\u201cCONTINGENCY COPY\u201d, and destroyed when the system is restored."],
             ["6", "Referred specimen",
              "A specimen sent to another laboratory for examination; travels with the printed "
              "Sample Referral Manifest."],
         ],
         [1.2, 4.3, 11.5]),
        ("h2", "5.2 Abbreviations"),
        ("tbl",
         ["No.", "Acronym", "Full Meaning"],
         [
             ["1", "EMR", "Electronic Medical Record"],
             ["2", "EQA", "External Quality Assessment"],
             ["3", "EDTA", "Ethylenediaminetetraacetic acid (tube anticoagulant)"],
             ["4", "IPPIS", "Integrated Personnel and Payroll Information System"],
             ["5", "IQC", "Internal Quality Control"],
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
            "(verbal) request is accepted only in an emergency: record who made it and when, "
            "and have it entered in the EMR within 24 hours.",
            "Identify the patient with at least two identifiers \u2014 full name plus date of birth "
            "or patient number \u2014 matched to the EMR request. Scan the wristband where one is "
            "worn. Never collect on a name alone.",
            "The request states the examinations requested and, where relevant, the patient\u2019s "
            "state (fasting, random, timed) and the exact time of collection.",
            "Each specimen receives its unique laboratory number at reception; the number is "
            "scanned onto the tube label and follows the specimen, aliquots and report in the "
            "LIS/EMR so that every result traces to one patient.",
        ]),
        ("h2", "6.2 Patient preparation and collection"),
        ("b", [
            "Confirm the request (examinations, tube types, fasting or timing requirements) "
            "before approaching the patient.",
            "Explain the procedure, confirm the patient\u2019s consent, and seat or lie the patient "
            "comfortably. Do not leave a patient unattended in a collection chair.",
            "For fasting specimens, confirm at least 8 hours without food or drink other than "
            "water, and record the fasting state in the EMR.",
            "Do not collect from an arm with a running infusion, a fistula, a mastectomy side, "
            "or burn/scarred skin unless no alternative exists \u2014 if used, note it in the EMR.",
        ]),
        ("h2", "6.3 Specimen types, volume, and storage"),
        ("tbl",
         ["Tube (cap colour / additive)", "Minimum volume", "Handling and storage"],
         [
             ["Blood culture bottles (yellow) \u2014 if requested", "Per manufacturer\u2019s fill mark",
              "Incubate promptly; do not refrigerate"],
             ["Light blue \u2014 sodium citrate 3.2%", "Fill to the mark (9:1 blood to citrate)",
              "Invert gently 8\u201310 times; transport upright; analyse within 24 h at 15\u201325 \u00b0C"],
             ["Red / gold \u2014 plain or clot activator", "3 mL",
              "Allow to clot (30 min); centrifuge; serum 8 h at 15\u201325 \u00b0C, 72 h at 2\u20138 \u00b0C"],
             ["Green \u2014 lithium heparin", "3 mL",
              "Invert gently; centrifuge promptly; as for serum"],
             ["Lavender \u2014 EDTA", "1 mL (2 mL paediatric)",
              "Invert gently 8\u201310 times; do not shake; examine for clots before testing"],
             ["Grey \u2014 fluoride/oxalate", "2 mL",
              "Invert gently; glucose preserved in the unseparated tube (24 h at 15\u201325 \u00b0C)"],
         ],
         [5.6, 4.1, 7.3]),
        ("h2", "6.4 Specimen acceptance and rejection criteria"),
        ("tbl",
         ["Criterion", "Acceptable", "Reject / request recollection"],
         [
             ["Identification", "Label complete; matches the EMR request; two identifiers present",
              "Unlabelled, mislabelled, or mismatch with the request"],
             ["Container", "Correct tube type for the test; intact and closed",
              "Wrong tube or additive; leaking or broken tube"],
             ["Volume", "At or above the minimum in 6.3",
              "Below the minimum; citrate or EDTA not filled to the mark"],
             ["Quality", "Clear serum or plasma; slight haemolysis noted on the report",
              "Gross haemolysis; clots in an EDTA tube; drawn from an infusion arm"],
             ["Timing", "Within the limits in 6.3; collection time recorded",
              "Limits exceeded; collection time unknown"],
             ["Transport", "Within 2 hours internally; manifest present for referrals",
              "Delayed beyond stability; referral without Sample Referral Manifest"],
         ],
         [3.2, 6.9, 6.9]),
        ("p", "Record every rejected specimen in the specimen rejection register and inform the "
              "requester at once so that a new specimen is collected. If a specimen that fails "
              "these criteria is irreplaceable or clinically critical, the Laboratory Manager "
              "or designee may authorise testing; the report then states the problem and the "
              "caution needed to interpret the result."),
    ]),
    ("7. Responsibilities and Authorisation", [
        ("tbl",
         ["Task", "Responsible Person(s)"],
         [
             ["Authorise this SOP and provide the resources needed to apply it",
              "Hospital Director General \u2014 Dr. D\u00e9ogratias KANEZA"],
             ["Review this SOP at least every two years; ensure that users are trained, "
              "competence-assessed and authorised; approve recollection exceptions",
              "Laboratory Manager \u2014 Kelvin SAMAHANI"],
             ["Maintain this SOP under document control; monitor indicators, nonconformities "
              "and corrective actions; maintain training and authorisation records",
              "Lab Quality Manager \u2014 Jean Paul NSENGIYUMVA"],
             ["Perform venipuncture; identify the patient; label at the bedside; record the "
              "collection in the EMR",
              "Authorised Laboratory Biotechnologists"],
             ["Receive specimens; decide acceptance and rejection (6.4); maintain the "
              "specimen rejection register",
              "Laboratory Receptionist / Senior Biotechnologist"],
             ["Manage exposure incidents, sharps and waste; keep PPE and vaccination current",
              "Health & Safety Focal Person; all laboratory staff"],
         ],
         [9.0, 8.0]),
        ("p", "Authorisation. Only staff trained on this SOP, assessed as competent and "
              "authorised by the Laboratory Manager perform venipuncture and label specimens "
              "without supervision. Students and staff in training work under the direct "
              "supervision of an authorised person, who checks and countersigns their work. "
              "Training, competence assessment and authorisation are recorded electronically "
              "in the personnel training module; acknowledgement of this SOP is recorded as "
              "described on the Signature Page (section 21). Competence is reassessed at the "
              "interval set in the competence assessment procedure and whenever this SOP "
              "changes."),
    ]),
    ("8. Equipment, Reagents, and Consumables", [
        ("h2", "8.1 Equipment"),
        ("b", [
            "Vacuum tube holders and safety-engineered needles 21G or 22G; butterfly sets for "
            "difficult veins and paediatric patients.",
            "Disposable single-patient tourniquets.",
            "Mobile label printer and barcode scanner linked to the EMR (collection stations).",
            "Cooled carriers (2\u20138 \u00b0C) and ambient transport boxes where required.",
        ]),
        ("h2", "8.2 Reagents, calibrators and controls"),
        ("p", "Not applicable \u2014 this procedure uses no reagents, calibrators or controls. "
              "Alcohol swabs and collection tubes are consumables (8.3)."),
        ("h2", "8.3 Consumables"),
        ("b", [
            "Evacuated tubes listed in 6.3, within expiry and stored as the manufacturer directs.",
            "70% isopropyl alcohol swabs, dry sterile gauze, adhesive plaster.",
            "Pre-printed barcode tube labels (EMR-generated); biohazard transport bags.",
            "Puncture-resistant sharps containers; clinical waste bags; 10% sodium hypochlorite; "
            "70% ethanol.",
        ]),
        ("h2", "8.4 Calibration and metrological traceability"),
        ("p", "Not applicable \u2014 no measuring equipment is used in this procedure. Centrifuges "
              "and analysers used downstream are controlled under the laboratory equipment "
              "management procedure (6.4 and 6.5 of ISO 15189:2022)."),
    ]),
    ("9. Environment and Biosafety Measures", [
        ("b", [
            "Treat every specimen as potentially infectious and apply standard precautions for "
            "every patient.",
            "Wear a laboratory coat and gloves; change gloves between patients; add eye "
            "protection where splashing is possible. Perform hand hygiene before and after "
            "every collection.",
            "Never recap needles. Activate the safety device and drop the assembly into the "
            "sharps container at the point of use.",
            "Hepatitis B vaccination is mandatory for all staff who collect blood.",
            "After any exposure (needle-stick or splash): wash the area immediately, inform the "
            "Laboratory Manager and follow the post-exposure procedure in the Laboratory Safety "
            "Manual. Record the event; raise a nonconformity where procedural failure "
            "contributed (7.5).",
            "Clean spills at once with 10% sodium hypochlorite; decontaminate benches with 70% "
            "ethanol. Discard used tubes and materials as infectious waste under the Laboratory "
            "Waste Management procedure.",
            "Keep all patient information confidential (ISO 15189:2022, 4.2): what is seen on "
            "the ward or screen is not repeated outside the work.",
        ]),
    ]),
    ("10. Procedural Steps", [
        ("h2", "10.1 Preparation and patient identification"),
        ("steps", [
            ("1.", "Open the collection worklist on the EMR station or tablet; verify the "
                   "pending requests."),
            ("2.", "Assemble the tubes in the order of draw (10.2), needle, swabs, gauze, "
                   "labels and sharps container within reach."),
            ("3.", "Identify the patient with two identifiers matched to the EMR request; scan "
                   "the wristband where present. Perform hand hygiene and put on gloves."),
        ]),
        ("h2", "10.2 Order of draw"),
        ("p", "When more than one tube is collected during one venipuncture, fill the tubes in "
              "the sequence below (consistent with CLSI GP41). Invert each filled tube gently "
              "8 to 10 times; never shake. Fill citrate and EDTA tubes to the marked volume."),
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
        ("h2", "10.3 Venipuncture"),
        ("steps", [
            ("1.", "Select the site: prefer the median cubital vein of the forearm. Avoid the "
                   "areas listed in 6.2."),
            ("2.", "Apply the tourniquet 3 to 4 cm above the site, tight enough to impede "
                   "venous return but not arterial flow. Do not leave it on for more than 1 minute."),
            ("3.", "Palpate the vein. Clean the site with a 70% isopropyl swab in a widening "
                   "circle over at least 2 cm and allow to air-dry for at least 30 seconds. Do "
                   "not touch, blow on or fan the site."),
            ("4.", "Anchor the vein, holding the skin taut below the puncture point. Insert the "
                   "needle bevel upwards at 15 to 30 degrees in one smooth motion."),
            ("5.", "When blood flows, fill the tubes in the order of draw (10.2). Under-filling "
                   "or over-filling citrate and EDTA tubes is a cause for rejection."),
            ("6.", "Release the tourniquet before withdrawing the needle. Withdraw, apply dry "
                   "gauze with firm pressure for 2 to 3 minutes, and apply plaster if needed."),
            ("7.", "Activate the needle safety device immediately and discard the assembly in "
                   "the sharps container at the point of use."),
            ("8.", "If blood does not flow: do not probe more than twice. Release the "
                   "tourniquet, withdraw, apply pressure, and use a second site or the other "
                   "arm. Record every difficult draw in the EMR."),
        ]),
        ("h2", "10.4 Labelling and identification"),
        ("b", [
            "Label every tube at the bedside, in the patient\u2019s presence, immediately after "
            "filling: scan the wristband, then scan/print the tube label from the EMR. Never "
            "pre-label tubes and never label away from the patient.",
            "Verify each label against the request before leaving the patient. The label "
            "carries the patient name, patient number, laboratory number, collection date and "
            "time, and the collector\u2019s identifier.",
            "For referred specimens, record the Sample Referral Manifest number on the tube "
            "label and in the EMR.",
            "Corrections: in the EMR only, traceable to the user account and time (never "
            "overwritten); on a printed manifest, single line through, correct entry, initial "
            "and date.",
        ]),
        ("h2", "10.5 Transport to the laboratory"),
        ("b", [
            "Internal transport: no paper accompanies the specimen \u2014 the EMR/LIS carries the "
            "request and tracking. Deliver within 2 hours at 15\u201325 \u00b0C, or at 2\u20138 \u00b0C in a "
            "cooled carrier where required.",
            "Do not use the pneumatic tube system for coagulation samples or blood culture "
            "bottles unless validated for them.",
            "Referred specimens: print the Sample Referral Manifest as required by the national "
            "referral system \u2014 the only paper that travels with a specimen. The sender and "
            "the receiving laboratory sign the manifest; a scanned copy is attached to the EMR "
            "episode after use.",
        ]),
        ("h2", "10.6 Post-collection"),
        ("b", [
            "Record in the EMR: collection time, collector, tube types, fasting state where "
            "relevant, and any note (difficult draw, haematoma, infusion arm).",
            "Dispose of waste and decontaminate the collection area as in section 9.",
        ]),
    ]),
    ("11. Quality Control and External Quality Assessment", [
        ("h2", "11.1 Quality indicators (internal monitoring)"),
        ("b", [
            "Specimen rejection rate: target at most 2% of specimens per month.",
            "Haemolysis rate: target at most 1% per month.",
            "Mislabelled specimens: target zero \u2014 every case is a nonconformity.",
            "Specimens delivered within 2 hours: target at least 95%.",
            "The Lab Quality Manager extracts these indicators monthly from the EMR/LIS, "
            "reviews them with the Laboratory Manager, and records the review.",
        ]),
        ("h2", "11.2 External quality assessment"),
        ("p", "The laboratory takes part in the pre-analytical or relevant EQA schemes that "
              "RBC/NRL organises or designates, and acts on unsatisfactory performance (11.3)."),
        ("h2", "11.3 Actions when monitoring fails"),
        ("p", "When an indicator misses its target or an EQA result is unsatisfactory: "
              "investigate the cause, define corrective action, and evaluate its "
              "effectiveness (ISO 15189:2022, 8.7). Where wrong-patient or mislabelled "
              "specimens were involved, perform a look-back to the patients affected and "
              "inform the Laboratory Manager immediately."),
    ]),
    ("12. Result Interpretation and Validation", [
        ("p", "Not applicable \u2014 this procedure produces no examination result. The decision "
              "on specimen fitness is made at reception (6.4). Interpretation and validation of "
              "results are covered by the examination SOP concerned and the result release "
              "procedure."),
    ]),
    ("13. Performance Characteristics and Limitations", [
        ("b", [
            "Tourniquet time beyond 1 minute causes haemoconcentration and raises protein and "
            "lipid results.",
            "Squeezing the puncture site causes haemolysis and tissue-fluid contamination.",
            "Difficult veins, paedemic collection and infusion arms limit volume and quality; "
            "record the limitation in the EMR where it may affect a result.",
            "This procedure controls the pre-examination phase only; method performance is "
            "stated in the examination SOP concerned.",
        ]),
    ]),
    ("14. Reference Intervals, Clinical Decision Limits, and Critical Results", [
        ("p", "Not applicable \u2014 no values are produced by this procedure. Reference intervals, "
              "decision limits and critical results are defined in the examination SOPs and the "
              "result release procedure."),
    ]),
    ("15. Reporting and Release of Results", [
        ("p", "Not applicable for results \u2014 reporting is covered by the examination SOPs and "
              "the result release procedure. The collection event, specimen status "
              "(accepted/rejected) and any recollection request are recorded in the EMR and are "
              "traceable to the collector\u2019s user account. Patient information seen during "
              "collection is confidential (ISO 15189:2022, 4.2)."),
    ]),
    ("16. Specimen Retention and Disposal", [
        ("b", [
            "Collected specimens are held and transported under the conditions in 6.3 and 10.5.",
            "Retention of specimens after examination, and their disposal, follow the "
            "examination SOP concerned and national requirements.",
            "Used tubes, swabs and dressings are discarded as infectious waste; sharps go into "
            "the puncture-resistant container at the point of use (Laboratory Waste Management "
            "procedure).",
        ]),
    ]),
    ("17. Risk Management, Nonconforming Work, and Contingency", [
        ("tbl",
         ["Risk", "Control measure"],
         [
             ["Wrong-patient identification",
              "Two identifiers at the bedside; wristband scan; label printed and scanned in the "
              "patient\u2019s presence; never pre-label"],
             ["Needlestick injury",
              "Safety-engineered devices; no recapping; sharps container at the point of use; "
              "hepatitis B vaccination; post-exposure procedure"],
             ["Haemolysed or clotted specimens",
              "Correct fill volumes; gentle inversion; order of draw; tourniquet under 1 "
              "minute; prompt transport"],
             ["Delayed transport affecting stability",
              "Two-hour delivery rule; scheduled collection runs; cooled carriers"],
             ["Mislabelling or lost traceability during EMR downtime",
              "Downtime procedure below: manual labels double-checked by a second person; "
              "back-entry within 24 hours"],
         ],
         [6.0, 11.0]),
        ("p", "Nonconforming work (ISO 15189:2022, 7.5): if any step was not followed \u2014 wrong "
              "tube, missed identification, unlabelled tube \u2014 stop, report to the Laboratory "
              "Manager, and record the event in the nonconformity log. The Laboratory Manager "
              "decides on the affected specimens, arranges recollection, and determines whether "
              "a look-back to released results is needed. Patient safety comes first."),
        ("p", "Contingency (ISO 15189:2022, 7.8): when the EMR/LIS is unavailable \u2014 use the "
              "approved downtime request form and manual tube labels; a second person "
              "double-checks every label; keep a written collection log; enter all data into "
              "the EMR within 24 hours of restoration; use stamped \u201cCONTINGENCY COPY\u201d "
              "printouts of this SOP and destroy them after restoration. Referral manifests "
              "continue to be printed as required by the national referral system."),
    ]),
    ("18. Records", [
        ("tbl",
         ["Record", "Where kept", "Retention"],
         [
             ["Collection events, labels, collector identity", "EMR / LIS (electronic)",
              "As set in the laboratory records policy and national requirements"],
             ["Specimen rejection register", "EMR / LIS (electronic)", "As above"],
             ["Nonconformity and corrective action records", "Quality module (electronic)", "As above"],
             ["Training, competence and authorisation (incl. section 21 acknowledgement)",
              "Personnel training module (electronic)", "Duration of employment plus the period "
              "required by national rules"],
             ["Exposure incident reports", "Health & Safety file", "As above"],
             ["Sample Referral Manifest (printed)", "Referral file; scanned copy in the EMR",
              "Per national referral policy"],
             ["Downtime collection log and forms (temporary)", "Scanned into the EMR after use",
              "Per records policy"],
         ],
         [6.5, 5.5, 5.0]),
        ("p", "Records are created at the time the activity is performed, are legible and "
              "traceable to the person or user account, and are retrievable (ISO 15189:2022, "
              "8.4; RliDHL-ADM-SOP-004)."),
    ]),
    ("19. Supporting Documents", [
        ("b", [
            "Document and Record Control (RliDHL-ADM-SOP-004).",
            "Pre-Examination Procedure (RliDHL-ADM-SOP-005).",
            "Laboratory Safety Manual (post-exposure procedure).",
            "Laboratory Waste Management procedure.",
            "Result Reporting and Release procedure.",
            "Sample Referral procedure and the national Sample Referral Manifest.",
            "Competence Assessment and Training procedure.",
        ]),
    ]),
    ("20. References", [
        ("b", [
            "ISO 15189:2022 \u2014 Medical laboratories: requirements for quality and competence "
            "(Clauses 4.2, 4.3, 5.4, 5.6, 6.2, 6.3, 7.2, 7.5, 7.8, 8.3, 8.4, 8.5, 8.7).",
            "CLSI GP41 \u2014 Collection of Diagnostic Venous Blood Specimens.",
            "CLSI H3 \u2014 Procedures for the Handling and Processing of Blood Specimens.",
            "WHO \u2014 Laboratory Quality Management System (LQMS) Handbook.",
            "Government of Rwanda paperless operations policy (in force since 01 August 2026).",
            "RBC/NRL guidance on specimen referral and EQA schemes.",
        ]),
    ]),
]

SIG_TITLE = "21. Signature Page \u2014 Training and Competence Acknowledgement (Form F-TRN-001)"
SIG_STATEMENT = (
    "In normal operation this acknowledgement is electronic: each member of staff opens this "
    "SOP in the document module and confirms \u201cI have read and understood this version\u201d; "
    "the training module records the user account, date and SOP version (ISO 15189:2022, 6.2 "
    "and 8.4). The table below is the contingency record, used only when the training module "
    "is unavailable or for staff without individual system accounts. I confirm that I have "
    "been trained on this SOP, understand it including the safety precautions, and agree to "
    "perform it as written. I must not perform it unsupervised until this acknowledgement is "
    "recorded."
)
SIG_HEADERS = ["S/No", "Full Name (print)", "IPPIS No.", "Designation", "Signature", "Date"]
SIG_ROWS = [[str(i), "", "", "", "", ""] for i in range(1, 16)]
SIG_WIDTHS = [1.2, 4.6, 2.6, 3.4, 3.0, 2.2]
SIG_TRAINER = ("Trainer / Supervisor: ______________________    Name (print): ______________    "
               "Date: ____________")
SIG_NOTE = (
    "Enter the IPPIS number (staff cards are not numbered). 15 rows are provided \u2014 at least "
    "12 staff fit on this page. If more rows are needed, attach continuation sheet F-TRN-001-2. "
    "File the completed form in the personnel training record (electronic scan) and reference "
    "it in the SOP training register maintained by the Lab Quality Manager."
)

AMEND_TITLE = "22. Amendment Page"
AMEND_STATEMENT = (
    "In normal operation, amendments are made only by revision of the electronic document "
    "through the document control system (RliDHL-ADM-SOP-004): a new version is authorised and "
    "the review history records the change. Handwritten amendments are permitted only on a "
    "stamped \u201cCONTINGENCY COPY\u201d during information-system downtime. Each such amendment "
    "must be authorised in writing by the Laboratory Manager, entered here word for word, and "
    "superseded by a formal revision within 30 days."
)
AMEND_HEADERS = ["Date", "Section", "Amendment (word for word)", "Reason", "Authorised by"]
AMEND_ROWS = [["", "", "", "", ""] for _ in range(6)]
AMEND_WIDTHS = [2.4, 2.0, 6.4, 3.4, 2.8]
AMEND_NOTE = (
    "Only the Laboratory Manager may authorise a handwritten amendment. The original text "
    "must remain legible. This page is retained with the document (ISO 15189:2022, 8.3)."
)

FOOTER_LEFT = (META["doc_no"] + "  |  Version " + META["version"] +
               "  |  Controlled electronic document \u2014 printed copies are uncontrolled")
HEADER_LEFT = META["doc_no"] + " \u2014 " + META["title"]
HEADER_RIGHT = "Version " + META["version"]

INK = "#1F4E79"
INK2 = "#2E74B5"
GREY = "#595959"

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
        add_field(p, "NUMPAGES", size=7.5)
        tab_stop(p, 17.0)
        bottom_border(p)

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
    sec.first_page_header.paragraphs[0].text = ""

    # -- cover --------------------------------------------------------------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if LOGO_PATH.exists():
        p.add_run().add_picture(str(LOGO_PATH), height=Cm(2.4))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(META["lab"])
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

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
    r = p.add_run(META["title"])
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    make_table(["Field", "Detail"], [[k, v] for k, v in DOC_CONTROL_ROWS], [5.0, 12.0],
               body_size=9.5, first_col_bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(28)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("Approval and Authorisation")
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(INK.lstrip("#"))

    make_table(["Role", "Name (print)", "Signature", "Date"],
               [[role, name, "", ""] for role, name in APPROVAL_ROWS],
               [4.2, 5.3, 4.5, 3.0], body_size=9.5, first_col_bold=True, row_height=1.05)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run(APPROVAL_STATEMENT)
    r.font.size = Pt(8.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    doc.add_page_break()

    # -- review history + controlled copies + TOC ---------------------------
    doc.add_heading("Review History", level=1)
    make_table(REVISION_HEADERS, REVISION_ROWS, [3.0, 1.8, 3.2, 9.0], body_size=9.5)

    doc.add_heading("Controlled Copies and Supersedes", level=1)
    for item in CONTROLLED_COPIES:
        bp = doc.add_paragraph()
        bp.paragraph_format.left_indent = Cm(0.6)
        bp.paragraph_format.first_line_indent = Cm(-0.4)
        bp.paragraph_format.space_after = Pt(2)
        bp.add_run("\u2022  " + item)

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

    # -- body sections 1-20 -------------------------------------------------
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

    # -- 21. Signature Page -------------------------------------------------
    doc.add_heading(SIG_TITLE, level=1)
    sp = doc.add_paragraph(SIG_STATEMENT)
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    make_table(SIG_HEADERS, SIG_ROWS, SIG_WIDTHS, body_size=9.5, row_height=0.82)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.add_run(SIG_TRAINER).font.size = Pt(9.5)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run(SIG_NOTE)
    r.font.size = Pt(8.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

    doc.add_page_break()

    # -- 22. Amendment Page -------------------------------------------------
    doc.add_heading(AMEND_TITLE, level=1)
    sp = doc.add_paragraph(AMEND_STATEMENT)
    sp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    make_table(AMEND_HEADERS, AMEND_ROWS, AMEND_WIDTHS, body_size=9.5, row_height=1.2)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run(AMEND_NOTE)
    r.font.size = Pt(8.5)
    r.font.italic = True
    r.font.color.rgb = RGBColor.from_string(GREY.lstrip("#"))

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

    ink = colors.HexColor(INK)
    ink2 = colors.HexColor(INK2)
    grey = colors.HexColor(GREY)
    line_grey = colors.HexColor("#BFBFBF")
    head_fill = colors.HexColor(INK)

    styles = {
        "lab": ParagraphStyle("lab", fontName="Helvetica-Bold", fontSize=14,
                              textColor=ink, alignment=1, spaceAfter=4),
        "kind": ParagraphStyle("kind", fontName="Helvetica-Bold", fontSize=11.5,
                               textColor=grey, alignment=1, spaceAfter=2),
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=17,
                                textColor=ink, alignment=1, spaceAfter=10),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=12,
                             textColor=ink, spaceBefore=10, spaceAfter=5),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10.5,
                             textColor=ink2, spaceBefore=7, spaceAfter=3),
        "p": ParagraphStyle("p", fontName="Helvetica", fontSize=9.5, leading=12.4,
                            textColor=colors.black, spaceAfter=4, alignment=4),
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
            ("BACKGROUND", (0, 0), (-1, 0), head_fill),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, line_grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    def logo_flowable(height_cm=2.2):
        if not LOGO_PATH.exists():
            return Spacer(1, 0.1 * cm)
        raw = LOGO_PATH.read_bytes()
        w, h = struct.unpack(">II", raw[16:24])
        width = height_cm * (w / h)
        img = Image(str(LOGO_PATH), width=width * cm, height=height_cm * cm)
        img.hAlign = "CENTER"
        return img

    story = []

    # -- cover --------------------------------------------------------------
    story += [
        Spacer(1, 0.4 * cm),
        logo_flowable(2.2),
        Spacer(1, 0.25 * cm),
        P(META["lab"], "lab"),
        P("STANDARD OPERATING PROCEDURE", "kind"),
        P(META["title"], "title"),
        make_pdf_table(["Field", "Detail"],
                       [[k, v] for k, v in DOC_CONTROL_ROWS],
                       [5.0, 12.0], first_col_bold=True),
        Spacer(1, 1.1 * cm),
        P("Approval and Authorisation", "h2"),
        make_pdf_table(["Role", "Name (print)", "Signature", "Date"],
                       [[role, name, "", ""] for role, name in APPROVAL_ROWS],
                       [4.2, 5.3, 4.5, 3.0], first_col_bold=True,
                       row_heights=[None] + [0.95 * cm] * len(APPROVAL_ROWS)),
        Spacer(1, 0.3 * cm),
        P(APPROVAL_STATEMENT, "small"),
        PageBreak(),
    ]

    # -- review history + controlled copies + TOC ---------------------------
    story.append(H1("Review History"))
    story.append(make_pdf_table(REVISION_HEADERS, REVISION_ROWS,
                                [3.0, 1.8, 3.2, 9.0],
                                row_heights=[None] + [0.9 * cm] * len(REVISION_ROWS)))
    story.append(H1("Controlled Copies and Supersedes"))
    for item in CONTROLLED_COPIES:
        story.append(P("\u2022  " + item, "b"))
    story.append(Paragraph(escape("Table of Contents"), styles["h1"]))
    toc = TableOfContents()
    toc.levelStyles = [ParagraphStyle("toc0", fontName="Helvetica", fontSize=9.5,
                                      leading=13.5, textColor=colors.black)]
    story.append(toc)
    story.append(PageBreak())

    # -- body 1-20 ----------------------------------------------------------
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

    # -- 21. Signature Page -------------------------------------------------
    story.append(H1(SIG_TITLE))
    story.append(P(SIG_STATEMENT))
    story.append(Spacer(1, 0.15 * cm))
    story.append(make_pdf_table(SIG_HEADERS, SIG_ROWS, SIG_WIDTHS,
                                row_heights=[None] + [0.82 * cm] * len(SIG_ROWS)))
    story.append(Spacer(1, 0.4 * cm))
    story.append(P(SIG_TRAINER, "p"))
    story.append(Spacer(1, 0.15 * cm))
    story.append(P(SIG_NOTE, "small"))

    story.append(PageBreak())

    # -- 22. Amendment Page -------------------------------------------------
    story.append(H1(AMEND_TITLE))
    story.append(P(AMEND_STATEMENT))
    story.append(Spacer(1, 0.15 * cm))
    story.append(make_pdf_table(AMEND_HEADERS, AMEND_ROWS, AMEND_WIDTHS,
                                row_heights=[None] + [1.2 * cm] * len(AMEND_ROWS)))
    story.append(Spacer(1, 0.3 * cm))
    story.append(P(AMEND_NOTE, "small"))

    PDF_PATH.parent.mkdir(parents=True, exist_ok=True)
    doc = SOPDoc(str(PDF_PATH), pagesize=A4,
                 leftMargin=2.0 * cm, rightMargin=2.0 * cm,
                 topMargin=2.0 * cm, bottomMargin=2.0 * cm,
                 title=META["doc_no"] + " \u2014 " + META["title"],
                 author="Ruli District Hospital Laboratory Department",
                 subject="ISO 15189:2022-compliant SOP \u2014 2026 standard layout")
    frame = Frame(doc.leftMargin, doc.bottomMargin, FRAME_W, PAGE_H - 4.0 * cm, id="main")
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])
    doc.multiBuild(story, canvasmaker=NumberedCanvas)
    print(f"PDF written:  {PDF_PATH.relative_to(ROOT)}  ({doc.page} pages)")


if __name__ == "__main__":
    build_docx()
    build_pdf()
