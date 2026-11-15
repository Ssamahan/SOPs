# SOPs — Ruli District Hospital Laboratory Department

Standard Operating Procedure materials built to **ISO 15189:2022** and the laboratory's
2026 standard SOP layout, designed for **paperless operation** (Rwanda, since 01 August 2026:
electronic requests and records in the EMR; only referred specimens travel with a printed
Sample Referral Manifest).

## Contents

| Path | Description |
|---|---|
| `docs/ISO-15189-2022-SOP-Guide.md` | How a compliant SOP must look: clause map, mandatory content, signatures, pagination & document control |
| `sop/RliDHL-HEM-SOP-001-VERS-002 - Venous Blood Sample Collection.pdf` | Print/screen-exact sample SOP (10 pages): cover with hospital logo and real signatories, review history + TOC, 20-section body, signature page (**18 rows**, IPPIS No. column, competence-assessment and authorisation blocks), amendment page |
| `sop/RliDHL-HEM-SOP-001-VERS-002 - Venous Blood Sample Collection.docx` | Editable Word version of the same document |
| `sop/RliDHL-HEM-SOP-005-VERS-001 - Sysmex XN-330 examination.{pdf,docx}` | Examination SOP for the Sysmex XN-330 (9 pages) — the analyser in service had none |
| `sop/RliDHL-HEM-SOP-006-VERS-001 - Sysmex XQ-320 examination.{pdf,docx}` | Examination SOP for the Sysmex XQ-320 (9 pages) |
| `sop/RliDHL-IMM-SOP-005-VERS-001 - Alere Pima CD4 examination.{pdf,docx}` | CD4 enumeration on the Alere Pima CD4 (9 pages) — absolute count and percentage, stability limits, critical thresholds |
| `sop/RliDHL-ADM-SOP-025-VERS-001 - External Quality Assessment and Proficiency Testing.{pdf,docx}` | **EQA/PT** (12 pages) — annual plan built from the Scope of Test list, enrolment, immediate testing of panels, routine-person/routine-method rule, submission, review within one month, ±2 SDI criteria, investigation by cause, verification of effectiveness, and the assessment of whether **patients** were affected. Written against the EQA the laboratory actually performs: HIV serology (DTS PT, RBC/NRL), TB on GeneXpert (participant 20044), biochemistry and FBC (NRL quarterly) |
| `sop/RliDHL-ADM-SOP-065-VERS-001 - Troubleshooting of laboratory analysers.{pdf,docx}` | **Error codes, causes and corrective action for every analyser** (16 pages) — Sysmex XN-330 and XQ-320, Roche cobas c 111, GeneXpert Dx (all assays), Alere Pima CD4. Each row: code or message, meaning, cause as the manufacturer states it, operator action, action if it persists, whether the module is withdrawn, who is responsible. Sourced from the manufacturers' own manuals; see section 18.2 |
| `sop/RliDHL-BIO-SOP-001-VERS-005 - Glucose Testing using Cobas c 111.{pdf,docx}` | Glucose examination SOP sample (11 pages) — reference intervals in both units, clinical decision limits, critical results with the escalation route, measurement uncertainty, comparability of results, interpretive notes |
| `assets/Ruli-District-Hospital-logo.png` | Hospital logo used on covers |
| `scripts/build_sop.py` | Regenerates both files from one content model; every code comes from `REGISTER` |
| `scripts/check_codes.py` | Fails if a serial is skipped or shared, if a reference does not resolve, or if a code is hand-typed |
| `docs/master-list.md` | Numbering rule, register extract, and the defects found in the Drive Master List |

## Conventions

- **Document codes carry the facility acronym**: `RliDHL-<AREA>-<TYPE>-<NNN>-VERS-<NNN>`
  (e.g. `RliDHL-HEM-SOP-001-VERS-002`). **Serials are sequential and do not skip**: the next
  document takes the next free number in its series, a number belongs to one document for
  life, and a withdrawn number is retired rather than re-used.
- **Forms carry their parent SOP's code**: `RliDHL-ADM-SOP-064-F01`. A free-standing
  `F-TRN-001` was the previous practice and is withdrawn — it tied the form to nothing.
- **Change a code in one place.** Edit `REGISTER` in `scripts/build_sop.py`, rebuild, then
  run `scripts/check_codes.py`. File names, cover tables, headers/footers and every
  cross-reference follow automatically; the checker fails the build if any reference is
  left dangling, and the build deletes output still named with a retired code.
- **Signature page capacity**: 18 rows on one page, so a section of **12 staff signs
  without turning over**; continuation sheet `RliDHL-ADM-SOP-064-F01-2` for larger teams. Competence is
  *assessed* (observation, records review, problem-solving, provided samples) and the
  person is *authorised* — acknowledgement alone is not sufficient (ISO 15189:2022, 6.2.2–6.2.3).
- **Page footer**: `Document no | Version` + `Page X of Y` on every page, with
  "Uncontrolled when printed unless stamped CONTROLLED COPY" on a second line of every page
  footer (8.3), so a detached page still shows its own status.
- **Signatories**: Prepared by the author (Lab Quality Manager Jean Paul NSENGIYUMVA),
  Reviewed by the Laboratory Manager (Kelvin SAMAHANI), Authorized by the Hospital's
  Director General (Dr. Déogratias KANEZA).
- **Employee IDs are IPPIS numbers** (staff cards are not numbered).
- **Paperless**: controlled electronic document; e-signatures and electronic acknowledgement
  in the training module; printed copies are contingency-only ("CONTINGENCY COPY");
  footer marks every page "printed copies are uncontrolled".
- Fresh page starts: cover · review history + controlled copies + TOC · body (1–20) ·
  signature page (21) · amendment page (22).

## Rebuilding

```bash
python3 -m venv .venv
.venv/bin/pip install python-docx reportlab pymupdf
.venv/bin/python scripts/build_sop.py     # rebuild from REGISTER
.venv/bin/python scripts/check_codes.py   # prove the numbering rules hold
```

`check_codes.py` exits non-zero on a skipped serial, a shared number, a reference to an
unregistered code, a form not bound to its parent SOP, or a hand-typed code. Run it before
committing any change that touches a document code.
