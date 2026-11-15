# SOPs — Ruli District Hospital Laboratory Department

Standard Operating Procedure materials built to **ISO 15189:2022** and the laboratory's
2026 standard SOP layout, designed for **paperless operation** (Rwanda, since 01 August 2026:
electronic requests and records in the EMR; only referred specimens travel with a printed
Sample Referral Manifest).

## Contents

| Path | Description |
|---|---|
| `docs/ISO-15189-2022-SOP-Guide.md` | How a compliant SOP must look: clause map, mandatory content, signatures, pagination & document control |
| `sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.pdf` | Print/screen-exact sample SOP (10 pages): cover with hospital logo and real signatories, review history + TOC, 20-section body, signature page (**18 rows**, IPPIS No. column, competence-assessment and authorisation blocks), amendment page |
| `sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.docx` | Editable Word version of the same document |
| `sop/RliDHL-BIO-SOP-001-VERS-005 - Glucose Testing using Cobas c 111.{pdf,docx}` | Glucose examination SOP sample (11 pages) — reference intervals in both units, clinical decision limits, critical results with the escalation route, measurement uncertainty, comparability of results, interpretive notes |
| `assets/Ruli-District-Hospital-logo.png` | Hospital logo used on covers |
| `scripts/build_sop.py` | Regenerates both files from one content model |

## Conventions

- **Document codes carry the facility acronym**: `RliDHL-XXX-SOP-NNN-VERS-NNN`
  (e.g. `RliDHL-IH-SOP-001-VERS-001`). Forms: `F-TRN-001` etc. as named in the SOPs.
- **Signature page capacity**: 18 rows on one page, so a section of **12 staff signs
  without turning over**; continuation sheet F-TRN-001-2 for larger teams. Competence is
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
.venv/bin/pip install python-docx reportlab
.venv/bin/python scripts/build_sop.py
```
