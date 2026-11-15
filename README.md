# SOPs — Ruli District Hospital Laboratory Department

Standard Operating Procedure materials built to **ISO 15189:2022** and the laboratory's
2026 standard SOP layout, designed for **paperless operation** (Rwanda, since 01 August 2026:
electronic requests and records in the EMR; only referred specimens travel with a printed
Sample Referral Manifest).

## Contents

| Path | Description |
|---|---|
| `docs/ISO-15189-2022-SOP-Guide.md` | How a compliant SOP must look: clause map, mandatory content, signatures, pagination & document control |
| `sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.pdf` | Print/screen-exact sample SOP (12 pages): cover with hospital logo and real signatories, review history + TOC, 22-section body, signature page (15 rows, IPPIS No. column), amendment page |
| `sop/RliDHL-IH-SOP-001-VERS-001 - Venous Blood Sample Collection.docx` | Editable Word version of the same document |
| `sop/RliDHL-BIO-SOP-001-VERS-005 - Glucose Testing using Cobas c 111.{pdf,docx}` | Glucose examination SOP sample — shows handling of reference intervals (both units), clinical decision limits, critical values and interpretive notes (causes of high/low glucose) |
| `assets/Ruli-District-Hospital-logo.png` | Hospital logo used on covers |
| `scripts/build_sop.py` | Regenerates both files from one content model |

## Conventions

- **Document codes carry the facility acronym**: `RliDHL-XXX-SOP-NNN-VERS-NNN`
  (e.g. `RliDHL-IH-SOP-001-VERS-001`). Forms: `F-TRN-001` etc. as named in the SOPs.
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
