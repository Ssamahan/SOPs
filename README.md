# SOPs

Standard Operating Procedure materials for the medical laboratory, built to the current
international standard **ISO 15189:2022** (Medical laboratories — requirements for quality
and competence).

## Contents

| Path | Description |
|---|---|
| [`docs/ISO-15189-2022-SOP-Guide.md`](docs/ISO-15189-2022-SOP-Guide.md) | How a compliant SOP must look: clause map, mandatory content, signature requirements, pagination & document-control rules |
| [`sop/SOP-PRE-001-Venous-Blood-Collection.pdf`](sop/SOP-PRE-001-Venous-Blood-Collection.pdf) | Print-exact sample SOP (7 pages) — cover with approval signatures, revision history + TOC, full procedure body, 15-row staff training signature sheet |
| [`sop/SOP-PRE-001-Venous-Blood-Collection.docx`](sop/SOP-PRE-001-Venous-Blood-Collection.docx) | Editable Word version of the same sample — use as the template for new SOPs |
| `scripts/build_sop.py` | Regenerates both files from one content model |

## Key conventions

- Fresh page starts: **cover** · **revision history + table of contents** · **body** ·
  **Annex A training & signature sheet**; body text flows continuously to save pages.
- Footer on every page: `SOP no | version | Page X of Y | uncontrolled-copy notice`.
- Signature capacity: Annex A provides **15 signature rows** (at least 12 staff per page,
  continuation sheet for more).

## Rebuilding

```bash
python3 -m venv .venv
.venv/bin/pip install python-docx reportlab
.venv/bin/python scripts/build_sop.py
```
