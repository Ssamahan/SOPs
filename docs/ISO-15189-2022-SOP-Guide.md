# How a Medical Laboratory SOP Must Look — ISO 15189:2022 (the current edition)

This guide is the "glimpse" of what the latest medical laboratory standard requires of a
Standard Operating Procedure (SOP), and how the sample in this repository implements it.

**The standard:** *ISO 15189:2022 — Medical laboratories: requirements for quality and
competence* (3rd edition, replacing ISO 15189:2012). It is the international standard used by
accreditation bodies to assess medical laboratories. The 2022 edition is **risk-based and
patient-centred**: patient welfare and rights are primary considerations (Clause 4.3), risk
thinking is built into the structure (5.6, 8.5), and competence must be *demonstrated*, not
merely documented (6.2).

> Always verify final requirements against the official standard text and your accreditation
> body's application documents. The sample content in `sop/` is illustrative (fictional lab,
> sample procedure) and must be adapted to your laboratory.

---

## 1. The ISO 15189:2022 structure at a glance

| Clause | Subject | What it means for your documents |
|---|---|---|
| 4.1–4.3 | Impartiality, confidentiality, requirements regarding patients | SOPs handling patient data state confidentiality duties; patient welfare is paramount |
| 5.1–5.6 | Legal entity, laboratory director, laboratory activities, structure & authority, objectives & policies, **risk management** | Who may sign what; scope of work; risk management is a governance duty |
| 6.1–6.8 | General, **personnel**, facilities & environment, equipment, calibration & metrological traceability, reagents & consumables, service agreements, externally provided products & services | Training + competence assessment + **authorisation records** for every procedure staff perform |
| 7.1–7.8 | General, **pre-examination**, **examination**, **post-examination** processes, nonconforming work, control of data & information management, complaints, **continuity & emergency preparedness planning** (new in 2022) | The technical heart: every test and every sample needs a written procedure |
| 8.1–8.9 | General, management system documentation, **control of documents (8.3)**, **control of records (8.4)**, actions on risks & opportunities (8.5), improvement (8.6), nonconformities & corrective actions (8.7), evaluations (8.8), management reviews (8.9) | How documents are numbered, approved, reviewed, updated, withdrawn — and how records are retained |

## 2. What ISO 15189:2022 requires INSIDE every SOP

An SOP must be written so that a trained employee can perform the task **without guessing**.
Minimum content elements (this is the checklist auditors use):

| # | Element | Why / clause |
|---|---|---|
| 1 | Document number, title, version, effective date, review date | Unique identification and version control (8.3) |
| 2 | Approval signatures: prepared / reviewed / approved (authorised by) | Documents are reviewed and approved by authorised personnel **before issue** (8.3) |
| 3 | Revision history table | Traceability of changes; obsolete versions identifiable (8.3) |
| 4 | Purpose | Links the SOP to the quality objective it serves |
| 5 | Scope — what is in and out | Defines where the procedure applies (5.3, 7.1) |
| 6 | Definitions and abbreviations | Removes ambiguity |
| 7 | Responsibilities — who does what | Authority and responsibility are defined (5.4) |
| 8 | Safety and precautions | Patient and personnel safety (4.3, 6.3) |
| 9 | Materials, equipment, reagents | Traceable, accepted consumables (6.4–6.6) |
| 10 | Step-by-step procedure — numbered, unambiguous | The core examination/pre-examination process (7.2–7.4) |
| 11 | Acceptance/rejection criteria and handling of nonconforming work | Nonconforming work is controlled (7.5) |
| 12 | Quality control / monitoring / quality indicators | Validity of results and evaluation (7.3, 8.8) |
| 13 | Risk considerations | Risk-based thinking (5.6, 8.5) — a hallmark of the 2022 edition |
| 14 | Documentation and records: what is recorded, corrections, retention | Records created at the time of the activity, legible, retrievable (8.4) |
| 15 | References (standard, CLSI/ISO methods, national rules) | Method traceability |
| 16 | Related documents and forms | Links the procedure into the wider management system |
| 17 | **Training & competence acknowledgement signature sheet** | Staff confirm they were trained and agree to comply; evidence of competence and authorisation (6.2); retained as a record (8.4) |

## 3. Signatures the standard expects

1. **Approval block (cover page)** — typically three signatories:
   *Prepared by* (author), *Reviewed by* (lab manager / senior scientist),
   *Approved by* (laboratory director or delegate — the person authorised to release the
   document for use). The document **must not be used before approval** (Clause 8.3).
2. **Staff training & competence sheet (Annex A)** — one row per staff member who will
   perform the procedure: printed name, staff ID, designation, **signature**, date. The
   sample provides **15 rows** so a section of **at least 12 staff can sign on one page**;
   continuation sheets are provided for larger teams. Competence must additionally be
   *assessed* (observation, records review) before authorisation — the signed sheet is the
   acknowledgement part of that training record (Clauses 6.2, 8.4).

## 4. Pagination rules used in this template

Sections that **must start on a fresh page**:

| Fresh page | Content |
|---|---|
| Page 1 | Cover: document-control table + approval signatures + controlled-copy statement |
| Page 2 | Revision history + distribution/control of copies + table of contents |
| Page 3 | Body (Purpose …) — flows continuously from here |
| Last page | Annex A: training & competence signature sheet (never share a page with body text — it is a form that is signed, scanned and filed) |

Everything else **flows continuously** so no page is wasted: headings are *kept-with-next*
style blocks, tables follow their introductions directly, and running header/footer carry
`Document no — title | Version | Effective date` and **`Page X of Y`** on every page
(so a photocopy shows if pages are missing — a document-control requirement of 8.3).
The footer also repeats the **"Uncontrolled when printed unless stamped CONTROLLED COPY"**
statement.

## 5. Document-control features visible in the sample

- Unique document number scheme: `SOP-XX-nnn` (process code + serial), forms `F-XX-nnn`.
- Version `1.0` at first issue; every later change increments it and is logged in the
  revision history with date, description, who prepared and who approved.
- **Next review date** on the cover (at least every 2 years, or earlier on change,
  incident, or nonconformity — the 2022 edition expects documents to remain suitable).
- Distribution and control-of-copies statement (master copy + controlled copies list).
- Table of contents with page numbers (auto-updating field in Word).
- Obsolete copies: withdrawn and destroyed or stamped `OBSOLETE` and retained for traceability.

## 6. What's in this repository

| File | What it is |
|---|---|
| `sop/SOP-PRE-001-Venous-Blood-Collection.pdf` | **The glimpse** — print-exact 7-page rendering of a complete, compliant SOP sample (approved cover, revision history + TOC, full procedure body, 15-row staff signature sheet) |
| `sop/SOP-PRE-001-Venous-Blood-Collection.docx` | The same document as an **editable Word file** — copy it, replace the sample content, keep the structure |
| `scripts/build_sop.py` | Generator that produces both files from one content model (edit content there and re-run) |
| `docs/ISO-15189-2022-SOP-Guide.md` | This guide |

### How to reuse the sample for your own SOPs

1. Copy `sop/SOP-PRE-001-….docx`, rename to your numbering scheme (`SOP-<AREA>-<nnn>`).
2. Replace the sample content section by section — keep the document-control table,
   approval block, revision history, and Annex A intact.
3. Update revision history (`1.0`, issue date, your name), fill the three approval rows,
   set effective/review dates.
4. In Word, open the document and **update fields** (Ctrl+A then F9) to rebuild the table
   of contents and `Page X of Y` footers.
5. Print the approval page for wet-ink signatures; the signed master goes to the Quality
   Office. Train all staff who will perform the procedure and have them sign Annex A
   **before** they work unsupervised.
6. Alternative: edit `scripts/build_sop.py` (content model at the top) and re-run
   `.venv/bin/python scripts/build_sop.py` to regenerate both outputs.

## 7. Quick anatomy of a compliant SOP (one-page picture)

```
┌───────────────────────────────────────────────┐  PAGE 1 (fresh)
│  [logo]  LABORATORY NAME                      │
│  STANDARD OPERATING PROCEDURE                 │
│  <Title of the procedure>                     │
│  ┌ Document control table ─────────────────┐  │  Doc no | title | version
│  │ number, title, version, effective,      │  │  effective | review | supersedes
│  │ review date, supersedes, owner          │  │
│  └─────────────────────────────────────────┘  │
│  ┌ Approval & authorisation ───────────────┐  │  Prepared / Reviewed / Approved
│  │ role | name | SIGNATURE | date          │  │  (wet signatures — never use yet)
│  └─────────────────────────────────────────┘  │
│  Controlled-copy statement (8.3, 8.4)         │
├───────────────────────────────────────────────┤  PAGE 2 (fresh)
│  Revision history table                       │
│  Distribution & control of copies             │
│  Table of contents (with page numbers)        │
├───────────────────────────────────────────────┤  PAGE 3+ (continuous flow)
│  1 Purpose        7 Procedure (7.1, 7.2 …)    │
│  2 Scope          8 Quality control           │
│  3 Definitions    9 Risk considerations       │
│  4 Responsibilities  10 Records               │
│  5 Safety         11 References               │
│  6 Materials      12 Related documents        │
├───────────────────────────────────────────────┤  LAST PAGE (fresh)
│  ANNEX A — Training & competence              │
│  "I have been trained … I agree to comply"    │
│  15 rows: no | name | ID | role | SIGN | date │  ≥ 12 staff fit on one page
│  Trainer signature block · continuation note   │
└───────────────────────────────────────────────┘
   Footer on every page: SOP no | version | "Page X of Y" | uncontrolled-copy notice
```

---

## Notes for the Ruli District Hospital rollout

- **Document codes**: every code carries the facility acronym — `RliDHL-XXX-SOP-NNN-VERS-NNN`
  (hyphens throughout), e.g. `RliDHL-IH-SOP-001-VERS-001`. Section codes are the three-letter
  codes of the Document and Record Control SOP (ADM, BAC, BIO, HEM, IMM, MOL, MYC, PAR, REC,
  SER, STC, SAF). A number belongs to one document for life and never changes at revision;
  only the version rises.
- **Employee ID on the signature page** = the **IPPIS number** (Integrated Personnel and
  Payroll Information System), since staff cards are not numbered. It is the unique staff
  identifier used in personnel and training records (ISO 15189:2022, 6.2).
- **Paperless operation (since 01 August 2026)**: SOPs are controlled electronic documents;
  authorisation and staff acknowledgement are electronic (user account + timestamp in the
  training module); records live in the EMR/LIS. The only routine print is the Sample Referral
  Manifest for referred specimens. Printed SOP copies are contingency-only, stamped
  "CONTINGENCY COPY", used during information-system downtime (ISO 15189:2022, 7.8) and
  destroyed afterwards — which is why every page carries "printed copies are uncontrolled".
  Design SOPs to be read on screen first: bookmarks/TOC links, no layout that depends on
  printing.

## Local conventions (Ruli District Hospital)

- **Hospital contact (from 14/11/2026): Tel (+250) 0781 316 368** — replaces (+250) 0782 000 853 on every document once it is updated or revised.
- **Normal Ranges list notation:** an entry written `>x` means any result under `x` is normal — `x` is the upper limit of the normal range (e.g. Gamma GT `>51 IU/L` → normal under 51 IU/L). Quote such entries as "up to x" or "< x" in SOPs.
- **Values source order:** newest approved local list first (Normal Ranges 2026, Critical Results, TAT 2026), then equipment/stock records (Laboratory Medical Equipment Presence List, RMS purchase orders via MediSell Rwanda Ltd), then manufacturer documentation.
