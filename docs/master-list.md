# Master List — document register for this repository

The Master List in Google Drive is the controlled register. This file is the
**working extract** of it that the SOPs in this repository depend on, so that a
code change is made in one place and verified automatically.

The build (`scripts/build_sop.py`) takes every code from `REGISTER` and the
check (`scripts/check_codes.py`) fails if a document contains a code that is not
registered here. Nothing in the generated SOPs carries a hand-typed code.

---

## 1. Numbering rule

```
RliDHL-<AREA>-<TYPE>-<NNN>-VERS-<NNN>        e.g. RliDHL-ADM-SOP-004-VERS-004
RliDHL-<AREA>-<TYPE>-<NNN>-F<NN>             e.g. RliDHL-ADM-SOP-060-F01
```

| Part | Rule |
|---|---|
| `AREA` | 3-letter section: `ADM` `BAC` `BIO` `HEM` `IH` `IMM` `MOL` `MYC` `PAR` `REC` `SER` `STC` `SAF` |
| `TYPE` | `MAN` manual · `GUID` guideline · `SOP` |
| `NNN` | Serial. **Sequential, no gaps, within one AREA+TYPE series.** Allocated once, held for life. Never changes at revision. A number released by a withdrawn document is not re-used. |
| `VERS` | Version, `001` at first issue. |
| `F<NN>` | Form. **Carries its parent SOP's code**, so a form cannot drift away from the SOP that controls it. |

Three consequences, all enforced by `scripts/check_codes.py`:

1. **No skipped numbers.** If a serial is missing from a series it is either a
   withdrawn document (recorded, number retired) or an error.
2. **No shared number.** Two documents must never answer to one code.
3. **No orphan reference.** Every code cited in an SOP, and every form cited,
   must exist in this register.

---

## 2. Crosscutting SOPs (ADM series)

Serials known to be in use, taken from the Master List:

| Serial | Code | Title | Vers | Status |
|---|---|---|---|---|
| 001 | RliDHL-ADM-SOP-001-VERS-004 | Laboratory Temperature Control | 004 | in force |
| 002 | RliDHL-ADM-SOP-002-VERS-002 | Documentation of Examination Procedure | 002 | in force |
| 004 | RliDHL-ADM-SOP-004-VERS-004 | Document and Record Control | 004 | in force — cited by every SOP here for 8.3 document control |
| 009 | RliDHL-ADM-SOP-009-VERS-002 | Disposal of sharps and needles | 002 | in force |
| 020 | RliDHL-ADM-SOP-020-VERS-004 | Pre-Examination Procedure | 004 | in force — cited by both SOPs here |
| 024 | RliDHL-ADM-SOP-024-VERS-004 | Quality Control and Quality Assurance | 004 | in force |
| 034 | RliDHL-ADM-SOP-034-VERS-003 | Laboratory Waste Management | 003 | in force |
| 053–059 | RliDHL-ADM-SOP-053…059-VERS-001 | Equipment-maintenance block (microscope, centrifuge, water bath, non-medical equipment, ICT/LIS, hematology analysers, GeneXpert) | 001 | issued 14/11/2026, awaiting approval |
| **060** | **RliDHL-ADM-SOP-060-VERS-001** | **Training, Competence Assessment and Authorisation of Personnel** | **001** | **NEW — see §4** |

**Next free ADM serial: 061.** Check the Master List again before allocating it.

### Defects found in the Master List while cross-checking

These are recorded, not silently corrected, because correcting them is a
document-control act that needs the Master List's own sign-off:

- **Serial 001 is used twice** — *Laboratory Temperature Control* and
  *Document and Record Control* both appear as `RliDHL-ADM-SOP-001-VERS-004`.
  Document and Record Control is 004 in the numbering actually in force (it is
  the code every other SOP cites for 8.3), so the row reading `001` for it is
  the error.
- **`RliDHL-ADM-SOP-034` is used twice** — *Laboratory Waste Management* and
  *Handling of biological spills*.
- **`RliDHL-ADM-SOP-020` and `024` are outside the 001–006 block** used by the
  first six crosscutting SOPs, so the series already runs non-contiguously.
  Adopting the no-gap rule means these gaps are now retired numbers, or the
  rows are renumbered once, at the next revision of each document.

---

## 3. Documents in this repository

| Code | Title | Pages | Notes |
|---|---|---|---|
| `RliDHL-IH-SOP-002-VERS-001` | Venous Blood Sample Collection by Venipuncture | 10 | **Serial corrected.** It was built as `IH-SOP-001`, which the Master List already assigns to *Vein Blood Collection and Preparation of Hematology sample* (HEM, v003). Two different documents would have shared one code. `IH-SOP-001` is taken, so this is `IH-SOP-002`. |
| `RliDHL-BIO-SOP-001-VERS-005` | Glucose Testing using Cobas c 111 | 11 | Serial 001 and version 005 confirmed against the Master List (v004 effective 16/03/2025). |

### Forms

| Form | Title | Parent SOP | Status |
|---|---|---|---|
| `RliDHL-ADM-SOP-060-F01` | Training and competence acknowledgement and authorisation record (section 21 of every SOP) | 060 | NEW |
| `RliDHL-ADM-SOP-060-F01-2` | Continuation sheet, for a section larger than one page | 060 | NEW |

The forms were previously cited as `F-TRN-001` / `F-TRN-001-2`, which broke both
rules: the `TRN` area code does not exist, and the form was not tied to a parent
SOP. They now take the code of the SOP that controls them.

---

## 4. Why 060 exists

The signature page in every SOP is the record of training, competence and
authorisation (ISO 15189:2022, 6.2.2, 6.2.3). An audit found that the Master
List contained **no SOP that governs training, competence assessment,
authorisation or the signing of SOPs** — so the record had no parent document,
and the forms had no owner.

`RliDHL-ADM-SOP-060` is that parent. It is a first issue, and it must be written,
approved and added to the Master List before the signature pages it controls can
be relied on as evidence.

Two existing ADM SOPs cover part of the ground and are **not** duplicates of it:

- `020` Pre-Examination Procedure — covers specimen collection, not people.
- `024` Quality Control and Quality Assurance — covers QC, not competence.

---

## 5. Propagation when a code changes

Changing a code is not a text edit. The order is:

1. Edit `REGISTER` in `scripts/build_sop.py`.
2. Rebuild: `.venv/bin/python scripts/build_sop.py` — file names, cover tables,
   running headers/footers, cross-reference lists and form codes all follow.
3. Check: `.venv/bin/python scripts/check_codes.py` — fails on an unregistered
   code, a gap, a duplicate or an orphan form.
4. Update this file and the **Master List** in Drive in the same act, so the
   register and the documents never disagree.
5. If the change is a renumber, raise it as a revision of each affected SOP and
   record it in that SOP's revision history — a code change is a controlled
   amendment, not a typo fix.
