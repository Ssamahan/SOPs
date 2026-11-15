# Master List extract — authoritative for document numbering

**Source of truth:** `MASTER LIST 2025.docx` (Drive `1cl3zHlZ2-SDW-1gUrrwrYVeRA7-YwPuI`).
This file is a working extract of it for the documents in this repository. If the two
disagree, the Drive **Master List** wins and this file is corrected.

> Read `SOP-PROGRAMME-HANDOFF.md` (§0 BLUF, §8 numbering, §11 open items) before touching
> anything here. A handoff doc that is not read is a handoff doc that is not worth writing.

---

## 1. Numbering rule

```
RliDHL-<AREA>-<TYPE>-<NNN>-VERS-<NNN>     e.g. RliDHL-ADM-SOP-004-VERS-001
RliDHL-<AREA>-<TYPE>-<NNN>-F<NN>          e.g. RliDHL-ADM-SOP-062-F01
```

| Part | Rule |
|---|---|
| `AREA` | Dept code. **The register uses exactly these:** `ADM` `BAC` `BIO` `HEM` `IMM` `MYC` `PAR` `SER` `STC`. **There is no `IH` section** — `IH` was introduced by this library, so `IH-SOP-002` creates a series rather than joining one. See §5. |
| `TYPE` | `MAN` · `GUID` · `SOP` |
| `NNN` | Serial, allocated once, held for life, never changed at revision. |
| `VERS` | Version, `001` at first issue. |
| `F<NN>` | Form, carrying its **parent SOP's** code so it cannot drift. |

**ISO 15189:2022 mandates uniqueness, not a scheme.** 8.3.2 a) requires unique
identification. Sequential-with-no-gaps is the laboratory's own convention on top of that,
and it is worth keeping because a gap is indistinguishable from a lost document.

### The rule that actually matters here

**Check the register before allocating a number. Never infer "next free" from a working
note.** The 053 error in the previous session, and the ADM-060 error in this one, were
both the same mistake: a plausible number taken from memory instead of from the register.

---

## 2. ADM crosscutting SOPs — the authoritative inventory

49 rows are registered. `scripts/check_codes.py` uses this table, so a number cannot be
allocated twice.

| # | Title | # | Title |
|---|---|---|---|
| 001 | All Document and record control | 025 | Calibration of equipment |
| 002 | Competency assessment | 026 | Pre-examination procedure |
| 003 | Confidentiality and Access Control | 027 | RliDHL Specimen management |
| 004 | Laboratory Temperature Control | 028 | Laboratory Advisory Services |
| 005 | Quality control and quality assurance | 029 | Internal Audit |
| 006 | Customer satisfaction Survey | 030 | Continual Improvement |
| 007 | Conflict of interest | 031 | Reporting and release of Results |
| 008 | **Authorization** — collides, see §3 | 032 | Correction of clerical errors |
| 009 | Service interruption and delay in service delivery | 033 | Laboratory personnel training |
| 010 | All Communication sop | 034 | Laboratory Waste Management — collides, see §3 |
| 011 | Service agreements | **035** | **never used** — see §3 |
| 012 | Purchasing and Inventory | 036 | Personnel records (files) |
| 013 | Lab Stock requisition by RliDHL staff | 037 | Training and Supervision of health centre staff |
| 014 | Laboratory Stock Management of Lab Consumables | 038 | Examination by referral laboratories |
| 015 | Resolution of Complaints | 039 | Equipment Selection, Verification and Validation |
| 016 | Management review | 040 | External service and suppliers |
| 017 | Identification and control of nonconformities, CAPA | 041 | Documentation of examination procedures |
| 018 | Accommodation and Environmental Conditions | 042 | Laboratory information system (non-computerised) |
| 019 | Decontamination of work space | 043 | Review of staff Performance |
| 020 | Decontamination of equipment | 044 | Personnel management |
| 021 | Laboratory Waste Management | 045 | RISK MANAGEMENT |
| 022 | Decontamination of specimens packaging materials | 046 | Ethical Conduct |
| 023 | Management of a post-exposition to an infectious material | 047 | LABORATORY CONTINGENCY PLAN |
| 024 | Equipment Management | 048 | MEASUREMENT UNCERTAINTY |
| | | 049 | BIOLOGICAL REFERENCE INTERVALS-1 — collides, see §3 |

**Above 049, allocated but NOT in the Master List** (the November 2026 equipment block,
built and uploaded, unsigned):

| # | Code | Title |
|---|---|---|
| 050 | RliDHL-ADM-SOP-050-VERS-001 | Use, Care and Preventive Maintenance of the Microscope |
| 051 | RliDHL-ADM-SOP-051-VERS-001 | Operation, Care and Preventive Maintenance of the Centrifuge |
| 052 | RliDHL-ADM-SOP-052-VERS-001 | Operation, Care and Preventive Maintenance of the Water Bath |
| 060 | RliDHL-ADM-SOP-060-VERS-001 | Management and Maintenance of Non-Medical Laboratory Equipment |
| 061 | RliDHL-ADM-SOP-061-VERS-001 | Maintenance of Laboratory ICT Equipment and the LIS |
| 062 | RliDHL-ADM-SOP-062-VERS-001 | Use, Care and Preventive Maintenance of Hematology Analysers |
| 063 | RliDHL-ADM-SOP-063-VERS-001 | Use, Care and Preventive Maintenance of the GeneXpert Dx System |

Forms in that block: `ADM-SOP-060-F01` (temperature monitoring), `-062-F01` (haematology
log), `-063-F01` (GeneXpert log).

### Why the block jumps 052 → 060

Files exist on disk above 52 that were never registered — notably
`52. Resolution of complaints` and `ADM-SOP-008-VERS-005`. So **053–059 cannot be proven
free**, and the pack starts again at 060, beyond anything the laboratory holds. **This is
why 060 is taken, and why the next free ADM serial is 064.**

---

## 3. Every SOP series in the register, as read on 15/11/2026

Read line by line from `MASTER LIST 2025.docx`. "Entries" counts distinct serials; the
"gap" column is what is absent from the sequence, not what is free.

| Series | Sections in the register | Serials seen | Gap | What the series actually holds |
|---|---|---|---|---|
| `ADM-SOP` | Crosscutting | 001–049 | 016, 017 (unreadable), 035 (never used) | Management SOPs: document control, complaints, training, audit |
| `BAC-SOP` | Bacteriology & parasitology | 001–006 | none | Culture and staining procedures |
| `BIO-SOP` | Biochemistry | 001–009 | none | Humastar 80 / Humalyzer 3500 analyses |
| `HEM-SOP` | Haematology | 001–004 | none | Blood collection, Sysmex and Humacount analyses |
| `IMM-SOP` | Immunology | 001–004 | none | FACSCount CD4 procedures |
| `MYC-SOP` | Mycobacteriology | 002–012 | **001** | GeneXpert / TB procedures |
| `PAR-SOP` | Parasitology | 001–010 | none | Parasitology procedures |
| `SER-SOP` | Serology | 001–007 | none | Serology procedures |
| `STC-SOP` | Stock / stores | 001–007 | none | Commodity and reagent management |

The `EXT-` codes throughout the register (`EXT-RliDHL-BIO-001-2016`, etc.) are
**manufacturer manuals**, not laboratory SOPs. They do not consume SOP serials.

### Two collisions this repository had not noticed

| # | Document | Registered to | Consequence |
|---|---|---|---|
| 1 | `RliDHL-BIO-SOP-001-VERS-005` — Glucose Testing using Cobas c 111 | `RliDHL-BIO-SOP 001-VERS 002` — *Use and preventive maintenance of Humastar 80*, 25/05/2016 | The glucose SOP holds a serial the register gives to a different, in-force document. Glucose in the registered series is **BIO-007**. |
| 2 | `RliDHL-IH-SOP-002-VERS-001` — Venous Blood Sample Collection | nothing — **no IH series exists** | `HEM-SOP-001-VERS-001` is *Vein blood collection and preparation of hematology sample* (16/08/2016). The new SOP very likely belongs there as a revision, not in a new department. |

Both are listed in `DECISIONS` in `scripts/build_sop.py` and printed by
`scripts/check_codes.py` after its passes. Neither is changed unilaterally: renumbering a
signed SOP, or creating a department code, is the Laboratory Manager's and Director
General's decision.

---

## 4. Defects recorded, not fixed

These need the Laboratory Manager and Director General. An agent must not renumber the
register.

| Defect | Detail | Consequence |
|---|---|---|
| **008 on two documents** | Legacy *Authorization* (008 v001) **and** *Use and Maintenance of Biomedical Equipment* **VERS-005** | **Already propagated:** all seven equipment SOPs cite `ADM-SOP-008-VERS-005` in their reference lists. A reader following that citation may open the wrong document. |
| **049 used twice** | *Biological Reference Intervals-1* and *Meet the Clinicians* | Two documents, one number |
| **034 used twice** | *Laboratory Waste Management* (registered twice as 021 and 034, and as 034 again for spills in the older list) | Ambiguous citation |
| **035 never used** | Gap at 035 | Indistinguishable from a lost document. Under the no-gap convention it should be declared retired. |
| **`52. Resolution of complaints`** | On disk, never registered | 052 cannot be proven free retroactively |
| **`ADM-FORM-014` "Signature page"** | Register holds a controlled form of that exact title, plus `ADM-FORM-001` *Staff competency assessment form*. Form `064-F01` covers the same ground. | Two controlled forms for one purpose. Either `064-F01` extends them or they are withdrawn — needs the signatories. |

### 008 is the one that matters

`ADM-SOP-008-VERS-005` is cited by all seven equipment SOPs. Whichever document keeps 008,
the other must be renumbered and **every** reference list updated in the same act. Until
that decision is taken, those seven SOPs are pointing at a number that identifies two
documents — which is exactly the failure this repository's checker exists to prevent.

---

## 5. Documents in this repository

| Code | Title | Pages | Note |
|---|---|---|---|
| `RliDHL-IH-SOP-002-VERS-001` | Venous Blood Sample Collection by Venipuncture | 10 | Checked 15/11/2026: no collision. (An earlier note in this file claimed HEM-SOP-008 owned IH-SOP-001. That is unverified and has been removed — see §6.) |
| `RliDHL-BIO-SOP-001-VERS-005` | Glucose Testing using Cobas c 111 | 11 | Matches the register |

### Forms

| Form | Title | Parent SOP | Status |
|---|---|---|---|
| `RliDHL-ADM-SOP-064-F01` | Training, competence assessment and authorisation record | 064 | NEW — serial reserved, SOP not yet written |
| `RliDHL-ADM-SOP-064-F01-2` | Continuation sheet, >18 staff | 064 | NEW |

---

## 6. Correction log for this file

| Date | Change |
|---|---|
| 15/11/2026 | **ADM-060 was wrongly allocated** to a training SOP. 060 is the *Non-Medical Laboratory Equipment* SOP from the November block. Moved to **064**, the next serial after 063. Its form moved with it, from `060-F01` to `064-F01`, which also released `060-F01` back to the temperature-monitoring form that owns it. |
| 15/11/2026 | ADM inventory replaced with the real 49-row list from `MASTER LIST 2025.docx`. The previous version claimed "001–052 in use, 053–059 equipment block", which was wrong on both counts. |
| 15/11/2026 | Read every series in the register, not just ADM. Found that **no IH series exists** and that `BIO-SOP-001` is registered to *Use and preventive maintenance of Humastar 80*, not to the glucose SOP. Both recorded as open decisions in `DECISIONS`, not silently renumbered. |
| 15/11/2026 | Found `ADM-FORM-014` *Signature page* and `ADM-FORM-001` *Staff competency assessment form* already in the register. Form `064-F01` overlaps both; recorded as an open decision. |
| 15/11/2026 | Removed the unverified claim that `IH-SOP-001` was taken by `HEM-SOP-008`. The IH allocation stands on the absence of `IH-SOP-001` in the register, not on that claim. |

---

## 7. When a code changes

1. Edit `REGISTER` in `scripts/build_sop.py`.
2. Rebuild — file names, cover tables, headers/footers, cross-references and form codes all
   follow.
3. Run `scripts/check_codes.py` — it fails on a serial already allocated in the register
   above, a skipped serial, a duplicate, a stale version, or an unbound form.
4. Regenerate the Master List insert (`scripts/build_masterlist_insert.py`) and update this
   file, so the three never disagree.
5. A renumber is a **controlled amendment**: raise it as a revision of each affected SOP and
   record it in that SOP's revision history. Not a typo fix.
