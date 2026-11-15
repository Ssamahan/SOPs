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
| `AREA` | Dept code. **The register uses exactly these:** `ADM` `BAC` `BIO` `HEM` `IMM` `MOL` `PAR` `SER` `STC` `TB`. **There is no `IH` section**, and this library no longer uses one — see §4. `MYC` was renamed `TB` and `MOL` was created on 15/11/2026 — see §3.1. |
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
| `BIO-SOP` | Biochemistry | 001–009 | none | Humastar 80 / Humalyzer 3500 analyses — **all decommissioned**, every row superseded. Only BIO-001 (glucose, Cobas c 111) is live |
| `HEM-SOP` | Haematology | 001–004 | none | Blood collection, Sysmex and Humacount analyses |
| `IMM-SOP` | Immunology | 001–004 | none | FACSCount CD4 procedures |
| `TB-SOP` | Tuberculosis (was Mycobacteriology) | 001–012 | none | GeneXpert / TB procedures. **Renamed from `MYC` on 15/11/2026; serials unchanged** |
| `MOL-SOP` | Molecular | 001–004 | none | HIV EID and HBV/HCV viral loads. **New series, 15/11/2026** — its three live documents were moved out of `MYC`, not renumbered |
| `PAR-SOP` | Parasitology | 001–010 | none | Parasitology procedures |
| `SER-SOP` | Serology | 001–007 | none | Serology procedures |
| `STC-SOP` | Stock / stores | 001–007 | none | Commodity and reagent management |

The `EXT-` codes throughout the register (`EXT-RliDHL-BIO-001-2016`, etc.) are
**manufacturer manuals**, not laboratory SOPs. They do not consume SOP serials.

### Examination SOPs written for the analysers in service

Three analysers were running with no examination SOP. `ADM-SOP-062` covers their
*maintenance* — start-up, QC, calibration, faults — and says nothing about specimen
acceptance, measurement, interpretation or reporting. ISO 15189:2022, **7.3.6 a)** requires
the laboratory to document its examination procedures, and no document did.

| New SOP | Analyser | Serial | Asset code | Section |
|---|---|---|---|---|
| `HEM-SOP-005-VERS-001` | Sysmex XN-330 | 17156 | `RULIDH/LAB/ME/0501` | Haematology |
| `HEM-SOP-006-VERS-001` | Sysmex XQ-320 | 12357 | `RULIDH/LAB/ME/0499` | Haematology |
| `IMM-SOP-005-VERS-001` | Alere Pima CD4 | `PIMA-D-005378` | `RULIDH/LAB/ME/0485` | Immunology |

Serials 005 and 006 in HEM and 005 in IMM are the lowest genuinely free numbers: HEM
001 is the venipuncture revision and 002–004 are withdrawn, IMM 001–004 are the legacy
FACSCount SOPs. Nothing is skipped.

**What these SOPs deliberately do not contain.** No measuring range, CV, reference interval
or interference limit is stated as a laboratory figure. Those are verified per instrument
under 7.3.2 and recorded in the method verification file; the reference intervals are
established and reviewed under 7.3.5 and held in the Normal Ranges list in force. The
manufacturer's own specifications and flagging rules are cited to the IFU under **7.3.6 d)**,
which permits instructions for use to be incorporated by reference. Writing a number into
an SOP that no one verified is the defect that section 13 of each document is built to
avoid.

### Equipment: what is actually in the laboratory

The decommissioning happened in **August 2026**. The equipment spreadsheet
`RDH EQUIPMENTS CODE25 JANV 2026.xlsx` is **stale** and must not be used to decide
equipment status — it still shows the Humastar 80 as *Operational*. The programme handoff's
inventory and the maintenance record are authoritative.

| Status | Equipment | Serial | Asset code |
|---|---|---|---|
| **In service** | Sysmex XN-330 | 17156 | `RULIDH/LAB/ME/0501` |
| **In service** | Sysmex XQ-320 | 12357 | `RULIDH/LAB/ME/0499` |
| **In service** | Alere Pima CD4 | `PIMA-D-005378` | `RULIDH/LAB/ME/0485` |
| **Awaiting repair** | Roche cobas c 111 (hospital's own) | 65053 | `RULIDH/LAB/ME/0486` |

**The cobas c 111 story, for the record.** The hospital's own analyser failed in **April
2026** with a Peltier element fault — *Error 5040, "Error cooling system. One of the Peltier
elements does not react anymore"* — reported to MediSell Rwanda Ltd on 18/04/2026. It was
shipped to MediSell for repair and is expected back.

While it is away the laboratory has been lent a replacement analyser by MediSell. That is a
**commercial and management circumstance, not an inventory item**: it carries no asset tag, it
appears on no register, and it is **not** test equipment of this laboratory. It is recorded
here for one reason only — the laboratory's biochemistry examination is currently unavailable,
because since **11/11/2026** the loaned unit will not enter Ready mode (a *Transfer Head Fluid
Sensor Calibration Error* and a *Level Detection Frequency Error*, persisting after every
corrective step in the service manual). **Biochemistry testing is at a standstill.**

**No SOP names the loaned unit.** A controlled procedure outlives a loan arrangement: written
around borrowed equipment it would be wrong the week the equipment went back. The glucose SOP
(BIO-001-VERS-005) describes the laboratory's **own** analyser, serial 65053, and handles the
outage through the contingency in section 17 — suspend, name the backup method on the report,
refer. The loan unit is raised with the supplier, not with document control.

| Decommissioned (Aug 2026) | Governing SOP | Action |
|---|---|---|
| Sysmex KX-21N | `HEM-SOP-004` | withdrawn, superseded in place |
| Sysmex XS-500i | — | no SOP row; appears only as an `EXT-` manual |
| Sysmex Pochi 100i | `HEM-SOP-002` | withdrawn. **No Pochi is in service** |
| Humacount 5 | `HEM-SOP-003` | withdrawn |
| Humastar 80 | `BIO-SOP-001` (VERS-002), `003`–`009` | withdrawn as a block |
| Humalyzer 3500 | `BIO-SOP-002` | withdrawn |

**Withdrawn documents are never renumbered.** Their serial is held for life; the document is
marked superseded and nothing is re-issued under that number. `WITHDRAWN` in
`scripts/build_sop.py` carries this, and `check_codes.py` fails if any SOP names a
decommissioned analyser outside a withdrawal clause — **and fails if any SOP so much as
mentions borrowed or loaned equipment.** A loan unit belongs in the equipment record and the
nonconformity file, never in a procedure.

### Three collisions, and how they were settled

Settled 15/11/2026 on the Laboratory Manager's instruction: no number is skipped, a doubled
number keeps 001 and the other document takes the number next to it, and existing controlled
forms are cited rather than withdrawn.

| # | Document | Registered to | Resolution |
|---|---|---|---|
| 1 | Glucose Testing using Cobas c 111 | `BIO-SOP-001-VERS-002` — *Use and preventive maintenance of Humastar 80* | **001 is kept** by the glucose SOP at VERS-005. The Humastar SOP is **withdrawn, not moved**: the analyser is decommissioned, so its serial stays with it. BIO-002 belongs to the Humalyzer SOP, also withdrawn. |
| 2 | Venous Blood Sample Collection | `HEM-SOP-001-VERS-001` — *Vein blood collection and preparation of hematology sample* | **`IH` is dropped.** The venous SOP becomes **`HEM-SOP-001-VERS-002`**, a revision of the registered document. No department was invented and no second SOP for drawing blood exists. |
| 3 | Signature page | `ADM-FORM-014` *Signature page*, `ADM-FORM-001` *Staff competency assessment form* | **Both are cited.** Form `064-F01` names them on its face and replaces neither, so nothing has to be withdrawn. |

The three are now recorded in `DECISIONS` in `scripts/build_sop.py` as resolved, with the
reasoning, so the numbers are not re-litigated later.


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
| `RliDHL-HEM-SOP-001-VERS-002` | Venous Blood Sample Collection by Venipuncture | 10 | Revision of the registered HEM-SOP-001. `IH` is not used — the series does not exist. |
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
| 15/11/2026 | **Corrected a serious error: decommissioned equipment treated as live.** The previous change moved *Use and preventive maintenance of Humastar 80* to `BIO-SOP-002`, and BIO-002 is the **Humalyzer 3500** SOP. Both analysers have been decommissioned; the register rows were read as live documents because the equipment status had not been checked. Both are now withdrawn in place, `WITHDRAWN` records why, and the checker fails on a decommissioned analyser named outside a withdrawal clause. |
| 15/11/2026 | **`MYC` renamed `TB`, and `MOL` created.** The register's "Mycobacteriology" section is the tuberculosis section, so `MYC 001–012` become `TB 001–012` — **same serials, same documents, same effective and review dates**. A rename is not a renumbering: no serial is freed, skipped or re-issued. HIV early infant diagnosis and the hepatitis B and C viral loads are molecular examinations and were being filed in a TB series, so they move to a new **`MOL`** series. The three live documents keep their serials — **`MOL-002`, `MOL-003`, `MOL-004`** — because a document that changes series has not changed identity. `MOL-001` is **held for life** by `RliDHL-MOL-SOP-001-VERS-003` (*HIV-1 Qualitative Testing*), which was absorbed into `MOL-002-VERS-002` and withdrawn; it is therefore **not free**, and `MOL` has no gap to explain. |
| 15/11/2026 | **`ADM-SOP-065` written: *Troubleshooting of laboratory analysers*.** None of the four examination SOPs carried the error codes of the analyser it governs — they told staff to record a code and consult a manufacturer's manual, and supplied neither. This SOP carries, per analyser, the code or message as displayed, its meaning, the cause the manufacturer gives, the operator action, the action if it persists, whether the module is withdrawn, and who is responsible. Every entry is taken from the manufacturer's own documentation (Sysmex XN-L/XN Series manuals and flagging guide; Roche cobas c 111 Operator's Manual ch. 9; Cepheid GeneXpert Dx Operator Manual ch. 9 and the GLI Module 9 training package; Abbott Pima User Guide §8–9). The examination SOPs now cite it. `check_codes.py` confirms 065 is free and that every reference resolves. |
| 15/11/2026 | **Why 065 and not 064:** 064 was taken by the SOP on the signing of SOPs, which this library issued. **Why one document and not four:** most failure modes are shared within an analyser family, and four copies of one table drift apart within a year. |
| 15/11/2026 | **Equipment status corrected from the maintenance record, not the spreadsheet.** Decommissioning was August 2026; `RDH EQUIPMENTS CODE25 JANV 2026.xlsx` is stale. Recorded that the hospital's cobas c 111 (serial 65053) has been with MediSell for repair since April 2026 after a Peltier element failure (Error 5040), and that a **loan** replacement supplied by MediSell is itself down since 11/11/2026 (Ready-mode failure). The glucose SOP was **reverted to its own-instrument form**: it names serial 65053 only and handles the outage through the section 17 contingency, because a controlled procedure must not be written around borrowed equipment. `check_codes.py` now fails on any SOP that mentions a loan or borrowed instrument. Version 005 was kept — its review date is 14/11/2028, so no re-version was triggered. |
| 15/11/2026 | **Three examination SOPs written** for the analysers in service — XN-330, XQ-320, Pima CD4 — because no document covered the examination itself (ISO 15189:2022, 7.3.6). Serials HEM-005, HEM-006, IMM-005, no skips. |
| 15/11/2026 | **Layout fix:** a bullet list could leave two or three lines alone on a page. Short lists are now kept together. This was found by measuring page fill, not by eye. |
| 15/11/2026 | **`IH` deleted.** Venous collection is `HEM-SOP-001-VERS-002`, a revision of the registered document. The invented department is gone from the register, the build and the documents. |
| 15/11/2026 | **`BIO-SOP-001` settled by keeping 001.** *Use and preventive maintenance of Humastar 80* is withdrawn, not renumbered — see the correction below. |
| 15/11/2026 | **`064-F01` now cites `ADM-FORM-014` and `ADM-FORM-001`** on the face of the signature page. Neither form is withdrawn. |
| 15/11/2026 | Removed the unverified claim that `IH-SOP-001` was taken by `HEM-SOP-008`. Superseded: the IH series is no longer used at all. |

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
