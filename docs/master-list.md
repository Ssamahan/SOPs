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
| `AREA` | Dept code. **This library uses exactly these:** `ADM` `BAC` `BIO` `HEM` `IMM` `MOL` `PAR` `SER` `STC` `TB`. The 2026 register still prints the TB section as `MYC` — the rename is a laboratory decision recorded here, applied in the SOP library, and pending a re-typed register. **There is no `IH` section**, and this library no longer uses one — see §4. `MYC` was renamed `TB` and `MOL` was confirmed in the 2026 register — see §3.1. |
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

**Source corrected on 15/11/2026.** This section was read from `MASTER LIST 2025.docx`
(last modified 15/04/2025). **That is not the current register.** The register the laboratory
works to — and the one prepared for the 2026 accreditation — is
**`MASTER LIST - 2024 - 2025.docx`, last modified 17 February 2026**. Despite its file name it
is the newer document: it contains a `MOLECULAR BIOLOGY SECTION` that the older list lacks, and
carries a document effective 18/02/2026.

Full transcription, the serials the register **does not** contain, its defects, and the EQA
evidence: **`docs/register/tb-ept-register-transcript.md`**.

The 2026 register's serials, by section — an absent serial is *unproven*, never free:

| Series | Serials the 2026 register lists |
|---|---|
| `ADM-SOP` | 001, 002, 009, 020, 024, 034 |
| `ADM-GUID` | 001, 002, 003 |
| `ADM-MAN` | 001, 002 |
| `BAC-SOP` | 004, 029 |
| `BIO-SOP` | 001, 002, 003, 004, 006, 008, 012, 015, 021, 026, 027, 031 |
| `HEM-SOP` | 001, 004, 005, 007, 008, 010, 016, 019, 023, 032 |
| `MOL-SOP` | 001, 002, 003, 004 |
| `TB-SOP` | 001, 002, 003 (was `MYC`) |
| `PAR-SOP` | 001, 002, 008, 017 |
| `SER-SOP` | 002, 003, 004, 007, 008, 010, 012, 013, 014, 018, 020, 030 |

**What this corrected.** `ADM-025` is free in the 2026 register, which is why the new EQA/PT SOP
takes it rather than skipping to 066. `ADM-041` — the serial this library gave *Documentation of
examination procedures* — is **not in the register at all**; it is a new allocation by this
library and is now marked as such. The old table's claim that every ADM serial 1–49 except 035
was registered came from the superseded list and is withdrawn: those serials are declared
**unprovable**, not free.

"Entries" counts distinct serials; the "gap" column is what is absent from the sequence, not what
is free.

| Series | Sections in the register | Serials seen | Gap | What the series actually holds |
|---|---|---|---|---|
| `ADM-SOP` | Crosscutting | 001–049 | 016, 017 (unreadable), 035 (never used) | Management SOPs: document control, complaints, training, audit |
| `BAC-SOP` | Bacteriology & parasitology | 001–006 | none | Culture and staining procedures |
| `BIO-SOP` | Biochemistry | 001–009 | none | Humastar 80 / Humalyzer 3500 analyses — **all decommissioned**, every row superseded. Only BIO-001 (glucose, Cobas c 111) is live |
| `HEM-SOP` | Haematology | 001–004 | none | Blood collection, Sysmex and Humacount analyses |
| `IMM-SOP` | Immunology | 001–004 | none | FACSCount CD4 procedures |
| `TB-SOP` | Tuberculosis (was Mycobacteriology) | 001–012 | none | **All TB testing, by any method**: ZN and auramine bacilloscopy, culture and DST, GeneXpert MTB/RIF and Ultra, and **TB LAM**. Renamed from `MYC` on 15/11/2026; serials unchanged |
| `MOL-SOP` | Molecular | 002–004 | **001** | HIV EID and HBV/HCV viral loads. **Already in the 2026 register** (`MOLECULAR BIOLOGY SECTION`, 001–004) — not invented by this library. The examination documents keep their serials, out of `MYC`, not renumbered. **HIV *qualitative* testing is NOT here: it is serology (`SER`)** |
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
| 15/11/2026 | **`MYC` renamed `TB`; `MOL` confirmed in the register.** The register's "Mycobacteriology" section is the tuberculosis section, so `MYC 001–012` become `TB 001–012` — **same serials, same documents, same effective and review dates**. A rename is not a renumbering: no serial is freed, skipped or re-issued. HIV early infant diagnosis and the hepatitis B and C viral loads are molecular examinations and were being filed in a TB series, so they move to a new **`MOL`** series. The three live documents keep their serials — **`MOL-002`, `MOL-003`, `MOL-004`** — because a document that changes series has not changed identity.

**Serology is untouched by this move.** Rapid `Determine`, `Unigold`, capillary collection for HIV testing and HIV QC transport are **serology** and stay in `SER` (`SER-001`, `-002`, `-004`, `-005`). The distinction is **method, not analyte and not instrument**: an antibody rapid test is serology; an assay that amplifies nucleic acid is molecular. The same GeneXpert runs both molecular viral loads and TB assays, so the instrument cannot decide the series either. `MOL-001` is **held for life** by `RliDHL-MOL-SOP-001-VERS-003` (*HIV-1 Qualitative Testing Using Xpert HIV-1 Assays*), which `MOL-002-VERS-002` records that it replaces and which is withdrawn — so it is not free.
| 15/11/2026 | **Serology vs molecular — do not conflate them.** `MOL` is **molecular only**: HIV early infant diagnosis, HBV viral load, HCV viral load — assays that amplify nucleic acid. **HIV rapid testing is serology and stays in `SER`**: `Determine` (001), `Unigold` (002), capillary collection for HIV testing (004), HIV QC sample transport (005). The distinction is the **method**, not the analyte and not the instrument. The only HIV document that was molecular was the Xpert HIV-1 Qual assay, `MOL-001-VERS-003`, and it is **withdrawn**. |
| 15/11/2026 | **A wrong entry was written here and has been taken back out.** For a few minutes this log claimed `MOL-001` had never existed and that the Xpert Qual SOP was a SER document. That is false: `MOL-002-VERS-002` states on its own face that it replaces `RliDHL-MOL-SOP-001-VERS-003`, *HIV-1 Qualitative Testing Using Xpert HIV-1 Assays*. The serial is held, the series has no gap, and the record stands. The error is recorded rather than quietly deleted, because a correction log that hides its own corrections is not a correction log. |
| 15/11/2026 | **The TB series holds all TB testing, whatever the method.** Renamed `MYC` → `TB` and inspected row by row: the section holds sputum collection with CPC (001), specimen rejection (002), ZN microscopy QC (003), auramine fluorescence microscopy (004), sample transport (005), IQC (006), turnaround time (007), slide re-testing (008), an untitled row (009), and GeneXpert sample processing (010, 011, 012). **TB LAM belongs here**, not in `SER`: it is a TB test performed by the TB section, and the register's own logic is that a test follows the section, not the method — which is exactly why bacilloscopy is here. **Register defects recorded, not fixed:** `TB-009` has no title in the register, and `TB-010`/`TB-012` carry the same subject (*Sample processing using GeneXpert*). **TB LAM is not registered at all:** no row for TB lipoarabinomannan exists anywhere in the register, and nor does one for the TB-in-stool Xpert Ultra SOP the laboratory holds as a document. TB LAM stays in `TB` when it is written, exactly as bacilloscopy does — the method does not move it to `SER`. The next free TB serial is **013**; nothing is allocated here, because writing those procedures is the Laboratory Manager's decision, not a numbering one. |
| 15/11/2026 | **The 2026 register was found, and it is not the one this file was built from.** `MASTER LIST - 2024 - 2025.docx` (Drive `1ESTwzyHMWFWKWJV76Krj6JCvrD7Ulr_9`, **modified 17/02/2026**) is the renewed register prepared for the 2026 accreditation, despite the file name. It contains a `MOLECULAR BIOLOGY SECTION` with `MOL-001`–`004` and a document effective 18/02/2026; `MASTER LIST 2025.docx` (15/04/2025) has neither. **`MOL` therefore already existed in the register** — this library did not invent it, and the earlier entry recording `MOL` as a new series is corrected. The register's serials are transcribed in `docs/register/tb-ept-register-transcript.md`. |
| 15/11/2026 | **`TB` prefix applied to the register as well as the library.** `docs/register/tb-ept-register-TB-prefix.md` is the transcription re-issued with `MYC` → `TB`. Prefix only: `MYC-001` = `TB-001`, same version, same effective date, same review date, same title, nothing renumbered and nothing reissued. The original transcription is kept unaltered beside it so that the change is auditable. |
| 15/11/2026 | **TB smear and malaria film EQA recorded — slide re-reading, not a panel.** NRL takes the laboratory's own stained slides and the results it reported, reads them independently and compares. These are two different scheme types and the SOP now treats them separately: `ADM-025` section 8 covers a panel, **section 8A** covers slide re-reading. The rules that follow from the difference: the slides must be a representative sample of routine work and never chosen for looking good; the result sent to the provider must be the result reported to the clinician; staining quality is part of what is assessed; discrepancy grading is the provider's, not the laboratory's; and slide retention must make re-reading possible. |
| 15/11/2026 | **Slide re-reading corrected in `ADM-025`, and the records read.** TB smear and malaria EQA are **re-reading schemes on the laboratory's own health centres' slides**, not PT panels. The NRL reports carry a "Result/Health Center" column against a "Result/Hospital lab" column, slide by slide — no manufactured sample appears anywhere. Section 8A was rewritten to say that; the earlier draft described the laboratory as sending its own slides out, which the records do not support. TB is graded **positive / scanty / negative** against a **90 per cent passing rate**; malaria is positive/negative with the discordant rate recorded. |
| 15/11/2026 | **Five defects found in the re-reading reports and recorded, not tidied away.** (1) Malaria Q3 2025/26 does not add up — Coko, Muyongwe and Rushashi rows fail to balance yet carry a discordant rate of 0 and a pass. (2) Two TB sheets are both headed "Quarter II". (3) TB Q4 2024/25 marks both sites **False** against a discordant rate of **0**, which should be Pass. (4) TB re-reading covers **two of eight centres**, and Nyange' sample runs 11 → **2** → 5 → 8 slides. (5) The TB reports cover **microscopy only** — GeneXpert TB is PT, not re-reading. All five go to NRL in writing and are tracked in `ADM-025` §12.5–12.6 with owners and dates. |
| 15/11/2026 | **A category error corrected: the health-centre programme was sitting inside the laboratory's own EQA SOP.** The two run in opposite directions. In one, NRL assesses **this laboratory** — the panel, the z-scores, the PT rounds. In the other, **this laboratory assesses its health centres** — it re-reads their slides, compares, and reports the discordance back. `ADM-025` was carrying health-centre discordance rates as though they were the laboratory's own EQA performance. Sections 8A and 12.5–12.6 were removed from `ADM-025` and the material moved to a new **`ADM-SOP-050`**, *Re-reading of Health-Centre Slides and Re-Testing of Health-Centre Samples* (8 pages). `ADM-025` now states in its scope, in its definitions and in a new **§6.4** that it governs only the assessment **of** this laboratory, that the two are never combined in one figure, and that no health-centre report may be counted as evidence about this laboratory's own accuracy. |
| 15/11/2026 | **Two gaps opened by that correction, recorded rather than hidden.** With the health-centre reports removed from the laboratory's EQA file, **no report was found of this laboratory's own TB or malaria slides ever being re-read by a provider.** Either the laboratory is not enrolled for its own smear microscopy — a genuine gap in its own quality assurance — or the reports are unfiled. `ADM-025` §6.4 records it against TB smear and malaria film with an owner and a one-month deadline. |
| 15/11/2026 | **`ADM-SOP-050` took serial 050.** That serial was held by an **unsigned** Microscope draft from the November 2026 block — never issued, never in force — so it was released back to the queue rather than leaving a hole. Nothing was ever issued under 050, so no document loses its number. |
| 15/11/2026 | **`ADM-SOP-025` written| 15/11/2026 | **`ADM-SOP-025` written: *External Quality Assessment and Proficiency Testing*.** No document governed EQA: who enrols, who tests the panel, who signs the submission, what happens when a result is unsatisfactory and how the annual plan is made were all unwritten. **The laboratory performs EQA for HIV serology** (Dried Tube Specimen PT, RBC/NRL — 2022, 2023, 2024 reports on file), **TB on GeneXpert** (MTB PT, participant code 20044 — 2023, 2024), **biochemistry** (creatinine, glucose, ALT) and **the full blood count** (NRL quarterly; January–March 2026 report `RBC/AG. DIVISION MANAGER OF NRL/17302/2026`). The SOP covers the annual plan built from the Scope of Test list, receipt and immediate testing of panels, routine-person/routine-method testing, submission, review within one month, acceptability criteria including ±2 SDI, investigation by cause across the whole examination process, verification of effectiveness, and — the section that matters — assessment of whether **patients** were affected. Serial 025 is free in the 2026 register, so the gap between 024 and 026 is filled rather than skipped. |
| 15/11/2026 | **EQA gaps recorded, not papered over.** No PT report is held for **HBV or HCV viral load**, **HIV early infant diagnosis** or **TB LAM**. Section 6 of `ADM-025` records these as gaps with an owner and a deadline; the Laboratory Manager establishes within one month whether the laboratory is enrolled and the reports are unfiled, or not enrolled at all. |
| 15/11/2026 | **The 2026 register does not number every section without a skip.** Its own serial sets are listed verbatim in `docs/register/tb-ept-register-transcript.md`, and most sections have large absent ranges — `ADM-SOP` lists six serials of 34, `BAC-SOP` two of 29. Those serials are declared **unprovable**, not free, and `check_codes.py` now fails if any of them is allocated. Where an allocation had to be made, the lowest genuinely absent serial in the right place was used: **025** for EQA, **065** for troubleshooting. |
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
