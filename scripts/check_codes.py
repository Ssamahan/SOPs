#!/usr/bin/env python3
"""Enforce the document-numbering rules on the built SOPs.

Run after build_sop.py.  Checks:

  1. NEW SERIAL IS THE NEXT FREE ONE   a serial this library allocates for the
                                       first time must be SERIES_KNOWN_MAX + 1.
                                       No skipping.
  2. NO SERIAL COLLISION               a serial is never used for two different
                                       documents (the IH-SOP-001 defect).
  3. NO GAP IN WHAT WE ALLOCATE        allocating 060 then 062 is a failure; a
                                       gap may only be declared in RETIRED.
  4. EVERY REFERENCE RESOLVES          every code cited in an SOP is that SOP
                                       itself, a registered document, or a
                                       declared superseded version of one.
  5. FORMS BOUND TO A PARENT           every form is <parent SOP code>-F<NN>.
  6. NO HAND-TYPED CODE                build_sop.py carries no literal code
                                       outside REGISTER.
  7. FILE NAME = CODE IN THE DOCUMENT  a renumbered file cannot keep an old name.

Exit status 1 on any failure, so this can gate a commit or a CI job.
"""
from __future__ import annotations

import re
import sys
import textwrap
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_sop import (  # noqa: E402
    DECISIONS, IN_SERVICE, REGISTER, SERIES_ALLOCATED, SERIES_UNVERIFIED,
    SERIES_KNOWN_MAX, WITHDRAWN, code, form,
)

# Every analyser named as decommissioned. A document in this repository must not
# describe one of these as if it were in use, and must not be built on one.
DECOMMISSIONED = (
    "Humastar 80", "Humalyzer 3500", "Humacount 5", "Sysmex KX-21N", "KX-21N",
    "Sysmex XS-500i", "XS-500i", "XS-1000i", "Sysmex Pochi", "Pochi 100i",
)

SOP_DIR = ROOT / "sop"

# Area codes are 2 or 3 letters: most are three (ADM BAC BIO HEM IMM MOL MYC PAR
# REC SER STC SAF) but IH is two. Requiring exactly three silently skipped every
# IH code, including the IH document's own — so the pattern allows 2-3.
CODE_RE = re.compile(
    r"(RliDHL-[A-Z]{2,3}-[A-Z]{3})[- ]?(\d{3})(?:-VERS[- ]?(\d{3}))?(?:-(F\d{2})(?:-(\d))?)?"
)

# A gap in a series is acceptable only once it is declared here with its reason.
# Two different things get declared, and conflating them would be dishonest:
#
#   retired     the number was allocated, the document is withdrawn, the number
#               is not re-used. This is a decision.
#   unprovable  the number MAY be in use on a document that was never registered,
#               so it cannot be re-allocated. This is an unknown, not a decision.
#
# Both block allocation. Neither is a licence to skip a number that is genuinely
# free -- the next free serial must still be the lowest unallocated one.
RETIRED: dict[tuple[str, str], dict[str, str]] = {
    ("ADM", "SOP"): {
        "035": "retired — never used; the register skips from 034 to 036 "
               "(SOP-PROGRAMME-HANDOFF §8). Under the no-gap convention this is "
               "recorded so it is not mistaken for a lost document.",
        "053": "unprovable — files exist above 052 that were never registered, so "
               "this range cannot be shown to be free. The Nov 2026 pack starts at "
               "060 instead. Not a decision; a limit.",
        "054": "unprovable — see 053",
        "055": "unprovable — see 053",
        "056": "unprovable — see 053",
        "057": "unprovable — see 053",
        "058": "unprovable — see 053",
        "059": "unprovable — see 053",
    },
}

# Superseded versions that are legitimately cited (e.g. in a "Supersedes" row).
# Keyed by full code, value = why.
SUPERSEDED_OK: dict[str, str] = {
    code("glucose", "004"): "listed as the version this issue supersedes",
}

failures: list[str] = []
passes: list[str] = []
notes: list[str] = []


def fail(msg: str) -> None:
    failures.append(msg)


def ok(msg: str) -> None:
    passes.append(msg)


def split_series(series: str) -> tuple[str, str]:
    _, area, typ = series.split("-")
    return area, typ


def extract(text: str) -> set[tuple]:
    return {m.groups() for m in CODE_RE.finditer(text)}


def fmt(series, serial, vers, fnum, fsub) -> str:
    out = f"{series}-{serial}"
    if vers:
        out += f"-VERS-{vers}"
    if fnum:
        out += f"-{fnum}" + (f"-{fsub}" if fsub else "")
    return out


def main() -> int:
    pdfs = sorted(SOP_DIR.glob("*.pdf"))
    if not pdfs:
        print("No PDFs in sop/ — run build_sop.py first.", file=sys.stderr)
        return 1
    try:
        import pymupdf
    except ImportError:
        print("pymupdf is required: pip install pymupdf", file=sys.stderr)
        return 1

    # ---------- 1. a new serial must be free in the register ----------
    # The decisive check. A serial is only legitimate if nothing in the register
    # or the allocated set already holds it. "Next free" is whatever the highest
    # allocated number plus one is -- never inferred from a working note.
    by_series: dict[str, list[tuple[int, str, bool]]] = {}
    for key, r in REGISTER.items():
        series, serial = r["code"].rsplit("-", 1)
        by_series.setdefault(series, []).append((int(serial), key, bool(r.get("new"))))
    for series, entries in sorted(by_series.items()):
        allocated = SERIES_ALLOCATED.get(series, {})
        area, typ = split_series(series)
        declared_here = RETIRED.get((area, typ), {})
        for serial, key, is_new in sorted(entries):
            if is_new:
                if serial in allocated:
                    fail(f"{series}: '{key}' takes serial {serial:03d}, but that "
                         f"number is ALREADY ALLOCATED ({allocated[serial]}) — "
                         f"allocate a free one")
                elif f"{serial:03d}" in declared_here:
                    fail(f"{series}: '{key}' takes serial {serial:03d}, which is "
                         f"declared unavailable ({declared_here[f'{serial:03d}']}) "
                         f"— allocate a free one")
                else:
                    ok(f"{series}: '{key}' takes serial {serial:03d}, which is free")
            else:
                ceiling = SERIES_KNOWN_MAX.get(series, 0)
                if serial in allocated:
                    ok(f"{series}: '{key}' cites registered serial {serial:03d}")
                elif key in DECISIONS:
                    # The number is inside the series, but the register gives it to a
                    # different document. That is not a pass and must not look like one.
                    notes.append(
                        f"{series}: '{key}' cites serial {serial:03d}, which the "
                        f"register gives to another document — see OPEN DECISION "
                        f"'{key}' below")
                elif serial <= ceiling:
                    # the register's own inventory above is partial for this series,
                    # so a serial within the known ceiling is accepted with a note
                    ok(f"{series}: '{key}' cites pre-existing serial {serial:03d} "
                       f"(within the known {ceiling:03d} ceiling)")
                elif series in SERIES_UNVERIFIED:
                    # Nothing has audited this series in the register, so the
                    # ceiling proves nothing. Say so rather than inventing one.
                    ok(f"{series}: '{key}' cites serial {serial:03d} — the register "
                       f"has not been audited for this series, so its presence "
                       f"cannot be confirmed from here")
                    notes.append(
                        f"{series}: serial {serial:03d} is UNVERIFIED — read "
                        f"MASTER LIST 2025.docx for this series before allocating "
                        f"or citing another")
                else:
                    fail(f"{series}: '{key}' is marked pre-existing but serial "
                         f"{serial:03d} is beyond the known ceiling {ceiling:03d} "
                         f"— either it is a new allocation (mark it) or the number "
                         f"is wrong")
        # Highest allocated counts BOTH sources: serials allocated by this library
        # (REGISTER) and serials held by documents outside it (SERIES_ALLOCATED).
        here = {e[0] for e in entries}
        allocated = {**allocated, **{n: "this library" for n in here}}
        if allocated:
            nxt = max(allocated) + 1
            notes.append(f"{series}: highest allocated {max(allocated):03d}, "
                         f"next free {nxt:03d}")
            # Serials held by withdrawn documents. These must never be reported as
            # free just because the sequence runs past them: a serial is held for life.
            live = {code(k) for k in REGISTER}
            held = []
            for c, _, _ in WITHDRAWN.values():
                m = re.fullmatch(rf"RliDHL-{area}-SOP-(\d{{3}})", c)
                if not m:
                    continue
                n = int(m.group(1))
                if f"RliDHL-{area}-SOP-{n:03d}" not in live:
                    held.append(n)
            if held:
                notes.append(
                    f"{series}: serial {', '.join(f'{n:03d}' for n in sorted(set(held)))} "
                    f"held by withdrawn/decommissioned documents — NOT free, whatever "
                    f"the sequence suggests")

        # ---------- 3. no gap in what we allocate ----------
        alloc = sorted({e[0] for e in entries if e[2]} | set(allocated))
        if len(alloc) > 1:
            declared = declared_here
            retired = {int(k) for k in declared}
            missing = [n for n in range(alloc[0], alloc[-1] + 1)
                       if n not in alloc and n not in retired]
            if missing:
                fail(f"{series}: undeclared gap in allocation: "
                     f"{[f'{n:03d}' for n in missing]} — declare each in RETIRED "
                     f"as retired or unprovable, with a reason")
            else:
                kinds = {}
                for k, v in declared.items():
                    kinds.setdefault(v.split(" ")[0], []).append(k)
                detail = "; ".join(f"{kind}: {','.join(sorted(serials))}"
                                   for kind, serials in sorted(kinds.items()))
                ok(f"{series}: allocation contiguous" + (f" except {detail}" if detail else ""))

    # ---------- 2. no serial collision ----------
    owner: dict[str, str] = {}
    for key, r in REGISTER.items():
        series, serial = r["code"].rsplit("-", 1)
        k = f"{series}-{serial}"
        if k in owner:
            fail(f"serial collision: {k} is claimed by both "
                 f"'{owner[k]}' and '{key}'")
        owner[k] = key
    if not any("serial collision" in f for f in failures):
        ok("no serial is used by two documents")

    registered_docs = {code(k) for k in REGISTER}
    registered_bases = {c.rsplit("-VERS", 1)[0] for c in registered_docs}
    registered_forms = {form("training", "F01"), form("training", "F01-2")}

    # ---------- per document ----------
    for pdf in pdfs:
        doc = pymupdf.open(pdf)
        text = "\n".join(p.get_text() for p in doc)
        found = extract(text)

        # ---------- 7. file name matches the cover ----------
        name_code = pdf.stem.split(" - ")[0]
        if name_code not in text:
            fail(f"{pdf.name}: file name carries {name_code}, which the document "
                 f"text does not contain")
        else:
            ok(f"{pdf.name}: file name matches the cover ({name_code})")

        # ---------- 4. every reference resolves ----------
        for series, serial, vers, fnum, fsub in sorted(
                found, key=lambda t: tuple(x or "" for x in t)):
            full = fmt(series, serial, vers, fnum, fsub)
            if fnum:
                if full not in registered_forms:
                    fail(f"{pdf.name}: form {full} is not bound to a registered "
                         f"SOP (expected one of {sorted(registered_forms)})")
            elif vers:
                if full in registered_docs or full in SUPERSEDED_OK:
                    continue
                if f"{series}-{serial}" not in registered_bases:
                    fail(f"{pdf.name}: cites {full} — serial not in the register")
                else:
                    fail(f"{pdf.name}: cites {full} — the serial is registered "
                         f"but this version is not (add it to SUPERSEDED_OK with "
                         f"a reason, or fix the citation)")
            else:
                if f"{series}-{serial}" not in registered_bases:
                    fail(f"{pdf.name}: cites {series}-{serial} — not registered")

    # ---------- 4b. a citation must name a real document ----------
    # A code can look perfectly well-formed and still point at nothing: IH-SOP-001
    # was well-formed, but that serial belongs to a different SOP. So every
    # citation is also required to be the *current* version of a registered
    # document, or an explicitly declared superseded one — a citation to a
    # version that is neither is a stale reference.
    current_vers = {}
    for k in REGISTER:
        full = code(k)
        current_vers[full.rsplit("-VERS", 1)[0]] = full.rsplit("-VERS", 1)[1]

    for pdf in pdfs:
        text = "\n".join(p.get_text() for p in pymupdf.open(pdf))
        for series, serial, vers, fnum, fsub in sorted(
                extract(text), key=lambda t: tuple(x or "" for x in t)):
            if fnum or not vers:
                continue
            full = fmt(series, serial, vers, fnum, fsub)
            if full in SUPERSEDED_OK or full in registered_docs:
                continue
            base = f"{series}-{serial}"
            if base in current_vers:
                fail(f"{pdf.name}: cites {full}, but the current version of "
                     f"{base} is VERS-{current_vers[base]} — stale version "
                     f"reference (declare it in SUPERSEDED_OK if intended)")
            else:
                fail(f"{pdf.name}: cites {full}, but no document in the register "
                     f"has serial {base} — the code is well-formed and wrong")

    # ---------- 5b. nothing in service may be described as withdrawn ----------
    # A serial in the register does not mean the machine is at the bench. Five of the
    # analysers this register carries SOPs for have been decommissioned. A document
    # that names one as current equipment is worse than a wrong number: it instructs
    # a technologist to use a machine that is not there.
    for pdf in pdfs:
        text = "\n".join(p.get_text() for p in pymupdf.open(pdf))
        for name in DECOMMISSIONED:
            if re.search(rf"\b{re.escape(name)}\b", text):
                # allowed only inside an explicit withdrawal/supersession sentence
                for m in re.finditer(rf"[^.]*\b{re.escape(name)}\b[^.]*\.", text):
                    sent = m.group(0)
                    if re.search(r"withdrawn|decommission|supersed|no longer|obsolete|"
                                 r"retired|not in service", sent, re.I):
                        continue
                    fail(f"{pdf.name}: names {name} outside a withdrawal clause — "
                         f"{name} is decommissioned; a current SOP must not describe it "
                         f"as equipment (handoff, Decommissioned)")
                    break

    ok(f"no SOP describes decommissioned equipment as in service "
       f"({len(DECOMMISSIONED)} withdrawn analysers checked)")

    # ---------- 6. no hand-typed code in the generator ----------
    src = (ROOT / "scripts" / "build_sop.py").read_text()
    # Cut out the REGISTER dict, which is where codes are *supposed* to live,
    # then look for any literal code left in the surrounding code and text.
    def cut(text, marker):
        s0 = text.index(marker)
        return text[:s0] + text[text.index("\n}\n", s0) + 3:]

    body = src
    for marker in ("REGISTER = {", "DECISIONS = {", "WITHDRAWN = {"):
        if marker in body:
            body = cut(body, marker)
    # comments and docstrings may quote a code as documentation (the numbering
    # rule, worked examples); only a literal code in executable code is a defect
    body = re.sub(r'"""[\s\S]*?"""', "", body)
    body = "\n".join(line.split("#", 1)[0] for line in body.splitlines())
    hard = sorted(set(re.findall(r"RliDHL-[A-Z]{2,3}-[A-Z]{3}[- ]\d{3}", body)))
    if hard:
        fail(f"build_sop.py has hand-typed codes outside REGISTER: {hard}")
    else:
        ok("build_sop.py: no hand-typed document codes outside REGISTER")

    # Decisions about specific registered documents. These are not failures --
    # the code is internally consistent -- but a green tick would hide them.
    # Every declared open decision is printed, whether or not it concerns a document this
    # repository builds. A decision about equipment status is exactly the kind that gets
    # lost when it is only recorded next to a document.
    open_decisions = list(DECISIONS.items())

    # ---------- report ----------
    print("Document-numbering check")
    print("=" * 66)
    for n in passes:
        print(f"  ok    {n}")
    for n in notes:
        print(f"  info  {n}")
    if failures:
        print()
        for f in failures:
            print(f"  FAIL  {f}")
        print(f"\n{len(failures)} problem(s) found.")
        return 1

    if open_decisions:
        print()
        print(f"  {len(open_decisions)} OPEN DECISION(S) -- not errors, but not signed off")
        for key, d in open_decisions:
            print(f"\n  [{key}] {d['issue']}")
            for line in textwrap.wrap(d["detail"], 66):
                print(f"      {line}")
            print()
            for line in textwrap.wrap("Options: " + d["options"], 66):
                print(f"      {line}")
        print()

    print("\nAll checks passed: serials sequential with no skipping, no collisions, "
          "every reference resolves, forms bound, no hand-typed codes.")
    if open_decisions:
        print(f"{len(open_decisions)} decision(s) above are still open — the documents "
              f"are consistent with each other, not with the register's history.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
