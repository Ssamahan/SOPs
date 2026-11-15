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
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_sop import REGISTER, SERIES_KNOWN_MAX, code, form  # noqa: E402

SOP_DIR = ROOT / "sop"

# Area codes are 2 or 3 letters: most are three (ADM BAC BIO HEM IMM MOL MYC PAR
# REC SER STC SAF) but IH is two. Requiring exactly three silently skipped every
# IH code, including the IH document's own — so the pattern allows 2-3.
CODE_RE = re.compile(
    r"(RliDHL-[A-Z]{2,3}-[A-Z]{3})[- ]?(\d{3})(?:-VERS[- ]?(\d{3}))?(?:-(F\d{2})(?:-(\d))?)?"
)

# A serial deliberately not in use, with the reason. A gap is only acceptable
# once it is declared here — that is what turns "missing" into "retired".
RETIRED: dict[tuple[str, str], dict[str, str]] = {}

# Superseded versions that are legitimately cited (e.g. in a "Supersedes" row).
# Keyed by full code, value = why.
SUPERSEDED_OK: dict[str, str] = {
    code("glucose", "004"): "listed as the version this issue supersedes",
}

failures: list[str] = []
passes: list[str] = []


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

    # ---------- 1. a new serial must be the next free number ----------
    by_series: dict[str, list[tuple[int, str, bool]]] = {}
    for key, r in REGISTER.items():
        series, serial = r["code"].rsplit("-", 1)
        by_series.setdefault(series, []).append((int(serial), key, bool(r.get("new"))))
    for series, entries in sorted(by_series.items()):
        known_max = SERIES_KNOWN_MAX.get(series, 0)
        new = sorted(e for e in entries if e[2])
        for serial, key, _ in new:
            expected = known_max + 1
            if serial != expected:
                fail(f"{series}: '{key}' takes serial {serial:03d}, but the next "
                     f"free serial is {expected:03d} (Master List runs to "
                     f"{known_max:03d}) — no skipping")
            else:
                ok(f"{series}: '{key}' takes the next free serial {serial:03d}")
                known_max = serial
        # ---------- 3. no gap in what we allocate ----------
        alloc = sorted({e[0] for e in entries if e[2]})
        if len(alloc) > 1:
            area, typ = split_series(series)
            retired = {int(s) for s in RETIRED.get((area, typ), {})}
            missing = [n for n in range(alloc[0], alloc[-1] + 1)
                       if n not in alloc and n not in retired]
            if missing:
                fail(f"{series}: gap in allocation: {[f'{n:03d}' for n in missing]} "
                     f"— declare them retired in RETIRED with a reason")
            else:
                ok(f"{series}: allocated serials contiguous")

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

    # ---------- 6. no hand-typed code in the generator ----------
    src = (ROOT / "scripts" / "build_sop.py").read_text()
    # Cut out the REGISTER dict, which is where codes are *supposed* to live,
    # then look for any literal code left in the surrounding code and text.
    start = src.index("REGISTER = {")
    end = src.index("\n}\n", start) + 3
    body = src[:start] + src[end:]
    # comments and docstrings may quote a code as documentation (the numbering
    # rule, worked examples); only a literal code in executable code is a defect
    body = re.sub(r'"""[\s\S]*?"""', "", body)
    body = "\n".join(line.split("#", 1)[0] for line in body.splitlines())
    hard = sorted(set(re.findall(r"RliDHL-[A-Z]{2,3}-[A-Z]{3}[- ]\d{3}", body)))
    if hard:
        fail(f"build_sop.py has hand-typed codes outside REGISTER: {hard}")
    else:
        ok("build_sop.py: no hand-typed document codes outside REGISTER")

    # ---------- report ----------
    print("Document-numbering check")
    print("=" * 66)
    for n in passes:
        print(f"  ok    {n}")
    if failures:
        print()
        for f in failures:
            print(f"  FAIL  {f}")
        print(f"\n{len(failures)} problem(s) found.")
        return 1
    print("\nAll checks passed: serials sequential with no skipping, no collisions, "
          "every reference resolves, forms bound, no hand-typed codes.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
