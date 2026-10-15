# -*- coding: utf-8 -*-
"""Deduplication + source-to-compilation reconciliation.

Answers the question "did every comment in every delegation's file actually
make it into the compilation, and if not, exactly why?" – per file and per
country – and produces an exportable diagnostic report.

Reconciliation arithmetic, per file:

    rows in source table
  − rows skipped by the extractor (blank / placeholder / unreadable)
  − duplicate rows removed
  = comments expected in the compilation

and that figure is then checked against what is actually in the compiled
output. Any mismatch is reported as an inconsistency with the detail needed
to chase it down.
"""
import re
from collections import Counter, OrderedDict
from typing import Dict, List, Tuple


def _norm(s: str) -> str:
    """Whitespace/case-insensitive text key."""
    return re.sub(r"\s+", " ", (s or "")).strip().lower()


def _key(c) -> Tuple[str, str, str]:
    return (
        _norm(c.jurisdiction),
        _norm(c.reference.plain() if c.reference else ""),
        _norm(c.comment.plain() if c.comment else ""),
    )


def dedupe_comments(comments: List) -> Tuple[List, List[Dict]]:
    """Remove duplicate comments.

    A duplicate is the *same delegation* saying the *same thing* about the
    *same reference* – i.e. an accidental double entry, or the same file
    uploaded twice. The first occurrence is kept.

    Identical text from two *different* delegations is NOT removed: two
    countries independently making the same point is normal and both must
    appear in the compilation. Those are reported separately as
    'cross_country' matches so you can eyeball them.

    Returns (kept, removed) where removed items carry the detail needed for
    the diagnostic report.
    """
    seen = {}
    kept, removed = [], []

    for c in comments:
        k = _key(c)
        if not k[2]:                       # empty comment text – never dedupe
            kept.append(c)
            continue
        if k in seen:
            first = seen[k]
            removed.append({
                "jurisdiction": c.jurisdiction,
                "reference": (c.reference.plain() if c.reference else "")[:80],
                "comment": (c.comment.plain() if c.comment else "")[:160],
                "removed_from": c.source_file,
                "removed_row": c.order_in_file,
                "kept_from": first.source_file,
                "kept_row": first.order_in_file,
            })
            continue
        seen[k] = c
        kept.append(c)

    return kept, removed


def cross_country_matches(comments: List) -> List[Dict]:
    """Identical comment text submitted by more than one delegation.
    Kept in the compilation, but worth a look."""
    by_text = {}
    for c in comments:
        t = _norm(c.comment.plain() if c.comment else "")
        if not t or len(t) < 40:          # ignore trivially short text
            continue
        by_text.setdefault(t, []).append(c)

    out = []
    for t, group in by_text.items():
        jurs = sorted({g.jurisdiction for g in group})
        if len(jurs) > 1:
            out.append({
                "delegations": jurs,
                "reference": (group[0].reference.plain()
                              if group[0].reference else "")[:60],
                "comment": (group[0].comment.plain()
                            if group[0].comment else "")[:160],
                "count": len(group),
            })
    return out


def build_report(parsed_files: List[Dict],
                 assignments: Dict[str, str],
                 compiled: List,
                 removed: List[Dict],
                 media: Dict[str, Dict] = None) -> Dict:
    """Build the full reconciliation report.

    parsed_files – the per-file dicts from app._parse_one (carry
                   'stats' with total_rows / extracted / skipped)
    assignments  – {filename: country actually used in the compilation}
    compiled     – the final comment list that was written out
    removed      – duplicate records from dedupe_comments()
    media        – {filename: media_probe result}
    """
    media = media or {}
    removed_by_file = Counter(r["removed_from"] for r in removed)
    compiled_by_file = Counter(c.source_file for c in compiled)
    compiled_by_country = Counter(c.jurisdiction for c in compiled)

    file_rows = []
    for pf in parsed_files:
        fn = pf["filename"]
        st = pf.get("stats") or {}
        total = st.get("total_rows", len(pf.get("comments") or []))
        extracted = st.get("extracted", len(pf.get("comments") or []))
        skipped = st.get("skipped", [])
        dups = removed_by_file.get(fn, 0)
        expected = extracted - dups
        actual = compiled_by_file.get(fn, 0)
        country = assignments.get(fn) or pf.get("jurisdiction_raw") \
            or pf.get("filename_guess") or "Unknown"

        notes = []
        if pf.get("error"):
            notes.append(f"FILE ERROR: {pf['error']}")
        if skipped:
            reasons = Counter(s["reason"].split("(")[0].strip()
                              for s in skipped)
            notes.append("; ".join(
                f"{n} row(s) skipped – {r}" for r, n in reasons.items()))
        if dups:
            notes.append(f"{dups} duplicate row(s) removed")
        if pf.get("kind") == "pdf":
            notes.append("ingested from PDF – formatting not preserved")
        m = media.get(fn) or {}
        if m.get("count"):
            n_rows = len(m.get("in_comment_rows") or [])
            notes.append(
                f"{m['count']} image/object(s) present"
                + (f", {n_rows} inside comment cells (NOT carried over)"
                   if n_rows else " (outside comment cells)"))
        if not pf.get("jurisdiction_raw"):
            notes.append(f"country not declared in file – assigned '{country}'")

        match = (expected == actual)
        if not match:
            notes.append(
                f"MISMATCH: expected {expected} in compilation but found "
                f"{actual} (difference {actual - expected:+d})")

        file_rows.append(OrderedDict([
            ("File name", fn),
            ("Country assigned", country),
            ("Source type", (pf.get("kind") or "docx").upper()),
            ("Rows in source table", total),
            ("Rows skipped", len(skipped)),
            ("Duplicates removed", dups),
            ("Expected in compilation", expected),
            ("Actually in compilation", actual),
            ("Match", "YES" if match else "NO"),
            ("Notes", " | ".join(notes)),
            ("_skipped_detail", skipped),
        ]))

    # ---- per-country reconciliation --------------------------------------
    country_expected = Counter()
    country_files = {}
    for fr in file_rows:
        c = fr["Country assigned"]
        country_expected[c] += fr["Expected in compilation"]
        country_files.setdefault(c, []).append(fr["File name"])

    country_rows = []
    for country in sorted(set(country_expected) | set(compiled_by_country)):
        exp = country_expected.get(country, 0)
        act = compiled_by_country.get(country, 0)
        match = (exp == act)
        country_rows.append(OrderedDict([
            ("Country", country),
            ("Files", ", ".join(country_files.get(country, []))),
            ("No. of files", len(country_files.get(country, []))),
            ("Expected comments", exp),
            ("Comments in compilation", act),
            ("Match", "YES" if match else "NO"),
            ("Difference", act - exp),
        ]))

    totals = {
        "files": len(file_rows),
        "rows_in_source": sum(f["Rows in source table"] for f in file_rows),
        "rows_skipped": sum(f["Rows skipped"] for f in file_rows),
        "duplicates_removed": len(removed),
        "expected": sum(f["Expected in compilation"] for f in file_rows),
        "compiled": len(compiled),
        "countries": len(country_rows),
        "file_mismatches": sum(1 for f in file_rows if f["Match"] == "NO"),
        "country_mismatches": sum(1 for c in country_rows if c["Match"] == "NO"),
    }
    totals["all_reconciled"] = (
        totals["file_mismatches"] == 0
        and totals["country_mismatches"] == 0
        and totals["expected"] == totals["compiled"]
    )

    manual = build_manual_checks(parsed_files, assignments, compiled, media)
    totals["manual_checks"] = len(manual)

    return {
        "files": file_rows,
        "countries": country_rows,
        "duplicates": removed,
        "cross_country": cross_country_matches(compiled),
        "manual_checks": manual,
        "totals": totals,
    }


def build_manual_checks(parsed_files, assignments, compiled, media=None):
    """The list a human must eyeball before the compilation is circulated.

    Everything here was compiled – nothing was dropped – but each item is
    something the tool cannot decide on its own: a delegation whose template
    was altered, a comment the tool could not file under a section, a PDF
    whose formatting cannot be verified automatically, an image that cannot
    be carried across, a country that had to be guessed.
    """
    media = media or {}
    out = []

    def add(priority, category, file, country, detail, action):
        out.append({
            "Priority": priority,
            "Check": category,
            "File": file,
            "Country": country,
            "What was found": detail,
            "What to do": action,
        })

    for pf in parsed_files:
        fn = pf["filename"]
        country = assignments.get(fn) or pf.get("jurisdiction_raw") or "?"
        st = pf.get("stats") or {}

        if pf.get("error"):
            add("HIGH", "Unreadable file", fn, country,
                f"The file could not be parsed: {pf['error']}",
                "Open the file by hand and re-save it as a .docx from the "
                "official template, then re-run.")

        if st.get("salvaged_tables"):
            add("HIGH", "Template altered", fn, country,
                f"The comment table is split across "
                f"{st['salvaged_tables'] + 1} tables – the file was converted "
                f"or edited outside the template. "
                f"{st.get('salvaged_rows', 0)} row(s) were recovered.",
                "Open the delegation's original file and confirm every comment "
                "made it into the compilation (compare against the count in "
                "'Check by File').")

        if pf.get("kind") == "pdf":
            add("MEDIUM", "PDF submission", fn, country,
                "Submitted as PDF, not the Word template. Paragraphs were "
                "rebuilt from the printed lines; bold/italic/colour are lost.",
                "Read this delegation's rows in the Word output and confirm "
                "the paragraphs broke in the right places.")

        if not pf.get("jurisdiction_raw"):
            add("HIGH", "Country not declared", fn, country,
                "The file does not state which delegation the comments are "
                f"from. '{country}' was used.",
                "Confirm with the delegation, or check the covering email.")

        m = media.get(fn) or {}
        rows_with_img = m.get("in_comment_rows") or []
        if rows_with_img:
            where = ", ".join(f"row {r['row']}" for r in rows_with_img[:8])
            add("HIGH", "Image in a comment", fn, country,
                f"{len(rows_with_img)} comment cell(s) contain an image or "
                f"embedded object ({where}). Images are not carried into the "
                "compiled Word document.",
                "Copy the image(s) from the delegation's file into the "
                "matching row of the compilation by hand.")

        for r in st.get("rescued", []):
            add("HIGH", "Row rescued (no section/type)", fn, country,
                f"Source row {r['row']} (ref: {r['reference'] or 'n.a.'}) had an "
                f"unreadable Section ({r['section_cell'] or 'blank'}) and Type "
                f"({r['type_cell'] or 'blank'}) but carried a real comment, so "
                "it was compiled rather than dropped.",
                "Decide which section it belongs to and move it out of "
                "'Unspecified'.")

        for sk in st.get("skipped", []):
            add("MEDIUM", "Row not compiled", fn, country,
                f"Source row {sk.get('row')} "
                f"(ref: {sk.get('reference') or 'n.a.'}) was skipped: "
                f"{sk.get('reason')}",
                "Confirm the row really was empty in the delegation's file.")

    for c in compiled:
        if c.bucket == "Unspecified":
            add("HIGH", "Cannot place comment", c.source_file,
                c.jurisdiction,
                f"Comment (ref: {(c.reference.plain() or 'n.a.')[:50]}) has no "
                "Section of report and its reference does not identify one. "
                "It is compiled in an 'Unspecified' section at the end.",
                "Decide which IO / Recommendation it belongs under and move it.")

    for x in cross_country_matches(compiled):
        add("LOW", "Same text, different delegations", "n.a.",
            ", ".join(x["delegations"]),
            f"Identical comment text submitted by {len(x['delegations'])} "
            f"delegations (ref: {x['reference'] or 'n.a.'}).",
            "Normal when delegations coordinate – confirm it is not the same "
            "file uploaded twice under two names.")

    rank = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
    out.sort(key=lambda r: (rank.get(r["Priority"], 9), r["Check"], r["File"]))
    return out
