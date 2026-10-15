# -*- coding: utf-8 -*-
"""Pre-compilation validation.

Runs a battery of checks over every uploaded file *before* anything is
compiled, and returns a structured issue list so the UI can show the user
exactly what needs attention (and get explicit confirmation) first.

Issue severities:
  * "error"   – blocks compilation of that item unless resolved/overridden
  * "warning" – compilation can proceed, but the user must acknowledge
  * "info"    – FYI only (e.g. late detection notice)

Issue kinds currently emitted:
  file_unreadable         error    file could not be parsed at all
  no_comments             error    no comment rows found in the file
  missing_jurisdiction    error    no country dropdown value; suggestion from
                                   filename supplied when possible (needs user
                                   confirmation via popup)
  jurisdiction_mismatch   warning  dropdown country != filename country
  unknown_jurisdiction    warning  jurisdiction not a FATF member / observer
  missing_section         warning  a row has no Section of report value
  missing_type            warning  a row has no Type value
  missing_reference       warning  a row has an empty MER reference cell
  empty_comment           warning  a row has an empty Comment cell
  placeholder_left        warning  a row still contains template placeholder
  duplicate_comment       warning  identical (jurisdiction, ref, text) rows
  duplicate_file          warning  two uploaded files normalise to the same
                                   jurisdiction (possible double submission)
  late_file               info     filename flags the file as late
  pdf_parsed              info     file was ingested via the PDF fallback
                                   parser (formatting fidelity is reduced)
"""
import os
import re
from collections import Counter
from typing import Dict, List, Optional

from extractor import VALID_SECTIONS, VALID_TYPES


# FATF members + common observers for the "unknown jurisdiction" sanity check.
KNOWN_JURISDICTIONS = {
    "Argentina","Australia","Austria","Belgium","Brazil","Canada","China",
    "Denmark","European Commission","Finland","France","Germany","Greece",
    "Gulf Co-operation Council","Hong Kong, China","Iceland","India",
    "Indonesia","Ireland","Israel","Italy","Japan","Korea","Luxembourg",
    "Malaysia","Mexico","Netherlands","New Zealand","Norway","Portugal",
    "Russian Federation","Saudi Arabia","Singapore","South Africa","Spain",
    "Sweden","Switzerland","Türkiye","Turkey","United Kingdom","United States",
    # frequent observers / FSRB secretariats that sometimes comment
    "IMF","World Bank","UNODC","OECD","Egmont Group","MONEYVAL","APG",
    "GAFILAT","MENAFATF","GIABA","EAG","ESAAMLG","CFATF","GABAC",
}

PLACEHOLDER_PATTERNS = [
    "choose an item",
    "click or tap here",
    "click here to enter",
    "enter comment",
]


def _placeholder_in(text: str) -> bool:
    low = (text or "").lower()
    return any(p in low for p in PLACEHOLDER_PATTERNS)


def make_issue(kind, severity, file, message, row=None, suggestion=None):
    d = {"kind": kind, "severity": severity, "file": file, "message": message}
    if row is not None:
        d["row"] = row
    if suggestion:
        d["suggestion"] = suggestion
    return d


def validate_files(parsed_files: List[dict]) -> Dict:
    """parsed_files: list of dicts:
        {
          "filename": str,
          "kind": "docx" | "pdf",
          "error": Optional[str],        # fatal parse error
          "comments": List[Comment],     # may be empty
          "jurisdiction_raw": str,       # what the file itself declared ("" if none)
          "filename_guess": str,         # country guessed from filename ("" if none)
          "is_late": bool,
        }

    Returns {"issues": [...], "files": [...per-file summary...]}
    """
    issues: List[dict] = []
    summaries: List[dict] = []

    # ---- per-file checks -------------------------------------------------
    for pf in parsed_files:
        fn = pf["filename"]
        comments = pf.get("comments") or []

        if pf.get("error"):
            issues.append(make_issue(
                "file_unreadable", "error", fn,
                f"File could not be read: {pf['error']}"))
            summaries.append({"filename": fn, "comments": 0,
                              "jurisdiction": "", "status": "unreadable"})
            continue

        if pf.get("kind") == "pdf":
            issues.append(make_issue(
                "pdf_template", "warning", fn,
                "This delegation returned a PDF instead of the Word template. "
                "It WILL be processed, but rich-text formatting (bold, "
                "italics, colours, bullets) inside comments cannot be "
                "recovered from a PDF, and table cells occasionally split "
                "awkwardly. Check this file's rows in the compiled output."))

        # embedded images / objects
        media = pf.get("media") or {}
        if media.get("count"):
            in_rows = media.get("in_comment_rows") or []
            sev = "warning" if in_rows else "info"
            issues.append(make_issue(
                "images_present", sev, fn,
                media.get("description") or
                f"{media['count']} embedded image(s) detected."))

        if not comments:
            issues.append(make_issue(
                "no_comments", "error", fn,
                "No comment rows could be extracted from this file. "
                "Check that it contains the standard 4-column comment table."))
            summaries.append({"filename": fn, "comments": 0,
                              "jurisdiction": pf.get("jurisdiction_raw", ""),
                              "status": "empty"})
            continue

        jur = (pf.get("jurisdiction_raw") or "").strip()
        guess = (pf.get("filename_guess") or "").strip()

        if not jur:
            if guess:
                issues.append(make_issue(
                    "missing_jurisdiction", "error", fn,
                    "No 'Comments from' country found inside the file. "
                    f"Filename suggests '{guess}' – please confirm.",
                    suggestion=guess))
            else:
                issues.append(make_issue(
                    "missing_jurisdiction", "error", fn,
                    "No 'Comments from' country found inside the file and "
                    "none could be guessed from the filename. Please assign "
                    "one before compiling.",
                    suggestion=""))
        else:
            from countries import canonical
            jur_c, guess_c = canonical(jur), canonical(guess)
            if guess and jur_c != guess_c and not (jur_c and guess_c and jur_c == guess_c):
                if (jur_c or jur).lower() != (guess_c or guess).lower():
                    issues.append(make_issue(
                        "jurisdiction_mismatch", "warning", fn,
                        f"File says the comments are from '{jur}' but the "
                        f"filename says '{guess}'. Confirm which delegation "
                        "these comments belong to in the dropdown – the whole "
                        "compilation is attributed on this."))
            if not jur_c:
                issues.append(make_issue(
                    "unknown_jurisdiction", "warning", fn,
                    f"'{jur}' is not a recognised FATF delegation, observer or "
                    "country name. Pick the right one from the dropdown."))

        # ---- structural integrity of the submitted template ------------
        st = pf.get("stats") or {}
        if st.get("salvaged_tables"):
            issues.append(make_issue(
                "template_altered", "warning", fn,
                f"This template's comment table is split across "
                f"{st['salvaged_tables'] + 1} separate tables (the header row "
                "sits in one table and the comments in another) – a sign the "
                "file was converted, re-saved or edited outside the official "
                "template. "
                f"{st.get('salvaged_rows', 0)} row(s) were recovered and ARE "
                "included, but check them against the original in the "
                "'Manual Checks' sheet."))

        for r in st.get("rescued", []):
            issues.append(make_issue(
                "rescued_row", "warning", fn,
                f"Row {r['row']} (ref: {r['reference'] or 'n.a.'}): neither the "
                f"Section ({r['section_cell'] or 'blank'!r}) nor the Type "
                f"({r['type_cell'] or 'blank'!r}) could be read, but the row "
                "has a real comment, so it WAS compiled – under 'Unspecified'. "
                "Assign it to a section by hand.", row=r["row"]))

        if st.get("total_rows") and not st.get("extracted"):
            issues.append(make_issue(
                "no_comments", "error", fn,
                f"The comment table was found ({st['total_rows']} rows) but "
                "no row produced a usable comment – every row was blank or "
                "still held placeholder text."))

        if pf.get("is_late"):
            issues.append(make_issue(
                "late_file", "info", fn,
                "Filename marks this submission as LATE – its comments will "
                "be shaded pink. Change it in the Timeliness dropdown if the "
                "filename is wrong."))

        # ---- per-row checks ----------------------------------------------
        for i, c in enumerate(comments, start=1):
            ref_txt = c.reference.plain().strip() if c.reference else ""
            com_txt = c.comment.plain().strip() if c.comment else ""

            if (not c.section or c.section == "Unspecified"):
                if c.bucket.startswith("Recommendation"):
                    issues.append(make_issue(
                        "missing_section", "info", fn,
                        f"Row {i}: 'Section of report' was empty, but the "
                        f"reference '{ref_txt[:30]}' identifies it as "
                        f"{c.bucket} – it will be compiled there.", row=i))
                else:
                    issues.append(make_issue(
                        "missing_section", "warning", fn,
                        f"Row {i}: no 'Section of report' value "
                        f"(ref: {ref_txt[:40] or 'n.a.'}) and the reference "
                        "does not identify a section – it will be listed "
                        "under 'Unspecified'.", row=i))
            if not c.ctype:
                issues.append(make_issue(
                    "missing_type", "warning", fn,
                    f"Row {i}: no 'Type' value (ref: {ref_txt[:40] or 'n.a.'}).",
                    row=i))
            if c.bucket == "Unspecified":
                issues.append(make_issue(
                    "unplaceable_row", "warning", fn,
                    f"Row {i}: cannot be placed under any MER section "
                    f"(ref: {ref_txt[:40] or 'n.a.'}). It will be compiled in an "
                    "'Unspecified' section at the end – decide where it "
                    "belongs before circulating.", row=i))
            if com_txt and len(com_txt) < 25:
                issues.append(make_issue(
                    "very_short_comment", "info", fn,
                    f"Row {i}: the comment is only {len(com_txt)} characters "
                    f"({com_txt[:30]!r}) – check nothing was truncated.",
                    row=i))
            if not ref_txt:
                issues.append(make_issue(
                    "missing_reference", "warning", fn,
                    f"Row {i}: MER reference cell is empty.", row=i))
            if not com_txt:
                issues.append(make_issue(
                    "empty_comment", "warning", fn,
                    f"Row {i}: Comment cell is empty "
                    f"(ref: {ref_txt[:40] or 'n.a.'}).", row=i))
            if _placeholder_in(ref_txt) or _placeholder_in(com_txt):
                issues.append(make_issue(
                    "placeholder_left", "warning", fn,
                    f"Row {i}: still contains template placeholder text.",
                    row=i))

        summaries.append({
            "filename": fn,
            "comments": len(comments),
            "jurisdiction": jur or (guess and f"(suggested) {guess}") or "",
            "status": "late" if pf.get("is_late") else "ok",
        })

    # ---- cross-file checks -----------------------------------------------
    # duplicate submissions: same jurisdiction in >1 file
    jur_files = {}
    for pf in parsed_files:
        j = (pf.get("jurisdiction_raw") or pf.get("filename_guess") or "").strip().lower()
        if j:
            jur_files.setdefault(j, []).append(pf["filename"])
    for j, fns in jur_files.items():
        if len(fns) > 1:
            issues.append(make_issue(
                "duplicate_file", "warning", ", ".join(fns),
                f"Multiple files appear to be from the same delegation "
                f"('{j.title()}'): {', '.join(fns)}. All their comments will "
                "be included – remove a file if one is an outdated version."))

    # duplicate comments: identical (jurisdiction, reference, comment text)
    def _norm(s: str) -> str:
        # whitespace-insensitive so a PDF line-wrapped copy still matches
        return re.sub(r"\s+", " ", (s or "")).strip().lower()

    seen = Counter()
    locs = {}
    for pf in parsed_files:
        for i, c in enumerate(pf.get("comments") or [], start=1):
            key = (
                _norm(pf.get("jurisdiction_raw") or pf.get("filename_guess") or ""),
                _norm(c.reference.plain() if c.reference else ""),
                _norm(c.comment.plain() if c.comment else ""),
            )
            if not key[2]:
                continue
            seen[key] += 1
            locs.setdefault(key, []).append(f"{pf['filename']} row {i}")
    for key, n in seen.items():
        if n > 1:
            issues.append(make_issue(
                "duplicate_comment", "warning", "; ".join(locs[key]),
                f"The same delegation submitted an identical comment {n}× "
                f"(ref: {key[1][:40] or 'n.a.'}). {n - 1} copy/copies will be "
                "REMOVED automatically – only the first is compiled. See the "
                "'Duplicates Removed' sheet of the diagnostic report."))

    return {"issues": issues, "files": summaries}


def has_blocking_errors(issues: List[dict]) -> bool:
    return any(i["severity"] == "error" for i in issues)
