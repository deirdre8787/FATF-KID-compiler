# -*- coding: utf-8 -*-
"""key_issue_document_compiler_vf – FATF MER comment compilation tool.

Two-phase flow:
  1. POST /validate  – parse every file (docx or pdf), run all checks and
     return the issue list + per-file country suggestions. Nothing compiled.
  2. POST /process   – same files re-submitted together with the confirmed
     jurisdiction assignments from the popup; compiles Excel + Word.
"""
import json
import os
import re
import tempfile
from datetime import datetime
from io import BytesIO

import pytz
from flask import Flask, render_template, request, send_file, jsonify

from extractor import extract_comments, extract_jurisdiction
from countries import (guess_from_filename as _guess_jurisdiction_from_filename,
                       resolve as _resolve_country, canonical as _canon_country,
                       grouped_choices, all_countries)
from pdf_extractor import extract_comments_pdf
from excel_export import build_excel
from word_export import build_word
from validator import validate_files, has_blocking_errors
from media_probe import probe as probe_media, describe as describe_media
from diagnostics import dedupe_comments, build_report
from diag_export import build_workbook as build_diag_workbook, add_sheets as add_diag_sheets


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

_LAST = {
    "excel": None, "word": None, "diag": None, "country": "",
    "failed": [], "ok_count": 0, "comment_count": 0, "when": "",
}


# "late" as a standalone token: no letter immediately before or after.
_LATE_RE = re.compile(r"(?<![A-Za-z])late(?![A-Za-z])", re.IGNORECASE)


def _filename_is_late(name: str) -> bool:
    if not name:
        return False
    stem = os.path.splitext(os.path.basename(name))[0]
    return bool(_LATE_RE.search(stem))


def _detect_country_from_files(files):
    """Assessed country = the country named in most filenames."""
    from collections import Counter
    counter = Counter()
    for f in files:
        name = getattr(f, "filename", "") or str(f)
        from countries import all_matches_in_filename
        for g in all_matches_in_filename(name):
            counter[g] += 1
    # The assessed country is the one named in (nearly) every filename;
    # a delegation's own name appears in only its own file.
    if counter and len(files) > 1:
        top, n = counter.most_common(1)[0]
        if n >= max(2, len(files) * 0.6):
            return top
    if counter:
        return counter.most_common(1)[0][0]
    return "Unknown"


def _parse_one(path: str, fn: str):
    """Parse a single uploaded file (docx or pdf) into the structure the
    validator expects. Never raises – errors are captured per-file."""
    ext = os.path.splitext(fn)[1].lower()
    rec = {
        "filename": fn,
        "kind": "pdf" if ext == ".pdf" else "docx",
        "error": None,
        "comments": [],
        "jurisdiction_raw": "",
        "filename_guess": _guess_jurisdiction_from_filename(fn) or "",
        "is_late": _filename_is_late(fn),
    }
    stats = {}
    try:
        if ext == ".docx":
            rec["comments"] = extract_comments(path, filename_hint=fn,
                                               stats=stats)
            rec["jurisdiction_raw"] = extract_jurisdiction(path) or ""
        elif ext == ".pdf":
            rec["comments"] = extract_comments_pdf(path, filename_hint=fn)
            rec["jurisdiction_raw"] = (
                rec["comments"][0].jurisdiction if rec["comments"] else "")
            n = len(rec["comments"])
            stats = {"total_rows": n, "extracted": n, "skipped": []}
        else:
            rec["error"] = f"Unsupported file type '{ext}'"
    except Exception as e:
        rec["error"] = str(e)
    rec["stats"] = stats

    # embedded images / objects (never carried into the compiled Word)
    try:
        m = probe_media(path, rec["kind"])
        m["description"] = describe_media(m, rec["kind"])
        rec["media"] = m
    except Exception:
        rec["media"] = {}
    return rec


def _collect_files(req):
    """Save uploads to a temp dir; returns (tmpdir_ctx, [(path, filename)])."""
    files = req.files.getlist("files")
    tmp = tempfile.TemporaryDirectory()
    out = []
    for f in files:
        fn = f.filename
        if not fn:
            continue
        ext = os.path.splitext(fn)[1].lower()
        if ext not in (".docx", ".pdf"):
            continue
        path = os.path.join(tmp.name, os.path.basename(fn))
        f.save(path)
        out.append((path, fn))
    return tmp, out


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", last=_LAST)


@app.route("/detect_country", methods=["POST"])
def detect_country():
    names = request.json.get("filenames", [])
    class _F: pass
    fake = []
    for n in names:
        f = _F(); f.filename = n; fake.append(f)
    return jsonify({"country": _detect_country_from_files(fake)})


@app.route("/countries")
def countries():
    """Full picker list – FATF delegations, observers, then all countries."""
    return jsonify({"groups": grouped_choices()})


@app.route("/validate", methods=["POST"])
def validate():
    """Phase 1: parse + validate everything, compile nothing."""
    tmp, saved = _collect_files(request)
    with tmp:
        if not saved:
            return jsonify({"ok": False,
                            "error": "No .docx or .pdf files uploaded."}), 400
        parsed = [_parse_one(path, fn) for path, fn in saved]
        report = validate_files(parsed)

        # Per-file review row for the confirmation popup. EVERY file gets a
        # country dropdown and a late/on-time dropdown – the operator confirms
        # all of them, not just the ones the tool was unsure about.
        assignments = []
        for pf in parsed:
            declared = pf["jurisdiction_raw"]
            declared_canon = _canon_country(declared) if declared else None
            suggested = pf["filename_guess"]
            chosen = declared_canon or suggested or ""
            mismatch = bool(declared_canon and suggested
                            and declared_canon != suggested)
            assignments.append({
                "filename": pf["filename"],
                "declared": declared,
                "declared_canonical": declared_canon or "",
                "suggested": suggested,
                "selected": chosen,
                "recognised": bool(declared_canon),
                "mismatch": mismatch,
                "needs_confirmation": not chosen,
                "late": bool(pf.get("is_late")),
                "comments": len(pf.get("comments") or []),
                "empty": not pf.get("comments"),
                "kind": pf.get("kind"),
                "error": pf.get("error") or "",
            })

        empties = [a["filename"] for a in assignments if a["empty"]]

        return jsonify({
            "ok": True,
            "issues": report["issues"],
            "files": report["files"],
            "assignments": assignments,
            "countries": grouped_choices(),
            "empty_files": empties,
            "blocking": has_blocking_errors(report["issues"]),
        })


@app.route("/process", methods=["POST"])
def process():
    """Phase 2: compile, using the user-confirmed jurisdiction map."""
    tmp, saved = _collect_files(request)
    with tmp:
        if not saved:
            return jsonify({"ok": False,
                            "error": "No .docx or .pdf files uploaded."}), 400

        # Per-file decisions confirmed in the popup:
        #   {"file.docx": {"country": "Korea", "late": false}, ...}
        # (the older {"file": "Country"} shape is still accepted)
        overrides, late_overrides = {}, {}
        raw = request.form.get("jurisdiction_overrides")
        if raw:
            try:
                data = json.loads(raw)
                for fn, val in (data or {}).items():
                    if isinstance(val, dict):
                        if val.get("country"):
                            overrides[fn] = val["country"]
                        if "late" in val:
                            late_overrides[fn] = bool(val["late"])
                    elif val:
                        overrides[fn] = val
            except Exception:
                overrides, late_overrides = {}, {}

        # Files the operator knowingly accepted as contributing nothing.
        accepted_empty = set()
        try:
            accepted_empty = set(json.loads(
                request.form.get("accepted_empty") or "[]"))
        except Exception:
            pass

        assessed = (request.form.get("assessed_country") or "").strip()
        if not assessed:
            class _F: pass
            fake = []
            for _, fn in saved:
                f = _F(); f.filename = fn; fake.append(f)
            assessed = _detect_country_from_files(fake) or "Unknown"

        all_comments = []
        failed = []
        parsed = []
        assignments = {}
        media_map = {}

        for path, fn in saved:
            rec = _parse_one(path, fn)
            parsed.append(rec)
            media_map[fn] = rec.get("media") or {}

            jur = _resolve_country(
                overrides.get(fn) or rec["jurisdiction_raw"]
                or rec["filename_guess"] or "Unknown")
            assignments[fn] = jur
            rec["assigned"] = jur
            rec["country_source"] = (
                "confirmed in popup" if overrides.get(fn)
                else "declared in file" if rec["jurisdiction_raw"]
                else "guessed from filename" if rec["filename_guess"]
                else "not identified")

            # operator's late/on-time choice wins over the filename rule
            late = late_overrides.get(fn, rec["is_late"])
            rec["is_late"] = late

            if rec["error"]:
                failed.append(f"{fn} ({rec['error']})")
                continue
            comments = rec["comments"]
            if not comments:
                failed.append(fn)
                continue

            for c in comments:
                c.on_time = not late
                c.source_file = fn
                c.jurisdiction = jur
            all_comments.extend(comments)

        # HARD BLOCK: a file that yielded nothing is never compiled silently.
        # Either it gets fixed and re-uploaded, or the operator explicitly
        # accepts that it contributes nothing.
        unacknowledged = [fn for fn in failed
                          if fn.split(" (")[0] not in accepted_empty]
        if unacknowledged:
            return jsonify({
                "ok": False,
                "blocked": True,
                "error": "These files produced no comments. Fix and re-upload "
                         "them, or tick 'accept as empty' for each one:",
                "failed": failed,
                "needs_acknowledgement": unacknowledged,
            }), 409

        if not all_comments:
            return jsonify({"ok": False, "error": "No comments extracted.",
                            "failed": failed}), 400

        # remove duplicate submissions (same delegation, same reference,
        # same text) before anything is written out
        all_comments, removed = dedupe_comments(all_comments)

        # re-validate on the final, deduplicated set so the warning list
        # exported alongside the diagnostics matches what was compiled
        report_issues = validate_files(parsed)["issues"]
        report = build_report(parsed, assignments, all_comments,
                              removed, media_map)

        excel_buf = build_excel(all_comments, assessed,
                                diagnostics=(report, report_issues))
        word_buf = build_word(all_comments, assessed)
        diag_buf = build_diag_workbook(report, report_issues, assessed)

        _LAST["excel"] = BytesIO(excel_buf.getvalue())
        _LAST["word"] = BytesIO(word_buf.getvalue())
        _LAST["diag"] = BytesIO(diag_buf.getvalue())
        _LAST["country"] = assessed
        _LAST["failed"] = failed
        _LAST["ok_count"] = len({c.source_file for c in all_comments})
        _LAST["comment_count"] = len(all_comments)
        _LAST["when"] = datetime.now().strftime("%Y-%m-%d %H:%M")

        t = report["totals"]
        return jsonify({
            "ok": True, "country": assessed,
            "comments": len(all_comments),
            "files_ok": _LAST["ok_count"], "failed": failed,
            "duplicates_removed": len(removed),
            "reconciled": t["all_reconciled"],
            "file_mismatches": t["file_mismatches"],
            "country_mismatches": t["country_mismatches"],
            "diagnostics": {
                "files": [{k: v for k, v in f.items()
                           if not k.startswith("_")} for f in report["files"]],
                "countries": report["countries"],
                "totals": t,
            },
        })


def _stamp():
    return datetime.now().strftime("%d_%m_%Y")


@app.route("/download/excel")
def download_excel():
    if not _LAST["excel"]:
        return "Nothing to download yet.", 400
    buf = BytesIO(_LAST["excel"].getvalue()); buf.seek(0)
    name = f"{_LAST['country'] or 'MER'}_Compiled_Analysis_{_stamp()}.xlsx"
    return send_file(buf, as_attachment=True, download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/download/word")
def download_word():
    if not _LAST["word"]:
        return "Nothing to download yet.", 400
    buf = BytesIO(_LAST["word"].getvalue()); buf.seek(0)
    name = f"{_LAST['country'] or 'MER'}_Compiled_Comments_{_stamp()}.docx"
    return send_file(buf, as_attachment=True, download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/download/diagnostics")
def download_diagnostics():
    if not _LAST["diag"]:
        return "Nothing to download yet.", 400
    buf = BytesIO(_LAST["diag"].getvalue()); buf.seek(0)
    name = f"{_LAST['country'] or 'MER'}_Diagnostic_Report_{_stamp()}.xlsx"
    return send_file(buf, as_attachment=True, download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    try:
        from waitress import serve
        print("Running on http://127.0.0.1:5050  (Ctrl+C to stop)")
        serve(app, host="127.0.0.1", port=5050, threads=8)
    except ImportError:
        app.run(host="127.0.0.1", port=5050, debug=False)
