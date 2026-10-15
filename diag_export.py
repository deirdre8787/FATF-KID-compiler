# -*- coding: utf-8 -*-
"""Write the diagnostic report + warning list to an Excel workbook.

Used two ways:
  * `add_sheets(wb, ...)` – appends the diagnostic sheets to the main
    analysis workbook, so the reconciliation always travels with the data;
  * `build_workbook(...)` – a standalone diagnostics file for circulation.
"""
from datetime import datetime
from io import BytesIO
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


TNR = "Times New Roman"

HDR   = PatternFill("solid", fgColor="FFD9D9D9")
GOOD  = PatternFill("solid", fgColor="FFDFFFD6")
BAD   = PatternFill("solid", fgColor="FFFFC7CE")
WARN  = PatternFill("solid", fgColor="FFFFF2CC")
INFO  = PatternFill("solid", fgColor="FFDEEBF7")
TITLE = PatternFill("solid", fgColor="FFDCE6F4")

THIN = Border(*(Side("thin") for _ in range(4)))

SEV_FILL = {"error": BAD, "warning": WARN, "info": INFO}


def _widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _header(ws, row=1):
    for cell in ws[row]:
        cell.fill = HDR
        cell.font = Font(name=TNR, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = THIN
    ws.row_dimensions[row].height = 30


def _body(ws, first_row=2):
    for r in range(first_row, ws.max_row + 1):
        for cell in ws[r]:
            cell.font = Font(name=TNR, size=11)
            cell.border = THIN
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _sheet_summary(wb, report: Dict, issues: List[Dict], country: str):
    ws = wb.create_sheet("Diagnostic Summary")
    t = report["totals"]
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 26

    rows = [
        ("Key Issue Document Compiler – diagnostic report", None),
        (f"Assessed country: {country}", None),
        (f"Generated: {datetime.now().strftime('%d %B %Y at %H:%M')}", None),
        ("", None),
        ("RECONCILIATION", None),
        ("Files processed", t["files"]),
        ("Delegations (countries)", t["countries"]),
        ("Rows found in source tables", t["rows_in_source"]),
        ("Rows skipped (blank / placeholder / unreadable)", t["rows_skipped"]),
        ("Duplicate rows removed", t["duplicates_removed"]),
        ("Comments expected in compilation", t["expected"]),
        ("Comments actually in compilation", t["compiled"]),
        ("", None),
        ("Items needing a manual check", t.get("manual_checks", 0)),
        ("Files with a count mismatch", t["file_mismatches"]),
        ("Countries with a count mismatch", t["country_mismatches"]),
        ("EVERYTHING RECONCILES",
         "YES" if t["all_reconciled"] else "NO – see below"),
        ("", None),
        ("WARNINGS RAISED", None),
        ("Errors", sum(1 for i in issues if i["severity"] == "error")),
        ("Warnings", sum(1 for i in issues if i["severity"] == "warning")),
        ("Notices", sum(1 for i in issues if i["severity"] == "info")),
    ]

    for i, (label, value) in enumerate(rows, start=1):
        ws.cell(row=i, column=1, value=label)
        if value is not None:
            ws.cell(row=i, column=2, value=value)

    ws.cell(row=1, column=1).font = Font(name=TNR, bold=True, size=14)
    for r in (2, 3):
        ws.cell(row=r, column=1).font = Font(name=TNR, italic=True, size=11)
    for r in (5, 18):
        ws.cell(row=r, column=1).font = Font(name=TNR, bold=True, size=12)
        ws.cell(row=r, column=1).fill = TITLE

    ok_cell = ws.cell(row=16, column=2)
    ok_cell.font = Font(name=TNR, bold=True)
    ok_cell.fill = GOOD if t["all_reconciled"] else BAD
    ws.cell(row=16, column=1).font = Font(name=TNR, bold=True)


def _sheet_by_file(wb, report: Dict):
    ws = wb.create_sheet("Check by File")
    cols = ["File name", "Country assigned", "Source type",
            "Rows in source table", "Rows skipped", "Duplicates removed",
            "Expected in compilation", "Actually in compilation", "Match",
            "Notes"]
    ws.append(cols)
    _header(ws)

    for fr in report["files"]:
        ws.append([fr[c] for c in cols])

    _widths(ws, [34, 20, 12, 14, 12, 13, 14, 14, 10, 70])
    _body(ws)

    for r in range(2, ws.max_row + 1):
        m = ws.cell(row=r, column=9)
        m.fill = GOOD if m.value == "YES" else BAD
        m.font = Font(name=TNR, bold=True)
        m.alignment = Alignment(horizontal="center", vertical="center")
        if m.value == "NO":
            for c in range(1, 11):
                ws.cell(row=r, column=c).fill = BAD
    ws.freeze_panes = "A2"


def _sheet_by_country(wb, report: Dict):
    ws = wb.create_sheet("Check by Country")
    cols = ["Country", "Files", "No. of files", "Expected comments",
            "Comments in compilation", "Match", "Difference"]
    ws.append(cols)
    _header(ws)

    for cr in report["countries"]:
        ws.append([cr[c] for c in cols])

    _widths(ws, [24, 46, 12, 16, 18, 10, 12])
    _body(ws)

    for r in range(2, ws.max_row + 1):
        m = ws.cell(row=r, column=6)
        m.fill = GOOD if m.value == "YES" else BAD
        m.font = Font(name=TNR, bold=True)
        m.alignment = Alignment(horizontal="center", vertical="center")
        if m.value == "NO":
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = BAD

    tr = ws.max_row + 1
    ws.cell(row=tr, column=1, value="TOTAL")
    ws.cell(row=tr, column=4,
            value=sum(c["Expected comments"] for c in report["countries"]))
    ws.cell(row=tr, column=5,
            value=sum(c["Comments in compilation"] for c in report["countries"]))
    for cell in ws[tr]:
        cell.font = Font(name=TNR, bold=True)
        cell.fill = TITLE
        cell.border = THIN
    ws.freeze_panes = "A2"


def _sheet_skipped(wb, report: Dict):
    rows = []
    for fr in report["files"]:
        for s in fr.get("_skipped_detail") or []:
            rows.append([fr["File name"], fr["Country assigned"],
                         s.get("row"), s.get("reference", ""),
                         s.get("reason", "")])
    if not rows:
        return
    ws = wb.create_sheet("Skipped Rows")
    ws.append(["File name", "Country", "Source row #",
               "Reference in row", "Why it was skipped"])
    _header(ws)
    for r in rows:
        ws.append(r)
    _widths(ws, [32, 18, 12, 34, 66])
    _body(ws)
    ws.freeze_panes = "A2"


def _sheet_duplicates(wb, report: Dict):
    dups = report.get("duplicates") or []
    ws = wb.create_sheet("Duplicates Removed")
    ws.append(["Delegation", "Reference", "Comment (truncated)",
               "Removed from file", "Removed row", "Kept from file",
               "Kept row"])
    _header(ws)
    for d in dups:
        ws.append([d["jurisdiction"], d["reference"], d["comment"],
                   d["removed_from"], d["removed_row"],
                   d["kept_from"], d["kept_row"]])
    if not dups:
        ws.append(["None", "", "", "", "", "", ""])
    _widths(ws, [20, 26, 70, 28, 12, 28, 10])
    _body(ws)
    ws.freeze_panes = "A2"


def _sheet_cross_country(wb, report: Dict):
    xs = report.get("cross_country") or []
    if not xs:
        return
    ws = wb.create_sheet("Same Text, Diff. Country")
    ws.append(["Delegations submitting identical text", "No.",
               "Reference", "Comment (truncated)"])
    _header(ws)
    for x in xs:
        ws.append([", ".join(x["delegations"]), x["count"],
                   x["reference"], x["comment"]])
    _widths(ws, [40, 8, 28, 76])
    _body(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = WARN
    ws.freeze_panes = "A2"


def _sheet_manual(wb, report: Dict):
    """The human to-do list – first sheet the team should open."""
    rows = report.get("manual_checks") or []
    ws = wb.create_sheet("Manual Checks")
    ws.append(["Priority", "Check", "File", "Country",
               "What was found", "What to do"])
    _header(ws)
    for r in rows:
        ws.append([r["Priority"], r["Check"], r["File"], r["Country"],
                   r["What was found"], r["What to do"]])
    if not rows:
        ws.append(["n.a.", "Nothing to check", "", "",
                   "No manual follow-up needed for this compilation.", ""])
    _widths(ws, [10, 30, 32, 20, 74, 66])
    _body(ws)
    prio = {"HIGH": BAD, "MEDIUM": WARN, "LOW": INFO}
    for r in range(2, ws.max_row + 1):
        f = prio.get(ws.cell(row=r, column=1).value)
        if f:
            ws.cell(row=r, column=1).fill = f
            ws.cell(row=r, column=1).font = Font(name=TNR, bold=True)
    ws.freeze_panes = "A2"


def _sheet_warnings(wb, issues: List[Dict]):
    ws = wb.create_sheet("Warnings")
    ws.append(["Severity", "Check", "File", "Detail"])
    _header(ws)

    order = {"error": 0, "warning": 1, "info": 2}
    for i in sorted(issues, key=lambda x: order.get(x["severity"], 9)):
        ws.append([i["severity"].upper(), i["kind"], i["file"], i["message"]])
    if not issues:
        ws.append(["n.a.", "none", "", "No issues raised."])

    _widths(ws, [12, 24, 34, 96])
    _body(ws)
    for r in range(2, ws.max_row + 1):
        sev = (ws.cell(row=r, column=1).value or "").lower()
        fill = SEV_FILL.get(sev)
        if fill:
            ws.cell(row=r, column=1).fill = fill
            ws.cell(row=r, column=1).font = Font(name=TNR, bold=True)
    ws.freeze_panes = "A2"


def add_sheets(wb, report: Dict, issues: List[Dict], country: str = ""):
    """Append every diagnostic sheet to an existing workbook."""
    _sheet_summary(wb, report, issues, country)
    _sheet_manual(wb, report)
    _sheet_by_file(wb, report)
    _sheet_by_country(wb, report)
    _sheet_skipped(wb, report)
    _sheet_duplicates(wb, report)
    _sheet_cross_country(wb, report)
    _sheet_warnings(wb, issues)


def build_workbook(report: Dict, issues: List[Dict],
                   country: str = "") -> BytesIO:
    """Standalone diagnostics workbook."""
    wb = Workbook()
    wb.remove(wb.active)
    add_sheets(wb, report, issues, country)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
