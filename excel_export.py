# -*- coding: utf-8 -*-
"""Build the analytical Excel workbook."""
import re
from io import BytesIO
from typing import List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from extractor import Comment, VALID_SECTIONS, VALID_TYPES
from sorter import bucket_sort_key


# ---- colours -------------------------------------------------------------
GREY   = PatternFill("solid", fgColor="FFD9D9D9")
GREEN  = PatternFill("solid", fgColor="FFDFFFD6")
RED    = PatternFill("solid", fgColor="FFFFD9E2")  # very light pastel pink for "Late"
BLUE   = PatternFill("solid", fgColor="FFE6F2FF")
PINK   = PatternFill("solid", fgColor="FFFFCCE5")
YELLOW = PatternFill("solid", fgColor="FFFFFFCC")
ORANGE = PatternFill("solid", fgColor="FFFFE0B3")
PURPLE = PatternFill("solid", fgColor="FFE6E6FF")

THIN  = Border(*(Side("thin") for _ in range(4)))
MED   = Border(*(Side("medium") for _ in range(4)))

TNR = "Times New Roman"

# FATF membership (for "no comment" sheet)
FATF_MEMBERS = sorted([
    "Argentina","Australia","Austria","Belgium","Brazil","Canada","China",
    "Denmark","European Commission","Finland","France","Germany","Greece",
    "Gulf Co-operation Council","Hong Kong, China","Iceland","India","Indonesia",
    "Ireland","Israel","Italy","Japan","Korea","Luxembourg","Malaysia",
    "Mexico","Netherlands","New Zealand","Norway","Portugal",
    "Russian Federation","Saudi Arabia","Singapore","South Africa","Spain",
    "Sweden","Switzerland","Türkiye","United Kingdom","United States",
])


def _fit_columns(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _style_header(ws, row=1):
    for cell in ws[row]:
        cell.fill = GREY
        cell.font = Font(name=TNR, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    ws.row_dimensions[row].height = 30


def _clean_for_excel(rich_text_obj) -> str:
    """RichText -> plain string for Excel cells."""
    return rich_text_obj.plain()


def _build_instructions(wb, assessed_country):
    ws = wb.create_sheet("Instructions", 0)
    ws.column_dimensions["A"].width = 120
    lines = [
        ("Reader’s Guide – MER Comment Compilation", Font(name=TNR, bold=True, size=14)),
        ("", None),
        (f"Assessed country: {assessed_country}", Font(name=TNR, italic=True, size=11)),
        ("", None),
        ("This workbook aggregates comments submitted by delegations on the "
         "draft Mutual Evaluation Report using the Simplified 5th Round ME comment template.",
         Font(name=TNR, size=11)),
        ("", None),
        ("Sheets", Font(name=TNR, bold=True, size=12)),
        ("  • Compiled – every comment as a row (filterable table).", Font(name=TNR, size=11)),
        ("  • By Section – counts of comments per MER section (Exec Summary, IO.1–11, Recommendations, …).", Font(name=TNR, size=11)),
        ("  • By Type – counts by the dropdown Type (Analysis, Key Findings, Ratings, …).", Font(name=TNR, size=11)),
        ("  • By Country – one row per delegation showing which sections and recommendations they touched.", Font(name=TNR, size=11)),
        ("  • Country × Section – cross-tab of delegation vs section.", Font(name=TNR, size=11)),
        ("  • Country × Type – cross-tab of delegation vs comment type.", Font(name=TNR, size=11)),
        ("  • No Comments – FATF members who submitted nothing.", Font(name=TNR, size=11)),
        ("  • Excl. Assessed Country – mirror of Summary/By Country without the assessed jurisdiction’s own comments.", Font(name=TNR, size=11)),
        ("", None),
        ("Conventions", Font(name=TNR, bold=True, size=12)),
        ("  • Green rows/cells = submitted on time. Red = late.", Font(name=TNR, size=11)),
        ("  • Purple highlight marks the most frequent type/IO per delegation.", Font(name=TNR, size=11)),
        ("  • TC Annex references (R.1, C.1.6, …) are bucketed into Recommendation N automatically.", Font(name=TNR, size=11)),
        ("  • Editorial = rows whose Type is ‘Non-substantive editorial changes’. Everything else = Substantive.", Font(name=TNR, size=11)),
    ]
    for i, (txt, font) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=txt)
        if font:
            ws.cell(row=i, column=1).font = font


def _build_compiled_sheet(wb, comments: List[Comment]):
    ws = wb.create_sheet("Compiled")
    headers = ["Section", "Bucket", "Type", "Substantive/Editorial",
               "MER reference", "Comment", "Jurisdiction",
               "Deadline status", "Source file", "Order"]
    ws.append(headers)
    _style_header(ws)

    for c in comments:
        ws.append([
            c.section,
            c.bucket,
            c.ctype,
            c.category,
            _clean_for_excel(c.reference),
            _clean_for_excel(c.comment),
            c.jurisdiction,
            "In time" if c.on_time else "Late",
            c.source_file,
            c.order_in_file,
        ])

    # Column widths
    _fit_columns(ws, [14, 22, 28, 18, 18, 60, 22, 14, 28, 8])

    # Font for data
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = Font(name=TNR, size=11)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Deadline colouring on col 8 (H), and tint the whole comment row
    # very light pastel pink when the row is late.
    for r in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=r, column=8)
        if status_cell.value == "In time":
            status_cell.fill = GREEN
        else:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=r, column=col).fill = RED

    # Make filterable Excel table
    if ws.max_row >= 2:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tbl = Table(displayName="CompiledTable", ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1",
                                            showRowStripes=False)
        ws.add_table(tbl)
        _style_header(ws)  # re-apply header styling after table conversion

    ws.freeze_panes = "A2"


def _build_by_section(wb, comments: List[Comment], title: str = "By Section"):
    ws = wb.create_sheet(title)
    ws.append(["Section / Bucket", "Total", "Substantive", "Editorial",
               "Unique jurisdictions", "Most common type"])
    _style_header(ws)

    # Bucket aggregation
    buckets = {}
    for c in comments:
        buckets.setdefault(c.bucket, []).append(c)

    rows = []
    for bucket in sorted(buckets.keys(), key=bucket_sort_key):
        items = buckets[bucket]
        total = len(items)
        subs = sum(1 for i in items if i.substantive)
        eds = total - subs
        juris = len({i.jurisdiction for i in items if i.jurisdiction})
        # Most common type among Substantive
        types = [i.ctype for i in items if i.substantive]
        most = ""
        if types:
            from collections import Counter
            most = Counter(types).most_common(1)[0][0]
        rows.append([bucket, total, subs, eds, juris, most])
        ws.append(rows[-1])

    # Totals
    ws.append([])
    totals_row = ws.max_row + 1
    ws.append(["Totals", sum(r[1] for r in rows), sum(r[2] for r in rows),
               sum(r[3] for r in rows), "", ""])
    for cell in ws[totals_row]:
        cell.fill = PINK
        cell.border = THIN
        cell.font = Font(name=TNR, bold=True)

    _fit_columns(ws, [30, 12, 14, 12, 22, 34])

    # Data font + top-5 highlight on total column
    for r in range(2, totals_row):
        for cell in ws[r]:
            cell.font = Font(name=TNR, size=11)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center")
    # top 5 total
    totals = sorted([(r, ws.cell(row=r, column=2).value) for r in range(2, totals_row)
                     if isinstance(ws.cell(row=r, column=2).value, int)],
                    key=lambda x: x[1], reverse=True)[:5]
    top_colors = ["FFFFB3B3","FFFFCC99","FFFFFFB3","FFCCFFB3","FFCCE5FF"]
    for i, (r, _) in enumerate(totals):
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=top_colors[i])

    ws.freeze_panes = "A2"


def _build_by_type(wb, comments: List[Comment]):
    ws = wb.create_sheet("By Type")
    ws.append(["Type", "Total", "Unique jurisdictions"])
    _style_header(ws)

    from collections import Counter, defaultdict
    counts = Counter(c.ctype for c in comments)
    jur_counts = defaultdict(set)
    for c in comments:
        jur_counts[c.ctype].add(c.jurisdiction)

    # Order: valid types first in their canonical order, then unknowns alphabetically
    order = VALID_TYPES + sorted(k for k in counts if k not in VALID_TYPES)
    for t in order:
        if t not in counts:
            continue
        ws.append([t, counts[t], len(jur_counts[t])])

    _fit_columns(ws, [40, 12, 22])
    for r in range(2, ws.max_row + 1):
        for cell in ws[r]:
            cell.font = Font(name=TNR, size=11)
            cell.border = THIN
    ws.freeze_panes = "A2"


def _build_by_country(wb, comments: List[Comment], title: str = "By Country"):
    ws = wb.create_sheet(title)

    headers = (["Country", "No. of comments", "No. of sections", "No. of Recs"]
               + [s for s in VALID_SECTIONS if s.startswith("IO.")]
               + ["Executive Summary", "Intro", "TC Annex", "Recs touched",
                  "Most common type", "In compilation?"])
    ws.append(headers)
    _style_header(ws)

    # Group by jurisdiction
    by_jur = {}
    for c in comments:
        by_jur.setdefault(c.jurisdiction or "Unknown", []).append(c)

    in_time = sorted(j for j, items in by_jur.items()
                     if any(i.on_time for i in items))
    late = sorted(j for j, items in by_jur.items()
                  if all(not i.on_time for i in items))

    def emit_country(jur):
        items = by_jur[jur]
        sections_touched = {i.section for i in items}
        recs_touched = sorted({i.rec_number for i in items if i.rec_number})
        # IO columns
        io_flags = []
        for s in VALID_SECTIONS:
            if s.startswith("IO."):
                io_flags.append(s if s in sections_touched else "")
        has_exec = "Executive Summary" in sections_touched
        has_intro = "Introduction to risk and context" in sections_touched
        has_tc = "TC Annex" in sections_touched
        recs_str = "/".join(f"R.{n}" for n in recs_touched)
        # Most common type (substantive)
        from collections import Counter
        types = [i.ctype for i in items if i.substantive]
        most = Counter(types).most_common(1)[0][0] if types else ""
        on_time = any(i.on_time for i in items)
        row = [
            jur,
            len(items),
            len(sections_touched),
            len(recs_touched),
            *io_flags,
            "Yes" if has_exec else "",
            "Yes" if has_intro else "",
            "Yes" if has_tc else "",
            recs_str,
            most,
            "Yes" if on_time else "No",
        ]
        ws.append(row)

    for jur in in_time:
        emit_country(jur)
    if late:
        ws.append([])
        for jur in late:
            emit_country(jur)

    # Style rows
    last_row = ws.max_row
    for r in range(2, last_row + 1):
        row_cells = ws[r]
        if not any(c.value for c in row_cells):
            continue
        for cell in row_cells:
            cell.font = Font(name=TNR, size=11)
            cell.border = THIN
        in_comp = ws.cell(row=r, column=ws.max_column)
        if in_comp.value == "No":
            in_comp.fill = RED
            ws.cell(row=r, column=1).font = Font(name=TNR, size=11, color="FF8B5A66")
        elif in_comp.value == "Yes":
            in_comp.fill = GREEN
        # Shade IO / section flag cells that are populated
        for col in range(5, 4 + 11 + 3 + 1):  # IO.1..IO.11 + 3 flags
            c = ws.cell(row=r, column=col)
            if c.value:
                c.fill = BLUE
                c.alignment = Alignment(horizontal="center")

    # Totals / averages (excluding late)
    non_late_rows = [r for r in range(2, last_row + 1)
                     if ws.cell(row=r, column=ws.max_column).value == "Yes"]
    if non_late_rows:
        ws.append([])
        tr = ws.max_row + 1
        ws.cell(row=tr, column=1, value="Totals (excl. late)")
        ws.cell(row=tr, column=2, value=sum(ws.cell(row=r, column=2).value or 0 for r in non_late_rows))
        ws.cell(row=tr, column=3, value=sum(ws.cell(row=r, column=3).value or 0 for r in non_late_rows))
        ws.cell(row=tr, column=4, value=sum(ws.cell(row=r, column=4).value or 0 for r in non_late_rows))
        for cell in ws[tr]:
            cell.fill = PINK
            cell.font = Font(name=TNR, bold=True)
            cell.border = THIN

        ar = tr + 1
        n = len(non_late_rows)
        ws.cell(row=ar, column=1, value="Average")
        ws.cell(row=ar, column=2, value=round(sum(ws.cell(row=r, column=2).value or 0 for r in non_late_rows) / n, 2))
        ws.cell(row=ar, column=3, value=round(sum(ws.cell(row=r, column=3).value or 0 for r in non_late_rows) / n, 2))
        ws.cell(row=ar, column=4, value=round(sum(ws.cell(row=r, column=4).value or 0 for r in non_late_rows) / n, 2))
        for cell in ws[ar]:
            cell.fill = YELLOW
            cell.border = THIN

        nr = ar + 2
        ws.cell(row=nr, column=1, value="Nb of delegations")
        ws.cell(row=nr, column=2, value=n)
        ws.cell(row=nr, column=1).fill = ORANGE
        ws.cell(row=nr, column=2).fill = ORANGE

    # Column widths
    widths = [25, 14, 14, 12] + [8]*11 + [14, 10, 12, 28, 28, 14]
    _fit_columns(ws, widths)
    ws.freeze_panes = "B2"


def _build_crosstab(wb, comments: List[Comment], dimension: str, title: str):
    """dimension='section' or 'type'."""
    ws = wb.create_sheet(title)

    # Rows = jurisdictions, Cols = dimension values
    jurs = sorted({c.jurisdiction for c in comments if c.jurisdiction})
    if dimension == "section":
        buckets_in_order = sorted({c.bucket for c in comments}, key=bucket_sort_key)
        cols = buckets_in_order
        key = lambda c: c.bucket
    else:
        cols = [t for t in VALID_TYPES if any(c.ctype == t for c in comments)]
        extras = sorted({c.ctype for c in comments if c.ctype not in VALID_TYPES})
        cols = cols + extras
        key = lambda c: c.ctype

    ws.append(["Country"] + cols + ["Total"])
    _style_header(ws)

    from collections import Counter
    totals_col = Counter()
    for jur in jurs:
        items = [c for c in comments if c.jurisdiction == jur]
        counts = Counter(key(c) for c in items)
        row = [jur]
        for col in cols:
            v = counts.get(col, 0)
            row.append(v if v else "")
            totals_col[col] += v
        row.append(sum(counts.values()))
        ws.append(row)

    # Totals row
    ws.append(["Totals"] + [totals_col[c] if totals_col[c] else "" for c in cols]
              + [sum(totals_col.values())])
    for cell in ws[ws.max_row]:
        cell.fill = PINK
        cell.font = Font(name=TNR, bold=True)
        cell.border = THIN

    # Styling
    for r in range(2, ws.max_row):
        for cell in ws[r]:
            cell.font = Font(name=TNR, size=11)
            cell.border = THIN
            if isinstance(cell.value, int) and cell.value > 0:
                cell.fill = BLUE
                cell.alignment = Alignment(horizontal="center")

    # Purple-highlight max per row (ignore Total col)
    for r in range(2, ws.max_row):
        max_val = 0
        max_col = None
        for col_idx in range(2, len(cols) + 2):
            v = ws.cell(row=r, column=col_idx).value
            if isinstance(v, int) and v > max_val:
                max_val = v
                max_col = col_idx
        if max_col is not None:
            ws.cell(row=r, column=max_col).fill = PURPLE
            ws.cell(row=r, column=max_col).font = Font(name=TNR, size=11, bold=True)

    widths = [25] + [14]*len(cols) + [10]
    _fit_columns(ws, widths)
    ws.freeze_panes = "B2"


def _build_no_comments(wb, comments: List[Comment]):
    ws = wb.create_sheet("No Comments")
    submitted = {c.jurisdiction for c in comments if c.jurisdiction}
    silent = sorted(set(FATF_MEMBERS) - submitted)
    ws.append(["FATF members with no comments submitted"])
    _style_header(ws)
    for m in silent:
        ws.append([m])
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1).font = Font(name=TNR, size=11)
    _fit_columns(ws, [40])


def build_excel(comments: List[Comment], assessed_country: str,
                diagnostics=None) -> BytesIO:
    """diagnostics: optional (report, issues) tuple – when supplied, the
    reconciliation and warning sheets are appended to the workbook."""
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    _build_instructions(wb, assessed_country)
    _build_compiled_sheet(wb, comments)
    _build_by_section(wb, comments, "By Section")
    _build_by_type(wb, comments)
    _build_by_country(wb, comments, "By Country")
    _build_crosstab(wb, comments, "section", "Country × Section")
    _build_crosstab(wb, comments, "type", "Country × Type")
    _build_no_comments(wb, comments)

    # Excluding assessed country
    excl = [c for c in comments
            if (c.jurisdiction or "").lower() != (assessed_country or "").lower()]
    if excl and excl != comments:
        _build_by_section(wb, excl, "By Section (excl. country)")
        _build_by_country(wb, excl, "By Country (excl. country)")

    if diagnostics:
        from diag_export import add_sheets
        report, issues = diagnostics
        add_sheets(wb, report, issues, assessed_country)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
