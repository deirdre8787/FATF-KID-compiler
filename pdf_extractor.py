# -*- coding: utf-8 -*-
"""PDF fallback ingestion.

Delegations occasionally return the comment template printed to PDF instead
of the .docx. This module extracts the same Comment records from such PDFs.

The hard part is that a PDF has no paragraphs – only lines that were
hard-wrapped by the printer at the column edge. Feeding those lines straight
into Word produces a cell where every visual line is its own paragraph, which
renders ragged-left with big gaps ("weird formatting, newlines all
throughout"). `_unwrap()` rebuilds real paragraphs from wrapped lines.

Limitations (surfaced to the user as a warning):
  * rich text (bold/italic/colour) is not recoverable from a PDF;
  * dropdown values appear as plain text, so they are normalised like free
    text;
  * tables split across pages are stitched back together.
"""
import re
from typing import List, Optional

from extractor import (
    Comment, RichText, Para, Run,
    _normalize_section, _normalize_type, _extract_rec_number,
)

_PLACEHOLDER = re.compile(
    r"choose an item|click or tap here|click here to enter|"
    r"choisissez un élément|entrez votre commentaire", re.IGNORECASE)

# A line that starts a new block rather than continuing the previous one.
_BULLET_START = re.compile(
    r"^\s*(?:[••‣◦⁃∙▪◦‣·]|[-–—]\s|"
    r"\(?[ivxlIVXL]+[\).]\s|\(?[a-zA-Z][\).]\s|\d+[\).]\s)")

# Header fragments of the template itself – never comment content.
_HEADER_WORDS = ("section of", "du rapport", "référence", "commentaire",
                 "reference -", "comment -", "type")


def _is_header_cell(text: str) -> bool:
    low = re.sub(r"\s+", " ", (text or "")).strip().lower()
    if not low:
        return False
    return sum(1 for h in _HEADER_WORDS if h in low) >= 2


def _unwrap(text: str) -> List[str]:
    """Turn hard-wrapped PDF lines back into paragraphs.

    A line continues the previous one unless it clearly starts a new block:
      * it begins with a bullet / numbering marker, or
      * the previous line ended a sentence AND was visibly short (i.e. it was
        the last line of a paragraph rather than a mid-sentence wrap).
    """
    raw = [l.rstrip() for l in (text or "").replace("\r", "").split("\n")]
    lines = [l for l in raw if l.strip()]
    if not lines:
        return []

    width = max(len(l) for l in lines)
    short = width * 0.80          # a full line is ~width; shorter = para end

    paras: List[str] = []
    cur = ""
    prev_line = ""
    for line in lines:
        st = line.strip()
        if not cur:
            cur, prev_line = st, st
            continue
        # judge the break on the PREVIOUS PHYSICAL LINE, not on the paragraph
        # built so far – an accumulated paragraph is never "short".
        prev_ends_sentence = bool(re.search(r"[.!?:;]['\"”’)\]]?$", prev_line))
        prev_is_short = len(prev_line) < short
        if _BULLET_START.match(line) or (prev_ends_sentence and prev_is_short):
            paras.append(cur)
            cur, prev_line = st, st
            continue
        prev_line = st
        if True:
            # mid-sentence wrap: join. Repair words split across the break.
            if cur.endswith("-") and not cur.endswith((" -", "--")):
                cur = cur[:-1] + st
            else:
                cur = cur + " " + st
    if cur:
        paras.append(cur)
    return paras


def _rt(text: str) -> RichText:
    """Plain-text RichText with PDF line-wrapping undone."""
    paras = []
    for p in _unwrap(text):
        is_bullet = bool(_BULLET_START.match(p))
        if is_bullet:
            p = _BULLET_START.sub("", p, count=1).strip()
        paras.append(Para(runs=[Run(text=p)], is_bullet=is_bullet))
    if not paras:
        paras = [Para(runs=[Run(text="")])]
    return RichText(paragraphs=paras)


def _flat(cell: Optional[str]) -> str:
    """Single-line value for Section / Type cells (never multi-paragraph)."""
    if not cell:
        return ""
    s = re.sub(r"\s+", " ", cell.replace("\r", "")).strip()
    if _PLACEHOLDER.search(s):
        return ""
    return s


def _clean(cell: Optional[str]) -> str:
    if not cell:
        return ""
    s = cell.replace("\r", "").strip()
    if _PLACEHOLDER.search(s):
        return ""
    return s


def _squeeze(row: List[Optional[str]]) -> List[str]:
    """Drop spacer columns pdfplumber invents, then normalise to 4 columns.

    Real rows are Section | Type | Reference | Comment. PDF grids often add
    empty separator columns, and the comment can be split across several.
    """
    cells = [(c or "").strip() for c in row]
    kept = [c for c in cells if c != ""]
    if len(kept) <= 1 and not any(cells):
        return []
    # If the row already has exactly 4 non-spacer columns, use them as-is.
    non_empty_idx = [i for i, c in enumerate(cells) if c != ""]
    if len(cells) == 4:
        return cells
    if len(kept) == 4:
        return kept
    if len(kept) < 4:
        # A continuation row (only comment text present, in the last column)
        out = ["", "", "", ""]
        if non_empty_idx:
            last = non_empty_idx[-1]
            text = cells[last]
            if last >= len(cells) - 2:
                out[3] = text
            else:
                out[min(last, 3)] = text
            for i in non_empty_idx[:-1]:
                out[min(i, 2)] = cells[i]
        return out
    # More than 4 populated columns: first three are section/type/reference,
    # everything after them belongs to the comment.
    return kept[:3] + [" ".join(kept[3:])]


def extract_jurisdiction_pdf(pdf) -> str:
    """Find the 'Comments from / Commentaires de' value."""
    from countries import canonical, is_plausible

    # Strictest patterns first. 'de(?![a-zà-ÿ])' stops 'commentaires des
    # délégations …' from being read as a country called 's délégations'.
    pats = [
        r"comments?\s+from\s*[-–—]?\s*commentaires?\s+de(?![a-zà-ÿ])\s*[:\-–—]?\s*(.+)",
        r"comments?\s+from\s*[:\-–—]?\s*(.+)",
        r"commentaires?\s+de(?![a-zà-ÿ])\s*[:\-–—]?\s*(.+)",
        r"delegation\s*[:\-–—]\s*(.+)",
    ]
    fallback = ""
    for page in pdf.pages[:4]:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            for pat in pats:
                m = re.search(pat, line, re.IGNORECASE)
                if not m:
                    continue
                cand = re.split(r"\s{2,}", m.group(1).strip())[0].strip(" :.-–—")
                if not cand or _PLACEHOLDER.search(cand):
                    continue
                resolved = canonical(cand)
                if resolved:                       # a real jurisdiction: done
                    return resolved
                if not fallback and is_plausible(cand):
                    fallback = cand
    return fallback


def extract_comments_pdf(pdf_path: str, filename_hint: str = "",
                         stats: Optional[dict] = None) -> List[Comment]:
    import pdfplumber

    if stats is not None:
        stats.setdefault("total_rows", 0)
        stats.setdefault("extracted", 0)
        stats.setdefault("skipped", [])

    comments: List[Comment] = []

    with pdfplumber.open(pdf_path) as pdf:
        jur = extract_jurisdiction_pdf(pdf)

        rows: List[List[str]] = []
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table or []:
                    cells = _squeeze(row)
                    if not cells or len(cells) < 4:
                        continue
                    rows.append([_clean(c) for c in cells])

        # stitch continuation rows (page breaks split a comment in two)
        stitched: List[List[str]] = []
        for sec, typ, ref, com in rows:
            if _is_header_cell(sec) or _is_header_cell(com) or \
               (_is_header_cell(ref) and not com):
                continue
            is_continuation = (not sec and not typ and (com or ref))
            if stitched and is_continuation:
                p = stitched[-1]
                if ref:
                    p[2] = (p[2] + "\n" + ref).strip()
                if com:
                    p[3] = (p[3] + "\n" + com).strip()
                continue
            stitched.append([sec, typ, ref, com])

        order = 0
        for row_no, (sec_raw, typ_raw, ref_raw, com_raw) in enumerate(stitched, start=1):
            if stats is not None:
                stats["total_rows"] += 1

            sec_flat, typ_flat = _flat(sec_raw), _flat(typ_raw)
            if not com_raw.strip():
                if stats is not None:
                    stats["skipped"].append({
                        "row": row_no,
                        "reason": "comment cell is blank or a placeholder",
                        "reference": _flat(ref_raw)[:60]})
                continue

            section = _normalize_section(sec_flat)
            ctype = _normalize_type(typ_flat)
            if section is None and ctype is None and not sec_flat and not typ_flat:
                if stats is not None:
                    stats["skipped"].append({
                        "row": row_no,
                        "reason": "no Section or Type value could be read "
                                  "(likely a stray table fragment)",
                        "reference": _flat(ref_raw)[:60]})
                continue

            ctype_final = ctype or typ_flat
            order += 1
            c = Comment(
                section=section or "Unspecified",
                ctype=ctype_final or "General",
                substantive=(ctype_final != "Non-substantive editorial changes"),
                reference=_rt(ref_raw).tidy(),
                comment=_rt(com_raw).tidy(),
                jurisdiction=jur,
                source_file=filename_hint or pdf_path,
                on_time=True,
                order_in_file=order,
            )
            if section == "TC Annex" or section is None:
                rec = _extract_rec_number(re.sub(r"\s+", " ", ref_raw or ""))
                if rec:
                    c.rec_number = rec
            comments.append(c)
            if stats is not None:
                stats["extracted"] += 1

    return comments
