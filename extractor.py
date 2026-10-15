# -*- coding: utf-8 -*-
"""
Extracts structured comments from the Simplified 5th Round ME comment template.

Template has 4 columns:
  0. Section of report  (Executive Summary, Introduction to risk and context,
                         IO.1..IO.11, TC Annex)
  1. Type               (General, Analysis, Rating, Key Findings,
                         Key Recommended Actions, Recommended Actions,
                         Non-substantive editorial changes)
  2. Reference          (free text: Para #, KF(a), KRA(a), RA(a), Table 1.2,
                         R.1, C.1.6, etc.)
  3. Comment            (rich text)

Rich text in the Comment + Reference cells is preserved as a list of paragraphs,
each a list of runs with formatting flags. This preserves bold, italic,
underline, color, super/subscript, line breaks, and bullets.
"""
import re
import zipfile
from xml.etree import ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Optional

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
ET.register_namespace("w", W)

PLACEHOLDERS = {
    "choose an item.", "choisissez un élément.", "choisissez un élement.",
    "insert reference here.", "insérez la référence ici.",
    "click or tap here to enter your comment", "entrez votre commentaire ici.",
    "click or tap here to enter your comment.",
    "click or tap here to enter your comment – entrez votre commentaire ici.",
}

VALID_SECTIONS = [
    "Executive Summary",
    "Introduction to risk and context",
    "IO.1", "IO.2", "IO.3", "IO.4", "IO.5", "IO.6",
    "IO.7", "IO.8", "IO.9", "IO.10", "IO.11",
    "TC Annex",
]

VALID_TYPES = [
    "General",
    "Analysis",
    "Rating -Weighting and conclusions",
    "Key Findings",
    "Key Recommended Actions",
    "Recommended Actions",
    "Non-substantive editorial changes",
]


@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None
    superscript: bool = False
    subscript: bool = False
    line_break: bool = False  # run is a <w:br/>


@dataclass
class Para:
    runs: List[Run] = field(default_factory=list)
    is_bullet: bool = False

    def plain(self) -> str:
        return "".join("\n" if r.line_break else r.text for r in self.runs)


@dataclass
class RichText:
    paragraphs: List[Para] = field(default_factory=list)

    def plain(self) -> str:
        return "\n".join(p.plain() for p in self.paragraphs).strip()

    def tidy(self) -> "RichText":
        """Drop empty leading/trailing paragraphs and collapse runs of blank
        paragraphs to at most one. Empty paragraphs are what make a compiled
        cell look 'spread out with newlines all through it'."""
        paras = []
        for p in self.paragraphs:
            if p.plain().strip():
                paras.append(p)
            elif paras and paras[-1].plain().strip():
                paras.append(p)          # keep a single separator blank
        while paras and not paras[-1].plain().strip():
            paras.pop()
        self.paragraphs = paras
        return self

    def is_placeholder(self) -> bool:
        t = self.plain().strip().lower()
        if not t:
            return True
        # Strip leading/trailing punctuation
        t = t.strip(" .–—-:")
        if t in {p.strip(" .–—-:") for p in PLACEHOLDERS}:
            return True
        # Generic "Click or tap..." catchall
        if "click or tap here" in t or "entrez votre commentaire" in t:
            return True
        if "choose an item" in t or "choisissez un élément" in t:
            return True
        if "insert reference here" in t or "insérez la référence" in t:
            return True
        return False

    def to_html(self) -> str:
        """Render as simple HTML for display/Excel-plain fallback (Excel writer
        will strip tags). Keeps formatting semantically."""
        import html as _html
        out = []
        for p in self.paragraphs:
            buf = []
            for r in p.runs:
                if r.line_break:
                    buf.append("<br/>")
                    continue
                t = _html.escape(r.text)
                if not t:
                    continue
                if r.bold:        t = f"<b>{t}</b>"
                if r.italic:      t = f"<i>{t}</i>"
                if r.underline:   t = f"<u>{t}</u>"
                if r.superscript: t = f"<sup>{t}</sup>"
                if r.subscript:   t = f"<sub>{t}</sub>"
                if r.color:       t = f'<span style="color:#{r.color}">{t}</span>'
                buf.append(t)
            para = "".join(buf).strip()
            if not para:
                continue
            if p.is_bullet:
                out.append(f"<li>{para}</li>")
            else:
                out.append(f"<p>{para}</p>")
        return "".join(out)


@dataclass
class Comment:
    section: str           # normalized Section of report
    ctype: str             # normalized Type
    substantive: bool      # True unless ctype == "Non-substantive editorial changes"
    reference: RichText
    comment: RichText
    # derived
    rec_number: Optional[int] = None   # for TC Annex -> Recommendation N
    jurisdiction: str = ""
    source_file: str = ""
    on_time: bool = True
    order_in_file: int = 0

    @property
    def category(self) -> str:
        return "Substantive" if self.substantive else "Editorial"

    @property
    def bucket(self) -> str:
        """Grouping key for the Word output + summaries.

        FATF convention: all 'Key Recommended Actions' comments are aggregated
        into a cross-IO 'KRA Roadmap' section, regardless of the IO they
        reference. TC Annex references are bucketed under 'Recommendation N'.
        """
        # KRA Roadmap aggregates all Key Recommended Actions across IOs
        if self.ctype == "Key Recommended Actions":
            return "KRA Roadmap"
        # Rec-number wins whenever we have one – covers TC Annex rows AND
        # rows whose Section dropdown was left empty but whose reference is
        # clearly a criteria/recommendation ref (R.5, C.1.5(d), c15.3, …).
        if self.rec_number is not None and self.section in ("TC Annex", "", "Unspecified", None):
            return f"Recommendation {self.rec_number}"
        if self.section == "Introduction to risk and context":
            return "Introduction"
        m = re.match(r"^IO\.(\d+)$", self.section or "")
        if m:
            return f"Immediate Outcome {int(m.group(1))}"
        return self.section or "Unspecified"

    @property
    def section_bucket(self) -> str:
        """The bucket this comment belongs to by its Section alone, ignoring
        the KRA Roadmap aggregation. Used to repeat Key Recommended Actions
        comments under their own Immediate Outcome in the Word compilation
        as well as in the KRA Roadmap."""
        if self.rec_number is not None and self.section in ("TC Annex", "", "Unspecified", None):
            return f"Recommendation {self.rec_number}"
        if self.section == "Introduction to risk and context":
            return "Introduction"
        m = re.match(r"^IO\.(\d+)$", self.section or "")
        if m:
            return f"Immediate Outcome {int(m.group(1))}"
        return self.section or "Unspecified"


# ---------------------------------------------------------------------------
# XML parsing helpers
# ---------------------------------------------------------------------------

def _q(tag):  # qualified tag
    return f"{{{W}}}{tag}"


def _parse_run(r_elem, num_fmt_map) -> List[Run]:
    """Convert a <w:r> element into one or more Run objects (line breaks
    split into separate Runs so we can preserve ordering)."""
    rPr = r_elem.find(_q("rPr"))
    bold = italic = underline = sup = sub = False
    color = None
    if rPr is not None:
        bold      = rPr.find(_q("b")) is not None
        italic    = rPr.find(_q("i")) is not None
        underline = rPr.find(_q("u")) is not None
        c = rPr.find(_q("color"))
        if c is not None:
            v = c.get(_q("val")) or c.get("val")
            if v and v.lower() not in ("auto", "000000"):
                color = v
        va = rPr.find(_q("vertAlign"))
        if va is not None:
            v = va.get(_q("val")) or va.get("val")
            if v == "superscript": sup = True
            if v == "subscript":   sub = True

    runs: List[Run] = []
    # Iterate children in order: <w:t>, <w:br/>, <w:tab/>
    for child in list(r_elem):
        tag = child.tag.split("}", 1)[-1]
        if tag == "t":
            if child.text:
                runs.append(Run(text=child.text, bold=bold, italic=italic,
                                underline=underline, color=color,
                                superscript=sup, subscript=sub))
        elif tag == "br":
            runs.append(Run(line_break=True))
        elif tag == "tab":
            runs.append(Run(text="\t"))
    return runs


def _parse_paragraph(p_elem) -> Para:
    """Convert a <w:p> into a Para, descending into <w:sdt> content."""
    para = Para()
    # Detect bullet/numbered list
    numPr = p_elem.find(f".//{_q('numPr')}")
    if numPr is not None:
        para.is_bullet = True

    # Walk all <w:r> in the paragraph, including runs nested in <w:sdt>
    for r in p_elem.iter(_q("r")):
        para.runs.extend(_parse_run(r, None))
    return para


def _cell_to_richtext(tc_elem) -> RichText:
    rt = RichText()
    for p in tc_elem.findall(f".//{_q('p')}"):
        # skip paragraphs inside nested tables? For cells we typically have only paragraphs
        para = _parse_paragraph(p)
        rt.paragraphs.append(para)
    return rt


def _cell_dropdown_value(tc_elem) -> Optional[str]:
    """Return the current selected value of a content-control dropdown in this
    cell, if one exists. Otherwise None."""
    for sdt in tc_elem.findall(f".//{_q('sdt')}"):
        # Current value is the text inside <w:sdtContent>
        content = sdt.find(_q("sdtContent"))
        if content is None:
            continue
        texts = []
        for t in content.findall(f".//{_q('t')}"):
            if t.text:
                texts.append(t.text)
        v = "".join(texts).strip()
        if v:
            return v
    return None


def _cell_plain_text(tc_elem) -> str:
    parts = []
    for t in tc_elem.findall(f".//{_q('t')}"):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------

def _normalize_section(raw: str) -> Optional[str]:
    if not raw:
        return None
    s = raw.strip()
    # strip French half after " - " or " – "
    s = re.split(r"\s+[–-]\s+", s)[0].strip()
    s_low = s.lower()

    # Direct dropdown matches
    for v in VALID_SECTIONS:
        if s_low == v.lower():
            return v

    # Fuzzy
    if "executive" in s_low and "summary" in s_low:
        return "Executive Summary"
    if "introduction" in s_low and "risk" in s_low:
        return "Introduction to risk and context"
    if "tc annex" in s_low or "technical compliance annex" in s_low or s_low == "annex":
        return "TC Annex"
    m = re.match(r"^io[.\s-]*(\d+)$", s_low)
    if m:
        n = int(m.group(1))
        if 1 <= n <= 11:
            return f"IO.{n}"
    # Delegations often type a Recommendation straight into the Section cell
    # ("Rec. 10", "R.10", "Recommendation 10", "c.15") instead of picking
    # "TC Annex". Treat those as TC Annex so they bucket to Recommendation N.
    m = re.match(r"^(?:r|rec|recommendation|c|criterion)[.\s]*(\d+)", s_low)
    if m:
        n = int(m.group(1))
        if 1 <= n <= 40:
            return "TC Annex"
    return None


def _normalize_type(raw: str) -> Optional[str]:
    if not raw:
        return None
    s = raw.strip().lower()
    s = re.split(r"\s+[–-]\s+", s)[0].strip()
    # map
    if "non-substantive" in s or "editorial" in s:
        return "Non-substantive editorial changes"
    if s.startswith("general") or s == "general":
        return "General"
    if "analysis" in s:
        return "Analysis"
    if "rating" in s or "weighting" in s or "conclusion" in s:
        return "Rating -Weighting and conclusions"
    if "key finding" in s:
        return "Key Findings"
    if "key recommended" in s:
        return "Key Recommended Actions"
    if "recommended action" in s:
        return "Recommended Actions"
    return None


REC_PATTERN = re.compile(r"\b(?:R|REC(?:OMMENDATION)?)\s*\.?\s*(\d+)\b", re.IGNORECASE)
# c.1.6 / C1.6 / C.15.3 / c2.1 – criterion N.x of Recommendation N.
# Also c.38 / C38 (a bare criterion group = Recommendation 38).
CRIT_PATTERN = re.compile(r"\bC\.?\s*(\d+)(?:\.\s*\d+)?(?:\s*\([a-z]\))?", re.IGNORECASE)


def _extract_rec_number(reference_plain: str) -> Optional[int]:
    """From a TC Annex Reference cell, pull the recommendation number.
    Handles 'R.1', 'R 1', 'Recommendation 10', 'C.1.6' (criterion of R.1)."""
    if not reference_plain:
        return None
    m = CRIT_PATTERN.search(reference_plain)
    if m:
        try:
            n = int(m.group(1))
            if 1 <= n <= 40:
                return n
        except ValueError:
            pass
    m = REC_PATTERN.search(reference_plain)
    if m:
        try:
            n = int(m.group(1))
            if 1 <= n <= 40:
                return n
        except ValueError:
            pass
    return None


# ---------------------------------------------------------------------------
# Main entry points
# ---------------------------------------------------------------------------

def _accept_jurisdiction(value: str) -> str:
    """Guard against a template whose delegation dropdown has been deleted.

    When the 'Comments from' label survives but its dropdown does not, the
    next content control in the document is often the comment table itself,
    so the 'jurisdiction' comes back as the whole table. Anything that is not
    a plausible, recognisable delegation name is rejected here and treated as
    'not declared', so the filename guess is used and the operator is asked to
    confirm it, instead of a wall of text landing in the Delegation column.
    """
    if not value:
        return ""
    v = re.sub(r"\s+", " ", value).strip(" :\t.\u00a0")
    if not v or len(v) > 60:
        return ""
    try:
        from countries import canonical, is_plausible
    except Exception:
        return v
    c = canonical(v)
    if c:
        return c
    return v if is_plausible(v) else ""


def extract_jurisdiction(docx_path: str) -> str:
    """Read the dropdown after 'Comments from - Commentaires de'.

    Strategy: walk the document XML in document order. Once we have seen
    the literal text 'Comments from', the next <w:sdt> dropdown's selected
    value is the jurisdiction. This is robust against templates that split
    the label into multiple runs or that place the dropdown immediately
    adjacent to the label without any whitespace.
    """
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml")
        tree = ET.fromstring(xml)
    except Exception:
        return ""

    # Pass 1: SDT-aware lookup
    saw_label = False
    label_buffer = ""
    candidate = None
    for el in tree.iter():
        tag = el.tag.split("}", 1)[-1]
        if tag == "t" and el.text:
            label_buffer += el.text
            if "Comments from" in label_buffer:
                saw_label = True
        if saw_label and tag == "sdt":
            content = el.find(_q("sdtContent"))
            if content is not None:
                texts = [t.text for t in content.findall(f".//{_q('t')}")
                         if t.text]
                v = "".join(texts).strip()
                low = v.lower()
                if v and low not in (
                    "choose an item",
                    "choose the name of your delegation",
                    "choisissez le nom de votre délégation",
                ):
                    candidate = v
                    break

    if candidate:
        candidate = re.split(r"\s+[–-]\s+", candidate)[0].strip(" :\t.")
        candidate = re.split(
            r"\b(?:Substantive|Editorial|Section of report|TEMPLATE|"
            r"Commentaires|Choose|Choisissez)\b",
            candidate,
            flags=re.IGNORECASE,
        )[0].strip(" :\t.")
        low = candidate.lower()
        bad = ("choose the name", "choose an item", "choisissez",
               "name of your delegation", "nom de votre délégation",
               "commentaires de")
        if candidate and not any(b in low for b in bad):
            accepted = _accept_jurisdiction(candidate)
            if accepted:
                return accepted

    # Pass 2: legacy fallback over concatenated text
    all_text = []
    for el in tree.iter():
        tag = el.tag.split("}", 1)[-1]
        if tag == "t" and el.text:
            all_text.append(el.text)
    joined = "".join(all_text)

    m = re.search(r"Comments from[^A-Za-z]*([A-Z][^\n\.\|]+)", joined)
    if m:
        jur = m.group(1).strip()
        jur = re.sub(r"^Commentaires\s+de\s+", "", jur, flags=re.IGNORECASE).strip()
        jur = re.split(r"\s+[–-]\s+", jur)[0].strip()
        jur = jur.strip(" :\t.")
        jur = re.split(r"\bCommentaires\b", jur, flags=re.IGNORECASE)[0].strip(" :\t.")
        jur = re.split(r"\b(?:Substantive|Editorial|Section of report|TEMPLATE)\b",
                       jur)[0].strip(" :\t.")
        low = jur.lower()
        bad = ("choose the name", "choose an item", "choisissez",
               "name of your delegation", "nom de votre délégation")
        if jur and not any(b in low for b in bad):
            accepted = _accept_jurisdiction(jur)
            if accepted:
                return accepted
    return ""


_LATE_RE = re.compile(r"(?<![A-Za-z])late(?![A-Za-z])", re.IGNORECASE)


def is_late_filename(filename: str) -> bool:
    """True if filename contains the token 'late' (not preceded/followed by a letter)."""
    if not filename:
        return False
    base = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    return bool(_LATE_RE.search(base))


def _guess_jurisdiction_from_filename(filename: str) -> str:
    """Fallback: scan filename for a FATF member country name."""
    try:
        import pycountry
        names = [c.name for c in pycountry.countries]
    except Exception:
        names = [
            "Argentina","Australia","Austria","Belgium","Brazil","Canada","China",
            "Denmark","Finland","France","Germany","Greece","Hong Kong","Iceland",
            "India","Indonesia","Ireland","Israel","Italy","Japan","Korea",
            "Luxembourg","Malaysia","Mexico","Netherlands","New Zealand","Norway",
            "Portugal","Russia","Saudi Arabia","Singapore","South Africa","Spain",
            "Sweden","Switzerland","Turkey","Türkiye","United Kingdom",
            "United States",
        ]
    low = filename.lower()
    best = ""
    for n in names:
        if n.lower() in low and len(n) > len(best):
            best = n
    return best


def _looks_like_data_table(rows, row_cells_fn) -> bool:
    """True when a header-less table is really a continuation of the comment
    table: 4+ columns, and at least one row whose first cell is a recognisable
    Section value (IO.3, TC Annex, Executive Summary…) or whose second cell is
    a recognisable Type, with a non-trivial comment in the last cell."""
    hits = 0
    for tr in rows:
        cells = row_cells_fn(tr)
        if len(cells) < 4:
            return False
        sec = (_cell_dropdown_value(cells[0]) or _cell_plain_text(cells[0]) or "").strip()
        typ = (_cell_dropdown_value(cells[1]) or _cell_plain_text(cells[1]) or "").strip()
        com = _cell_plain_text(cells[3]) or ""
        if not com.strip() or len(com.strip()) < 15:
            continue
        if _normalize_section(sec) or _normalize_type(typ):
            hits += 1
    return hits > 0


def extract_comments(docx_path: str, filename_hint: str = "",
                     stats: Optional[dict] = None) -> List[Comment]:
    """Parse one filled template and return Comment records.

    If `stats` is a dict it is populated with reconciliation data:
        total_rows  – data rows found in the comment table(s)
        extracted   – rows turned into Comment records
        skipped     – [{"row": n, "reason": str, "reference": str}, ...]
    so the caller can prove every source row is accounted for.
    """
    if stats is not None:
        stats.setdefault("total_rows", 0)
        stats.setdefault("extracted", 0)
        stats.setdefault("skipped", [])

    def _skip(n, reason, ref=""):
        if stats is not None:
            stats["skipped"].append({"row": n, "reason": reason,
                                     "reference": ref[:60]})

    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    tree = ET.fromstring(xml)

    jur = extract_jurisdiction(docx_path)
    if not jur:
        jur = _guess_jurisdiction_from_filename(filename_hint or docx_path)

    results: List[Comment] = []
    order = 0

    for tbl in tree.findall(f".//{_q('tbl')}"):
        # Use a deep search so we also find rows wrapped inside <w:sdt>
        # repeating-section content controls (this is how the new template's
        # "add new box" button creates rows). We exclude rows that belong to
        # any nested table within this table's cells.
        all_rows = tbl.findall(f".//{_q('tr')}")
        # Filter to rows whose nearest <w:tbl> ancestor is THIS tbl
        rows = []
        for tr in all_rows:
            anc = tr.getparent() if hasattr(tr, "getparent") else None
            # ElementTree does not expose getparent; instead walk the tree once
            # to identify ancestor tables. Build a parent map lazily.
            rows.append(tr)
        # Build parent map to filter nested-table rows
        parent_map = {c: p for p in tbl.iter() for c in p}
        def _is_nested(tr):
            cur = parent_map.get(tr)
            while cur is not None:
                if cur is tbl:
                    return False
                if cur.tag == _q("tbl"):
                    return True
                cur = parent_map.get(cur)
            return False
        rows = [tr for tr in rows if not _is_nested(tr)]
        if not rows:
            continue

        # Cells may also be wrapped in <w:sdt>, so use a deep search and then
        # filter to those whose nearest <w:tr> ancestor is this row.
        def _row_cells(tr):
            all_tc = tr.findall(f".//{_q('tc')}")
            row_parent_map = {c: p for p in tr.iter() for c in p}
            def _own_cell(tc):
                cur = row_parent_map.get(tc)
                while cur is not None:
                    if cur is tr:
                        return True
                    if cur.tag == _q("tr"):
                        return False
                    cur = row_parent_map.get(cur)
                return True
            return [tc for tc in all_tc if _own_cell(tc)]

        # Identify the data table. Normally the header row says "Section of
        # report". But templates that have been converted, re-saved or edited
        # frequently end up with the header row in ONE table and the data rows
        # in a SEPARATE table immediately after it (seen in the wild). If we
        # only accepted header-bearing tables those submissions would silently
        # compile as zero comments, so a headerless 4-column table whose rows
        # look like real comment rows is salvaged instead.
        header_text = "".join(_cell_plain_text(tc) for tc in _row_cells(rows[0])).lower()
        has_header = "section of report" in header_text

        if has_header:
            data_rows = rows[1:]
        else:
            if not _looks_like_data_table(rows, _row_cells):
                continue
            data_rows = rows                      # every row is data
            if stats is not None:
                stats.setdefault("salvaged_tables", 0)
                stats["salvaged_tables"] += 1
                stats.setdefault("salvaged_rows", 0)
                stats["salvaged_rows"] += len(rows)

        for row_no, row in enumerate(data_rows, start=1):
            cells = _row_cells(row)
            if stats is not None:
                stats["total_rows"] += 1
            if len(cells) < 4:
                _skip(row_no, "row has fewer than 4 cells (malformed table)")
                continue

            sec_raw  = _cell_dropdown_value(cells[0]) or _cell_plain_text(cells[0])
            type_raw = _cell_dropdown_value(cells[1]) or _cell_plain_text(cells[1])
            ref_rt   = _cell_to_richtext(cells[2]).tidy()
            com_rt   = _cell_to_richtext(cells[3]).tidy()

            # Skip placeholder/empty rows
            if com_rt.is_placeholder() and ref_rt.is_placeholder():
                _skip(row_no, "empty row (both reference and comment are "
                              "blank or still the template placeholder)")
                continue
            if com_rt.is_placeholder():
                _skip(row_no, "comment cell is blank or still the template "
                              "placeholder", ref_rt.plain())
                continue
            # If only the reference is placeholder/empty, blank it out so the
            # placeholder text does not leak into outputs.
            if ref_rt.is_placeholder():
                ref_rt = RichText()

            section = _normalize_section(sec_raw)
            ctype = _normalize_type(type_raw)
            # SAFETY NET: a row that carries real comment text is NEVER
            # dropped, even when both dropdowns are unreadable. Losing a
            # delegation's comment silently is the worst outcome this tool
            # can produce; an unplaceable comment is compiled under
            # "Unspecified" and raised as a manual check instead.
            if section is None and ctype is None:
                if len(com_rt.plain().strip()) >= 15:
                    if stats is not None:
                        stats.setdefault("rescued", []).append({
                            "row": row_no,
                            "reference": ref_rt.plain()[:60],
                            "section_cell": sec_raw[:40],
                            "type_cell": type_raw[:40],
                        })
                else:
                    _skip(row_no,
                          "neither 'Section of report' nor 'Type' could be read "
                          f"(section cell: {sec_raw[:25]!r}, type cell: "
                          f"{type_raw[:25]!r}) and the comment is empty",
                          ref_rt.plain())
                    continue

            substantive = (ctype != "Non-substantive editorial changes")

            rec_num = None
            # A Recommendation number typed into the Section cell
            # ("Rec. 10") is as authoritative as one in the Reference cell.
            if section == "TC Annex":
                rec_num = _extract_rec_number(sec_raw or "")
            if rec_num is None and (section == "TC Annex" or section is None):
                # criteria/rec refs (R.5, C.1.6, c2.1 …) identify the bucket
                # even when the Section dropdown was left empty
                rec_num = _extract_rec_number(ref_rt.plain())

            order += 1
            results.append(Comment(
                section=section or "Unspecified",
                ctype=ctype or "General",
                substantive=substantive,
                reference=ref_rt,
                comment=com_rt,
                rec_number=rec_num,
                jurisdiction=jur,
                source_file=filename_hint or docx_path,
                order_in_file=order,
            ))
            if stats is not None:
                stats["extracted"] += 1

    return results
