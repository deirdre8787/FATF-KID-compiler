# -*- coding: utf-8 -*-
"""Build the compiled Word document by injecting content into the bundled
FATF skeleton. Strips identifying metadata and applies uniform table styling.

Final-version requirements:

* Author = 'LUO Deirdre, FATF/RPU' (everything else stripped)
* app.xml: Company blank, Template = 'Normal.dotm',
           word/character/page counts zeroed
* custom.xml: emptied
* customXml/item2.xml: SharePoint Meeting/Date/Distribution metadata blanked
* customXml/item4.xml: OECD content type schema blanked
* customXml/item5.xml: OECD document cote (FATF/ME(2026)11) blanked

Table layout (fixed for every bucket):
    col 1 = 2.50 cm  (MER reference)
    col 2 = 18.80 cm (Comment)
    col 3 = 2.07 cm  (Delegation)

* Header row 'MER reference | Comment | Delegation' shaded light blue (#DEEBF7)
* 'Substantive issues' / 'Editorial issues' divider rows shaded light grey (#D9D9D9)
* Late comments shaded very light pastel pink (#FDEDF1)
* Cols 1 + 3 left-aligned, col 2 justified
"""
import os
import re
import zipfile
from datetime import datetime
from io import BytesIO
from typing import List, Optional

from extractor import Comment, RichText, Run, Para
from sorter import bucket_sort_key

SKELETON = os.path.join(os.path.dirname(__file__), "fatf_skeleton.docx")

# Column widths in twips (1 cm = 567 twips)
COL1_TWIPS = 1418   # 2.50 cm  - MER reference
COL2_TWIPS = 10660  # 18.80 cm - Comment
COL3_TWIPS = 1174   # 2.07 cm  - Delegation
TOTAL_TWIPS = COL1_TWIPS + COL2_TWIPS + COL3_TWIPS

# The KRA Roadmap table carries one extra column at the far left showing the
# Immediate Outcome each Key Recommended Action was filed against. The width
# is taken out of the Comment column so the table's overall width – and the
# page margins – stay identical to every other table in the document.
KRA_IO_TWIPS   = 1021                      # 1.80 cm - Immediate Outcome
KRA_COM_TWIPS  = COL2_TWIPS - KRA_IO_TWIPS  # 17.00 cm - Comment

# Shading
HEADER_BLUE = "DEEBF7"   # top heading row
DIVIDER_GREY = "D9D9D9"  # Substantive/Editorial divider rows
LATE_PINK = "FFD9E2"     # very light pastel pink for late comments

AUTHOR = "deirdre.luo@fatf-gafi.org"


# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _run_xml(run: Run) -> str:
    if run.line_break:
        return '<w:r><w:br/></w:r>'
    text = run.text or ""
    if not text:
        return ""
    rpr = []
    if run.bold:        rpr.append('<w:b/><w:bCs/>')
    if run.italic:      rpr.append('<w:i/><w:iCs/>')
    if run.underline:   rpr.append('<w:u w:val="single"/>')
    if run.color:       rpr.append(f'<w:color w:val="{_esc(run.color)}"/>')
    if run.superscript: rpr.append('<w:vertAlign w:val="superscript"/>')
    if run.subscript:   rpr.append('<w:vertAlign w:val="subscript"/>')
    rpr_xml = ("<w:rPr>" + "".join(rpr) + "</w:rPr>") if rpr else ""
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<w:r>{rpr_xml}<w:t{preserve}>{_esc(text)}</w:t></w:r>'


def _para_xml(para: Para, justify: str = "left") -> str:
    runs = "".join(_run_xml(r) for r in para.runs)
    if not runs:
        runs = '<w:r><w:t xml:space="preserve"></w:t></w:r>'
    # Uniform, compact spacing in every cell. Without this, paragraphs keep
    # whatever spacing the delegation's own template had, so some comments
    # render airy/ragged and others tight.
    ppr_parts = ['<w:spacing w:before="0" w:after="60" w:line="240"'
                 ' w:lineRule="auto"/>',
                 f'<w:jc w:val="{justify}"/>']
    if para.is_bullet:
        ppr_parts.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    return f'<w:p><w:pPr>{"".join(ppr_parts)}</w:pPr>{runs}</w:p>'


def _richtext_xml(rt: RichText, justify: str = "left") -> str:
    if not rt or not rt.paragraphs:
        return (f'<w:p><w:pPr><w:jc w:val="{justify}"/></w:pPr>'
                '<w:r><w:t xml:space="preserve"></w:t></w:r></w:p>')
    return "".join(_para_xml(p, justify) for p in rt.paragraphs)


def _plain_para(text: str, justify: str = "left", bold: bool = False) -> str:
    rpr = "<w:rPr><w:b/><w:bCs/></w:rPr>" if bold else ""
    return (f'<w:p><w:pPr><w:jc w:val="{justify}"/></w:pPr>'
            f'<w:r>{rpr}<w:t xml:space="preserve">{_esc(text)}</w:t></w:r></w:p>')


# ---- Cell / row builders --------------------------------------------------

def _tc(content_xml: str, width_twips: int, shade_hex: Optional[str] = None) -> str:
    shd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{shade_hex}"/>'
           if shade_hex else "")
    return (
        f'<w:tc><w:tcPr>'
        f'<w:tcW w:w="{width_twips}" w:type="dxa"/>'
        f'{shd}'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:left w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:right w:val="single" w:sz="4" w:color="auto"/>'
        f'</w:tcBorders>'
        f'<w:vAlign w:val="top"/>'
        f'</w:tcPr>{content_xml}</w:tc>'
    )


def _io_label(c: Comment) -> str:
    """Short Immediate Outcome label for the KRA Roadmap's extra column."""
    m = re.match(r"^IO\.(\d+)$", (c.section or "").strip())
    if m:
        return f"IO.{int(m.group(1))}"
    own = c.section_bucket or ""
    m = re.match(r"^Immediate Outcome (\d+)$", own)
    if m:
        return f"IO.{int(m.group(1))}"
    return own or (c.section or "")


def _io_sort_num(c: Comment) -> int:
    """Immediate Outcome number for sorting the KRA Roadmap (IO.1 → IO.11).
    Rows with no IO sort last."""
    m = re.match(r"^IO\.(\d+)$", (c.section or "").strip())
    if m:
        return int(m.group(1))
    m = re.match(r"^Immediate Outcome (\d+)$", c.section_bucket or "")
    if m:
        return int(m.group(1))
    return 999


def _header_row(with_io: bool = False) -> str:
    cells = ""
    if with_io:
        cells += _tc(_plain_para("IO", "left", bold=True),
                     KRA_IO_TWIPS, HEADER_BLUE)
    cells += (
        _tc(_plain_para("MER reference", "left", bold=True),
            COL1_TWIPS, HEADER_BLUE) +
        _tc(_plain_para("Comment", "center", bold=True),
            KRA_COM_TWIPS if with_io else COL2_TWIPS, HEADER_BLUE) +
        _tc(_plain_para("Delegation", "left", bold=True),
            COL3_TWIPS, HEADER_BLUE)
    )
    return f'<w:tr><w:trPr><w:tblHeader/></w:trPr>{cells}</w:tr>'


def _divider_row(label: str, with_io: bool = False) -> str:
    cell = (
        f'<w:tc><w:tcPr>'
        f'<w:tcW w:w="{TOTAL_TWIPS}" w:type="dxa"/>'
        f'<w:gridSpan w:val="{4 if with_io else 3}"/>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{DIVIDER_GREY}"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:left w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:bottom w:val="single" w:sz="4" w:color="auto"/>'
        f'<w:right w:val="single" w:sz="4" w:color="auto"/>'
        f'</w:tcBorders>'
        f'</w:tcPr>'
        f'{_plain_para(label, "left", bold=True)}'
        f'</w:tc>'
    )
    return f'<w:tr>{cell}</w:tr>'


def _data_row(c: Comment, with_io: bool = False) -> str:
    shade = LATE_PINK if not c.on_time else None
    ref_xml = _richtext_xml(c.reference, justify="left")
    com_xml = _richtext_xml(c.comment, justify="both")
    jur_xml = _plain_para(c.jurisdiction or "", "left")
    out = '<w:tr>'
    if with_io:
        out += _tc(_plain_para(_io_label(c), "left"), KRA_IO_TWIPS, shade)
    out += (
        _tc(ref_xml, COL1_TWIPS, shade)
        + _tc(com_xml, KRA_COM_TWIPS if with_io else COL2_TWIPS, shade)
        + _tc(jur_xml, COL3_TWIPS, shade)
    )
    return out + '</w:tr>' 


def _table_open(with_io: bool = False) -> str:
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblStyle w:val="OECD"/>'
        f'<w:tblW w:w="{TOTAL_TWIPS}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"'
        ' w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
        '<w:tblGrid>'
        + (f'<w:gridCol w:w="{KRA_IO_TWIPS}"/>' if with_io else '')
        + f'<w:gridCol w:w="{COL1_TWIPS}"/>'
        f'<w:gridCol w:w="{KRA_COM_TWIPS if with_io else COL2_TWIPS}"/>'
        f'<w:gridCol w:w="{COL3_TWIPS}"/>'
        '</w:tblGrid>'
    )


def _table_close() -> str:
    return '</w:tbl>'


def _section_heading(title: str) -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="Title3"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{_esc(title)}</w:t></w:r></w:p>'
    )


def _empty_p() -> str:
    return '<w:p/>'


def _note_para(text: str) -> str:
    """Small italic grey note under a section heading."""
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="80"/>'
        '<w:jc w:val="left"/></w:pPr>'
        '<w:r><w:rPr><w:i/><w:iCs/><w:sz w:val="18"/><w:szCs w:val="18"/>'
        '<w:color w:val="595959"/></w:rPr>'
        f'<w:t xml:space="preserve">{_esc(text)}</w:t></w:r></w:p>'
    )


# ---------------------------------------------------------------------------
# Body composition
# ---------------------------------------------------------------------------

KRA_BUCKET = "KRA Roadmap"

KRA_NOTE = ("The comments below are Key Recommended Actions. Each one is also "
            "reproduced under the Immediate Outcome it relates to, so it "
            "appears twice in this compilation.")


def _group(comments: List[Comment]):
    buckets = {}
    for c in comments:
        buckets.setdefault(c.bucket, []).append(c)
        # A Key Recommended Actions comment belongs to the cross-IO KRA
        # Roadmap AND to the Immediate Outcome it was filed against, so it is
        # repeated there. The KRA Roadmap section carries a note saying so.
        if c.bucket == KRA_BUCKET:
            own = c.section_bucket
            if own and own != KRA_BUCKET and own != "Unspecified":
                buckets.setdefault(own, []).append(c)
    for bucket in sorted(buckets.keys(), key=bucket_sort_key):
        items = buckets[bucket]
        if bucket == KRA_BUCKET:
            # The KRA Roadmap runs across every IO, so order it by the IO in
            # its first column (IO.1 → IO.11), then by delegation. Anything
            # without an IO (e.g. a TC Annex reference) sorts to the end.
            items.sort(key=lambda c: (
                0 if c.substantive else 1,
                _io_sort_num(c),
                (c.jurisdiction or "").lower(),
                c.order_in_file,
            ))
        else:
            items.sort(key=lambda c: (
                0 if c.substantive else 1,
                (c.jurisdiction or "").lower(),
                c.order_in_file,
            ))
        subs = [c for c in items if c.substantive]
        eds = [c for c in items if not c.substantive]
        yield bucket, subs, eds


def _build_body_xml(comments: List[Comment]) -> str:
    out = []
    for bucket, subs, eds in _group(comments):
        out.append(_section_heading(bucket))
        if bucket == KRA_BUCKET:
            out.append(_note_para(KRA_NOTE))
        is_kra = (bucket == KRA_BUCKET)
        out.append(_table_open(is_kra))
        out.append(_header_row(is_kra))
        if subs:
            out.append(_divider_row("Substantive issues", is_kra))
            for c in subs:
                out.append(_data_row(c, is_kra))
        if eds:
            out.append(_divider_row("Editorial issues", is_kra))
            for c in eds:
                out.append(_data_row(c, is_kra))
        out.append(_table_close())
        out.append(_empty_p())
    return "".join(out)


# ---------------------------------------------------------------------------
# Metadata scrubbing
# ---------------------------------------------------------------------------

def _new_core_xml(country: str) -> bytes:
    title = f"COMPILATION OF DELEGATION COMMENTS – {country}" if country else "COMPILATION OF DELEGATION COMMENTS"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:title>{_esc(title)}</dc:title>'
        '<dc:subject></dc:subject>'
        f'<dc:creator>{_esc(AUTHOR)}</dc:creator>'
        '<cp:keywords></cp:keywords>'
        '<dc:description></dc:description>'
        f'<cp:lastModifiedBy>{_esc(AUTHOR)}</cp:lastModifiedBy>'
        '<cp:revision>1</cp:revision>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{_now_iso()}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{_now_iso()}</dcterms:modified>'
        '</cp:coreProperties>'
    ).encode("utf-8")


def _now_iso() -> str:
    """UTC timestamp of the moment the document is generated."""
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _new_app_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties '
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Template>Normal.dotm</Template>'
        '<TotalTime>0</TotalTime>'
        '<Pages>0</Pages>'
        '<Words>0</Words>'
        '<Characters>0</Characters>'
        '<Application>Microsoft Office Word</Application>'
        '<DocSecurity>0</DocSecurity>'
        '<Lines>0</Lines>'
        '<Paragraphs>0</Paragraphs>'
        '<ScaleCrop>false</ScaleCrop>'
        '<Company></Company>'
        '<LinksUpToDate>false</LinksUpToDate>'
        '<CharactersWithSpaces>0</CharactersWithSpaces>'
        '<SharedDoc>false</SharedDoc>'
        '<HyperlinksChanged>false</HyperlinksChanged>'
        '<AppVersion>16.0000</AppVersion>'
        '</Properties>'
    ).encode("utf-8")


def _empty_custom_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties '
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>'
    ).encode("utf-8")


def _blank_item2_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<p:properties '
        'xmlns:p="http://schemas.microsoft.com/office/2006/metadata/properties" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<documentManagement/>'
        '</p:properties>'
    ).encode("utf-8")


def _blank_item4_xml() -> bytes:
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>'


def _blank_item5_xml() -> bytes:
    # Original is UTF-16 LE with BOM; keep that encoding.
    xml = (
        '<?xml version="1.0" encoding="utf-16"?>'
        '<ref-cotes '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'Version="3.9" '
        'xmlns="http://www.oecd.org/oneauthor/2016/cotes"/>'
    )
    return xml.encode("utf-16")


# tolerate ')' or ']' as the closing char (skeleton field codes render
# inconsistently in raw <w:t> text), optional revision suffix
_COTE_RE = re.compile(r"FATF/ME\(\d+[\)\]]\d+(?:/[A-Z0-9]+)?")
_WT_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)")


def _scrub_cotes(s: str) -> str:
    """Remove document cotes ('FATF/ME(2026)11'), including cotes whose
    characters are split across multiple <w:t> runs.

    Approach: concatenate all <w:t> text of the part, find cote spans in the
    concatenated string, then map those spans back onto the individual runs
    and delete just those characters. XML structure is left untouched.
    """
    # collect runs: list of (match, start_in_concat, text)
    runs = []
    concat = []
    pos = 0
    for m in _WT_RE.finditer(s):
        text = m.group(2)
        runs.append((m, pos, text))
        concat.append(text)
        pos += len(text)
    full = "".join(concat)

    spans = [mm.span() for mm in _COTE_RE.finditer(full)]
    if not spans:
        return s
    # extend each span over immediately-surrounding brackets and stray
    # trailing digits (field-code renderings split cotes unpredictably)
    ext = []
    for a, b in spans:
        if a > 0 and full[a-1] == "[":
            a -= 1
        while b < len(full) and full[b] in "0123456789":
            b += 1
        if b < len(full) and full[b] == "]":
            b += 1
        ext.append((a, b))
    spans = ext

    def _in_span(i):
        return any(a <= i < b for a, b in spans)

    # rebuild each run's text minus scrubbed characters
    out = []
    last_end = 0
    for m, start, text in runs:
        new_text = "".join(
            ch for k, ch in enumerate(text) if not _in_span(start + k))
        out.append(s[last_end:m.start()])
        out.append(m.group(1) + new_text + m.group(3))
        last_end = m.end()
    out.append(s[last_end:])
    return "".join(out)
_DATE_RE = re.compile(
    r"\b\d{1,2}\s*[-–]\s*\d{1,2}\s+(January|February|March|April|May|June|July|"
    r"August|September|October|November|December)\s+20\d{2}\b"
    r"|"
    r"\b\d{1,2}\s+(January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\s+20\d{2}\b")


# Content controls (SDTs) whose text must be blanked on the cover page.
# Values are the <w:tag w:val="…"> identifiers used inside fatf_skeleton.docx.
_BLANK_SDT_TAGS = {
    "txtInfomationMeeting",   # "9-10 February 2026, Hotel Camino Real …"
    "txtContacts",            # names / phone numbers / emails
}


def _generation_date_str() -> str:
    """Human date of the day the document is generated, e.g. '15 July 2026'."""
    return datetime.now().strftime("%-d %B %Y") if os.name != "nt" else \
        datetime.now().strftime("%#d %B %Y")


def _blank_sdt_content(xml: str) -> str:
    """For every SDT whose <w:tag w:val="…"> is in _BLANK_SDT_TAGS, delete
    the visible text inside its <w:sdtContent> (paragraphs / formatting are
    preserved so the placeholder still renders as an empty content control,
    just like the skeleton's own already-blank 'Directorate', 'Committee'
    and 'Cancel / Replace' controls)."""
    # walk depth-first through <w:sdt>…</w:sdt> pairs
    stack = []
    ranges = []
    for m in re.finditer(r"<w:sdt>|</w:sdt>", xml):
        if m.group() == "<w:sdt>":
            stack.append(m.start())
        elif stack:
            start = stack.pop()
            ranges.append((start, m.end()))

    targets = []
    for a, b in ranges:
        seg = xml[a:b]
        tag_m = re.search(r'<w:tag w:val="([^"]+)"', seg)
        if tag_m and tag_m.group(1) in _BLANK_SDT_TAGS:
            targets.append((a, b))

    if not targets:
        return xml

    # process in reverse so earlier indices stay valid
    for a, b in sorted(targets, reverse=True):
        seg = xml[a:b]
        # blank every <w:t …>…</w:t> inside this SDT
        blanked = re.sub(
            r"(<w:t(?:\s[^>]*)?>)[^<]*(</w:t>)",
            r"\1\2",
            seg,
        )
        xml = xml[:a] + blanked + xml[b:]
    return xml


def _replace_country(data: bytes, country: str) -> bytes:
    """Replace the skeleton's country, hardcoded dates (-> generation date),
    document cote (-> removed) and blank the cover-page meeting/contact
    content controls, in any body/header/footer XML part."""
    try:
        s = data.decode("utf-8")
    except UnicodeDecodeError:
        return data
    if country:
        s = s.replace("SINGAPORE", country.upper())
        s = s.replace("Singapore", country)
    s = _blank_sdt_content(s)          # remove meeting/contacts text
    s = _DATE_RE.sub(_generation_date_str(), s)
    s = _scrub_cotes(s)
    return s.encode("utf-8")


# ---------------------------------------------------------------------------
# Main entrypoint
# ---------------------------------------------------------------------------

def build_word(comments: List[Comment], assessed_country: str = "") -> BytesIO:
    body_xml = _build_body_xml(comments)
    out_buf = BytesIO()

    with zipfile.ZipFile(SKELETON, "r") as zin:
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)

                if name == "word/document.xml":
                    # ORDER MATTERS. Scrub the skeleton shell FIRST, then
                    # inject the compiled body. The scrub does a global
                    # search/replace of the skeleton's own country name and
                    # of hardcoded dates; running it after injection rewrote
                    # delegation names and comment text (a comment from
                    # Singapore, or any comment mentioning the skeleton's
                    # country or a date, was silently overwritten with the
                    # assessed country). Delegation content must never be
                    # touched by the shell scrub.
                    data = _replace_country(data, assessed_country)
                    s = data.decode("utf-8", errors="replace")
                    # Inject body before the final sectPr (which lives as a
                    # direct child of <w:body>, after the last paragraph).
                    idx = s.rfind("<w:sectPr")
                    if idx > 0:
                        s = s[:idx] + body_xml + s[idx:]
                    data = s.encode("utf-8")

                elif name == "docProps/core.xml":
                    data = _new_core_xml(assessed_country or "")
                elif name == "docProps/app.xml":
                    data = _new_app_xml()
                elif name == "docProps/custom.xml":
                    data = _empty_custom_xml()
                elif name == "customXml/item2.xml":
                    data = _blank_item2_xml()
                elif name == "customXml/item4.xml":
                    data = _blank_item4_xml()
                elif name == "customXml/item5.xml":
                    data = _blank_item5_xml()
                elif name.startswith("word/header") or name.startswith("word/footer"):
                    data = _replace_country(data, assessed_country)

                zout.writestr(name, data)

    out_buf.seek(0)
    return out_buf
