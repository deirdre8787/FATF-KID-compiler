# -*- coding: utf-8 -*-
"""Detect embedded images (and other non-text objects) in delegation files.

Images matter because the compiler carries **text only** into the compiled
Word document: any screenshot, pasted chart, drawn shape or embedded object
inside a delegation's comment cell is silently lost. Rather than let that
happen quietly, we detect it here and raise a warning naming the exact rows,
so the compiler can add them back by hand.

Returns, for both file types:
    {
      "count": int,                 # total images found in the file
      "in_comment_rows": [ {row, kind, reference}, ... ],
      "elsewhere": int,             # images outside the comment table
                                    # (logos, letterheads – usually harmless)
      "kinds": {"drawing": n, "pict": n, "object": n, ...},
    }
"""
import zipfile
from typing import Dict, List
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag):
    return f"{{{W}}}{tag}"


# Elements that render as a picture / embedded object inside a Word cell.
_PICTURE_TAGS = {
    _q("drawing"): "drawing",       # modern DrawingML picture / chart
    _q("pict"): "pict",             # legacy VML picture
    _q("object"): "object",         # embedded OLE object (Excel range, etc.)
}


def _cell_text(tc) -> str:
    return "".join(t.text or "" for t in tc.iter(_q("t")))


def probe_docx(path: str) -> Dict:
    """Find images in a .docx, flagging which comment rows contain them."""
    out = {"count": 0, "in_comment_rows": [], "elsewhere": 0, "kinds": {}}

    try:
        with zipfile.ZipFile(path) as z:
            media = [n for n in z.namelist()
                     if n.startswith("word/media/") and not n.endswith("/")]
            out["media_files"] = len(media)
            xml = z.read("word/document.xml")
    except Exception:
        return out

    try:
        tree = ET.fromstring(xml)
    except Exception:
        return out

    # count every picture element in the body
    all_pics = []
    for tag, label in _PICTURE_TAGS.items():
        found = tree.findall(f".//{tag}")
        if found:
            out["kinds"][label] = len(found)
        all_pics.extend((el, label) for el in found)
    out["count"] = len(all_pics)
    if not all_pics:
        return out

    pic_set = {id(el) for el, _ in all_pics}
    label_by_id = {id(el): lbl for el, lbl in all_pics}

    # walk the comment table exactly like the extractor does, so row numbers
    # in the warning line up with the row numbers the extractor reports
    accounted = set()
    for tbl in tree.findall(f".//{_q('tbl')}"):
        rows = tbl.findall(f".//{_q('tr')}")
        if not rows:
            continue
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

        def _row_cells(tr):
            return [tc for tc in tr.findall(f".//{_q('tc')}")]

        header = "".join(_cell_text(tc) for tc in _row_cells(rows[0])).lower()
        if "section of report" not in header:
            continue

        for row_no, tr in enumerate(rows[1:], start=1):
            cells = _row_cells(tr)
            if len(cells) < 4:
                continue
            ref = _cell_text(cells[2]).strip()
            for tag, label in _PICTURE_TAGS.items():
                for el in tr.iter(tag):
                    if id(el) in pic_set and id(el) not in accounted:
                        accounted.add(id(el))
                        out["in_comment_rows"].append({
                            "row": row_no,
                            "kind": label,
                            "reference": ref[:60],
                        })

    out["elsewhere"] = out["count"] - len(accounted)
    return out


def probe_pdf(path: str) -> Dict:
    """Find images in a PDF. Row attribution is not reliable in PDFs, so we
    report per page instead."""
    out = {"count": 0, "in_comment_rows": [], "elsewhere": 0,
           "kinds": {}, "pages": []}
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for pno, page in enumerate(pdf.pages, start=1):
                n = len(page.images or [])
                if n:
                    out["pages"].append({"page": pno, "images": n})
                    out["count"] += n
    except Exception:
        return out
    if out["count"]:
        out["kinds"]["image"] = out["count"]
        out["elsewhere"] = out["count"]
    return out


def probe(path: str, kind: str) -> Dict:
    return probe_pdf(path) if kind == "pdf" else probe_docx(path)


def describe(result: Dict, kind: str) -> str:
    """Human sentence for the warning list."""
    n = result.get("count", 0)
    if not n:
        return ""
    rows = result.get("in_comment_rows") or []
    if rows:
        where = ", ".join(
            f"row {r['row']}" + (f" ({r['reference']})" if r["reference"] else "")
            for r in rows[:6])
        more = f" and {len(rows) - 6} more" if len(rows) > 6 else ""
        return (f"{n} embedded image/object(s) found, {len(rows)} of them "
                f"inside comment cells ({where}{more}). Images are NOT "
                "carried into the compiled Word document – add them manually "
                "after compiling.")
    if kind == "pdf":
        pages = result.get("pages") or []
        where = ", ".join(f"p.{p['page']}" for p in pages[:8])
        return (f"{n} image(s) detected in this PDF ({where}). If any of them "
                "sit inside comment cells they will NOT appear in the "
                "compiled Word document.")
    return (f"{n} embedded image(s) found, none inside comment cells "
            "(likely a letterhead or logo) – nothing will be lost.")
