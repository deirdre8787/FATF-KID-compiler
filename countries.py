# -*- coding: utf-8 -*-
"""Canonical jurisdiction names, aliases and the picker list.

One source of truth for: the UI dropdown, filename guessing, PDF/DOCX
jurisdiction detection and the validator. Stops 'Korea' / 'Republic of
Korea' / 'KOR' being treated as three delegations, and stops free text
from ever reaching the compilation.
"""
import re
import unicodedata
from typing import List, Optional

# FATF members + observers / FSRBs that submit comments. Order matters:
# this is what the dropdown pins at the top.
FATF_DELEGATIONS: List[str] = [
    "Argentina", "Australia", "Austria", "Belgium", "Brazil", "Canada",
    "China", "Denmark", "European Commission", "Finland", "France",
    "Germany", "Greece", "Gulf Co-operation Council", "Hong Kong, China",
    "Iceland", "India", "Indonesia", "Ireland", "Israel", "Italy", "Japan",
    "Korea", "Luxembourg", "Malaysia", "Mexico", "Netherlands",
    "New Zealand", "Norway", "Portugal", "Russian Federation",
    "Saudi Arabia", "Singapore", "South Africa", "Spain", "Sweden",
    "Switzerland", "Türkiye", "United Kingdom", "United States",
]

OBSERVERS: List[str] = [
    "APG", "CFATF", "EAG", "ESAAMLG", "GABAC", "GAFILAT", "GIABA",
    "MENAFATF", "MONEYVAL", "Egmont Group", "IMF", "World Bank", "UNODC",
    "OECD", "United Nations", "Interpol", "Europol", "Basel Committee",
    "European Central Bank", "Asian Development Bank",
]

# alias (normalised) -> canonical name
_ALIASES = {
    "usa": "United States", "us": "United States",
    "u s a": "United States", "u s": "United States",
    "united states of america": "United States",
    "america": "United States",
    "uk": "United Kingdom", "great britain": "United Kingdom",
    "britain": "United Kingdom", "england": "United Kingdom",
    "united kingdom of great britain and northern ireland": "United Kingdom",
    "korea republic of": "Korea", "republic of korea": "Korea",
    "south korea": "Korea", "korea rep": "Korea", "kor": "Korea",
    "rok": "Korea",
    "turkey": "Türkiye", "turkiye": "Türkiye", "türkiye": "Türkiye",
    "hong kong": "Hong Kong, China", "hongkong": "Hong Kong, China",
    "hong kong china": "Hong Kong, China", "hksar": "Hong Kong, China",
    "hong kong sar": "Hong Kong, China", "hong kong sar china": "Hong Kong, China",
    "russia": "Russian Federation", "russian federation": "Russian Federation",
    "gcc": "Gulf Co-operation Council",
    "gulf cooperation council": "Gulf Co-operation Council",
    "ec": "European Commission", "eu": "European Commission",
    "european union": "European Commission",
    "netherlands the": "Netherlands", "the netherlands": "Netherlands",
    "holland": "Netherlands",
    "china prc": "China", "peoples republic of china": "China",
    "people s republic of china": "China", "prc": "China",
    "india s": "India", "indias": "India",
    "new zealand aotearoa": "New Zealand",
    "czechia": "Czechia", "czech republic": "Czechia",
    "ivory coast": "Côte d'Ivoire",
    "cape verde": "Cabo Verde",
    "swaziland": "Eswatini",
    "burma": "Myanmar",
    "macao": "Macao, China", "macau": "Macao, China",
    "chinese taipei": "Chinese Taipei", "taiwan": "Chinese Taipei",
}

# Words that are never a jurisdiction, used to reject bad PDF matches.
_JUNK_WORDS = (
    "delegation", "compilation", "document", "template", "comment",
    "commentaire", "question", "issue", "please", "veuillez", "section",
    "rapport", "principle", "principe", "click", "choose", "item",
)


def _norm(s: str) -> str:
    """Lowercase, strip accents/punctuation, squeeze spaces."""
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip().lower()


def _all_world() -> List[str]:
    try:
        import pycountry
        names = []
        for c in pycountry.countries:
            names.append(getattr(c, "common_name", None) or c.name)
        return names
    except Exception:
        return []


def all_countries() -> List[str]:
    """Every selectable jurisdiction: FATF delegations first, then observers,
    then every other country alphabetically. No duplicates."""
    seen, out = set(), []
    for group in (FATF_DELEGATIONS, OBSERVERS):
        for n in group:
            if _norm(n) not in seen:
                seen.add(_norm(n)); out.append(n)
    rest = []
    for n in _all_world():
        canon = _ALIASES.get(_norm(n), n)
        if _norm(canon) not in seen:
            seen.add(_norm(canon)); rest.append(canon)
    return out + sorted(rest)


def grouped_choices():
    """[(group label, [names])] for an <optgroup> dropdown."""
    fatf = list(FATF_DELEGATIONS)
    obs = list(OBSERVERS)
    known = {_norm(x) for x in fatf + obs}
    rest = sorted({_ALIASES.get(_norm(n), n) for n in _all_world()
                   if _norm(_ALIASES.get(_norm(n), n)) not in known})
    return [("FATF delegations", fatf),
            ("Observers / FSRBs", obs),
            ("All other countries", rest)]


_CANON_INDEX = None


def _index():
    global _CANON_INDEX
    if _CANON_INDEX is None:
        idx = {}
        for n in all_countries():
            idx[_norm(n)] = n
        idx.update({k: v for k, v in _ALIASES.items()})
        _CANON_INDEX = idx
    return _CANON_INDEX


def canonical(name: str) -> Optional[str]:
    """Resolve any spelling to the canonical jurisdiction name, else None."""
    if not name:
        return None
    n = _norm(name)
    if not n:
        return None
    idx = _index()
    if n in idx:
        return idx[n]
    # bilingual label halves: 'Singapore Singapour', 'Korea Coree'
    parts = [p.strip() for p in re.split(r"\s+[-–—/]\s+|\s{2,}", name) if p.strip()]
    for p in parts:
        pn = _norm(p)
        if pn in idx:
            return idx[pn]
    # 'comments from X', 'X delegation'
    n2 = re.sub(r"^(comments?|commentaires?)\s+(from|de|des)\s+", "", n)
    n2 = re.sub(r"\s+(delegation|delegations)$", "", n2).strip()
    if n2 in idx:
        return idx[n2]
    return None


def is_plausible(name: str) -> bool:
    """Reject obviously-not-a-country strings (PDF text bleed)."""
    if not name or len(name) > 60:
        return False
    n = _norm(name)
    if not n or len(n) < 2:
        return False
    if any(w in n for w in _JUNK_WORDS):
        return False
    return len(n.split()) <= 6


def resolve(name: str) -> str:
    """Canonical name when recognised, else the cleaned original."""
    c = canonical(name)
    if c:
        return c
    return re.sub(r"\s+", " ", (name or "")).strip()


def guess_from_filename(filename: str) -> str:
    """Longest canonical country name appearing in the filename."""
    if not filename:
        return ""
    stem = re.sub(r"\.[A-Za-z0-9]+$", "", filename)
    stem_n = " " + _norm(stem.replace("_", " ").replace("-", " ")) + " "
    best = ""
    for name in all_countries():
        for variant in {name} | {k for k, v in _ALIASES.items() if v == name}:
            vn = _norm(variant)
            if not vn or len(vn) < 3:
                continue
            if f" {vn} " in stem_n and len(vn) > len(_norm(best)):
                best = name
    if best:
        return best
    # possessive forms: "Indias_comments" -> India
    m = re.findall(r"[A-Za-zÀ-ÿ]+", stem)
    for w in m:
        if len(w) > 3 and w.lower().endswith("s"):
            c = canonical(w[:-1])
            if c:
                return c
    return ""


def all_matches_in_filename(filename: str) -> List[str]:
    """Every canonical country named in a filename (longest-first, no
    overlaps). Used to spot the assessed country: it is the one that recurs
    across all the delegation filenames."""
    if not filename:
        return []
    stem = re.sub(r"\.[A-Za-z0-9]+$", "", filename)
    hay = " " + _norm(stem.replace("_", " ").replace("-", " ")) + " "
    found, taken = [], []
    cands = []
    for name in all_countries():
        for variant in {name} | {k for k, v in _ALIASES.items() if v == name}:
            vn = _norm(variant)
            if len(vn) >= 2:
                cands.append((vn, name))
    for vn, name in sorted(cands, key=lambda x: -len(x[0])):
        idx = hay.find(f" {vn} ")
        if idx < 0:
            continue
        span = (idx, idx + len(vn) + 2)
        if any(not (span[1] <= a or span[0] >= b) for a, b in taken):
            continue
        taken.append(span)
        if name not in found:
            found.append(name)
    return found
