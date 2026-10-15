# -*- coding: utf-8 -*-
"""Bucket ordering for the Word + Excel summaries."""
import re


def bucket_sort_key(bucket: str):
    b = (bucket or "").strip()
    if b == "Executive Summary":
        return (0, 0, "")
    if b == "KRA Roadmap":
        return (1, 0, "")
    if b == "Introduction":
        return (2, 0, "")
    m = re.match(r"^Immediate Outcome\s+(\d+)$", b)
    if m:
        return (3, int(m.group(1)), "")
    m = re.match(r"^Recommendation\s+(\d+)$", b)
    if m:
        return (4, int(m.group(1)), "")
    if b == "TC Annex":
        return (5, 0, "")
    return (6, 0, b.lower())
