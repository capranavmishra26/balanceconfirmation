"""Formatting helpers: Indian number format and ordinal dates."""
from __future__ import annotations

import datetime as dt


def inr(amount: float) -> str:
    """1234567.5 -> '12,34,567.50' (absolute value)."""
    s = f"{abs(round(amount, 2)):.2f}"
    whole, dec = s.split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [tail])
    return f"{whole}.{dec}"


def ordinal_suffix(day: int) -> str:
    if 11 <= day % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def long_date(d: dt.date) -> str:
    """'30th September, 2026' (plain text)."""
    return f"{d.day}{ordinal_suffix(d.day)} {d:%B}, {d.year}"


def long_date_markup(d: dt.date) -> str:
    """Same as long_date but with a superscript suffix (ReportLab/HTML markup)."""
    return f"{d.day}<super>{ordinal_suffix(d.day)}</super> {d:%B}, {d.year}"


def long_date_html(d: dt.date) -> str:
    return f"{d.day}<sup>{ordinal_suffix(d.day)}</sup> {d:%B}, {d.year}"
