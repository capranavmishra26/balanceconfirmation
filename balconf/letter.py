"""Generate the balance confirmation letter as a PDF (ReportLab)."""
from __future__ import annotations

import datetime as dt
import os
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from .fmt import inr, long_date_markup

FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
FONT_CANDIDATES = [("calibri.ttf", "calibrib.ttf"), ("arial.ttf", "arialbd.ttf"), ("segoeui.ttf", "segoeuib.ttf")]
RUPEE = "\u20b9"


@lru_cache(maxsize=1)
def _fonts() -> tuple[str, str, str]:
    """Register a Windows TrueType font that has the rupee sign.
    Returns (regular, bold, currency_prefix)."""
    fallback = None
    for reg, bold in FONT_CANDIDATES:
        rp, bp = FONT_DIR / reg, FONT_DIR / bold
        if not (rp.exists() and bp.exists()):
            continue
        try:
            pdfmetrics.registerFont(TTFont("BC-Regular", str(rp)))
            pdfmetrics.registerFont(TTFont("BC-Bold", str(bp)))
            pdfmetrics.registerFontFamily("BC", normal="BC-Regular", bold="BC-Bold",
                                          italic="BC-Regular", boldItalic="BC-Bold")
        except Exception:
            continue
        face = pdfmetrics.getFont("BC-Regular").face
        if ord(RUPEE) in face.charToGlyph:
            return "BC-Regular", "BC-Bold", RUPEE + " "
        fallback = fallback or ("BC-Regular", "BC-Bold", "Rs. ")
    return fallback or ("Helvetica", "Helvetica-Bold", "Rs. ")


def balance_phrase(net: float, cfg: dict) -> str:
    if net < 0:
        return cfg["phrase_credit"]
    if net > 0:
        return cfg["phrase_debit"]
    return cfg["phrase_nil"]


def build_letter_pdf(
    out_path: Path,
    *,
    vendor_code: str,
    vendor_name: str,
    net: float,
    as_at: dt.date,
    letter_date: dt.date,
    cfg: dict,
) -> Path:
    regular, bold, cur = _fonts()
    base = ParagraphStyle("base", fontName=regular, fontSize=10.5, leading=14)
    body = ParagraphStyle("body", parent=base, alignment=TA_JUSTIFY)
    bold_st = ParagraphStyle("bold", parent=base, fontName=bold)

    e = escape
    as_at_m = long_date_markup(as_at)
    amount = f"{cur}{inr(net)}"
    company = e(cfg["company_name"])

    story = [
        Paragraph(f"<b>Date:</b> {letter_date:%d-%m-%Y}", base),
        Spacer(1, 10),
        Paragraph("To,", base),
        Spacer(1, 6),
        Paragraph(f"M/s {e(vendor_name)}", bold_st),
        Paragraph(f"Vendor Code: {e(vendor_code)}", base),
        Spacer(1, 14),
        Paragraph(f"<b><u>Subject: Confirmation of Outstanding Balance as at {as_at_m}</u></b>", base),
        Spacer(1, 14),
        Paragraph(f"<b>{e(cfg['salutation'])}</b>", base),
        Spacer(1, 8),
        Paragraph(
            "As part of our financial closing and reporting process, we request you to confirm the "
            f"balance appearing in our books of account in your name as at {as_at_m}", body),
        Spacer(1, 10),
        Paragraph(e(cfg["sign_off"]), base),
        Spacer(1, 40),
        Paragraph(f"<b>{company}</b>", base),
        Spacer(1, 4),
        HRFlowable(width="100%", thickness=0.7, color="black", spaceBefore=4, spaceAfter=14),
        Paragraph("To,", base),
    ]
    for line in [cfg["addressee"], cfg["company_name"], *cfg["address_lines"]]:
        story.append(Paragraph(e(line), base))
    story += [
        Spacer(1, 12),
        Paragraph(
            f"The balance {e(balance_phrase(net, cfg))} of {amount} as at {as_at_m} "
            "is correct except as noted below:", body),
        Spacer(1, 34),
        HRFlowable(width="85%", thickness=0.6, color="black", hAlign="LEFT", spaceAfter=30),
        HRFlowable(width="85%", thickness=0.6, color="black", hAlign="LEFT", spaceAfter=40),
    ]
    sig = Table([[Paragraph("Date:", base), Paragraph("Company Seal &amp; Signature", base)]],
                colWidths=["60%", "40%"])
    sig.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("ALIGN", (1, 0), (1, 0), "RIGHT")]))
    story.append(sig)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(out_path), pagesize=A4, leftMargin=25 * mm, rightMargin=25 * mm,
                            topMargin=20 * mm, bottomMargin=18 * mm,
                            title=f"Balance Confirmation - {vendor_name}", author=cfg["company_name"])
    doc.build(story)
    return out_path
