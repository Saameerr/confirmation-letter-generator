"""PDF letter generation"""
import io
from functools import lru_cache
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from PIL import Image as PILImage
from reportlab.platypus import (Flowable, HRFlowable, KeepTogether, PageBreak,
                                Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

NAVY = colors.HexColor("#1F3864")
GREY = colors.HexColor("#404040")
LIGHT = colors.HexColor("#EAF0F8")
LINE = colors.HexColor("#808080")
FONT, BOLD = "Helvetica", "Helvetica-Bold"

CONTENT_W = 15.9 * cm - 12  # frame has 6pt padding each side


# ----------------------------------------------------------------- numbers
def money(x):
    """1234.5 -> '1,234.50'; zero -> '-'"""
    return "-" if abs(x) < 0.005 else f"{x:,.2f}"


def calc(rec):
    """All the arithmetic in one place.
    Sign convention: positive = we owe the party (payable), negative = party owes us."""
    nature = rec["nature"]
    has_p = nature in ("Purchase", "Purchase and Sales")
    has_s = nature in ("Sales", "Purchase and Sales")
    p_tax, p_exe, p_vat = (rec["p_tax"], rec["p_exe"], rec["p_vat"]) if has_p else (0, 0, 0)
    s_tax, s_exe, s_vat = (rec["s_tax"], rec["s_exe"], rec["s_vat"]) if has_s else (0, 0, 0)
    paid = rec["paid"] if has_p else 0.0
    recv = rec["recv"] if has_s else 0.0
    opening = rec["open_amt"] if rec["open_type"] == "Payable by us" else -rec["open_amt"]
    p_total = p_tax + p_exe + p_vat
    s_total = s_tax + s_exe + s_vat
    closing = opening + p_total - s_total - paid + recv
    return dict(has_p=has_p, has_s=has_s, p_tax=p_tax, p_exe=p_exe, p_vat=p_vat,
                s_tax=s_tax, s_exe=s_exe, s_vat=s_vat, p_total=p_total,
                s_total=s_total, paid=paid, recv=recv, opening=opening,
                closing=closing)


def status_words(closing):
    if closing > 0.004:
        return "payable by us to you", "payable"
    if closing < -0.004:
        return "receivable by us from you", "receivable"
    return "nil", ""


# ------------------------------------------------------------------ styles
def _styles():
    base = dict(fontName=FONT, fontSize=10.5, leading=14.5, textColor=colors.black)
    return {
        "title": ParagraphStyle("title", fontName=BOLD, fontSize=25, leading=29, textColor=NAVY),
        "addr": ParagraphStyle("addr", fontName=FONT, fontSize=10.5, leading=14, textColor=GREY),
        "body": ParagraphStyle("body", spaceAfter=7, **base),
        "just": ParagraphStyle("just", alignment=TA_JUSTIFY, spaceAfter=7, **base),
        "tight": ParagraphStyle("tight", **base),
        "right": ParagraphStyle("right", alignment=TA_RIGHT, **base),
        "cell": ParagraphStyle("cell", fontName=FONT, fontSize=10, leading=13),
        "cellb": ParagraphStyle("cellb", fontName=BOLD, fontSize=10, leading=13),
        "cellr": ParagraphStyle("cellr", fontName=FONT, fontSize=10, leading=13, alignment=TA_RIGHT),
        "cellrb": ParagraphStyle("cellrb", fontName=BOLD, fontSize=10, leading=13, alignment=TA_RIGHT),
        "head": ParagraphStyle("head", fontName=BOLD, fontSize=10, leading=13, textColor=colors.white),
        "headr": ParagraphStyle("headr", fontName=BOLD, fontSize=10, leading=13,
                                textColor=colors.white, alignment=TA_RIGHT),
    }


def _p(text, style):
    return Paragraph(escape(str(text)), style)


# ------------------------------------------------------------------- table
def _amount_table(rec, c, S):
    nature = rec["nature"]
    op_word = "" if abs(c["opening"]) < 0.005 else (" payable" if c["opening"] > 0 else " receivable")
    _, cl_word = status_words(c["closing"])
    cl_label = "Closing balance" + (f" {cl_word}" if cl_word else "")
    opening_label = "Opening balance" + op_word
    closing_amt = money(abs(c["closing"])) if abs(c["closing"]) >= 0.005 else "0.00"

    style = [
        ("GRID", (0, 0), (-1, -1), 0.6, LINE),
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]

    if nature == "Purchase and Sales":
        widths = [6.5 * cm, 4.7 * cm, CONTENT_W - 11.2 * cm]
        rows = [
            [_p("Details", S["head"]), _p("Purchase", S["headr"]), _p("Sales", S["headr"])],
            [_p(opening_label, S["cell"]), _p(money(abs(c["opening"])), S["cellr"]), ""],
            [_p("Taxable", S["cell"]), _p(money(c["p_tax"]), S["cellr"]), _p(money(c["s_tax"]), S["cellr"])],
            [_p("Tax-exempt", S["cell"]), _p(money(c["p_exe"]), S["cellr"]), _p(money(c["s_exe"]), S["cellr"])],
            [_p("VAT", S["cell"]), _p(money(c["p_vat"]), S["cellr"]), _p(money(c["s_vat"]), S["cellr"])],
            [_p("Total", S["cellb"]), _p(money(c["p_total"]), S["cellrb"]), _p(money(c["s_total"]), S["cellrb"])],
            [_p(cl_label, S["cellb"]), _p(closing_amt, S["cellrb"]), ""],
        ]
        style += [
            ("SPAN", (1, 1), (2, 1)), ("SPAN", (1, 6), (2, 6)),
            ("BACKGROUND", (0, 5), (-1, 6), LIGHT),
        ]
    else:
        word = "purchase" if nature == "Purchase" else "sales"
        key = "p" if nature == "Purchase" else "s"
        widths = [10.6 * cm, CONTENT_W - 10.6 * cm]
        rows = [
            [_p("Details", S["head"]), _p("Amount (Rs.)", S["headr"])],
            [_p(opening_label, S["cell"]), _p(money(abs(c["opening"])), S["cellr"])],
            [_p(f"Taxable {word}", S["cell"]), _p(money(c[key + "_tax"]), S["cellr"])],
            [_p(f"Tax-exempt {word}", S["cell"]), _p(money(c[key + "_exe"]), S["cellr"])],
            [_p("VAT", S["cell"]), _p(money(c[key + "_vat"]), S["cellr"])],
            [_p(f"Total {word}", S["cellb"]), _p(money(c[key + "_total"]), S["cellrb"])],
            [_p(cl_label, S["cellb"]), _p(closing_amt, S["cellrb"])],
        ]
        style += [("BACKGROUND", (0, 5), (-1, 6), LIGHT)]

    t = Table(rows, colWidths=widths, hAlign="LEFT")
    t.setStyle(TableStyle(style))
    return t


# --------------------------------------------------------------- signature
@lru_cache(maxsize=16)
def _prepared_bytes(path, mtime):
    """Load a stamp/signature image, make a plain white background transparent
    and crop away empty margins, so the visible part prints as large as possible."""
    img = PILImage.open(path).convert("RGBA")
    alpha = img.getchannel("A")
    if alpha.getextrema()[0] >= 250:  # fully opaque (e.g. a JPG): drop the white paper
        gray = img.convert("L")
        alpha = gray.point(lambda v: max(0, min(255, int((240 - v) * 255 / 90))))
        img.putalpha(alpha)
    box = img.getchannel("A").point(lambda v: 255 if v > 20 else 0).getbbox()
    if box:
        img = img.crop(box)
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def prepared_image(path):
    """Processed image bytes (as printed on the letter)."""
    p = Path(path)
    return _prepared_bytes(str(p), p.stat().st_mtime)


class StampSign(Flowable):
    """Stamp and signature drawn close together; the signature overlaps the stamp."""

    def __init__(self, stamp, sign, stamp_h, sign_h, overlap, max_w):
        super().__init__()
        self.items = []  # (ImageReader, w, h)
        for src, hgt in ((stamp, stamp_h), (sign, sign_h)):
            if src:
                reader = ImageReader(io.BytesIO(prepared_image(src)))
                iw, ih = reader.getSize()
                self.items.append([reader, iw * hgt / ih, hgt])
        self.overlap = overlap if len(self.items) == 2 else 0.0
        total = self._total_width()
        self.scale = min(1.0, max_w / total) if total else 1.0

    def _total_width(self):
        if not self.items:
            return 0.0
        w = self.items[0][1]
        if len(self.items) == 2:
            w = w * (1 - self.overlap) + self.items[1][1]
        return w

    def wrap(self, availW, availH):
        h = max((i[2] for i in self.items), default=0.0)
        self.width, self.height = self._total_width() * self.scale, h * self.scale
        return self.width, self.height

    def draw(self):
        c, s, x = self.canv, self.scale, 0.0
        for n, (reader, w, h) in enumerate(self.items):
            y = (self.height - h * s) / 2
            if n == 1:
                y -= 0.12 * h * s  # signature sits a little lower, like a real one
            c.drawImage(reader, x, max(y, 0), width=w * s, height=h * s, mask="auto")
            x += w * s * (1 - self.overlap) if n == 0 else 0


def _sign_table(rec, S, company, stamp_paths):
    paths = list(stamp_paths)
    stamp = paths[0] if len(paths) > 0 and paths[0] else None
    sign = paths[1] if len(paths) > 1 and paths[1] else None
    top_left, row_h = "", 2.4 * cm
    if stamp or sign:
        top_left = StampSign(stamp, sign, company.get("stamp_h", 2.8) * cm,
                             company.get("sign_h", 2.2) * cm,
                             company.get("overlap", 35) / 100.0, 9.0 * cm)
        row_h = max(company.get("stamp_h", 2.8), company.get("sign_h", 2.2)) * cm + 0.3 * cm
    rows = [
        [top_left, "", ""],
        [_p("Authorised Signature & Seal", S["tight"]), "", _p("Authorised Signature & Seal", S["tight"])],
        [_p(f"For, {company['name'].upper()}", S["tight"]), "", _p(f"For, {rec['party'].upper()}", S["tight"])],
    ]
    t = Table(rows, colWidths=[6.2 * cm, 2.3 * cm, CONTENT_W - 8.5 * cm], rowHeights=[row_h, None, None], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ("LINEABOVE", (0, 1), (0, 1), 0.8, colors.black),
        ("LINEABOVE", (2, 1), (2, 1), 0.8, colors.black),
        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 1), (-1, 1), 3), ("TOPPADDING", (0, 0), (-1, 0), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    return t


# ------------------------------------------------------------------ letter
def letter_flowables(rec, company, stamp_paths=()):
    S = _styles()
    c = calc(rec)
    f = []

    f.append(Paragraph(escape(company["name"].upper()), S["title"]))
    f.append(Paragraph(escape(company["address"]), S["addr"]))
    contact = f"Phone: {company['phone']}   |   PAN No.: {company['pan']}"
    if company.get("email"):
        contact += f"   |   Email: {company['email']}"
    f.append(Paragraph(escape(contact), S["addr"]))
    f.append(HRFlowable(width="100%", thickness=1.8, color=NAVY, spaceBefore=4, spaceAfter=10))

    # ref / date line
    date_txt = f"Date : {rec['date_ad']}"
    if rec["date_bs"]:
        date_txt += f"   |   Miti : {rec['date_bs']}"
    left = _p(f"Ref No.: {rec['ref']}", S["tight"]) if rec["ref"] else ""
    rd = Table([[left, _p(date_txt, S["right"])]], colWidths=[6 * cm, CONTENT_W - 6 * cm])
    rd.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    f.append(rd)
    f.append(Spacer(1, 8))

    f.append(Paragraph("To,", S["tight"]))
    f.append(Paragraph(f"<b>M/S {escape(rec['party'].upper())}</b>", S["tight"]))
    if rec["address"]:
        f.append(Paragraph(escape(rec["address"]), S["tight"]))
    if rec["pan"]:
        f.append(Paragraph(f"PAN No.: {escape(rec['pan'])}", S["tight"]))
    f.append(Spacer(1, 10))

    f.append(Paragraph(f"<u><b>Subject: Confirmation of Transaction and Balance for F.Y. {escape(rec['fy'])}</b></u>", S["body"]))
    f.append(Paragraph("Dear Sir/Madam,", S["body"]))

    goods = {"Purchase": "purchased from you", "Sales": "supplied to you",
             "Purchase and Sales": "purchased from and supplied to you"}[rec["nature"]]
    asat = escape(rec["asat"])
    st_full, _ = status_words(c["closing"])
    if st_full == "nil":
        intro = (f"We write to inform you that, as per our books of account, no amount is due from either party "
                 f"as at the {asat}, in respect of goods {goods} during the fiscal year {escape(rec['fy'])}. "
                 f"The details are set out below.")
    else:
        intro = (f"We write to inform you that, as per our books of account, an amount of "
                 f"<b>Rs. {money(abs(c['closing']))}</b> is {st_full} as at the {asat}, in respect of goods "
                 f"{goods} during the fiscal year {escape(rec['fy'])}. The details are set out below.")
    f.append(Paragraph(intro, S["just"]))
    f.append(Paragraph(f"<b>Nature of Transaction:</b> {escape(rec['nature'])}", S["body"]))
    f.append(_amount_table(rec, c, S))
    f.append(Spacer(1, 12))

    f.append(Paragraph(
        "Kindly confirm the above balance by signing and stamping in the space provided below and returning "
        "this letter to us. If the amount does not tally with your records, please send us your statement of "
        "account immediately so that we can reconcile it at our end.", S["just"]))
    f.append(Paragraph(
        "If we do not receive a reply within fifteen (15) days of the date of this letter, the above balance "
        "will be deemed to have been confirmed and to agree with your books of account.", S["just"]))
    f.append(Paragraph("Thanking you !", S["body"]))
    f.append(Spacer(1, 4))
    f.append(KeepTogether([_sign_table(rec, S, company, stamp_paths)]))
    return f


def build_pdf(recs, company, stamp_paths=()):
    """recs: list of records -> one PDF (one letter per page, more if needed)."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2.55 * cm, rightMargin=2.55 * cm,
                            topMargin=1.5 * cm, bottomMargin=1.4 * cm,
                            title="Account Confirmation", author=company["name"])
    story = []
    for i, r in enumerate(recs):
        if i:
            story.append(PageBreak())
        story += letter_flowables(r, company, stamp_paths)
    doc.build(story)
    return buf.getvalue()
