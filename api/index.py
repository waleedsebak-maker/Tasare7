import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime, date

from flask import Flask, request, send_file, jsonify, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.utils.datetime import from_excel
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.utils import ImageReader
from PIL import Image
import arabic_reshaper
from bidi.algorithm import get_display
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
STAMPS = ROOT / "stamps"
LEGACY = ASSETS / "legacy_templates"
AR = ASSETS / "NotoSansArabic-Regular.ttf"
ARB = ASSETS / "NotoSansArabic-Bold.ttf"
LATIN = ASSETS / "DejaVuSans.ttf"

pdfmetrics.registerFont(TTFont("Arabic", str(AR)))
pdfmetrics.registerFont(TTFont("ArabicBold", str(ARB)))
pdfmetrics.registerFont(TTFont("Latin", str(LATIN)))
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

app = Flask(__name__, template_folder=str(ROOT / "templates"), static_folder=str(ROOT / "static"))
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

COMPANIES = {
    "unlimited": {"name": "ان ليميتد ايجيبت ترافيل", "english": "Unlimited Egypt Travel", "license": "2257"},
    "memory": {"name": "ميمورى تريب تورز", "english": "Memory Trip Tours", "license": "1586"},
}
PERMITS = {
    "cairo": "توصيلة القاهرة",
    "luxor_transfer": "توصيلة الأقصر",
    "luxor_overnight": "مبيت الأقصر",
    "hurghada": "توصيلة الغردقة من الأقصر",
}
ALLOWED = {
    "unlimited": {"cairo", "luxor_transfer", "luxor_overnight"},
    "memory": {"cairo", "luxor_transfer", "luxor_overnight", "hurghada"},
}

# Program defaults tied to the selected company + permit type.
PROGRAM_PRESETS = {
    ("unlimited", "cairo"): "زعفرانه – خان – ازهر – مصر القديمه – عشاء",
    ("unlimited", "luxor_transfer"): "85 – دندره – ملوك – ممنون – مرسي سياحي",
    ("unlimited", "luxor_overnight"): "85 – دندره – ملوك – ممنون – مرسي سياحي",
    ("memory", "cairo"): "زعفرانه – خان – ازهر – مصر القديمه – عشاء",
    ("memory", "luxor_transfer"): "85 – دندره – ملوك – ممنون – مرسي سياحي",
    ("memory", "luxor_overnight"): "85 – دندره – ملوك – ممنون – مرسي سياحي",
    ("memory", "hurghada"): "دندره 85",
}


DOCX_TEMPLATES = {
    ("unlimited", "cairo"): STAMPS / "ان ليميتيد ايجيبت تصريح القاهره(1).docx",
    ("memory", "cairo"): STAMPS / "تصريح ميموري القاهره الأساسي.docx",
    ("unlimited", "luxor_overnight"): STAMPS / "ان ليميتيد ايجيبت تصريح الاقصر مبيت.docx",
    ("memory", "luxor_overnight"): STAMPS / "تصريح ميموري الاقصر مبيت الأساسي.docx",
    ("memory", "hurghada"): STAMPS / "ميموري توصيله الغردقه.docx",
}
POLICE_TEMPLATE = STAMPS / "4_5859343746985894309.docx"

LEGACY_PDF_TEMPLATES = {
    ("unlimited", "luxor_transfer"): LEGACY / "unlimited_luxor_transfer.pdf",
    ("memory", "luxor_transfer"): LEGACY / "memory_luxor_transfer.pdf",
    ("memory", "hurghada"): LEGACY / "memory_hurghada.pdf",
}

ALIASES = {
    "ID": ["id", "no", "no.", "م", "رقم", "الرقم"],
    "Chinese name": [
        "name",
        "ch name", "chinese name",
        "الاسم بالصيني", "الاسم الصينى",
        "الصينى", "中文姓名", "chinese"
    ],
    "English name": [
        "spelling",
        "english name",
        "english",
        "name spelling",
        "romanized name",
        "romanized",
        "الاسم بالانجليزي",
        "الاسم بالإنجليزي",
        "الاسم الانجليزي",
        "الاسم الإنجليزي",
        "الاسم باللاتيني"
    ],
    "Surname": [
        "surname", "family name",
        "اسم العائلة", "العائلة", "family"
    ],
    "Given name": [
        "given name", "first name",
        "الاسم الأول", "الاسم الاول", "given"
    ],
    "Sex": ["sex", "gender", "النوع", "الجنس"],
    "DOB": [
        "dob", "date of birth",
        "تاريخ الميلاد", "birth date"
    ],
    "Passport": [
        "passport", "passport no", "passportno",
        "رقم الجواز", "جواز السفر", "passport number"
    ],
    "Expiry": [
        "expiry", "passport expiry",
        "date of expiry", "انتهاء الجواز",
        "تاريخ انتهاء الجواز", "expiry date"
    ],
    "Room": [
        "room", "room no", "الغرفة",
        "رقم الغرفة", "option", "room type"
    ],
    "Note": ["note", "notes", "ملاحظات", "ملحوظة"],
}


def norm(v):
    s = str(v or "").strip().lower()
    s = s.replace("\n", " ").replace("_", " ")
    s = re.sub(r"\s+", " ", s)
    return s


def fmt_date(v):
    if v in (None, ""):
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    try:
        if isinstance(v, (int, float)) and 20000 < float(v) < 80000:
            return from_excel(v).strftime("%d/%m/%Y")
    except Exception:
        pass
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m:
        return f"{m.group(3).zfill(2)}/{m.group(2).zfill(2)}/{m.group(1)}"
    return s.replace("-", "/")


def ar(text):
    return get_display(arabic_reshaper.reshape(str(text or "")))


def draw_rtl(c, text, x, y, max_width=None, size=10, bold=True):
    font = "ArabicBold" if bold else "Arabic"
    raw = ar(text)
    if max_width:
        while size > 6 and pdfmetrics.stringWidth(raw, font, size) > max_width:
            size -= 0.5
    c.setFont(font, size)
    c.drawRightString(x, y, raw)


def find_header_row(rows):
    best = None
    best_score = 0
    for i, row in enumerate(rows[:30]):
        vals = [norm(v) for v in row]
        score = 0
        for opts in ALIASES.values():
            if any(v in opts for v in vals):
                score += 1
        if score > best_score:
            best_score, best = score, i
    return best if best_score >= 2 else 0


def read_rooming(stream):
    wb = load_workbook(stream, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    hi = find_header_row(rows)
    headers = [norm(v) for v in rows[hi]]

    idx = {}
    for canonical, opts in ALIASES.items():
        for i, h in enumerate(headers):
            if h in opts:
                idx[canonical] = i
                break

    out = []
    for raw in rows[hi + 1:]:
        if not any(v not in (None, "") for v in raw):
            continue

        item = {}
        for key, col in idx.items():
            item[key] = raw[col] if col < len(raw) else ""

        english = str(item.get("English name") or "").strip()
        chinese = str(item.get("Chinese name") or "").strip()

        if not chinese and not english:
            continue

        if not english:
            surname = str(item.get("Surname") or "").strip()
            given = str(item.get("Given name") or "").strip()
            if surname or given:
                english = "/".join(x for x in (surname, given) if x)

        item["English name"] = english
        item["ID"] = item.get("ID") or len(out) + 1

        for k in ("DOB", "Expiry"):
            item[k] = fmt_date(item.get(k))

        for k in item:
            if item[k] is None:
                item[k] = ""

        out.append(item)

    return out


def replace_cell(cell, value, size=None, bold=True):
    cell.text = str(value)
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.name = "Arial"
            if size:
                from docx.shared import Pt
                run.font.size = Pt(size)
            run.bold = bool(bold)


def fill_docx_template(template_path, data, output_path):
    """Fill the approved DOCX template without changing its original layout."""
    doc = Document(str(template_path))
    starts = [0, 4, 8]
    if len(doc.tables) < 12:
        raise ValueError("الاستمارة الرسمية لا تحتوي على نسخ الإخطار الثلاث المطلوبة")

    for base in starts:
        t1 = doc.tables[base + 1]
        t3 = doc.tables[base + 3]

        # Actual value cells in the approved notification templates.
        replace_cell(t1.cell(1, 1), str(data["pax"]), 16, True)
        replace_cell(t1.cell(2, 1), data["guide"], 14, True)
        replace_cell(t1.cell(2, 3), data["phone"], 13, True)

        # Trip section: preserve the original table and checkbox positions.
        replace_cell(t3.cell(0, 1), data["permit"], 14, True)
        replace_cell(t3.cell(0, 4), data["program"], 11, True)

        direction = (
            "القاهرة"
            if data["permit_key"] == "cairo"
            else (
                "الأقصر"
                if data["permit_key"] in {"luxor_transfer", "luxor_overnight"}
                else "الغردقة"
            )
        )
        checkpoint = "كمين الجونة" if data["permit_key"] == "cairo" else "كمين سفاجا - قنا"

        replace_cell(
            t3.cell(1, 1),
            ("☑ " if direction == "القاهرة" else "☐ ")
            + "القاهرة    الإسكندرية    شرم الشيخ"
        )
        replace_cell(
            t3.cell(1, 3),
            ("☑ " if direction == "الأقصر" else "☐ ")
            + "الأقصر                    أسوان"
        )
        replace_cell(
            t3.cell(2, 1),
            ("☑ " if checkpoint == "كمين الجونة" else "☐ ")
            + "كمين الجونة"
        )
        replace_cell(
            t3.cell(2, 3),
            ("☑ " if checkpoint != "كمين الجونة" else "☐ ")
            + "كمين سفاجا - قنا       كمين مرسى علم - إدفو"
        )

    doc.save(str(output_path))


def fill_hurghada_docx_template(template_path, data, output_path):
    """Fill the approved Memory Hurghada/Luxor Word form only in its blank fields.
    The original checkbox/approval/stamp layout is preserved verbatim.
    """
    doc = Document(str(template_path))
    starts = [0, 4, 8]
    for base in starts:
        t0 = doc.tables[base]
        t1 = doc.tables[base + 1]
        t3 = doc.tables[base + 3]

        # Header date cell from the original form.
        replace_cell(t0.cell(0, 0), f"تاريخ الرحله: {data['trip_date']}")

        # Company information: these are the blank cells in the original form.
        replace_cell(t1.cell(1, 1), str(data['pax']))
        replace_cell(t1.cell(2, 1), data['guide'])
        replace_cell(t1.cell(2, 3), data['phone'])

        # Trip program: preserve the original trip type, direction and checkpoint
        # checkboxes already embedded in the approved template.
        replace_cell(t3.cell(0, 4), data['program'])

    doc.save(str(output_path))

def libreoffice_convert(src, outdir):
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice غير مثبت على الخادم. ثبته لتفعيل تحويل Word إلى PDF.")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
    pdf = Path(outdir) / (Path(src).stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError("فشل تحويل Word إلى PDF")
    return pdf


def overlay_docx_notice_fields(c, data):
    c.setFillColor(colors.black)
    c.setFont("Latin", 9)
    c.drawRightString(575, 738, data["file_no"])
    c.setFont("ArabicBold", 13)
    c.drawString(60, 738, data["trip_date"])
    if data.get("issue_date"):
        draw_rtl(c, f"تحريراً في : {data['issue_date']}", 205, 150, 160, 8)


def overlay_legacy_notice_fields(c, data, w, h):
    """Fill blank cells on the old PDF stamps (page size ~612x936)."""
    c.setFillColor(colors.black)

    # Header bar: file number on the right, trip date on the left.
    c.setFont("Latin", 10)
    c.drawRightString(w - 40, 768, data["file_no"])
    c.setFont("Latin", 10)
    c.drawString(248, 768, data["trip_date"])

    # Company table value cells
    draw_rtl(c, str(data["pax"]), 452, 712, 160, 11, True)
    draw_rtl(c, data["guide"], 452, 696, 165, 10, True)
    draw_rtl(c, data["phone"], 136, 696, 80, 8, True)

    # Trip type + program
    draw_rtl(c, data["permit"], 468, 526, 140, 9, True)
    program = data.get("program") or ""
    draw_rtl(c, program, 208, 526, 150, 7)

    if data.get("issue_date"):
        draw_rtl(c, f"تحريراً في : {data['issue_date']}", w - 40, 358, 220, 8)


def overlay_notice_fields(pdf_bytes, data):
    src = PdfReader(io.BytesIO(pdf_bytes))
    out = PdfWriter()

    is_hg = data.get("permit_key") == "hurghada"

    for page_no, page in enumerate(src.pages[:3], 1):
        ov = io.BytesIO()
        c = canvas.Canvas(ov, pagesize=A4)

        # =========================================================
        # Memory Trip Tours - Hurghada transfer
        # Preserve the original PDF completely and only add values
        # into the blank/value areas.
        # =========================================================
        if is_hg:
            c.setFillColorRGB(0, 0, 0)

            # Common header
            c.setFont("Latin", 9)
            c.drawRightString(575, 738, str(data.get("file_no", "")))
            c.drawString(60, 738, str(data.get("trip_date", "")))

            if page_no == 1:
                # العدد
                draw_rtl(c, str(data.get("pax", "")), 205, 145, 55, 11, True)

                # اسم المرشد
                draw_rtl(c, str(data.get("guide", "")), 205, 130, 160, 9, False)

                # الهاتف
                draw_rtl(c, str(data.get("phone", "")), 205, 115, 150, 9, False)

                # نوع التصريح
                draw_rtl(c, "توصيلة الغردقة", 205, 100, 160, 9, True)

                # البرنامج
                draw_rtl(c, str(data.get("program", "")), 205, 84, 180, 8, False)

            elif page_no in (2, 3):
                # بيانات شركة السياحة
                draw_rtl(c, str(data["company"]["name"]), 515, 653, 260, 10, True)
                draw_rtl(c, str(data.get("pax", "")), 535, 620, 60, 10, True)

                # المرشد
                draw_rtl(c, str(data.get("guide", "")), 450, 608, 220, 9, False)

                # تليفون المرشد
                draw_rtl(c, str(data.get("phone", "")), 195, 608, 150, 9, False)

                # نوع التصريح
                draw_rtl(c, "توصيلة الغردقة", 500, 574, 190, 9, True)

                # البرنامج
                draw_rtl(c, str(data.get("program", "")), 245, 458, 235, 8, False)

                # الفندق
                if data.get("hotel_to"):
                    draw_rtl(c, str(data["hotel_to"]), 245, 430, 235, 8, False)

                # checkpoint / route information
                draw_rtl(c, "الأقصر ← الغردقة", 500, 414, 190, 9, True)

        else:
            # Existing behavior for all other permit types
            c.setFont("Latin", 9)
            c.drawRightString(575, 738, data["file_no"])
            c.drawString(60, 738, data["trip_date"])
            draw_rtl(
                c,
                f"تحريراً في : {data['issue_date']}",
                205,
                150,
                160,
                8
            )

        c.save()
        ov.seek(0)
        page.merge_page(PdfReader(ov).pages[0])
        out.add_page(page)

    buf = io.BytesIO()
    out.write(buf)
    buf.seek(0)
    return buf

def build_notices(data, workdir):
    key = (data["company_key"], data["permit_key"])
    if key in DOCX_TEMPLATES and DOCX_TEMPLATES[key].exists():
        filled = Path(workdir) / "filled.docx"
        if key == ("memory", "hurghada"):
            fill_hurghada_docx_template(DOCX_TEMPLATES[key], data, filled)
        else:
            fill_docx_template(DOCX_TEMPLATES[key], data, filled)
        pdf = libreoffice_convert(filled, workdir)
        raw = pdf.read_bytes()
        return overlay_notice_fields(raw, data, legacy=False)
    if key in LEGACY_PDF_TEMPLATES and LEGACY_PDF_TEMPLATES[key].exists():
        raw = LEGACY_PDF_TEMPLATES[key].read_bytes()
        return overlay_notice_fields(raw, data, legacy=True)
    raise ValueError("لا توجد استمارة رسمية معتمدة لهذا النوع في مجلد stamps/assets.")


def make_room_page(data, rows, copy_no):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    BLACK = colors.black
    WHITE = colors.white

    # العنوان
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        "كشف أسماء السائحين / نيم ليست",
        W - 28, H - 32, 400, 16, True
    )

    c.setFillColor(BLACK)
    c.setFont("Latin", 8)
    copies = 1 if data.get("permit_key") == "hurghada" else 3
    c.drawString(28, H - 32, f"Copy {copy_no}/{copies}")

    # الشركة ورقم الملف والعدد
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        f"{data['company']['english']} | "
        f"{data['company']['license']} | "
        f"{data['file_no']} | {data['pax']} Pax",
        W - 28, H - 52, 520, 8
    )

    # المرشد والتليفون - كبير وواضح
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        f"المرشد: {data['guide']}   الهاتف: {data['phone']}",
        W - 28, H - 70, 520, 11, True
    )

    # التاريخ - كبير وواضح
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        f"تاريخ الرحلة: {data['trip_date']}",
        W - 28, H - 88, 250, 12, True
    )

    # نوع التصريح - أكبر وBold
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        data["permit"],
        300, H - 88, 250, 13, True
    )

    # الفنادق
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        (
            f"فندق الغردقة: {data['hotel_to']}"
            if data.get("permit_key") == "hurghada" and data.get("hotel_to")
            else f"{data['hotel_from']}  ←  {data['hotel_to']}"
        ),
        W - 28, H - 103, 520, 8
    )

    # جدول النيم ليست
    headers = [
        "م",
        "الاسم بالصيني",
        "الاسم بالإنجليزي",
        "النوع",
        "تاريخ الميلاد",
        "رقم الجواز",
        "انتهاء الجواز",
        "الغرفة",
        "ملاحظات"
    ]

    widths = [22, 92, 100, 34, 65, 74, 65, 45, 55]

    total = sum(widths)
    x0 = (W - total) / 2
    top = H - 120
    available = top - 52

    if len(rows) > 38:
        raise ValueError(
            "عدد الأسماء كبير جداً لصفحة نيم ليست واحدة."
        )

    rh = max(
        10.0,
        min(19.0, available / max(len(rows) + 1, 1))
    )

    header_font = max(4.8, min(6.2, rh * 0.32))
    body_font = max(4.5, min(5.5, rh * 0.29))

    # Header أبيض + حدود سوداء
    x = x0

    for h, w in zip(headers, widths):
        c.setFillColor(WHITE)
        c.setStrokeColor(BLACK)
        c.setLineWidth(.35)
        c.rect(x, top-rh, w, rh, stroke=1, fill=1)

        c.setFillColor(BLACK)
        c.setFont("ArabicBold", header_font)
        c.drawCentredString(
            x + w/2,
            top - (rh * 0.68),
            ar(h)
        )

        x += w

    # الصفوف
    y = top - rh

    for r in rows:
        x = x0

        vals = [
            r.get("ID", ""),
            r.get("Chinese name", ""),
            r.get("English name", ""),
            r.get("Sex", ""),
            r.get("DOB", ""),
            r.get("Passport", ""),
            r.get("Expiry", ""),
            r.get("Room", ""),
            r.get("Note", "")
        ]

        for j, (v, w) in enumerate(zip(vals, widths)):
            c.setFillColor(WHITE)
            c.setStrokeColor(BLACK)
            c.setLineWidth(.35)
            c.rect(x, y-rh, w, rh, stroke=1, fill=1)

            c.setFillColor(BLACK)

            if j == 1:
                c.setFont("STSong-Light", body_font)
            else:
                c.setFont("Latin", body_font)

            text_value = str(v or "")
            c.drawCentredString(
                x + w/2,
                y - (rh * 0.68),
                text_value[:24]
            )

            x += w

        y -= rh

        if y < 50:
            break

    # الإقرار
    c.setFillColor(BLACK)
    draw_rtl(
        c,
        f"إقرار من الشركة: نقر نحن شركة "
        f"{data['company']['name']} ترخيص رقم "
        f"{data['company']['license']} بأن البيانات الواردة أعلاه صحيحة.",
        W - 28, 26, 540, 7, True
    )

    c.save()
    buf.seek(0)
    return buf

def make_image_page(img_bytes, landscape_page=False):
    page_size = landscape(A4) if landscape_page else A4
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=page_size); W, H = page_size
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    maxw, maxh = W - 24, H - 24
    scale = min(maxw/im.width, maxh/im.height)
    iw, ih = im.width*scale, im.height*scale
    tmp = io.BytesIO(); im.save(tmp, format="PNG"); tmp.seek(0)
    c.drawImage(ImageReader(tmp), (W-iw)/2, (H-ih)/2, width=iw, height=ih, preserveAspectRatio=True, mask="auto")
    c.save(); buf.seek(0); return buf



def make_room_shot_page(img_bytes):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    maxw, maxh = W - 16, H - 16
    scale = min(maxw / float(im.width), maxh / float(im.height))
    iw, ih = im.width * scale, im.height * scale
    tmp = io.BytesIO(); im.save(tmp, format="PNG"); tmp.seek(0)
    c.drawImage(ImageReader(tmp), (W-iw)/2, (H-ih)/2, width=iw, height=ih, preserveAspectRatio=True, mask="auto")
    c.save(); buf.seek(0); return buf

def make_program_page(img_bytes):
    return make_image_page(img_bytes, landscape_page=True)


def extra_page_from_upload(file_storage):
    raw = file_storage.read()
    name = (file_storage.filename or "").lower()

    # الورقة الإضافية: صورة فقط، ممنوع PDF.
    if name.endswith(".pdf") or raw[:4] == b"%PDF":
        raise ValueError("الورقة الإضافية يجب أن تكون صورة فقط JPG أو PNG أو WEBP")

    try:
        img = Image.open(io.BytesIO(raw))
        img.verify()
    except Exception:
        raise ValueError("الورقة الإضافية يجب أن تكون صورة صحيحة JPG أو PNG أو WEBP")

    return make_image_page(raw, landscape_page=False)

def fill_police_notice(template_path, data, output_path):
    doc = Document(str(template_path))
    company = data["company"]
    hotel = data.get("hotel_to") or ""
    program = data.get("program") or "توصيلة الغردقة"
    if hotel and hotel not in program:
        program = f"توصيله الغردقه – {hotel}"
    elif data.get("permit_key") == "hurghada":
        program = f"توصيله الغردقه – {hotel}".rstrip(" –")

    if len(doc.paragraphs) > 4:
        replace_cell_like = doc.paragraphs[4]
        replace_cell_like.text = (
            f"نخطر سيادتكم علما نحن شركة {company['name']} ترخيص رقم {company['license']}"
        )
    if len(doc.paragraphs) > 13:
        issue = data.get("issue_date") or data.get("trip_date") or ""
        doc.paragraphs[13].text = f"تحريرا في يوم {issue}"

    if doc.tables:
        t = doc.tables[0]
        replace_cell(t.cell(1, 0), f"صيني\nالعدد {data['pax']}")
        replace_cell(t.cell(1, 1), program)
        replace_cell(t.cell(1, 2), "الأسم / ")
    doc.save(str(output_path))


def build_police_pdf(data, workdir):
    if not POLICE_TEMPLATE.exists():
        raise ValueError("ملف إخطار الشرطة غير موجود في stamps")
    filled = Path(workdir) / "police.docx"
    fill_police_notice(POLICE_TEMPLATE, data, filled)
    return libreoffice_convert(filled, workdir)


def merge_pages(notices, rooms, program_bytes=None):
    writer = PdfWriter()
    for p in PdfReader(notices).pages[:3]: writer.add_page(p)
    if program_bytes:
        writer.add_page(PdfReader(make_program_page(program_bytes)).pages[0])
    for i in range(1, 4):
        writer.add_page(PdfReader(rooms[i-1]).pages[0])
    expected = 7 if program_bytes else 6
    if len(writer.pages) != expected:
        raise RuntimeError(f"خطأ في عدد الصفحات: {len(writer.pages)} بدل {expected}")
    out = io.BytesIO(); writer.write(out); out.seek(0); return out


def _first_page(src):
    if hasattr(src, "read"):
        src.seek(0)
        reader = PdfReader(src)
    else:
        reader = PdfReader(str(src))
    if not reader.pages:
        raise ValueError("صفحة ناقصة أثناء تجميع ملف الغردقة")
    return reader.pages[0]


def merge_hurghada_pages(notices, room_page, program_bytes, police_pdf, extra_pdf):
    """3 إخطارات + روم ليست واحدة + إدارة البرامج + إخطار شرطة الأقصر + الورقة الإضافية."""
    writer = PdfWriter()
    for p in PdfReader(notices).pages[:3]:
        writer.add_page(p)
    writer.add_page(_first_page(room_page))
    writer.add_page(_first_page(make_program_page(program_bytes)))
    writer.add_page(_first_page(police_pdf))
    writer.add_page(_first_page(extra_pdf))
    if len(writer.pages) != 7:
        raise RuntimeError(f"خطأ في عدد صفحات الغردقة: {len(writer.pages)} بدل 7")
    out = io.BytesIO(); writer.write(out); out.seek(0); return out


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/health")
def health():
    return jsonify(ok=True, libreoffice=bool(shutil.which("libreoffice") or shutil.which("soffice")))


@app.get("/api/rooming-template")
def rooming_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Rooming"
    headers = ["ID", "Chinese name", "English name", "Sex", "DOB", "Passport", "Expiry", "Room", "Note"]
    ws.append(headers)
    ws.append([1, "张伟", "ZHANG/WEI", "M", "15/01/1990", "E12345678", "20/05/2030", "101", ""])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name="nime_list_template.xlsx")


@app.post("/api/preview-rooming")
def preview_rooming():
    try:
        if not room or not room.filename:
            raise ValueError("ارفع ملف النيم ليست Excel")
        rows = read_rooming(io.BytesIO(room.read()))
        preview = []
        for r in rows:
            preview.append({
                "id": r.get("ID", ""),
                "chinese": r.get("Chinese name", ""),
                "english": r.get("English name", ""),
                "sex": r.get("Sex", ""),
                "dob": r.get("DOB", ""),
                "passport": r.get("Passport", ""),
                "expiry": r.get("Expiry", ""),
                "room": r.get("Room", ""),
                "note": r.get("Note", ""),
            })
        return jsonify(ok=True, count=len(preview), rows=preview)
    except Exception as e:
        return jsonify(error=str(e)), 400


@app.post("/api/generate")
def generate():
    try:
        f = request.form
        company_key = f.get("company", "unlimited")
        permit_key = f.get("permit", "cairo")
        if company_key not in COMPANIES or permit_key not in ALLOWED[company_key]:
            raise ValueError("الشركة أو نوع التصريح غير صحيح")
        pax_raw = str(f.get("pax", "")).strip()
        pax = int(pax_raw) if pax_raw.isdigit() else 0
        file_no = f.get("file_no", "").strip()
        guide = f.get("guide", "").strip()
        phone = f.get("phone", "").strip()
        trip = fmt_date(f.get("trip_date", ""))
        issue = fmt_date(f.get("issue_date", ""))
        program = f.get("program", "").strip() or PROGRAM_PRESETS.get((company_key, permit_key), "")
        hotel_to = f.get("hotel_to", "").strip()
        if hotel_to and hotel_to not in program:
            program = program.rstrip(' -–—') + ' – ' + hotel_to
        if not file_no or not guide or not trip or not program:
            raise ValueError("أكمل رقم الملف والمرشد وتاريخ الرحلة والبرنامج")
        room_shot = request.files.get("rooming_img")
        room = request.files.get("rooming")

        if permit_key == "hurghada":
            # Memory Hurghada uses the uploaded rooming-list screenshot.
            if not room_shot or not room_shot.filename:
                raise ValueError("ارفع سكرين النيم ليست للغردقة")
            rows = []
        else:
            # Cairo/Luxor continue to use the Excel rooming list.
            if not room or not room.filename:
                raise ValueError("ارفع ملف النيم ليست Excel")
            rows = read_rooming(io.BytesIO(room.read()))
            if len(rows) != pax:
                raise ValueError(
                    f"العدد المعلن {pax} لا يساوي عدد الأسماء في النيم ليست ({len(rows)})"
                )

        program_file = request.files.get("program_image")
        program_bytes = program_file.read() if program_file and program_file.filename else None
        extra_file = request.files.get("extra_page")
        company = COMPANIES[company_key]
        data = {"company": company, "company_key": company_key, "permit_key": permit_key, "permit": PERMITS[permit_key], "file_no": file_no, "pax": pax, "guide": guide, "phone": phone, "trip_date": trip, "issue_date": issue, "program": program, "hotel_from": f.get("hotel_from", "").strip(), "hotel_to": hotel_to}
        with tempfile.TemporaryDirectory() as td:
            notices = build_notices(data, td)
            safe = re.sub(r"[^A-Za-z0-9_-]+", "_", file_no) or "permit"
            suffix = {"cairo":"cairo_transfer","luxor_transfer":"luxor_transfer","luxor_overnight":"luxor_overnight","hurghada":"hurghada_transfer"}[permit_key]
            main_name = f"{safe}_{suffix}.pdf"
            if permit_key == "hurghada":
                if not program_bytes:
                    raise ValueError("ارفع صورة إدارة البرامج. ملف الغردقة لازم 7 ورقات")
                if not extra_file or not extra_file.filename:
                    raise ValueError("ارفع الورقة الإضافية. ملف الغردقة لازم 7 ورقات")
                extra_pdf = extra_page_from_upload(extra_file)
                police_pdf = build_police_pdf(data, td)
                room_one = make_room_shot_page(room_shot.read()) if room_shot else make_room_page(data, rows, 1)
                final = merge_hurghada_pages(notices, room_one, program_bytes, police_pdf, extra_pdf)
            else:
                room_pages = [make_room_page(data, rows, i) for i in range(1, 4)]
                final = merge_pages(notices, room_pages, program_bytes)
            return send_file(final, mimetype="application/pdf", as_attachment=True, download_name=main_name)
    except Exception as e:
        return jsonify(error=str(e)), 400
