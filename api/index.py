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
from openpyxl.worksheet.page import PageMargins
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
    "ID": [
        "id", "no", "no.", "م", "رقم", "الرقم"
    ],
    "Chinese name": [
        "ch name", "chinese name",
        "الاسم بالصيني", "الاسم الصينى",
        "الصينى", "中文姓名",
        "chinese", "旅客姓名"
    ],
    "Surname": [
        "surname", "family name",
        "اسم العائلة", "العائلة",
        "family", "英文姓"
    ],
    "Given name": [
        "given name", "first name",
        "الاسم الأول", "الاسم الاول",
        "given", "英文名"
    ],
    "Sex": [
        "sex", "gender",
        "النوع", "الجنس", "性别"
    ],
    "DOB": [
        "dob", "date of birth",
        "تاريخ الميلاد", "birth date",
        "出生日期"
    ],
    "Passport": [
        "passport", "passport no",
        "passportno", "رقم الجواز",
        "جواز السفر", "passport number",
        "护照号码"
    ],
    "Expiry": [
        "expiry", "passport expiry",
        "date of expiry", "انتهاء الجواز",
        "تاريخ انتهاء الجواز", "expiry date",
        "有效期"
    ],
    "Room": [
        "room", "room no",
        "الغرفة", "رقم الغرفة",
        "option", "room type",
        "分房编号"
    ],
    "Note": [
        "note", "notes",
        "ملاحظات", "ملحوظة",
        "重要备注"
    ],
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

    # DDMMYYYY مثل 23092026
    m = re.match(r"^(\d{2})(\d{2})(\d{4})$", s)
    if m:
        return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"

    # YYYYMMDD مثل 20260923
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"

    # YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
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
    raw = stream.read() if hasattr(stream, "read") else bytes(stream)

    def load_xlsx(data):
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
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
        for raw_row in rows[hi + 1:]:
            if not any(v not in (None, "") for v in raw_row):
                continue

            item = {}
            for key, col in idx.items():
                item[key] = raw_row[col] if col < len(raw_row) else ""

            if not item.get("Passport") and not item.get("Chinese name") and not item.get("Given name"):
                continue

            item["ID"] = item.get("ID") or len(out) + 1

            for k in ("DOB", "Expiry"):
                item[k] = fmt_date(item.get(k))

            for k in item:
                if item[k] is None:
                    item[k] = ""

            out.append(item)

        return out

    try:
        return load_xlsx(raw)
    except Exception:
        # دعم Excel القديم .xls تلقائياً عن طريق LibreOffice
        soffice = shutil.which("libreoffice") or shutil.which("soffice")
        if not soffice:
            raise ValueError("لا يمكن قراءة ملف Excel القديم .xls لأن LibreOffice غير متوفر")

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            work = Path(td)
            src = work / "rooming_source.xls"
            src.write_bytes(raw)

            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to", "xlsx",
                    "--outdir", str(work),
                    str(src)
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=90
            )

            converted = work / "rooming_source.xlsx"
            if not converted.exists():
                raise ValueError("تعذر تحويل ملف Excel القديم .xls")

            return load_xlsx(converted.read_bytes())

def replace_cell(cell, value, size=14, bold=True):
    cell.text = str(value)
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.name = "Arial"
            if size:
                from docx.shared import Pt
                run.font.size = Pt(size)
            run.bold = bool(bold)



def _set_transport_cell(cell, label, value):
    """Keep template label and write exactly one value."""
    value = str(value or "").strip()
    if not value:
        return

    paragraph = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()

    original = " ".join(cell.text.split())
    pos = original.find(label)

    if pos >= 0:
        base = original[:pos] + label
    else:
        base = label

    rpr = None
    if paragraph.runs:
        try:
            rpr = deepcopy(paragraph.runs[0]._r.get_or_add_rPr())
        except Exception:
            rpr = None

    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)

    r1 = paragraph.add_run(base.strip())
    r2 = paragraph.add_run(" " + value)

    # تكبير وتغليظ القيمة المدخلة فقط
    try:
        from docx.shared import Pt

        font_sizes = {
            "العدد": 14,
            "المرشد": 14,
            "الهاتف": 14,
            "التصريح": 14,
            "البرنامج": 13,
            "السائق": 14,
            "رقم الهاتف": 14,
            "شركة النقل": 13,
            "رقم السيارة": 13,
        }

        if label in font_sizes:
            r2.font.size = Pt(font_sizes[label])
            r2.bold = True
    except Exception:
        pass

    # Increase ONLY the inserted value text.
    # Original labels, cells, borders and layout remain unchanged.
    try:
        from docx.shared import Pt

        font_sizes = {
            "العدد": 14,
            "المرشد": 14,
            "الهاتف": 14,
            "التصريح": 14,
            "البرنامج": 13,
            "السائق": 14,
            "رقم الهاتف": 14,
            "شركة النقل": 13,
            "رقم السيارة": 13,
        }

        if label in font_sizes:
            r2.font.size = Pt(font_sizes[label])
    except Exception:
        pass


    if rpr is not None:
        try:
            r1._r.get_or_add_rPr().append(deepcopy(rpr))
            r2._r.get_or_add_rPr().append(deepcopy(rpr))
        except Exception:
            pass

def _fill_transport_tables(doc, data):
    """
    Fill only the existing transport fields in the official template.
    Tables 2, 6 and 10 are the three notification copies.
    No new rows/cells/pages are created.
    """
    transport_tables = (2, 6, 10)

    for ti in transport_tables:
        if ti >= len(doc.tables):
            continue

        table = doc.tables[ti]

        # Company name: label R0C0, value R0C1
        company = str(data.get("transport_company", "") or "").strip()
        if company and len(table.rows) > 0 and len(table.rows[0].cells) > 1:
            replace_cell(table.cell(0, 1), company)

        # Vehicle number: label R1C2, value R1C3
        vehicle = str(data.get("vehicle_number", "") or "").strip()
        if vehicle and len(table.rows) > 1 and len(table.rows[1].cells) > 3:
            replace_cell(table.cell(1, 3), vehicle)

        # Primary driver: label/value live in the same original cells.
        driver = str(data.get("driver_name", "") or "").strip()
        if driver and len(table.rows) > 3 and len(table.rows[3].cells) > 0:
            _set_transport_cell(table.cell(3, 0), "السائق", driver)

        # Driver phone: label/value live in the same original cell.
        phone = str(data.get("driver_phone", "") or "").strip()
        if phone and len(table.rows) > 3 and len(table.rows[3].cells) > 1:
            _set_transport_cell(table.cell(3, 1), "الهاتف", phone)


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
        _set_transport_cell(t1.cell(1, 1), "العدد", str(data["pax"]))
        _set_transport_cell(t1.cell(2, 1), "المرشد", data["guide"])
        _set_transport_cell(t1.cell(2, 3), "الهاتف", data["phone"])

        # Trip section: preserve the original table and checkbox positions.
        _set_transport_cell(t3.cell(0, 1), "التصريح", data["permit"])
        _set_transport_cell(t3.cell(0, 4), "البرنامج", data["program"])

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

    _fill_transport_tables(doc, data)
    # Apply the same approved transport-field method to all permit types.
    # This keeps the original labels/cells and their formatting in all 3 notices.
    _fill_transport_tables(doc, data)

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


def _format_display_date(value):
    """
    Convert supported date formats to DD/MM/YYYY.
    Examples:
      2026-09-20 -> 20/09/2026
      20092026   -> 20/09/2026
      20/09/2026 -> 20/09/2026
    """
    value = str(value or "").strip()

    if not value:
        return ""

    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", value)
    if m:
        y, mo, d = m.groups()
        return f"{d}/{mo}/{y}"

    m = re.fullmatch(r"(\d{2})(\d{2})(\d{4})", value)
    if m:
        d, mo, y = m.groups()
        return f"{d}/{mo}/{y}"

    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", value)
    if m:
        return value

    return value


def _draw_bold_text(canvas_obj, text, x, y, size=12, align="right"):
    """
    Draw added fields in strong black bold text.
    """
    try:
        canvas_obj.setFillColorRGB(0, 0, 0)
        canvas_obj.setFont("ArabicBold", size)
    except Exception:
        # Fallback if ArabicBold is not registered in this version.
        canvas_obj.setFillColorRGB(0, 0, 0)
        canvas_obj.setFont("Helvetica-Bold", size)

    if align == "right":
        canvas_obj.drawRightString(x, y, str(text))
    else:
        canvas_obj.drawString(x, y, str(text))


def overlay_notice_fields(pdf_bytes, data, legacy=None):
    src = PdfReader(io.BytesIO(pdf_bytes))
    out = PdfWriter()

    for page in src.pages[:3]:
        ov = io.BytesIO()
        c = canvas.Canvas(ov, pagesize=A4)

        # رقم الملف: أعلى اليمين فقط
        c.setFont("Latin", 11)
        c.drawRightString(575, 806, str(data["file_no"]))

        # تاريخ الرحلة: أعلى اليسار
        c.setFont("Latin", 9)
        c.drawString(28, 806, str(data["trip_date"]))

        # تاريخ التحرير
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
        return overlay_notice_fields(raw, data, legacy=False)
    raise ValueError("لا توجد استمارة رسمية معتمدة لهذا النوع في مجلد stamps/assets.")


def render_raw_rooming_pdf(excel_bytes, workdir):
    '''
    Render the uploaded Excel rooming list as a PDF using LibreOffice.

    No column meaning is interpreted:
    - no translation
    - no renaming
    - no Surname/Given-name reconstruction
    - no date normalization
    - no passport-number parsing
    - no semantic aliases

    Only temporary print settings are applied to a temporary copy so the
    visible room-list area fits on one A4 page. The uploaded file itself
    is never modified.
    '''
    workdir = Path(workdir)
    src = workdir / "rooming_source.xlsx"
    src.write_bytes(excel_bytes)

    try:
        wb = load_workbook(str(src), data_only=False)
        ws = wb.active
    except Exception as exc:
        raise ValueError(f"تعذر فتح ملف النيم ليست Excel: {exc}")

    first_row = None
    last_row = 0
    last_col = 0

    for row in ws.iter_rows():
        nonempty = 0
        row_no = row[0].row if row else 0

        for cell in row:
            if cell.value not in (None, ""):
                nonempty += 1
                last_col = max(last_col, cell.column)

        if nonempty:
            last_row = max(last_row, row_no)

        if first_row is None and nonempty >= 2:
            first_row = row_no

    if first_row is None:
        first_row = 1

    if last_row < first_row or last_col < 1:
        raise ValueError("ملف النيم ليست فارغ أو لا يحتوي بيانات.")

    from openpyxl.utils import get_column_letter
    ws.print_area = f"A{first_row}:{get_column_letter(last_col)}{last_row}"

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    from openpyxl.worksheet.page import PageMargins
    ws.page_margins = PageMargins(
        left=0.10,
        right=0.10,
        top=1.55,
        bottom=0.42,
        header=0.03,
        footer=0.03,
    )
    ws.print_options.horizontalCentered = True

    temp_xlsx = workdir / "rooming_render.xlsx"
    wb.save(str(temp_xlsx))

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice غير مثبت على الخادم.")

    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(workdir),
            str(temp_xlsx),
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=90,
        text=True,
    )

    pdf_path = workdir / "rooming_render.pdf"
    if not pdf_path.exists():
        raise RuntimeError(
            "فشل تحويل النيم ليست إلى PDF: "
            + (result.stderr or result.stdout or "")
        )

    pdf_bytes = pdf_path.read_bytes()
    reader = PdfReader(io.BytesIO(pdf_bytes))
    if len(reader.pages) < 1:
        raise ValueError("ملف النيم ليست لم ينتج صفحة صالحة.")

    return pdf_bytes


def make_room_page(data, rooming_pdf_bytes, copy_no):
    '''
    Fixed approved header + original Excel room-list rendering + declaration.
    The source room-list page is placed underneath without translating or
    rebuilding its cells.
    '''
    src_reader = PdfReader(io.BytesIO(rooming_pdf_bytes))
    if not src_reader.pages:
        raise ValueError("لا توجد صفحة روم ليست صالحة.")

    src_page = src_reader.pages[0]

    overlay_buf = io.BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=A4)
    W, H = A4
    BLACK = colors.black

    c.setFillColor(BLACK)
    draw_rtl(
        c,
        "كشف أسماء السائحين / نيم ليست",
        W - 28, H - 32, 400, 16, True
    )

    c.setFont("Latin", 8)
    copies = 1 if data.get("permit_key") == "hurghada" else 3
    c.drawString(28, H - 32, f"Copy {copy_no}/{copies}")

    draw_rtl(
        c,
        f"{data['company']['english']} | "
        f"{data['company']['license']} | "
        f"{data['file_no']} | {data['pax']} Pax",
        W - 28, H - 52, 520, 8
    )

    draw_rtl(
        c,
        f"المرشد: {data['guide']}   الهاتف: {data['phone']}",
        W - 28, H - 70, 520, 11, True
    )

    draw_rtl(
        c,
        f"تاريخ الرحلة: {data['trip_date']}",
        W - 28, H - 88, 250, 12, True
    )

    draw_rtl(
        c,
        data["permit"],
        300, H - 88, 250, 13, True
    )

    draw_rtl(
        c,
        (
            f"فندق الغردقة: {data['hotel_to']}"
            if data.get("permit_key") == "hurghada" and data.get("hotel_to")
            else f"{data['hotel_from']}  ←  {data['hotel_to']}"
        ),
        W - 28, H - 103, 520, 8
    )

    draw_rtl(
        c,
        f"إقرار من الشركة: نقر نحن شركة "
        f"{data['company']['name']} ترخيص رقم "
        f"{data['company']['license']} بأن البيانات الواردة أعلاه صحيحة.",
        W - 28, 26, 540, 7, True
    )

    c.save()
    overlay_buf.seek(0)
    overlay_page = PdfReader(overlay_buf).pages[0]

    src_page.merge_page(overlay_page)

    out = PdfWriter()
    out.add_page(src_page)

    buf = io.BytesIO()
    out.write(buf)
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
@app.post("/api/preview-rooming")
def preview_rooming():
    try:
        room = request.files.get("rooming")
        if not room or not room.filename:
            raise ValueError("ارفع ملف النيم ليست Excel")

        raw = room.read()
        if not raw:
            raise ValueError("ملف النيم ليست فارغ.")

        wb = load_workbook(io.BytesIO(raw), data_only=False, read_only=True)
        ws = wb.active

        nonempty_rows = 0
        for row in ws.iter_rows(values_only=True):
            if any(v not in (None, "") for v in row):
                nonempty_rows += 1

        return jsonify(
            ok=True,
            count=nonempty_rows,
            rows=[],
            raw_copy=True,
        )

    except Exception as e:
        return jsonify(error=str(e)), 400


@app.post("/api/generate")
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

        program = (
            f.get("program", "").strip()
            or PROGRAM_PRESETS.get((company_key, permit_key), "")
        )

        hotel_to = f.get("hotel_to", "").strip()
        if hotel_to and hotel_to not in program:
            program = program.rstrip(" -–—") + " – " + hotel_to

        if not file_no or not guide or not trip or not program:
            raise ValueError("أكمل رقم الملف والمرشد وتاريخ الرحلة والبرنامج")

        room_shot = request.files.get("rooming_img")
        room = request.files.get("rooming")
        rooming_excel_bytes = None

        if permit_key == "hurghada":
            if not room_shot or not room_shot.filename:
                raise ValueError("ارفع سكرين النيم ليست للغردقة")
        else:
            if not room or not room.filename:
                raise ValueError("ارفع ملف النيم ليست Excel")

            rooming_excel_bytes = room.read()
            if not rooming_excel_bytes:
                raise ValueError("ملف النيم ليست فارغ.")

        program_file = request.files.get("program_image")
        program_bytes = (
            program_file.read()
            if program_file and program_file.filename
            else None
        )

        extra_file = request.files.get("extra_page")
        company = COMPANIES[company_key]

        data = {
            "company": company,
            "company_key": company_key,
            "permit_key": permit_key,
            "permit": PERMITS[permit_key],
            "file_no": file_no,
            "pax": pax,
            "guide": guide,
            "phone": phone,
            "trip_date": trip,
            "issue_date": issue,
            "program": program,
            "hotel_from": f.get("hotel_from", "").strip(),
            "hotel_to": hotel_to,
            "transport_company": f.get("transport_company", "").strip(),
            "vehicle_number": f.get("vehicle_number", "").strip(),
            "driver_name": f.get("driver_name", "").strip(),
            "driver_phone": f.get("driver_phone", "").strip(),
        }

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            notices = build_notices(data, td)

            safe = re.sub(r"[^A-Za-z0-9_-]+", "_", file_no) or "permit"
            suffix = {
                "cairo": "cairo_transfer",
                "luxor_transfer": "luxor_transfer",
                "luxor_overnight": "luxor_overnight",
                "hurghada": "hurghada_transfer",
            }[permit_key]

            main_name = f"{safe}_{suffix}.pdf"

            if permit_key == "hurghada":
                if not program_bytes:
                    raise ValueError(
                        "ارفع صورة إدارة البرامج. ملف الغردقة لازم 7 ورقات"
                    )
                if not extra_file or not extra_file.filename:
                    raise ValueError(
                        "ارفع الورقة الإضافية. ملف الغردقة لازم 7 ورقات"
                    )

                extra_pdf = extra_page_from_upload(extra_file)
                police_pdf = build_police_pdf(data, td)

                room_one = (
                    make_room_shot_page(room_shot.read())
                    if room_shot
                    else None
                )

                final = merge_hurghada_pages(
                    notices,
                    room_one,
                    program_bytes,
                    police_pdf,
                    extra_pdf,
                )

            else:
                rooming_pdf_bytes = render_raw_rooming_pdf(
                    rooming_excel_bytes, td
                )

                room_pages = [
                    make_room_page(data, rooming_pdf_bytes, i)
                    for i in range(1, 4)
                ]

                final = merge_pages(
                    notices,
                    room_pages,
                    program_bytes,
                )

            return send_file(
                final,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=main_name,
            )

    except Exception as e:
        return jsonify(error=str(e)), 400
