import io
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime, date

from flask import Flask, request, send_file, jsonify, render_template
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
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
    "hurghada": "توصيلة الغردقة",
}
ALLOWED = {
    "unlimited": {"cairo", "luxor_transfer", "luxor_overnight"},
    "memory": {"cairo", "luxor_transfer", "luxor_overnight", "hurghada"},
}

DOCX_TEMPLATES = {
    ("unlimited", "cairo"): STAMPS / "ان ليميتيد ايجيبت تصريح القاهره(1).docx",
    ("memory", "cairo"): STAMPS / "تصريح ميموري القاهره الأساسي.docx",
    ("unlimited", "luxor_overnight"): STAMPS / "ان ليميتيد ايجيبت تصريح الاقصر مبيت.docx",
    ("memory", "luxor_overnight"): STAMPS / "تصريح ميموري الاقصر مبيت الأساسي.docx",
}
LEGACY_PDF_TEMPLATES = {
    ("unlimited", "luxor_transfer"): LEGACY / "unlimited_luxor_transfer.pdf",
    ("memory", "luxor_transfer"): LEGACY / "memory_luxor_transfer.pdf",
    ("memory", "hurghada"): LEGACY / "memory_hurghada.pdf",
}

ALIASES = {
    "ID": ["id", "no", "no.", "م", "رقم", "الرقم", "no"],
    "Chinese name": ["ch name", "chinese name", "الاسم بالصيني", "الاسم الصينى", "الصينى", "中文姓名", "chinese"],
    "Surname": ["surname", "family name", "اسم العائلة", "العائلة", "family"],
    "Given name": ["given name", "first name", "الاسم الأول", "الاسم الاول", "given"],
    "Sex": ["sex", "gender", "النوع", "الجنس"],
    "DOB": ["dob", "date of birth", "تاريخ الميلاد", "birth date"],
    "Passport": ["passport", "passport no", "passportno", "رقم الجواز", "جواز السفر", "passport number"],
    "Expiry": ["expiry", "passport expiry", "date of expiry", "انتهاء الجواز", "تاريخ انتهاء الجواز", "expiry date"],
    "Room": ["room", "room no", "الغرفة", "رقم الغرفة", "option", "room type"],
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


def draw_rtl(c, text, x, y, max_width=None, size=10, bold=False):
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
        # Skip obvious title/footer rows.
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


def replace_cell(cell, value):
    cell.text = str(value)
    for p in cell.paragraphs:
        for run in p.runs:
            run.font.name = "Arial"


def fill_docx_template(template_path, data, output_path):
    """Fill only known variable cells while leaving the original document tables/layout intact."""
    doc = Document(str(template_path))
    # Each approved DOCX template has 3 copies, with table groups beginning at 0/4/8.
    starts = [0, 4, 8] if len(doc.tables) >= 12 else []
    for base in starts:
        try:
            t0, t1, t3 = doc.tables[base], doc.tables[base + 1], doc.tables[base + 3]
            # Date-only header table.
            if t0.rows and t0.cells:
                if "تاريخ" in t0.cell(0, 0).text:
                    replace_cell(t0.cell(0, 0), f"تاريخ الرحلة : {data['trip_date']}")
            # Company information.
            replace_cell(t1.cell(1, 1), str(data["pax"]))
            replace_cell(t1.cell(2, 1), data["guide"])
            replace_cell(t1.cell(2, 3), data["phone"])
            # Trip section.
            replace_cell(t3.cell(0, 1), data["permit"])
            replace_cell(t3.cell(0, 4), data["program"])
            direction = "القاهرة" if data["permit_key"] == "cairo" else ("الأقصر" if data["permit_key"] in {"luxor_transfer", "luxor_overnight"} else "الغردقة")
            checkpoint = "كمين الجونة" if data["permit_key"] == "cairo" else "كمين سفاجا - قنا"
            # Preserve the original labels/check-box positions, but set the chosen text with a check marker.
            replace_cell(t3.cell(1, 1), ("☑ " if direction == "القاهرة" else "☐ ") + "القاهرة    الإسكندرية    شرم الشيخ")
            replace_cell(t3.cell(1, 3), ("☑ " if direction == "الأقصر" else "☐ ") + "الأقصر                    أسوان")
            replace_cell(t3.cell(2, 1), ("☑ " if checkpoint == "كمين الجونة" else "☐ ") + "كمين الجونة")
            replace_cell(t3.cell(2, 3), ("☑ " if checkpoint != "كمين الجونة" else "☐ ") + "كمين سفاجا - قنا       كمين مرسى علم - إدفو")
        except Exception:
            continue
    # Replace visible text in paragraphs/tables only where it is an exact placeholder/value, never the form labels.
    old_names = {"ان ليميتد ايجيبت ترافيل", "ميمورى تريب تورز"}
    for p in doc.paragraphs:
        if p.text.strip() in old_names:
            continue
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


def overlay_notice_fields(pdf_bytes, data):
    src = PdfReader(io.BytesIO(pdf_bytes))
    out = PdfWriter()
    for page in src.pages[:3]:
        ov = io.BytesIO()
        c = canvas.Canvas(ov, pagesize=A4)
        # Header overlay is intentionally placed over the blank header fields of the approved form.
        c.setFont("Latin", 9)
        c.drawRightString(575, 738, data["file_no"])
        c.drawString(60, 738, data["trip_date"])
        # Independent issue date in the approval area.
        draw_rtl(c, f"تحريراً في : {data['issue_date']}", 205, 150, 160, 8)
        c.save(); ov.seek(0)
        page.merge_page(PdfReader(ov).pages[0])
        out.add_page(page)
    buf = io.BytesIO(); out.write(buf); buf.seek(0); return buf


def build_notices(data, workdir):
    key = (data["company"], data["permit_key"])
    if key in DOCX_TEMPLATES and DOCX_TEMPLATES[key].exists():
        filled = Path(workdir) / "filled.docx"
        fill_docx_template(DOCX_TEMPLATES[key], data, filled)
        pdf = libreoffice_convert(filled, workdir)
        raw = pdf.read_bytes()
    elif key in LEGACY_PDF_TEMPLATES and LEGACY_PDF_TEMPLATES[key].exists():
        raw = LEGACY_PDF_TEMPLATES[key].read_bytes()
    else:
        raise ValueError("لا توجد استمارة رسمية معتمدة لهذا النوع في مجلد stamps/assets.")
    return overlay_notice_fields(raw, data)


def make_room_page(data, rows, copy_no):
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=A4); W, H = A4
    draw_rtl(c, "كشف أسماء السائحين / نيم ليست", W - 28, H - 32, 360, 15, True)
    c.setFont("Latin", 8); c.drawString(28, H - 32, f"Copy {copy_no}/3")
    draw_rtl(c, f"{data['company']['english']} | {data['company']['license']} | {data['file_no']} | {data['pax']} Pax", W - 28, H - 52, 520, 8)
    draw_rtl(c, f"المرشد: {data['guide']}   الهاتف: {data['phone']}   تاريخ الرحلة: {data['trip_date']}   {data['permit']}", W - 28, H - 68, 520, 8)
    draw_rtl(c, f"{data['hotel_from']}  ←  {data['hotel_to']}", W - 28, H - 84, 520, 8)
    headers = ["م", "الاسم بالصيني", "اسم العائلة", "الاسم الأول", "النوع", "تاريخ الميلاد", "رقم الجواز", "انتهاء الجواز", "الغرفة", "ملاحظات"]
    widths = [22, 92, 63, 70, 28, 62, 70, 62, 43, 47]
    total = sum(widths); x0 = (W - total) / 2; top = H - 103
    available = top - 52
    if len(rows) > 38:
        raise ValueError("عدد الأسماء كبير جداً لصفحة نيم ليست واحدة؛ راجع العدد أو استخدم قالباً معتمداً مختلفاً.")
    rh = max(10.0, min(19.0, available / max(len(rows) + 1, 1)))
    header_font = max(4.5, min(5.8, rh * 0.30))
    body_font = max(4.2, min(5.2, rh * 0.27))
    c.setLineWidth(.35); c.setFont("ArabicBold", header_font)
    x = x0
    for h, w in zip(headers, widths):
        c.rect(x, top-rh, w, rh); c.drawCentredString(x+w/2, top-(rh*0.68), ar(h)); x += w
    y = top-rh
    for r in rows:
        x = x0
        vals = [r.get("ID", ""), r.get("Chinese name", ""), r.get("Surname", ""), r.get("Given name", ""), r.get("Sex", ""), r.get("DOB", ""), r.get("Passport", ""), r.get("Expiry", ""), r.get("Room", ""), r.get("Note", "")]
        for j, (v, w) in enumerate(zip(vals, widths)):
            c.rect(x, y-rh, w, rh)
            s = str(v or "")
            if j == 1:
                c.setFont("STSong-Light", body_font)
            else:
                c.setFont("Latin", body_font)
            c.drawCentredString(x+w/2, y-(rh*0.68), s[:22])
            x += w
        y -= rh
        if y < 50:
            break
    draw_rtl(c, f"إقرار من الشركة: نقر نحن شركة {data['company']['name']} ترخيص رقم {data['company']['license']} بأن البيانات الواردة أعلاه صحيحة.", W-28, 26, 540, 7, True)
    c.save(); buf.seek(0); return buf


def make_program_page(img_bytes):
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=landscape(A4)); W, H = landscape(A4)
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    maxw, maxh = W - 24, H - 24
    scale = min(maxw/im.width, maxh/im.height)
    iw, ih = im.width*scale, im.height*scale
    tmp = io.BytesIO(); im.save(tmp, format="PNG"); tmp.seek(0)
    c.drawImage(tmp, (W-iw)/2, (H-ih)/2, width=iw, height=ih, preserveAspectRatio=True, mask="auto")
    c.save(); buf.seek(0); return buf


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


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/health")
def health():
    return jsonify(ok=True, libreoffice=bool(shutil.which("libreoffice") or shutil.which("soffice")))


@app.post("/api/generate")
def generate():
    try:
        f = request.form
        company_key = f.get("company", "unlimited")
        permit_key = f.get("permit", "cairo")
        if company_key not in COMPANIES or permit_key not in ALLOWED[company_key]:
            raise ValueError("الشركة أو نوع التصريح غير صحيح")
        pax = int(f.get("pax", "0"))
        if pax < 1: raise ValueError("أدخل عدد السائحين")
        file_no = f.get("file_no", "").strip()
        guide = f.get("guide", "").strip()
        phone = f.get("phone", "").strip()
        trip = fmt_date(f.get("trip_date", ""))
        issue = fmt_date(f.get("issue_date", ""))
        program = f.get("program", "").strip()
        hotel_to = f.get("hotel_to", "").strip()
        if hotel_to and hotel_to not in program:
            program = program.rstrip(' -–—') + ' – ' + hotel_to
        if not file_no or not guide or not trip or not program:
            raise ValueError("أكمل رقم الملف والمرشد وتاريخ الرحلة والبرنامج")
        room = request.files.get("rooming")
        if not room or not room.filename: raise ValueError("ارفع ملف النيم ليست Excel")
        rows = read_rooming(io.BytesIO(room.read()))
        if len(rows) != pax:
            raise ValueError(f"العدد المعلن {pax} لا يساوي عدد الأسماء في النيم ليست ({len(rows)})")
        program_file = request.files.get("program_image")
        program_bytes = program_file.read() if program_file and program_file.filename else None
        company = COMPANIES[company_key]
        data = {"company": company, "company_key": company_key, "permit_key": permit_key, "permit": PERMITS[permit_key], "file_no": file_no, "pax": pax, "guide": guide, "phone": phone, "trip_date": trip, "issue_date": issue, "program": program, "hotel_from": f.get("hotel_from", "").strip(), "hotel_to": hotel_to}
        with tempfile.TemporaryDirectory() as td:
            notices = build_notices(data, td)
            room_pages = [make_room_page(data, rows, i) for i in range(1,4)]
            final = merge_pages(notices, room_pages, program_bytes)
            safe = re.sub(r"[^A-Za-z0-9_-]+", "_", file_no) or "permit"
            suffix = {"cairo":"cairo_transfer","luxor_transfer":"luxor_transfer","luxor_overnight":"luxor_overnight","hurghada":"hurghada_transfer"}[permit_key]
            return send_file(final, mimetype="application/pdf", as_attachment=True, download_name=f"{safe}_{suffix}.pdf")
    except Exception as e:
        return jsonify(error=str(e)), 400
