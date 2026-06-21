import os
import re
import tempfile
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
import requests
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse

# PDF
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfgen import canvas
import qrcode

processed_message_ids = set()
KAZAN_SESSIONS = {}

app = FastAPI()

VERIFY_TOKEN = os.getenv("VERIFY_TOKEN", "pramid_verify_2026")
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_NUMBER_ID = os.getenv("PHONE_NUMBER_ID")

RADYATOR_FILE = "fiyat_listesi.xlsx"
KOMBI_FILE = "kombi_listesi.xlsx"
KAZAN_FILE = os.getenv("KAZAN_FILE", "KAZAN_BOT_AB_BACA_DUZENLI.xlsx")
VAT_RATE = 0.20
GRAPH_VERSION = "v20.0"


@app.get("/")
def home():
    return {"status": "PRAMID WhatsApp Bot çalışıyor"}


@app.get("/webhook")
def verify_webhook(request: Request):
    params = request.query_params
    if params.get("hub.mode") == "subscribe" and params.get("hub.verify_token") == VERIFY_TOKEN:
        return PlainTextResponse(params.get("hub.challenge"))
    return PlainTextResponse("Forbidden", status_code=403)


@app.post("/webhook")
async def receive_message(request: Request):
    data = await request.json()
    print("Gelen veri:", data, flush=True)

    try:
        value = data["entry"][0]["changes"][0]["value"]
        if "messages" not in value:
            print("Mesaj değil, durum bildirimi geldi.", flush=True)
            return {"status": "ok"}

        message = value["messages"][0]
        message_id = message.get("id")
        if message_id in processed_message_ids:
            print("Tekrarlanan mesaj, cevap verilmedi:", message_id, flush=True)
            return {"status": "ok"}
        processed_message_ids.add(message_id)

        sender = message["from"]
        text = message.get("text", {}).get("body", "")

        result = create_reply(sender, text)
        if isinstance(result, dict):
            if result.get("text"):
                send_whatsapp_message(sender, result["text"])
            if result.get("document_path"):
                send_whatsapp_document(sender, result["document_path"], result.get("filename", "pramid-teklif.pdf"), result.get("caption", ""))
        else:
            send_whatsapp_message(sender, result)

    except Exception as e:
        print("Hata:", e, flush=True)

    return {"status": "ok"}


# -------------------------
# Ortak yardımcı fonksiyonlar
# -------------------------

def normalize_text(text):
    text = str(text).lower()
    replacements = {
        "ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c",
        "kw": " kw ", "kW": " kw "
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def money_tl(value):
    return f"{int(round(value)):,}".replace(",", ".") + " TL"


def money_eur(value):
    s = f"{float(value):,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"


def to_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.replace("€", "").replace("TL", "").strip().replace(".", "").replace(",", ".")
        return float(value)
    except Exception:
        return default


def extract_quantity(text):
    clean = normalize_text(text)
    patterns = [r"(\d+)\s*(adet|ad|tane|tn)", r"(adet|ad|tane|tn)\s*(\d+)"]
    for pattern in patterns:
        match = re.search(pattern, clean)
        if match:
            nums = re.findall(r"\d+", match.group(0))
            if nums:
                return int(nums[0]), clean.replace(match.group(0), " ")

    nums = [int(n) for n in re.findall(r"\d+", clean)]
    if len(nums) >= 3:
        if nums[0] <= 50:
            return nums[0], clean
        if nums[-1] <= 50:
            return nums[-1], clean
    return 1, clean


# -------------------------
# Radyatör
# -------------------------

def normalize_dimension(n):
    n = int(n)
    if n in [50, 60, 90]:
        return n * 10
    if n in [100, 120, 140, 160, 180, 200]:
        return n * 10
    return n


def find_radiator_measure(line):
    qty, clean = extract_quantity(line)
    clean = clean.replace("x", " ").replace("*", " ").replace("/", " ").replace(" a ", " ")
    nums = [int(n) for n in re.findall(r"\d+", clean)]
    if len(nums) >= 3:
        if nums[0] == qty:
            nums = nums[1:]
        elif nums[-1] == qty:
            nums = nums[:-1]
    if len(nums) < 2:
        return None, qty
    d1 = normalize_dimension(nums[0])
    d2 = normalize_dimension(nums[1])
    return f"{min(d1, d2)}X{max(d1, d2)}", qty


def load_radiators():
    df = pd.read_excel(RADYATOR_FILE)
    df.columns = [str(c).strip().upper() for c in df.columns]
    products = []
    for _, row in df.iterrows():
        olcu = str(row["ÖLÇÜ"]).strip().upper().replace(" ", "")
        liste = to_float(row["LİSTE FİYATI"])
        nakit_iskonto = to_float(row["NAKİT İSKONTO"])
        kart_iskonto = to_float(row["KART İSKONTO"])
        products.append({
            "olcu": olcu,
            "nakit": round(liste * (1 - nakit_iskonto / 100)),
            "kart": round(liste * (1 - kart_iskonto / 100)),
        })
    return products


# -------------------------
# Kombi
# -------------------------

def load_kombis():
    df = pd.read_excel(KOMBI_FILE)
    df.columns = [str(c).strip().upper() for c in df.columns]
    products = []
    for _, row in df.iterrows():
        name = str(row["ÜRÜN ADI"]).strip()
        kw = str(row["ÜRÜN KW"]).strip()
        full_name = f"{name} {kw} KW"
        products.append({
            "name": full_name,
            "search": normalize_text(full_name),
            "nakit": to_float(row["NAKİT FİYAT"]),
            "kart": to_float(row["KART FİYAT"]),
        })
    return products


def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def find_kombi(text):
    qty, clean = extract_quantity(text)
    try:
        kombis = load_kombis()
    except Exception as e:
        print("Kombi Excel okunamadı:", e, flush=True)
        return None, qty

    best_product = None
    best_score = 0
    for product in kombis:
        score = similarity(clean, product["search"])
        words = product["search"].split()
        score += sum(1 for w in words if w in clean) * 0.12
        if score > best_score:
            best_score = score
            best_product = product
    if best_score >= 0.45:
        return best_product, qty
    return None, qty


# -------------------------
# Kazan botu
# -------------------------

ALLOWED_KW = [50, 65, 100, 125, 150]


def is_yes(text):
    clean = normalize_text(text)
    return clean in ["evet", "e", "var", "yes", "1"] or "evet" in clean or "var" in clean


def is_no(text):
    clean = normalize_text(text)
    return clean in ["hayir", "h", "yok", "no", "2"] or "hayir" in clean or "yok" in clean


def is_kazan_message(text):
    clean = normalize_text(text)
    if any(w in clean for w in ["kazan", "felis", "eca felis", "fl"]):
        return True
    nums = [int(n) for n in re.findall(r"\d+", clean)]
    return " kw " in f" {clean} " and any(n in ALLOWED_KW for n in nums)


def parse_kazan_request(text):
    clean = normalize_text(text)
    nums = [int(n) for n in re.findall(r"\d+", clean)]
    kw = None

    # Önce "125 kw" gibi açık yazımları yakala
    kw_match = re.search(r"(50|65|100|125|150)\s*kw", clean)
    if kw_match:
        kw = int(kw_match.group(1))
    else:
        for n in nums:
            if n in ALLOWED_KW:
                kw = n
                break

    qty, _ = extract_quantity(text)

    # "2 125 kazan" gibi yazımlarda adet genelde güç olmayan sayı olur
    if qty == 1 and kw is not None:
        non_kw_nums = [n for n in nums if n != kw]
        if non_kw_nums:
            qty_candidate = non_kw_nums[0]
            if 1 <= qty_candidate <= 100:
                qty = qty_candidate

    return qty, kw


def norm_cols(df):
    # Sütun adlarını elle değiştirmiyoruz; sadece baş/son boşlukları temizliyoruz.
    # Böylece ListeFiyatı / Liste Fiyatı gibi Excel yazım farklarını yardımcı fonksiyonlarla yakalayabiliyoruz.
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df



def normalize_header_name(name):
    text = str(name).strip().upper()
    tr_map = str.maketrans({
        "İ": "I", "I": "I", "Ş": "S", "Ğ": "G", "Ü": "U", "Ö": "O", "Ç": "C",
        "ı": "I", "ş": "S", "ğ": "G", "ü": "U", "ö": "O", "ç": "C",
    })
    text = text.translate(tr_map)
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def col_value(row, names, default=None):
    wanted = {normalize_header_name(n) for n in names}
    for col in row.index:
        if normalize_header_name(col) in wanted:
            value = row[col]
            if pd.isna(value):
                return default
            return value
    return default


def df_filter_eq(df, col_names, expected):
    wanted = {normalize_header_name(n) for n in col_names}
    real_col = None
    for col in df.columns:
        if normalize_header_name(col) in wanted:
            real_col = col
            break
    if real_col is None:
        raise KeyError(f"Excel sütunu bulunamadı: {', '.join(col_names)}")
    return df[df[real_col].astype(str).str.upper().str.strip() == str(expected).upper().strip()]


def load_kazan_excel():
    xls = pd.ExcelFile(KAZAN_FILE)
    data = {
        "kazanlar": norm_cols(pd.read_excel(xls, "KAZANLAR")),
        "aksesuarlar": norm_cols(pd.read_excel(xls, "AKSESUARLAR")),
        "pompalar": norm_cols(pd.read_excel(xls, "POMPA_SETLERİ")),
        "iskontolar": norm_cols(pd.read_excel(xls, "ISKONTOLAR")),
        "carpan": norm_cols(pd.read_excel(xls, "EKIPMAN_CARPANI")),
        "kurallar": norm_cols(pd.read_excel(xls, "KURALLAR")),
    }
    return data


def get_discount(data, tip):
    df = data["iskontolar"]
    row = df_filter_eq(df, ["TİP", "TIP"], tip)
    if row.empty:
        return 0.0
    return to_float(col_value(row.iloc[0], ["İSKONTO (%)", "ISKONTO (%)", "İskonto (%)"], 0))


def get_equipment_multiplier(data, kazan_adedi):
    for _, row in data["carpan"].iterrows():
        mn = int(to_float(col_value(row, ["MİN KAZAN", "MIN KAZAN"], 0)))
        mx = int(to_float(col_value(row, ["MAX KAZAN"], 0)))
        if mn <= kazan_adedi <= mx:
            return int(to_float(col_value(row, ["ÇARPAN", "CARPAN"], 1), 1))
    return 1


def get_rule_products(data, rule_name):
    df = data["kurallar"]
    rows = df_filter_eq(df, ["KURAL"], rule_name)
    urun_col = next((c for c in rows.columns if normalize_header_name(c) == normalize_header_name("ÜRÜN")), None)
    if urun_col is None:
        return []
    return [str(v).strip() for v in rows[urun_col].tolist() if str(v).strip() and not pd.isna(v)]


def get_accessory(data, key):
    df = data["aksesuarlar"]
    row = df_filter_eq(df, ["ANAHTAR"], key)
    if row.empty:
        raise ValueError(f"Aksesuar bulunamadı: {key}")
    row = row.iloc[0]
    return {
        "code": str(col_value(row, ["ÜRÜN KODU", "URUN KODU"], str(key))).strip(),
        "short": str(key),
        "name": str(col_value(row, ["ÜRÜN ADI", "URUN ADI"], "")).strip(),
        "list_price": to_float(col_value(row, ["LİSTE FİYATI (€)", "LISTE FIYATI (€)", "ListeFiyatı (€)", "Liste Fiyatı (€)"], 0)),
    }


def get_pump(data, kw):
    df = data["pompalar"]
    if kw in [50, 65]:
        wanted = "50/65"
    elif kw in [100, 125]:
        wanted = "100/125"
    else:
        wanted = "150"
    row = df_filter_eq(df, ["GRUP"], wanted)
    if row.empty:
        raise ValueError(f"Pompa seti bulunamadı: {wanted}")
    row = row.iloc[0]
    return {
        "code": str(col_value(row, ["ÜRÜN KODU", "URUN KODU"], "")).strip(),
        "short": f"POMPA {wanted}",
        "name": str(col_value(row, ["ÜRÜN ADI", "URUN ADI"], "")).strip(),
        "list_price": to_float(col_value(row, ["LİSTE FİYATI (€)", "LISTE FIYATI (€)", "ListeFiyatı (€)", "Liste Fiyatı (€)"], 0)),
    }


def add_line(items, product, qty):
    qty = int(qty)
    if qty <= 0:
        return
    items.append({
        "code": product.get("code", ""),
        "short": product.get("short", ""),
        "name": product["name"],
        "qty": qty,
        "list_unit": float(product["list_price"]),
    })


def build_kazan_quote(kazan_adedi, kw, hot_water, has_boiler):
    data = load_kazan_excel()
    kazanlar = data["kazanlar"]
    guc_col = next((c for c in kazanlar.columns if normalize_header_name(c) in [normalize_header_name("GÜÇ (KW)"), normalize_header_name("GUC (KW)")]), None)
    if guc_col is None:
        raise KeyError("KAZANLAR sayfasında Güç (kW) sütunu bulunamadı.")
    row = kazanlar[kazanlar[guc_col].astype(int) == int(kw)]
    if row.empty:
        raise ValueError(f"{kw} kW kazan Excel'de bulunamadı.")
    row = row.iloc[0]

    nakit_disc = get_discount(data, "NAKIT")
    kart_disc = get_discount(data, "KART")
    multiplier = get_equipment_multiplier(data, kazan_adedi)

    items = []
    warnings = []

    add_line(items, {
        "code": str(col_value(row, ["ÜRÜN KODU", "URUN KODU"], "")).strip(),
        "short": str(col_value(row, ["ÜRÜN KISALTMASI", "URUN KISALTMASI"], "")).strip(),
        "name": str(col_value(row, ["ÜRÜN ADI", "URUN ADI"], "")).strip(),
        "list_price": to_float(col_value(row, ["LİSTEFİYATI (€)", "LISTEFIYATI (€)", "LİSTE FİYATI (€)", "LISTE FIYATI (€)", "ListeFiyatı (€)", "Liste Fiyatı (€)"], 0)),
    }, kazan_adedi)

    # Pompa seti kazan adedi kadar eklenir, ekipman çarpanı uygulanmaz.
    add_line(items, get_pump(data, kw), kazan_adedi)

    # Her teklifte gelen ekipmanlar 16 kazanlık çarpana göre artar.
    for key in get_rule_products(data, "HER_TEKLIF"):
        add_line(items, get_accessory(data, key), multiplier)

    # Opsiyonel ekipmanlar: 16 kazanlık çarpanla eklenir.
    # İstersen burada multiplier yerine 1 yazarak sabit yapabiliriz.
    if hot_water:
        for key in get_rule_products(data, "SICAK_SU"):
            add_line(items, get_accessory(data, key), multiplier)

    if has_boiler:
        for key in get_rule_products(data, "BOYLER_VAR"):
            add_line(items, get_accessory(data, key), multiplier)

    if kazan_adedi == 1:
        rule = "TEK_KAZAN_50_65" if kw in [50, 65] else "TEK_KAZAN_100_125_150"
        for key in get_rule_products(data, rule):
            p = get_accessory(data, key)
            if p["list_price"] <= 0:
                warnings.append(f"{key} fiyatı Excel'de boş görünüyor.")
            add_line(items, p, 1)
    else:
        for key in get_rule_products(data, "COKLU_KAZAN"):
            add_line(items, get_accessory(data, key), kazan_adedi)

    for item in items:
        item["list_total"] = item["list_unit"] * item["qty"]
        item["nakit_unit_kdv"] = item["list_unit"] * (1 - nakit_disc / 100) * (1 + VAT_RATE)
        item["kart_unit_kdv"] = item["list_unit"] * (1 - kart_disc / 100) * (1 + VAT_RATE)
        item["nakit_total_kdv"] = item["nakit_unit_kdv"] * item["qty"]
        item["kart_total_kdv"] = item["kart_unit_kdv"] * item["qty"]

    total_list = sum(i["list_total"] for i in items)

    total_nakit_haric = sum(
        i["list_unit"] * (1 - nakit_disc / 100) * i["qty"]
        for i in items
    )
    total_kart_haric = sum(
        i["list_unit"] * (1 - kart_disc / 100) * i["qty"]
        for i in items
    )

    total_nakit = total_nakit_haric * (1 + VAT_RATE)
    total_kart = total_kart_haric * (1 + VAT_RATE)

    return {
        "kw": kw,
        "kazan_adedi": kazan_adedi,
        "hot_water": hot_water,
        "has_boiler": has_boiler,
        "multiplier": multiplier,
        "items": items,
        "nakit_disc": nakit_disc,
        "kart_disc": kart_disc,
        "total_list": total_list,
        "total_nakit_haric": total_nakit_haric,
        "total_kart_haric": total_kart_haric,
        "total_nakit": total_nakit,
        "total_kart": total_kart,
        "warnings": warnings,
    }


def short_kazan_text(quote):
    lines = [
        "Kazan teklifiniz hazırlanmıştır.",
        f"Kazan: {quote['kazan_adedi']} adet {quote['kw']} kW E.C.A. Felis",
        f"Nakit Toplam: {money_eur(quote['total_nakit'])} KDV Dahil",
        f"Kart Toplam: {money_eur(quote['total_kart'])} KDV Dahil",
    ]
    if quote["warnings"]:
        lines.append("")
        lines.append("Not: " + " ".join(quote["warnings"]))
    return "\n".join(lines)


def setup_pdf_fonts():
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    if all(Path(p).exists() for p in candidates):
        pdfmetrics.registerFont(TTFont("DejaVu", candidates[0]))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", candidates[1]))
        return "DejaVu", "DejaVu-Bold"
    return "Helvetica", "Helvetica-Bold"



class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
        self.setTitle("PRAMID İNŞAAT FİYAT TEKLİFİ")
        self.setAuthor("PRAMID İNŞAAT")
        self.setSubject("Fiyat Teklifi")
        self.setCreator("PRAMID WhatsApp Fiyatlandırma Botu")

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)

        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.setFont("Helvetica", 7)
            self.setFillColor(colors.HexColor("#6B7280"))
            self.drawCentredString(105 * mm, 8 * mm, f"{self._pageNumber}/{total_pages}")
            super().showPage()

        super().save()


def _safe_code(item):
    code = str(item.get("code", "")).strip()
    if code and code.lower() != "nan":
        return code
    return str(item.get("short", "")).strip() or "-"


def _draw_pdf_background(c, document, teklif_no, regular_font, navy, orange, text_muted, qr_label="QR teklif doğrulama"):
    c.saveState()

    # Üst kurumsal bant
    c.setFillColor(navy)
    c.rect(0, 287 * mm, 210 * mm, 10 * mm, fill=1, stroke=0)

    c.setFillColor(orange)
    c.rect(0, 285.8 * mm, 210 * mm, 1.2 * mm, fill=1, stroke=0)

    # QR kod
    try:
        qr_text = f"PRAMID Teklif No: {teklif_no}"
        qr_img = qrcode.make(qr_text)
        qr_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        qr_file.close()
        qr_img.save(qr_file.name)
        c.drawImage(
            qr_file.name,
            178 * mm,
            17 * mm,
            width=20 * mm,
            height=20 * mm,
            preserveAspectRatio=True,
            mask="auto",
        )
        try:
            os.remove(qr_file.name)
        except Exception:
            pass
    except Exception:
        pass

    c.setFont(regular_font, 6.5)
    c.setFillColor(text_muted)
    c.drawRightString(198 * mm, 14 * mm, qr_label)

    c.restoreState()


def generate_kazan_pdf(quote):
    regular_font, bold_font = setup_pdf_fonts()

    filename = f"pramid-kazan-teklifi-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
    path = os.path.join(tempfile.gettempdir(), filename)

    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=14 * mm,
        bottomMargin=16 * mm,
    )

    navy = colors.HexColor("#0F2742")
    orange = colors.HexColor("#F97316")
    light_bg = colors.HexColor("#F8FAFC")
    mid_bg = colors.HexColor("#EEF2F7")
    border = colors.HexColor("#D8DEE8")
    text_dark = colors.HexColor("#111827")
    text_muted = colors.HexColor("#6B7280")
    white = colors.white

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "PramidTitle",
        parent=styles["Title"],
        fontName=bold_font,
        fontSize=20,
        leading=24,
        textColor=navy,
        alignment=TA_RIGHT,
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        "PramidSubtitle",
        parent=styles["Normal"],
        fontName=bold_font,
        fontSize=12,
        leading=15,
        textColor=orange,
        alignment=TA_RIGHT,
    )

    normal = ParagraphStyle(
        "NormalTR",
        parent=styles["Normal"],
        fontName=regular_font,
        fontSize=8.5,
        leading=11,
        textColor=text_dark,
    )

    small = ParagraphStyle(
        "SmallTR",
        parent=normal,
        fontSize=7.5,
        leading=10,
        textColor=text_muted,
    )

    bold = ParagraphStyle(
        "BoldTR",
        parent=normal,
        fontName=bold_font,
        textColor=text_dark,
    )

    table_head = ParagraphStyle(
        "TableHead",
        parent=normal,
        fontName=bold_font,
        fontSize=8,
        leading=10,
        textColor=white,
        alignment=TA_CENTER,
    )

    cell = ParagraphStyle(
        "Cell",
        parent=normal,
        fontSize=7.8,
        leading=10,
    )

    cell_right = ParagraphStyle(
        "CellRight",
        parent=cell,
        alignment=TA_RIGHT,
    )

    summary_label = ParagraphStyle(
        "SummaryLabel",
        parent=normal,
        fontName=bold_font,
        fontSize=8,
        leading=10,
        textColor=text_muted,
    )

    summary_value = ParagraphStyle(
        "SummaryValue",
        parent=normal,
        fontName=bold_font,
        fontSize=12,
        leading=15,
        textColor=navy,
        alignment=TA_RIGHT,
    )

    story = []

    now = datetime.now()
    teklif_no = f"PRM-KZN-{now.strftime('%Y%m%d')}-{now.strftime('%H%M%S')}"
    tarih = now.strftime("%d.%m.%Y")

    header = Table(
        [[
            Paragraph(
                "<b>PRAMID</b><br/><font size='8'>İNŞAAT</font>",
                ParagraphStyle(
                    "LogoText",
                    parent=normal,
                    fontName=bold_font,
                    fontSize=18,
                    leading=18,
                    textColor=navy,
                )
            ),
            [
                Paragraph("PRAMID İNŞAAT", title_style),
                Paragraph("FİYAT TEKLİFİ", subtitle_style),
                Paragraph(f"Teklif No: {teklif_no}<br/>Tarih: {tarih}", small),
            ],
        ]],
        colWidths=[65 * mm, 120 * mm],
    )

    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(header)

    story.append(Table(
        [[""]],
        colWidths=[185 * mm],
        rowHeights=[2.2 * mm],
        style=TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), orange),
        ])
    ))

    story.append(Spacer(1, 5 * mm))

    customer_box = Table(
        [[
            Paragraph(
                "<b>TEKLİF BİLGİLERİ</b><br/>"
                f"E.C.A. Felis Kazan Teklifi<br/>"
                f"{quote['kazan_adedi']} adet {quote['kw']} kW<br/>"
                f"Kullanım: {'Isıtma + Sıcak su' if quote['hot_water'] else 'Isıtma'}<br/>"
                f"Boyler: {'Var' if quote['has_boiler'] else 'Yok'}",
                normal
            ),
            Paragraph(
                "<b>FİYAT DETAYI</b><br/>"
                f"KDV Oranı: %{int(VAT_RATE * 100)}<br/>"
                f"Nakit İskonto: %{quote['nakit_disc']:.0f}<br/>"
                f"Kart İskonto: %{quote['kart_disc']:.0f}<br/>"
                "Para Birimi: Euro (€)",
                normal
            )
        ]],
        colWidths=[115 * mm, 70 * mm],
    )

    customer_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), light_bg),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(customer_box)
    story.append(Spacer(1, 6 * mm))

    data = [[
        Paragraph("Kod", table_head),
        Paragraph("Ürün Adı", table_head),
        Paragraph("Adet", table_head),
        Paragraph("Birim Fiyat (€)", table_head),
        Paragraph("Toplam Fiyat (€)", table_head),
    ]]

    for item in quote["items"]:
        data.append([
            Paragraph(_safe_code(item), cell),
            Paragraph(str(item["name"]), cell),
            Paragraph(str(item["qty"]), cell_right),
            Paragraph(money_eur(item["list_unit"]), cell_right),
            Paragraph(money_eur(item["list_total"]), cell_right),
        ])

    table = Table(
        data,
        colWidths=[27 * mm, 83 * mm, 15 * mm, 30 * mm, 30 * mm],
        repeatRows=1,
    )

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, 0), bold_font),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, light_bg]),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, orange),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story.append(table)
    story.append(Spacer(1, 7 * mm))

    summary_data = [
        [Paragraph("Nakit Toplam", summary_label), Paragraph(money_eur(quote["total_nakit_haric"]), summary_value)],
        [Paragraph("Nakit KDV Dahil Toplam", summary_label), Paragraph(money_eur(quote["total_nakit"]), summary_value)],
        [Paragraph("Kart Toplam", summary_label), Paragraph(money_eur(quote["total_kart_haric"]), summary_value)],
        [Paragraph("Kart KDV Dahil Toplam", summary_label), Paragraph(money_eur(quote["total_kart"]), summary_value)],
    ]

    summary_table = Table(summary_data, colWidths=[65 * mm, 45 * mm], hAlign="RIGHT")

    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), mid_bg),
        ("BOX", (0, 0), (-1, -1), 0.8, navy),
        ("LINEBELOW", (0, 0), (-1, 2), 0.4, border),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(summary_table)

    if quote["warnings"]:
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("Not: " + " ".join(quote["warnings"]), small))

    story.append(Spacer(1, 6 * mm))

    footer_text = (
        "Bu teklif Euro (€) bazlıdır. KDV oranı %20 olarak hesaplanmıştır. "
        "Teklif geçerlilik süresi 7 gündür. Nihai stok ve fiyat teyidi için PRAMID İnşaat ile görüşünüz.<br/>"
        "PRAMID İNŞAAT | WhatsApp Fiyatlandırma Sistemi"
    )

    story.append(Paragraph(footer_text, small))

    def page_canvas(c, document):
        _draw_pdf_background(
            c,
            document,
            teklif_no,
            regular_font,
            navy,
            orange,
            text_muted,
            "QR teklif doğrulama / arşiv kodu",
        )

    doc.build(
        story,
        onFirstPage=page_canvas,
        onLaterPages=page_canvas,
        canvasmaker=NumberedCanvas,
    )

    return path


def generate_radyator_pdf(quote):
    regular_font, bold_font = setup_pdf_fonts()

    filename = f"pramid-radyator-teklifi-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
    path = os.path.join(tempfile.gettempdir(), filename)

    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=14 * mm,
        bottomMargin=16 * mm,
    )

    navy = colors.HexColor("#0F2742")
    orange = colors.HexColor("#F97316")
    light_bg = colors.HexColor("#F8FAFC")
    mid_bg = colors.HexColor("#EEF2F7")
    border = colors.HexColor("#D8DEE8")
    text_dark = colors.HexColor("#111827")
    text_muted = colors.HexColor("#6B7280")
    white = colors.white

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "RadTitle",
        parent=styles["Title"],
        fontName=bold_font,
        fontSize=20,
        leading=24,
        textColor=navy,
        alignment=TA_RIGHT,
        spaceAfter=2,
    )

    subtitle_style = ParagraphStyle(
        "RadSubtitle",
        parent=styles["Normal"],
        fontName=bold_font,
        fontSize=12,
        leading=15,
        textColor=orange,
        alignment=TA_RIGHT,
    )

    normal = ParagraphStyle(
        "RadNormal",
        parent=styles["Normal"],
        fontName=regular_font,
        fontSize=8.5,
        leading=11,
        textColor=text_dark,
    )

    small = ParagraphStyle(
        "RadSmall",
        parent=normal,
        fontSize=7.5,
        leading=10,
        textColor=text_muted,
    )

    table_head = ParagraphStyle(
        "RadTableHead",
        parent=normal,
        fontName=bold_font,
        fontSize=8,
        leading=10,
        textColor=white,
        alignment=TA_CENTER,
    )

    cell = ParagraphStyle(
        "RadCell",
        parent=normal,
        fontSize=7.8,
        leading=10,
    )

    cell_right = ParagraphStyle(
        "RadCellRight",
        parent=cell,
        alignment=TA_RIGHT,
    )

    summary_label = ParagraphStyle(
        "RadSummaryLabel",
        parent=normal,
        fontName=bold_font,
        fontSize=8,
        leading=10,
        textColor=text_muted,
    )

    summary_value = ParagraphStyle(
        "RadSummaryValue",
        parent=normal,
        fontName=bold_font,
        fontSize=12,
        leading=15,
        textColor=navy,
        alignment=TA_RIGHT,
    )

    story = []

    now = datetime.now()
    teklif_no = f"PRM-RAD-{now.strftime('%Y%m%d')}-{now.strftime('%H%M%S')}"
    tarih = now.strftime("%d.%m.%Y")

    header = Table(
        [[
            Paragraph(
                "<b>PRAMID</b><br/><font size='8'>İNŞAAT</font>",
                ParagraphStyle(
                    "RadLogoText",
                    parent=normal,
                    fontName=bold_font,
                    fontSize=18,
                    leading=18,
                    textColor=navy,
                )
            ),
            [
                Paragraph("PRAMID İNŞAAT", title_style),
                Paragraph("FİYAT TEKLİFİ", subtitle_style),
                Paragraph(f"Teklif No: {teklif_no}<br/>Tarih: {tarih}", small),
            ],
        ]],
        colWidths=[65 * mm, 120 * mm],
    )

    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    story.append(header)

    story.append(Table(
        [[""]],
        colWidths=[185 * mm],
        rowHeights=[2.2 * mm],
        style=TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), orange),
        ])
    ))

    story.append(Spacer(1, 5 * mm))

    info_box = Table(
        [[
            Paragraph(
                "<b>TEKLİF BİLGİLERİ</b><br/>"
                f"Ürün Grubu: {quote.get('title', 'Radyatör Teklifi')}<br/>"
                f"Kalem Sayısı: {len(quote.get('items', []))}<br/>"
                "Para Birimi: TL",
                normal
            ),
            Paragraph(
                "<b>FİYAT DETAYI</b><br/>"
                "Fiyatlar KDV dahil hesaplanmıştır.<br/>"
                "Nakit ve kart toplamları ayrı gösterilmiştir.<br/>"
                "Teklif geçerlilik süresi: 7 gün",
                normal
            )
        ]],
        colWidths=[115 * mm, 70 * mm],
    )

    info_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), light_bg),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(info_box)
    story.append(Spacer(1, 6 * mm))

    data = [[
        Paragraph("Ürün", table_head),
        Paragraph("Adet", table_head),
        Paragraph("Nakit Birim", table_head),
        Paragraph("Kart Birim", table_head),
        Paragraph("Nakit Toplam", table_head),
        Paragraph("Kart Toplam", table_head),
    ]]

    for item in quote.get("items", []):
        data.append([
            Paragraph(str(item.get("name", "")), cell),
            Paragraph(str(item.get("qty", "")), cell_right),
            Paragraph(money_tl(item.get("nakit_unit", 0)), cell_right),
            Paragraph(money_tl(item.get("kart_unit", 0)), cell_right),
            Paragraph(money_tl(item.get("nakit_total", 0)), cell_right),
            Paragraph(money_tl(item.get("kart_total", 0)), cell_right),
        ])

    table = Table(
        data,
        colWidths=[63 * mm, 14 * mm, 27 * mm, 27 * mm, 27 * mm, 27 * mm],
        repeatRows=1,
    )

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), white),
        ("FONTNAME", (0, 0), (-1, 0), bold_font),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, light_bg]),
        ("BOX", (0, 0), (-1, -1), 0.6, border),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, orange),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story.append(table)
    story.append(Spacer(1, 7 * mm))

    summary_data = [
        [Paragraph("Nakit KDV Dahil Toplam", summary_label), Paragraph(money_tl(quote.get("total_nakit", 0)), summary_value)],
        [Paragraph("Kart KDV Dahil Toplam", summary_label), Paragraph(money_tl(quote.get("total_kart", 0)), summary_value)],
    ]

    summary_table = Table(summary_data, colWidths=[65 * mm, 45 * mm], hAlign="RIGHT")

    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), mid_bg),
        ("BOX", (0, 0), (-1, -1), 0.8, navy),
        ("LINEBELOW", (0, 0), (-1, 0), 0.4, border),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    story.append(summary_table)

    not_found = quote.get("not_found", [])
    if not_found:
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("Bulunamayan ölçüler: " + ", ".join(not_found), small))

    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(
        "Bu teklif TL bazlıdır ve KDV dahil fiyatlar üzerinden hazırlanmıştır. "
        "Teklif geçerlilik süresi 7 gündür. Nihai stok ve fiyat teyidi için PRAMID İnşaat ile görüşünüz.<br/>"
        "PRAMID İNŞAAT | WhatsApp Fiyatlandırma Sistemi",
        small
    ))

    def page_canvas(c, document):
        _draw_pdf_background(
            c,
            document,
            teklif_no,
            regular_font,
            navy,
            orange,
            text_muted,
            "QR teklif doğrulama / arşiv kodu",
        )

    doc.build(
        story,
        onFirstPage=page_canvas,
        onLaterPages=page_canvas,
        canvasmaker=NumberedCanvas,
    )

    return path

def start_kazan_flow(sender, text):
    qty, kw = parse_kazan_request(text)
    KAZAN_SESSIONS[sender] = {"step": "ASK_KAZAN", "qty": qty, "kw": kw}
    if not kw:
        return "Kazan teklifi için kazan adedi ve gücünü yazınız.\nÖrnek: 2 adet 125 kW"
    KAZAN_SESSIONS[sender]["step"] = "ASK_USE"
    return "Kullanım amacı nedir?\n1 - Sadece ısıtma\n2 - Isıtma + sıcak su"


def continue_kazan_flow(sender, text):
    state = KAZAN_SESSIONS.get(sender, {})
    step = state.get("step")
    clean = normalize_text(text)

    if clean in ["iptal", "vazgec", "cancel"]:
        KAZAN_SESSIONS.pop(sender, None)
        return "Kazan teklif işlemi iptal edildi."

    if step == "ASK_KAZAN":
        qty, kw = parse_kazan_request(text)
        if not kw:
            return "Kazan gücünü anlayamadım. Lütfen örnek gibi yazınız: 2 adet 125 kW"
        state["qty"] = qty
        state["kw"] = kw
        state["step"] = "ASK_USE"
        return "Kullanım amacı nedir?\n1 - Sadece ısıtma\n2 - Isıtma + sıcak su"

    if step == "ASK_USE":
        if "2" in clean or "sicak" in clean or "sıcak" in text.lower():
            state["hot_water"] = True
        elif "1" in clean or "isitma" in clean:
            state["hot_water"] = False
        else:
            return "Lütfen seçim yapınız:\n1 - Sadece ısıtma\n2 - Isıtma + sıcak su"
        state["step"] = "ASK_BOILER"
        return "Sistemde boyler var mı?\nEvet / Hayır"

    if step == "ASK_BOILER":
        if is_yes(text):
            state["has_boiler"] = True
        elif is_no(text):
            state["has_boiler"] = False
        else:
            return "Boyler durumunu anlayamadım. Lütfen Evet veya Hayır yazınız."

        try:
            quote = build_kazan_quote(
                kazan_adedi=int(state["qty"]),
                kw=int(state["kw"]),
                hot_water=bool(state["hot_water"]),
                has_boiler=bool(state["has_boiler"]),
            )
            pdf_path = generate_kazan_pdf(quote)
            KAZAN_SESSIONS.pop(sender, None)
            return {
                "document_path": pdf_path,
                "filename": os.path.basename(pdf_path),
                "caption": short_kazan_text(quote),
            }
        except Exception as e:
            print("Kazan teklif hatası:", e, flush=True)
            KAZAN_SESSIONS.pop(sender, None)
            return f"Kazan teklifi hazırlanırken hata oluştu: {e}"

    KAZAN_SESSIONS.pop(sender, None)
    return start_kazan_flow(sender, text)


# -------------------------
# Ana cevap üretici
# -------------------------

def create_reply(sender: str, text: str):
    text_lower = normalize_text(text)

    if sender in KAZAN_SESSIONS:
        return continue_kazan_flow(sender, text)

    if is_kazan_message(text):
        return start_kazan_flow(sender, text)

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    radiator_results = []
    kombi_results = []
    total_nakit = 0
    total_kart = 0
    not_found = []

    try:
        radiator_products = load_radiators()
        for line in lines:
            measure, qty = find_radiator_measure(line)
            if not measure:
                continue
            product = next((p for p in radiator_products if p["olcu"] == measure), None)
            if not product:
                not_found.append(measure)
                continue
            nakit_total = product["nakit"] * qty
            kart_total = product["kart"] * qty
            total_nakit += nakit_total
            total_kart += kart_total
            radiator_results.append(
                f"Radyatör {measure}\n"
                f"Adet: {qty}\n"
                f"Nakit Birim: {money_tl(product['nakit'])} KDV Dahil\n"
                f"Kart Birim: {money_tl(product['kart'])} KDV Dahil\n"
                f"Nakit Toplam: {money_tl(nakit_total)} KDV Dahil\n"
                f"Kart Toplam: {money_tl(kart_total)} KDV Dahil"
            )
    except Exception as e:
        print("Radyatör Excel okunamadı:", e, flush=True)

    kombi = None
    kombi_qty = 0
    
    # Eğer mesajdan radyatör bulunduysa kombi arama yapma
    if not radiator_results:
        # Kombi aramasını satır satır yapıyoruz.
# Ama bir satır radyatör olarak algılandıysa o satırda kombi aramıyoruz.
for line in lines:
    measure, _ = find_radiator_measure(line)

    if measure:
        continue

    kombi, kombi_qty = find_kombi(line)

    if kombi:
        kombi_nakit_total = kombi["nakit"] * kombi_qty
        kombi_kart_total = kombi["kart"] * kombi_qty

        total_nakit += kombi_nakit_total
        total_kart += kombi_kart_total

        kombi_results.append(
            f"{kombi['name']}\n"
            f"Adet: {kombi_qty}\n"
            f"Nakit Birim: {money_tl(kombi['nakit'])} KDV Dahil\n"
            f"Kart Birim: {money_tl(kombi['kart'])} KDV Dahil\n"
            f"Nakit Toplam: {money_tl(kombi_nakit_total)} KDV Dahil\n"
            f"Kart Toplam: {money_tl(kombi_kart_total)} KDV Dahil"
        )

    if radiator_results or kombi_results:
        pdf_items = []

        for line in lines:
            measure, qty = find_radiator_measure(line)
            if not measure:
                continue
            product = next((p for p in radiator_products if p["olcu"] == measure), None)
            if not product:
                continue
            pdf_items.append({
                "name": f"E.C.A. Panel Radyatör {measure}",
                "qty": qty,
                "nakit_unit": product["nakit"],
                "kart_unit": product["kart"],
                "nakit_total": product["nakit"] * qty,
                "kart_total": product["kart"] * qty,
            })

        for kombi_item_text in kombi_results:
            if kombi:
                pdf_items.append({
                    "name": kombi["name"],
                    "qty": kombi_qty,
                    "nakit_unit": kombi["nakit"],
                    "kart_unit": kombi["kart"],
                    "nakit_total": kombi["nakit"] * kombi_qty,
                    "kart_total": kombi["kart"] * kombi_qty,
                })

        quote = {
            "title": "Radyatör / Kombi Teklifi" if kombi_results else "Radyatör Teklifi",
            "items": pdf_items,
            "total_nakit": total_nakit,
            "total_kart": total_kart,
            "not_found": not_found,
        }

        pdf_path = generate_radyator_pdf(quote)

        reply = (
            "Fiyat teklifiniz hazırlanmıştır.\n\n"
            f"Nakit Toplam: {money_tl(total_nakit)} KDV Dahil\n"
            f"Kart Toplam: {money_tl(total_kart)} KDV Dahil"
        )

        if not_found:
            reply += "\n\nBulunamayan ölçüler:\n" + "\n".join(not_found)

        return {
            "text": reply,
            "document_path": pdf_path,
            "filename": os.path.basename(pdf_path),
            "caption": reply,
        }


    if "merhaba" in text_lower or "selam" in text_lower or text_lower == "sa":
        return (
            "Merhaba, Pramid İnşaat fiyat sistemine hoş geldiniz.\n\n"
            "Radyatör için adet ve ölçü yazabilirsiniz:\n"
            "3 tane 100 60\n"
            "600x1000 2 adet\n\n"
            "Kazan için örnek:\n"
            "2 adet 125 kW kazan"
        )

    return (
        "Fiyat verebilmem için ürün bilgisini yazınız.\n\n"
        "Radyatör örnek:\n"
        "3 tane 100 60\n"
        "600x1000 2 adet\n\n"
        "Kazan örnek:\n"
        "2 adet 125 kW kazan"
    )


# -------------------------
# WhatsApp gönderimleri
# -------------------------

def send_whatsapp_message(to: str, body: str):
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{PHONE_NUMBER_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    parts = [body[i:i+3000] for i in range(0, len(body), 3000)]
    for part in parts:
        payload = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": part}}
        response = requests.post(url, headers=headers, json=payload)
        print("WhatsApp mesaj cevap:", response.status_code, response.text, flush=True)


def upload_whatsapp_media(file_path: str):
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{PHONE_NUMBER_ID}/media"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    with open(file_path, "rb") as f:
        files = {"file": (os.path.basename(file_path), f, "application/pdf")}
        data = {"messaging_product": "whatsapp", "type": "application/pdf"}
        response = requests.post(url, headers=headers, data=data, files=files)
    print("WhatsApp medya upload:", response.status_code, response.text, flush=True)
    response.raise_for_status()
    return response.json()["id"]


def send_whatsapp_document(to: str, file_path: str, filename: str, caption: str = ""):
    media_id = upload_whatsapp_media(file_path)
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{PHONE_NUMBER_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "document",
        "document": {"id": media_id, "filename": filename, "caption": caption[:1000]},
    }
    response = requests.post(url, headers=headers, json=payload)
    print("WhatsApp PDF cevap:", response.status_code, response.text, flush=True)
