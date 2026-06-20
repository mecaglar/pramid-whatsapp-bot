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
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

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
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


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
    row = df[df["TİP"].astype(str).str.upper().str.strip() == tip]
    if row.empty:
        return 0.0
    return to_float(row.iloc[0]["İSKONTO (%)"])


def get_equipment_multiplier(data, kazan_adedi):
    for _, row in data["carpan"].iterrows():
        mn = int(to_float(row["MİN KAZAN"]))
        mx = int(to_float(row["MAX KAZAN"]))
        if mn <= kazan_adedi <= mx:
            return int(to_float(row["ÇARPAN"], 1))
    return 1


def get_rule_products(data, rule_name):
    df = data["kurallar"]
    rows = df[df["KURAL"].astype(str).str.upper().str.strip() == rule_name]
    return [str(v).strip() for v in rows["ÜRÜN"].tolist()]


def get_accessory(data, key):
    df = data["aksesuarlar"]
    row = df[df["ANAHTAR"].astype(str).str.upper().str.strip() == str(key).upper().strip()]
    if row.empty:
        raise ValueError(f"Aksesuar bulunamadı: {key}")
    row = row.iloc[0]
    return {
        "code": str(row.get("ÜRÜN KODU", "")).strip() if not pd.isna(row.get("ÜRÜN KODU", "")) else str(key),
        "short": str(key),
        "name": str(row["ÜRÜN ADI"]).strip(),
        "list_price": to_float(row["LİSTE FİYATI (€)"] if "LİSTE FİYATI (€)" in row else row.get("LISTE FIYATI (€)", 0)),
    }


def get_pump(data, kw):
    df = data["pompalar"]
    if kw in [50, 65]:
        wanted = "50/65"
    elif kw in [100, 125]:
        wanted = "100/125"
    else:
        wanted = "150"
    row = df[df["GRUP"].astype(str).str.strip() == wanted]
    if row.empty:
        raise ValueError(f"Pompa seti bulunamadı: {wanted}")
    row = row.iloc[0]
    return {
        "code": str(row.get("ÜRÜN KODU", "")).strip() if not pd.isna(row.get("ÜRÜN KODU", "")) else "",
        "short": f"POMPA {wanted}",
        "name": str(row["ÜRÜN ADI"]).strip(),
        "list_price": to_float(row["LİSTE FİYATI (€)"] if "LİSTE FİYATI (€)" in row else row.get("LISTE FIYATI (€)", 0)),
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
    row = kazanlar[kazanlar["GÜÇ (KW)"].astype(int) == int(kw)]
    if row.empty:
        raise ValueError(f"{kw} kW kazan Excel'de bulunamadı.")
    row = row.iloc[0]

    nakit_disc = get_discount(data, "NAKIT")
    kart_disc = get_discount(data, "KART")
    multiplier = get_equipment_multiplier(data, kazan_adedi)

    items = []
    warnings = []

    add_line(items, {
        "code": str(row["ÜRÜN KODU"]).strip(),
        "short": str(row.get("ÜRÜN KISALTMASI", "")).strip(),
        "name": str(row["ÜRÜN ADI"]).strip(),
        "list_price": to_float(row["LİSTEFİYATI (€)"]),
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

    total_nakit = sum(i["nakit_total_kdv"] for i in items)
    total_kart = sum(i["kart_total_kdv"] for i in items)

    return {
        "kw": kw,
        "kazan_adedi": kazan_adedi,
        "hot_water": hot_water,
        "has_boiler": has_boiler,
        "multiplier": multiplier,
        "items": items,
        "nakit_disc": nakit_disc,
        "kart_disc": kart_disc,
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


def generate_kazan_pdf(quote):
    regular_font, bold_font = setup_pdf_fonts()
    filename = f"pramid-kazan-teklifi-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
    path = os.path.join(tempfile.gettempdir(), filename)

    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    title = ParagraphStyle("Title", parent=styles["Title"], fontName=bold_font, fontSize=16, alignment=TA_CENTER, spaceAfter=8)
    normal = ParagraphStyle("NormalTR", parent=styles["Normal"], fontName=regular_font, fontSize=8.5, leading=11)
    bold = ParagraphStyle("BoldTR", parent=normal, fontName=bold_font)
    right = ParagraphStyle("RightTR", parent=normal, alignment=TA_RIGHT)

    story = []
    story.append(Paragraph("PRAMID İNŞAAT FİYAT TEKLİFİ", title))
    story.append(Paragraph(f"Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}", right))
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(f"E.C.A. Felis Kazan Teklifi - {quote['kazan_adedi']} adet {quote['kw']} kW", bold))
    story.append(Paragraph(f"Kullanım: {'Isıtma + Sıcak su' if quote['hot_water'] else 'Isıtma'} | Boyler: {'Var' if quote['has_boiler'] else 'Yok'}", normal))
    story.append(Spacer(1, 4 * mm))

    data = [[
        Paragraph("Kod", bold), Paragraph("Ürün", bold), Paragraph("Adet", bold),
        Paragraph("Liste Birim", bold), Paragraph("Nakit Toplam", bold), Paragraph("Kart Toplam", bold),
    ]]
    for item in quote["items"]:
        code = item["code"] if item["code"] and item["code"] != "nan" else item["short"]
        data.append([
            Paragraph(str(code), normal),
            Paragraph(item["name"], normal),
            Paragraph(str(item["qty"]), normal),
            Paragraph(money_eur(item["list_unit"]), normal),
            Paragraph(money_eur(item["nakit_total_kdv"]), normal),
            Paragraph(money_eur(item["kart_total_kdv"]), normal),
        ])

    table = Table(data, colWidths=[24*mm, 70*mm, 14*mm, 25*mm, 28*mm, 28*mm], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), bold_font),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
    ]))
    story.append(table)
    story.append(Spacer(1, 6 * mm))

    totals = Table([
        [Paragraph("Nakit Toplam KDV Dahil", bold), Paragraph(money_eur(quote["total_nakit"]), bold)],
        [Paragraph("Kart Toplam KDV Dahil", bold), Paragraph(money_eur(quote["total_kart"]), bold)],
    ], colWidths=[55*mm, 40*mm], hAlign="RIGHT")
    totals.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F4F6")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D1D5DB")),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(totals)

    if quote["warnings"]:
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("Not: " + " ".join(quote["warnings"]), normal))

    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("Fiyatlar teklif amaçlıdır. Nihai stok ve fiyat teyidi için PRAMID İnşaat ile görüşünüz.", normal))

    def add_page_number(canvas, document):
        canvas.saveState()
        canvas.setFont(regular_font, 7)
        canvas.drawRightString(200 * mm, 8 * mm, f"Sayfa {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
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
                "text": short_kazan_text(quote),
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

    kombi, kombi_qty = find_kombi(text)
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
        reply = "Fiyat Teklifi\n\n"
        if radiator_results:
            reply += "RADYATÖR\n\n" + "\n\n".join(radiator_results)
        if kombi_results:
            if radiator_results:
                reply += "\n\n"
            reply += "KOMBİ\n\n" + "\n\n".join(kombi_results)
        reply += "\n\nGenel Toplam\n" + f"Nakit: {money_tl(total_nakit)} KDV Dahil\n" + f"Kart: {money_tl(total_kart)} KDV Dahil"
        if not_found:
            reply += "\n\nBulunamayan ölçüler:\n" + "\n".join(not_found)
        return reply

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
