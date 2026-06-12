import os
import requests
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from difflib import SequenceMatcher
processed_message_ids = set()

app = FastAPI()

VERIFY_TOKEN = os.getenv("VERIFY_TOKEN", "pramid_verify_2026")
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_NUMBER_ID = os.getenv("PHONE_NUMBER_ID")


@app.get("/")
def home():
    return {"status": "PRAMID WhatsApp Bot çalışıyor"}


@app.get("/webhook")
def verify_webhook(request: Request):
    params = request.query_params

    if (
        params.get("hub.mode") == "subscribe"
        and params.get("hub.verify_token") == VERIFY_TOKEN
    ):
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

        reply = create_reply(text)
        send_whatsapp_message(sender, reply)

    except Exception as e:
        print("Hata:", e, flush=True)

    return {"status": "ok"}


PRODUCTS = {
    "eca proteus premix 24": {
        "name": "E.C.A. Proteus Premix 24 kW",
        "price": "36.900 TL + KDV",
    },
    "eca proteus premix 28": {
        "name": "E.C.A. Proteus Premix 28 kW",
        "price": "39.900 TL + KDV",
    },
    "airfel kombi": {
        "name": "Airfel Kombi",
        "price": "31.500 TL + KDV",
    },
}


import pandas as pd
import re

EXCEL_FILE = "fiyat_listesi.xlsx"


def load_products():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = [str(c).strip().upper() for c in df.columns]

    products = []

    for _, row in df.iterrows():
        olcu = str(row["ÖLÇÜ"]).strip().upper().replace(" ", "")
        liste = float(row["LİSTE FİYATI"])
        nakit_iskonto = float(row["NAKİT İSKONTO"])
        kart_iskonto = float(row["KART İSKONTO"])

        nakit = liste * (1 - nakit_iskonto / 100)
        kart = liste * (1 - kart_iskonto / 100)

        products.append({
            "olcu": olcu,
            "nakit": round(nakit),
            "kart": round(kart),
        })

    return products


def money(value):
    return f"{int(value):,}".replace(",", ".") + " TL"


def normalize_dimension(n):
    n = int(n)

    if n in [50, 60, 90]:
        return n * 10

    if n in [100, 120, 140, 160, 180, 200]:
        return n * 10

    return n


def find_measure_from_line(line):
    clean = line.lower()
    clean = clean.replace("x", " ")
    clean = clean.replace("*", " ")
    clean = clean.replace("/", " ")
    clean = clean.replace("a", " ")

    nums = re.findall(r"\d+", clean)
    nums = [int(n) for n in nums]

    if len(nums) < 2:
        return None, 1

    qty = 1

    if len(nums) >= 3:
        qty = nums[0]
        dims = nums[1:3]
    else:
        dims = nums[0:2]

    d1 = normalize_dimension(dims[0])
    d2 = normalize_dimension(dims[1])

    small = min(d1, d2)
    big = max(d1, d2)

    measure = f"{small}X{big}"

    return measure, qty


RADYATOR_FILE = "fiyat_listesi.xlsx"
KOMBI_FILE = "kombi_listesi.xlsx"


def money(value):
    return f"{int(round(value)):,}".replace(",", ".") + " TL"


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


def extract_quantity(text):
    clean = normalize_text(text)

    patterns = [
        r"(\d+)\s*(adet|ad|tane|tn)",
        r"(adet|ad|tane|tn)\s*(\d+)"
    ]

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


def normalize_dimension(n):
    n = int(n)

    if n in [50, 60, 90]:
        return n * 10

    if n in [100, 120, 140, 160, 180, 200]:
        return n * 10

    return n


def find_radiator_measure(line):
    qty, clean = extract_quantity(line)

    clean = clean.replace("x", " ")
    clean = clean.replace("*", " ")
    clean = clean.replace("/", " ")
    clean = clean.replace(" a ", " ")

    nums = [int(n) for n in re.findall(r"\d+", clean)]

    # adet başta/sonda yazıldıysa ölçü için adet sayısını ele
    if len(nums) >= 3:
        if nums[0] == qty:
            nums = nums[1:]
        elif nums[-1] == qty:
            nums = nums[:-1]

    if len(nums) < 2:
        return None, qty

    d1 = normalize_dimension(nums[0])
    d2 = normalize_dimension(nums[1])

    small = min(d1, d2)
    big = max(d1, d2)

    return f"{small}X{big}", qty


def load_radiators():
    df = pd.read_excel(RADYATOR_FILE)
    df.columns = [str(c).strip().upper() for c in df.columns]

    products = []

    for _, row in df.iterrows():
        olcu = str(row["ÖLÇÜ"]).strip().upper().replace(" ", "")
        liste = float(row["LİSTE FİYATI"])
        nakit_iskonto = float(row["NAKİT İSKONTO"])
        kart_iskonto = float(row["KART İSKONTO"])

        nakit = liste * (1 - nakit_iskonto / 100)
        kart = liste * (1 - kart_iskonto / 100)

        products.append({
            "olcu": olcu,
            "nakit": round(nakit),
            "kart": round(kart),
        })

    return products


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
            "nakit": float(row["NAKİT FİYAT"]),
            "kart": float(row["KART FİYAT"]),
        })

    return products


def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def find_kombi(text):
    qty, clean = extract_quantity(text)
    kombis = load_kombis()

    best_product = None
    best_score = 0

    for product in kombis:
        score = similarity(clean, product["search"])

        # Ürün adı içinden parçalı yakalama
        words = product["search"].split()
        matched_words = sum(1 for w in words if w in clean)
        score += matched_words * 0.12

        if score > best_score:
            best_score = score
            best_product = product

    if best_score >= 0.45:
        return best_product, qty

    return None, qty


def create_reply(text: str) -> str:
    text_lower = normalize_text(text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    radiator_products = load_radiators()

    radiator_results = []
    kombi_results = []

    total_nakit = 0
    total_kart = 0
    not_found = []

    # Radyatörleri satır satır ara
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
            f"Nakit Birim: {money(product['nakit'])} KDV Dahil\n"
            f"Kart Birim: {money(product['kart'])} KDV Dahil\n"
            f"Nakit Toplam: {money(nakit_total)} KDV Dahil\n"
            f"Kart Toplam: {money(kart_total)} KDV Dahil"
        )

    # Kombiyi de ayrıca ara
    kombi, kombi_qty = find_kombi(text)

    if kombi:
        kombi_nakit_total = kombi["nakit"] * kombi_qty
        kombi_kart_total = kombi["kart"] * kombi_qty

        total_nakit += kombi_nakit_total
        total_kart += kombi_kart_total

        kombi_results.append(
            f"{kombi['name']}\n"
            f"Adet: {kombi_qty}\n"
            f"Nakit Birim: {money(kombi['nakit'])} KDV Dahil\n"
            f"Kart Birim: {money(kombi['kart'])} KDV Dahil\n"
            f"Nakit Toplam: {money(kombi_nakit_total)} KDV Dahil\n"
            f"Kart Toplam: {money(kombi_kart_total)} KDV Dahil"
        )

    if radiator_results or kombi_results:
        reply = "Fiyat Teklifi\n\n"

        if radiator_results:
            reply += "RADYATÖR\n\n"
            reply += "\n\n".join(radiator_results)

        if kombi_results:
            if radiator_results:
                reply += "\n\n"
            reply += "KOMBİ\n\n"
            reply += "\n\n".join(kombi_results)

        reply += (
            "\n\nGenel Toplam\n"
            f"Nakit: {money(total_nakit)} KDV Dahil\n"
            f"Kart: {money(total_kart)} KDV Dahil"
        )

        if not_found:
            reply += "\n\nBulunamayan ölçüler:\n"
            reply += "\n".join(not_found)

        return reply

    if "merhaba" in text_lower or "selam" in text_lower or text_lower == "sa":
        return (
            "Merhaba, Pramid İnşaat fiyat botuna hoş geldiniz.\n\n"
            "Radyatör için adet ve ölçü yazabilirsiniz:\n"
            "3 tane 100 60\n"
            "600x1000 2 adet\n\n"
            "Kombi için ürün adını yazabilirsiniz:\n"
            "Proteus Premix 24\n"
            "Citius 28"
        )

    return (
        "Fiyat verebilmem için ürün bilgisini yazınız.\n\n"
        "Radyatör örnek:\n"
        "3 tane 100 60\n"
        "600x1000 2 adet\n\n"
        "Kombi örnek:\n"
        "Proteus Premix 24\n"
        "Confeo 30"
    )

def send_whatsapp_message(to: str, body: str):
    url = f"https://graph.facebook.com/v20.0/{PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json",
    }

    parts = [body[i:i+3000] for i in range(0, len(body), 3000)]

    for part in parts:
        payload = {
            "messaging_product": "whatsapp",
            "to": to,
            "type": "text",
            "text": {"body": part},
        }

        response = requests.post(url, headers=headers, json=payload)
        print("WhatsApp cevap:", response.status_code, response.text, flush=True)
