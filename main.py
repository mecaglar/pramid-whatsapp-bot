import os
import requests
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse

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
        message = data["entry"][0]["changes"][0]["value"]["messages"][0]
        sender = message["from"]
        text = message.get("text", {}).get("body", "")

        reply = create_reply(text)
        send_whatsapp_message(sender, reply)

    except Exception as e:
        print("Hata:", e)

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


def create_reply(text: str) -> str:
    products = load_products()

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    results = []
    total_nakit = 0
    total_kart = 0
    not_found = []

    for line in lines:
        measure, qty = find_measure_from_line(line)

        if not measure:
            continue

        product = next((p for p in products if p["olcu"] == measure), None)

        if not product:
            not_found.append(measure)
            continue

        nakit_total = product["nakit"] * qty
        kart_total = product["kart"] * qty

        total_nakit += nakit_total
        total_kart += kart_total

        results.append(
            f"{measure}\n"
            f"Adet: {qty}\n"
            f"Nakit Birim: {money(product['nakit'])} KDV Dahil\n"
            f"Kart Birim: {money(product['kart'])} KDV Dahil\n"
            f"Nakit Toplam: {money(nakit_total)} KDV Dahil\n"
            f"Kart Toplam: {money(kart_total)} KDV Dahil"
        )

    if results:
        reply = "Radyatör Fiyat Teklifi\n\n"
        reply += "\n\n".join(results)

        reply += (
            "\n\nGenel Toplam\n"
            f"Nakit: {money(total_nakit)} KDV Dahil\n"
            f"Kart: {money(total_kart)} KDV Dahil"
        )

        if not_found:
            reply += "\n\nBulunamayan ölçüler:\n"
            reply += "\n".join(not_found)

        return reply

    if "merhaba" in text.lower() or "selam" in text.lower() or "sa" == text.lower().strip():
        return (
            "Merhaba, Pramid İnşaat radyatör fiyat botuna hoş geldiniz.\n\n"
            "Ölçü ve adet bilgisi yazabilirsiniz.\n\n"
            "Örnek:\n"
            "3 tane 100 60\n"
            "2 adet 2000 1200\n"
            "5 adet 90 a 100"
        )

    return (
        "Radyatör fiyatı verebilmem için adet ve ölçü yazınız.\n\n"
        "Örnek:\n"
        "3 tane 100 60\n"
        "2 adet 2000 1200\n"
        "5 adet 90 a 100"
    )


def send_whatsapp_message(to: str, body: str):
    url = f"https://graph.facebook.com/v20.0/{PHONE_NUMBER_ID}/messages"

    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json",
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "text",
        "text": {"body": body},
    }

    response = requests.post(url, headers=headers, json=payload)
    print("WhatsApp cevap:", response.status_code, response.text, flush=True)
