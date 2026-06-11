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


def create_reply(text: str) -> str:
    text = text.lower().strip()

    for keyword, product in PRODUCTS.items():
        if keyword in text:
            return (
                f"{product['name']}\n"
                f"Fiyat: {product['price']}\n\n"
                "Adet ve teslimat ilini yazarsanız toplam tutarı da hesaplayabilirim."
            )

    if "fiyat" in text:
        return (
            "Fiyat verebilmem için ürün adını yazar mısınız?\n\n"
            "Örnek: ECA Proteus Premix 24 fiyat"
        )

    if "merhaba" in text or "selam" in text:
        return (
            "Merhaba, Pramid İnşaat WhatsApp fiyat botuna hoş geldiniz.\n\n"
            "Fiyat almak için ürün adını yazabilirsiniz.\n"
            "Örnek: ECA Proteus Premix 24 fiyat"
        )

    return (
        "Mesajınızı aldım.\n\n"
        "Fiyat almak için ürün adını yazabilirsiniz.\n"
        "Örnek: ECA Proteus Premix 24 fiyat"
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
