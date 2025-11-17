import telebot
import requests
import json
import os   # TOKENni environmentdan olish uchun

# TOKEN endi kodda bo'lmaydi!
TOKEN = os.getenv("TOKEN")

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


def format_answer(imei: str, data: dict) -> str:
    d = data.get("data", {}) or {}

    model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
    fullname = d.get("standardisedFullName") or model

    reg = d.get("registrationStatus", {}) or {}
    status_code = (reg.get("status") or "").upper()

    # Русский вариант + emoji
    if status_code == "REGISTERED":
        status_emoji = "✅"
        status_text = "Зарегистрирован"
    elif status_code == "UNREGISTERED":
        status_emoji = "❌"
        status_text = "Не зарегистрирован"
    else:
        status_emoji = "ℹ️"
        status_text = "Неизвестно"

    text = f"""
📱 <b>IMEI:</b> <code>{imei}</code>

<b>Модель:</b> {model}
<b>Полное название:</b> {fullname}

<b>Статус:</b> {status_emoji} {status_text}
"""
    return text.strip()


@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(
        message,
        "Здравствуйте! 👋\nОтправьте один или несколько IMEI (каждый с новой строки)."
    )


@bot.message_handler(content_types=['text'])
def check_imei(message):
    lines = [l.strip().replace(" ", "") for l in message.text.split("\n") if l.strip()]
    imeis = [l for l in lines if l.isdigit() and 14 <= len(l) <= 17]

    if not imeis:
        bot.reply_to(
            message,
            "❌ Это не похоже на IMEI.\nОтправьте 14–17 значный IMEI."
        )
        return

    for imei in imeis:
        try:
            r = requests.get(API_URL.format(imei=imei), timeout=10)
        except:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: ошибка подключения.")
            continue

        if r.status_code != 200:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: HTTP ошибка {r.status_code}")
            continue

        try:
            data = r.json()
        except:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: неверный ответ от сервера.")
            continue

        if data.get("status") != "SUCCESS":
            bot.reply_to(
                message,
                f"❌ IMEI <b>{imei}</b>: {data.get('message', 'Ошибка')}"
            )
            continue

        bot.reply_to(message, format_answer(imei, data))


print("Bot zapущен...")
bot.infinity_polling(skip_pending=True)
