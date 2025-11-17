[bot.py](https://github.com/user-attachments/files/23573203/bot.py)
import telebot
import requests
import json

# 🔑 Ваш токен сюда
TOKEN = "8492570853:AAHk4o4uDr8lQuGyfVBCnZyd-szvghOQS1A"
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"

def format_answer(imei: str, data: dict) -> str:
    d = data.get("data", {}) or {}

    model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
    fullname = d.get("standardisedFullName") or model

    reg = d.get("registrationStatus", {}) or {}
    status_code = (reg.get("status") or "").upper()

    # Русский вариант статусa + emoji
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
        "Здравствуйте! 👋\n"
        "Отправьте один или несколько IMEI (каждый с новой строки)."
    )


@bot.message_handler(content_types=['text'])
def check_imei(message):
    lines = [l.strip().replace(" ", "") for l in message.text.split("\n") if l.strip()]
    imeis = [l for l in lines if l.isdigit() and 14 <= len(l) <= 17]

    if not imeis:
        bot.reply_to(
            message,
            "❌ Это не похоже на IMEI.\n"
            "Отправьте 14–17 значный IMEI."
        )
        return

    # Har bir IMEI uchun alohida javob
    for imei in imeis:
        try:
            r = requests.get(API_URL.format(imei=imei), timeout=10)
        except:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: ошибка подключения к серверу.")
            continue

        if r.status_code != 200:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: HTTP ошибка {r.status_code}")
            continue

        try:
            data = r.json()
        except:
            bot.reply_to(message, f"❌ IMEI <b>{imei}</b>: неверный ответ сервера.")
            continue

        if data.get("status") != "SUCCESS":
            bot.reply_to(
                message,
                f"❌ IMEI <b>{imei}</b>: {data.get('message', 'Неизвестная ошибка')}"
            )
            continue

        bot.reply_to(message, format_answer(imei, data))


print("Bot запущен...")
bot.infinity_polling(skip_pending=True)
