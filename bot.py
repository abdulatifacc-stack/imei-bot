import os
import csv
import json
import requests
import telebot

# TOKENni Render environmentdan olamiz
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("TOKEN env var is not set")

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


def parse_imei(line: str):
    """Qator ichidan faqat raqamlarni olib, IMEI uzunligini tekshiradi."""
    digits = "".join(ch for ch in line if ch.isdigit())
    if 14 <= len(digits) <= 17:
        return digits
    return None


def call_api(imei: str) -> dict:
    """imei.kg API ga so'rov yuboradi va JSON qaytaradi."""
    r = requests.get(API_URL.format(imei=imei), timeout=10)
    r.raise_for_status()
    return r.json()


def format_answer(imei: str, data: dict) -> str:
    """Bitta IMEI uchun chiroyli ruscha javob."""
    d = data.get("data", {}) or {}

    model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
    fullname = d.get("standardisedFullName") or model

    reg = d.get("registrationStatus", {}) or {}
    status_code = (reg.get("status") or "").upper()

    if status_code == "REGISTERED":
        emoji = "✅"
        status_text = "Зарегистрирован"
    elif status_code == "UNREGISTERED":
        emoji = "❌"
        status_text = "Не зарегистрирован"
    else:
        emoji = "ℹ️"
        status_text = "Неизвестно"

    text = (
        f"📱 <b>IMEI:</b> <code>{imei}</code>\n\n"
        f"<b>Модель:</b> {model}\n"
        f"<b>Полное название:</b> {fullname}\n\n"
        f"<b>Статус:</b> {emoji} {status_text}"
    )
    return text


@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(
        message,
        "Здравствуйте! 👋\n"
        "Отправьте IMEI для проверки.\n\n"
        "Для Excel отчёта:\n"
        "/excel и ниже списком IMEI."
    )


@bot.message_handler(commands=['excel'])
def excel_handler(message):
    """
    /excel
    3540...
    3535...
    kabi kelgan xabardan Excel (CSV) fayl tayyorlaydi.
    """
    lines = message.text.splitlines()[1:]  # birinchi qator /excel, qolganlari IMEI

    imeis = []
    for line in lines:
        imei = parse_imei(line)
        if imei:
            imeis.append(imei)

    if not imeis:
        bot.reply_to(
            message,
            "Пожалуйста, отправьте команду так:\n"
            "/excel\n3540...\n3535...\n..."
        )
        return

    results = []

    for imei in imeis:
        try:
            data = call_api(imei)
            if data.get("status") != "SUCCESS":
                results.append((imei, "Ошибка", "", "Ошибка"))
                continue

            d = data.get("data", {}) or {}
            model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
            fullname = d.get("standardisedFullName") or model

            reg = d.get("registrationStatus", {}) or {}
            status_code = (reg.get("status") or "").upper()

            if status_code == "REGISTERED":
                status_text = "Зарегистрирован"
            elif status_code == "UNREGISTERED":
                status_text = "Не зарегистрирован"
            else:
                status_text = "Неизвестно"

            results.append((imei, model, fullname, status_text))
        except Exception:
            results.append((imei, "Ошибка", "", "Ошибка"))

    if not results:
        bot.reply_to(message, "IMEI не найдены. Проверьте список.")
        return

    # CSV fayl (Excel ochadi)
    filepath = "/tmp/imei_results.csv"
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["IMEI", "Модель", "Полное имя", "Статус"])
        for row in results:
            writer.writerow(row)

    with open(filepath, "rb") as f:
        bot.send_document(message.chat.id, f, caption="📄 IMEI отчёт (Excel)")


@bot.message_handler(content_types=['text'])
def imei_handler(message):
    """Oddiy xabar – IMEI(lar)ni tekshirish."""
    lines = [l for l in message.text.splitlines() if l.strip()]
    imeis = []

    for line in lines:
        imei = parse_imei(line)
        if imei:
            imeis.append(imei)

    if not imeis:
        bot.reply_to(
            message,
            "Это не похоже на IMEI.\n"
            "Отправьте 14–17 значный IMEI."
        )
        return

    for imei in imeis:
        try:
            data = call_api(imei)
            if data.get("status") != "SUCCESS":
                bot.reply_to(
                    message,
                    f"❌ IMEI <b>{imei}</b>: {data.get('message', 'Ошибка')}"
                )
                continue

            text = format_answer(imei, data)
            bot.reply_to(message, text)
        except Exception:
            bot.reply_to(
                message,
                f"❌ IMEI <b>{imei}</b>: ошибка сервера."
            )


print("Bot ishga tushdi...")
bot.infinity_polling(skip_pending=True)
