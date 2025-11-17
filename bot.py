import telebot
import requests
import csv
import io
import os

# TOKEN Render Environment'dan olinadi
TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


# ---------- IMEI haqidagi ma'lumotni olish ----------
def get_imei_info(imei: str):
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=10)
        data = r.json()
    except Exception:
        return {
            "imei": imei,
            "model": "Ошибка",
            "fullname": "Ошибка",
            "status_code": "ERROR",
        }

    d = data.get("data", {}) or {}

    model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
    fullname = d.get("standardisedFullName") or model

    reg = d.get("registrationStatus", {}) or {}
    status_code = (reg.get("status") or "").upper()

    return {
        "imei": imei,
        "model": model,
        "fullname": fullname,
        "status_code": status_code,
    }


def status_text_and_emoji(status_code: str):
    if status_code == "REGISTERED":
        return "✅", "Зарегистрирован"
    elif status_code == "UNREGISTERED":
        return "❌", "Не зарегистрирован"
    elif status_code == "ERROR":
        return "❌", "Ошибка запроса"
    else:
        return "ℹ️", "Неизвестно"


# ---------- /start ----------
@bot.message_handler(commands=["start"])
def start(message):
    text = (
        "Здравствуйте! 👋\n"
        "Отправьте один IMEI — я покажу информацию.\n\n"
        "Чтобы получить Excel по нескольким IMEI, напишите:\n"
        "<code>/excel 354058249103607 353548756116480 3535...</code>"
    )
    bot.reply_to(message, text)


# ---------- /excel – avtomatik fayl ----------
@bot.message_handler(commands=["excel"])
def excel_handler(message):
    parts = message.text.split()
    imeis = [p.strip() for p in parts[1:] if p.strip().isdigit()]

    if not imeis:
        bot.reply_to(message, "❌ Укажите IMEI после команды /excel")
        return

    # CSV (Excel) tayyorlaymiz
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["IMEI", "Модель", "Полное название", "Статус"])

    for imei in imeis:
        info = get_imei_info(imei)
        emoji, status_ru = status_text_and_emoji(info["status_code"])
        writer.writerow([info["imei"], info["model"], info["fullname"], f"{emoji} {status_ru}"])

    output.seek(0)
    file_bytes = io.BytesIO(output.getvalue().encode("utf-8"))
    file_bytes.name = "imei_result.csv"

    bot.send_document(message.chat.id, file_bytes, caption="📄 Excel (CSV) по указанным IMEI")


# ---------- Donalik IMEI (oddiy xabar) ----------
@bot.message_handler(func=lambda m: m.text and m.text.strip().isdigit())
def single_imei(message):
    imei = message.text.strip()
    info = get_imei_info(imei)
    emoji, status_ru = status_text_and_emoji(info["status_code"])

    text = (
        f"📱 <b>IMEI:</b> <code>{info['imei']}</code>\n\n"
        f"<b>Модель:</b> {info['model']}\n"
        f"<b>Полное название:</b> {info['fullname']}\n\n"
        f"<b>Статус:</b> {emoji} {status_ru}"
    )
    bot.reply_to(message, text)


# ---------- Botni ishga tushirish ----------
bot.infinity_polling(skip_pending=True)
