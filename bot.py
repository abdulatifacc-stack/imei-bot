import os
import json
import requests
import telebot

from openpyxl import Workbook
from openpyxl.styles import PatternFill

# TOKEN environmentdan olinadi (Render → Environment → TOKEN)
TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


def fetch_imei_data(imei: str) -> dict:
    """
    IMEI bo'yicha API'dan ma'lumot olib keladi.
    Xato bo'lsa ham dict qaytaradi.
    """
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=10)
        return r.json()
    except Exception:
        return {"status": "ERROR", "message": "Request failed"}


def parse_imei(data: dict) -> dict:
    """
    API javobini qulay formatga keltiramiz.
    Qaytadi: {
      "model": ...,
      "fullname": ...,
      "status_code": "REGISTERED"/"UNREGISTERED"/"UNKNOWN",
      "status_text": "Зарегистрирован"/...,
    }
    """
    # Agar SUCCESS bo'lmasa, xato deb hisoblaymiz
    if data.get("status") != "SUCCESS":
        return {
            "model": "Неизвестно",
            "fullname": "Неизвестно",
            "status_code": "UNKNOWN",
            "status_text": "Ошибка или неверный IMEI",
        }

    d = data.get("data", {}) or {}

    model = (
        d.get("gsmaMarketingName")
        or d.get("standardisedFullName")
        or "Неизвестно"
    )
    fullname = d.get("standardisedFullName") or model

    reg = d.get("registrationStatus", {}) or {}
    status_code = (reg.get("status") or "").upper()

    if status_code == "REGISTERED":
        status_text = "Зарегистрирован"
    elif status_code == "UNREGISTERED":
        status_text = "Не зарегистрирован"
    else:
        status_code = "UNKNOWN"
        status_text = "Неизвестно"

    return {
        "model": model,
        "fullname": fullname,
        "status_code": status_code,
        "status_text": status_text,
    }


def format_text_answer(imei: str, info: dict) -> str:
    """
    Bitta IMEI uchun Telegramga chiqadigan matn.
    """
    model = info["model"]
    fullname = info["fullname"]
    status_code = info["status_code"]
    status_text = info["status_text"]

    if status_code == "REGISTERED":
        emoji = "✅"
    elif status_code == "UNREGISTERED":
        emoji = "❌"
    else:
        emoji = "ℹ️"

    text = f"""
📱 <b>IMEI:</b> <code>{imei}</code>

<b>Модель:</b> {model}
<b>Полное название:</b> {fullname}

<b>Статус:</b> {emoji} {status_text}
"""
    return text.strip()


def create_excel_xlsx(imei_list, info_map) -> str:
    """
    IMEIlar bo'yicha xlsx fayl yaratadi.
    Ustunlar: Полное название / IMEI / Статус
    Ranglar: yashil/qizil/ko'k.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "IMEI Report"

    # Sarlavha
    ws.append(["Полное название", "IMEI", "Статус"])

    # Ranglar
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    blue_fill = PatternFill(start_color="C6D9F1", end_color="C6D9F1", fill_type="solid")

    for imei in imei_list:
        info = info_map.get(imei) or {
            "fullname": "Неизвестно",
            "status_code": "UNKNOWN",
            "status_text": "Неизвестно",
        }

        fullname = info.get("fullname", "Неизвестно")
        status_code = info.get("status_code", "UNKNOWN")
        status_text = info.get("status_text", "Неизвестно")

        row = ws.max_row + 1
        ws.append([fullname, imei, status_text])

        # Rangni statusga qarab beramiz
        if status_code == "REGISTERED":
            fill = green_fill
        elif status_code == "UNREGISTERED":
            fill = red_fill
        else:
            fill = blue_fill

        # Faqat status ustuniga rang
        ws[f"C{row}"].fill = fill

    file_path = "imei_result.xlsx"
    wb.save(file_path)
    return file_path


@bot.message_handler(commands=["start"])
def start_handler(message):
    text = (
        "Ассалому алейкум!\n\n"
        "Бу бот IMEIни текшириб беради.\n"
        "Просто отправьте IMEI (14–15 цифр).\n\n"
        "📄 Для Excel отправьте:\n"
        "<code>/excel</code> и далее список IMEI построчно."
    )
    bot.reply_to(message, text)


@bot.message_handler(commands=["excel"])
def excel_handler(message):
    """
    /excel 3540... 3535...
    yoki:
    /excel
    3540...
    3535...
    """
    # /excel ni olib tashlaymiz va qolgan hamma whitespace bo'yicha bo'lamiz
    text = message.text
    parts = text.split()
    imeis = [p.strip() for p in parts[1:] if p.strip()]

    if not imeis:
        bot.reply_to(
            message,
            "❗ IMEIлар ёзинг.\nМисол:\n"
            "<code>/excel\n3540...\n3535...\n3536...</code>",
        )
        return

    info_map = {}

    # Har bir IMEI uchun alohida javob + info_mapga saqlash
    for imei in imeis:
        data = fetch_imei_data(imei)
        info = parse_imei(data)
        info_map[imei] = info

        answer = format_text_answer(imei, info)
        bot.send_message(message.chat.id, answer)

    # Xlsx fayl yaratamiz
    file_path = create_excel_xlsx(imeis, info_map)
    with open(file_path, "rb") as f:
        bot.send_document(
            message.chat.id,
            f,
            caption="📄 Excel (XLSX) по указанным IMEI",
        )


@bot.message_handler(content_types=["text"])
def single_imei_handler(message):
    """
    Oddiy text kelganda – bitta IMEI deb hisoblaymiz.
    """
    imei = message.text.strip()

    # Juda oddiy filter: faqat raqam va uzunligi 14–17 oralig'ida
    if not imei.isdigit() or not (10 <= len(imei) <= 20):
        bot.reply_to(message, "❗ Пожалуйста, отправьте корректный IMEI (только цифры).")
        return

    data = fetch_imei_data(imei)
    info = parse_imei(data)
    answer = format_text_answer(imei, info)
    bot.reply_to(message, answer)


print("Bot ishga tushdi...")
bot.infinity_polling(skip_pending=True)
