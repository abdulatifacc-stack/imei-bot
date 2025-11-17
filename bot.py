import os
import re
import json
import requests
import telebot

from openpyxl import Workbook
from openpyxl.styles import PatternFill

# TOKEN берём из переменной окружения (Render → Environment → TOKEN)
TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


def fetch_imei_data(imei: str) -> dict:
    """Запрос к API по IMEI. Всегда возвращает dict."""
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=10)
        return r.json()
    except Exception:
        return {"status": "ERROR", "message": "Request failed"}


def parse_imei(data: dict) -> dict:
    """
    Приводим ответ API к удобному виду.

    Возвращаем:
    {
      "model": ...,
      "fullname": ...,
      "status_code": "REGISTERED"/"UNREGISTERED"/"UNKNOWN",
      "status_text": "Зарегистрирован"/...
    }
    """
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
    """Формируем текстовый ответ для одного IMEI."""
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
    Создаём XLSX-файл по списку IMEI.

    Колонки:
      - Полное название
      - IMEI
      - Статус

    Цвета:
      - зелёный: зарегистрирован
      - красный: не зарегистрирован
      - синий: неизвестно / ошибка
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "IMEI Report"

    # Заголовки
    ws.append(["Полное название", "IMEI", "Статус"])

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

        if status_code == "REGISTERED":
            fill = green_fill
        elif status_code == "UNREGISTERED":
            fill = red_fill
        else:
            fill = blue_fill

        ws[f"C{row}"].fill = fill

    file_path = "imei_result.xlsx"
    wb.save(file_path)
    return file_path


@bot.message_handler(commands=["start"])
def start_handler(message):
    text = (
        "Здравствуйте!\n\n"
        "Этот бот проверяет IMEI по базе imei.kg.\n"
        "Просто отправьте один или несколько IMEI (каждый на новой строке или через пробел).\n\n"
        "📄 Чтобы получить Excel (XLSX), отправьте команду:\n"
        "<code>/excel</code> или <code>/exel</code>, а ниже список IMEI построчно."
    )
    bot.reply_to(message, text)


@bot.message_handler(commands=["excel", "exel"])
def excel_handler(message):
    """
    Команда /excel или /exel:

    /excel
    3540...
    3535...

    — для этих IMEI бот отправит отдельные ответы + XLSX-файл.
    """
    text = message.text

    # Убираем саму команду и собираем все числа 10–20 знаков
    # (вдруг человек напишет в несколько строк)
    all_numbers = re.findall(r"\d{10,20}", text)
    imeis = [n.strip() for n in all_numbers]

    if not imeis:
        bot.reply_to(
            message,
            "❗ Укажите IMEI после команды.\nПример:\n"
            "<code>/excel\n3540...\n3535...\n3536...</code>",
        )
        return

    info_map = {}

    # Для каждого IMEI — отдельный текстовый ответ
    for imei in imeis:
        data = fetch_imei_data(imei)
        info = parse_imei(data)
        info_map[imei] = info

        answer = format_text_answer(imei, info)
        bot.send_message(message.chat.id, answer)

    # Создаём XLSX-файл
    file_path = create_excel_xlsx(imeis, info_map)
    with open(file_path, "rb") as f:
        bot.send_document(
            message.chat.id,
            f,
            caption="📄 Excel (XLSX) по указанным IMEI",
        )


@bot.message_handler(content_types=["text"])
def text_handler(message):
    """
    Любой текст БЕЗ команды /excel:
    - Если нашли 1 IMEI → один ответ.
    - Если нашли несколько IMEI → по каждому отдельный ответ.
    - Excel НЕ отправляем.
    """
    text = message.text

    # Ищем все последовательности цифр длиной 10–20
    imeis = re.findall(r"\d{10,20}", text)

    if not imeis:
        bot.reply_to(
            message,
            "❗ Пожалуйста, отправьте IMEI (только цифры).\n"
            "Можно один или несколько IMEI в одном сообщении."
        )
        return

    for imei in imeis:
        data = fetch_imei_data(imei)
        info = parse_imei(data)
        answer = format_text_answer(imei, info)
        bot.send_message(message.chat.id, answer)


print("Бот запущен...")
bot.infinity_polling(skip_pending=True)
