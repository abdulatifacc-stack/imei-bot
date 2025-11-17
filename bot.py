import os
import re
import requests
import telebot

from openpyxl import Workbook
from openpyxl.styles import PatternFill

# TOKEN берём из переменной окружения (Render → Environment → TOKEN)
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Переменная окружения TOKEN не задана!")

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


# --------- API и парсинг ответа ---------

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
    """
    Bir dona IMEI uchun javob.
    Tepada 📱 + IMEI (oddiy matn),
    pastda kichik code-blok – ustiga bossa oson kopiya bo‘ladi.
    """
    model = info["model"]
    fullname = info["fullname"]
    status_code = info["status_code"]
    status_text = info["status_text"]

    # Emoji статуса
    if status_code == "REGISTERED":
        status_emoji = "✅"
    elif status_code == "UNREGISTERED":
        status_emoji = "❌"
    else:
        status_emoji = "ℹ️"

    text = f"""
📱 {imei}

<pre><code>Полное название: {fullname}
IMEI: {imei}
Статус: {status_emoji} {status_text}</code></pre>
"""
    return text.strip()


# --------- Создание Excel (XLSX) ---------

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


# --------- Handlers ---------

@bot.message_handler(commands=["start"])
def start_handler(message):
    text = (
        "Здравствуйте! 👋\n\n"
        "Этот бот проверяет IMEI по базе imei.kg.\n"
        "Просто отправьте один или несколько IMEI в одном сообщении "
        "(каждый на новой строке или через пробел).\n\n"
        "📄 Для Excel (XLSX) используйте команду:\n"
        "<code>/excel</code> или <code>/exel</code>, затем список IMEI."
    )
    bot.reply_to(message, text)


@bot.message_handler(commands=["excel", "exel"])
def excel_handler(message):
    """
    /excel или /exel:

    /excel
    3540...
    3535...

    → БОТ ТОЛЬКО отправляет XLSX-файл, без отдельных ответов по каждому IMEI.
    """
    text = message.text

    # Ищем все последовательности цифр длиной 10–20 (IMEI)
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

    # Только собираем данные для файла, НИЧЕГО не отправляем поштучно
    for imei in imeis:
        data = fetch_imei_data(imei)
        info = parse_imei(data)
        info_map[imei] = info

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
    Любой текст БЕЗ /excel:
    - Ищем все IMEI (10–20 цифр).
    - По каждому IMEI отправляем отдельный ответ на русском.
    - НИКАКОГО Excel здесь нет.
    """
    text = message.text

    # Если это другая команда (например /start), выходим
    if text.startswith("/"):
        return

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


