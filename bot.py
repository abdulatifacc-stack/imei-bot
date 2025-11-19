import os
import re
import json
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


# --------- KG статус (SIM / срок регистрации) ---------

def extract_kg_status_short(raw_data: dict) -> dict:
    """
    Из полного JSON imei.kg вытащить короткий KG-статус.
    Ищем по тексту знакомые русские фразы, которые ты показывал на скринах.

    Возвращает:
    {
        "kg_status_short": "🇰🇬 KG: ...",
        "kg_color": "RED"/"BLUE"/"GREEN"/"NONE"
    }
    """
    try:
        text = json.dumps(raw_data, ensure_ascii=False)
        lower = text.lower()

        # 1) SIM хали умуман тушмаган (қизил)
        # "Устройство не использовалось ни в одной из сетей мобильных операторов связи Кыргызской Республики."
        if "не использовалось ни в одной из сетей мобильных операторов связи кыргызской республики" in lower:
            return {
                "kg_status_short": "🇰🇬 KG: SIM в сетях КР ещё не работала. 🔴",
                "kg_color": "RED",
            }

        # 2) "Истёк 30-дневный триал период. Требуется регистрация устройства." (кўк)
        if "истёк 30-дневный триал период" in lower:
            return {
                "kg_status_short": "🇰🇬 KG: истёк 30-дневный период, нужна регистрация. 🔵",
                "kg_color": "BLUE",
            }

        # 3) "Срок регистрации устройства до 20.11.2025" (кўк, санали вариант)
        m = re.search(r"Срок регистрации устройства до\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            date = m.group(1)
            return {
                "kg_status_short": f"🇰🇬 KG: до {date}, потом нужна регистрация. 🔵",
                "kg_color": "BLUE",
            }

        # 4) Агар жуда аниқ "зарегистрирован в сетях ..." бўлса (ихтиёрий, яшил)
        if "зарегистрирован в сетях" in lower:
            return {
                "kg_status_short": "🇰🇬 KG: зарегистрирован, можно пользоваться. 🟢",
                "kg_color": "GREEN",
            }

        # Агар ҳеч нарса тушунарли чиқмаса
        return {
            "kg_status_short": "🇰🇬 KG: статус не удалось определить, проверьте на imei.kg. ⚪️",
            "kg_color": "NONE",
        }
    except Exception:
        return {
            "kg_status_short": "🇰🇬 KG: ошибка при разборе статуса, проверьте на imei.kg. ⚪️",
            "kg_color": "NONE",
        }


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
      "status_text": "Зарегистрирован"/...,
      "kg_status_short": "🇰🇬 KG: ...",
      "kg_color": "RED"/"BLUE"/"GREEN"/"NONE"
    }
    """
    # KG-статус тасаввуридан қатьи назар, бутун JSONдан излаймиз
    kg_info = extract_kg_status_short(data)
    kg_status_short = kg_info["kg_status_short"]
    kg_color = kg_info["kg_color"]

    if data.get("status") != "SUCCESS":
        return {
            "model": "Неизвестно",
            "fullname": "Неизвестно",
            "status_code": "UNKNOWN",
            "status_text": "Ошибка или неверный IMEI",
            "kg_status_short": kg_status_short,
            "kg_color": kg_color,
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
        "kg_status_short": kg_status_short,
        "kg_color": kg_color,
    }


def format_text_answer(imei: str, info: dict) -> str:
    """
    Bir dona IMEI uchun javob.
    Tepada 📱 + IMEI (oddiy matn),
    pastda kichik code-blok – ustiga bossa oson kopiya bo‘ladi.
    Ichida KG qisqa jb ham bor.
    """
    model = info["model"]
    fullname = info["fullname"]
    status_code = info["status_code"]
    status_text = info["status_text"]
    kg_status_short = info.get("kg_status_short", "")

    # Emoji статуса регистрации (imei.kg API)
    if status_code == "REGISTERED":
        status_emoji = "✅"
    elif status_code == "UNREGISTERED":
        status_emoji = "❌"
    else:
        status_emoji = "ℹ️"

    kg_line = f"\n{kg_status_short}" if kg_status_short else ""

    text = f"""
📱 {imei}

<pre><code>Полное название: {fullname}
IMEI: {imei}
Статус: {status_emoji} {status_text}{kg_line}</code></pre>
"""
    return text.strip()


# --------- Создание Excel (XLSX) ---------

def create_excel_xlsx(imei_list, info_map) -> str:
    """
    Создаём XLSX-файл по списку IMEI.

    Колонки:
      - Полное название
      - IMEI
      - Статус регистрации (REGISTERED/UNREGISTERED/UNKNOWN)
      - KG статус (коротко, JB)

    Цвета:
      - зелёный: зарегистрирован (регистрация)
      - красный: не зарегистрирован (регистрация)
      - синий: неизвестно / ошибка (регистрация)
      + для KG-статуса:
        - RED  → ячейка синхронно красная
        - BLUE → ячейка синяя
        - GREEN→ ячейка зелёная
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "IMEI Report"

    # Заголовки
    ws.append(["Полное название", "IMEI", "Статус регистрации", "KG статус"])

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    blue_fill = PatternFill(start_color="C6D9F1", end_color="C6D9F1", fill_type="solid")

    for imei in imei_list:
        info = info_map.get(imei) or {
            "fullname": "Неизвестно",
            "status_code": "UNKNOWN",
            "status_text": "Неизвестно",
            "kg_status_short": "",
            "kg_color": "NONE",
        }

        fullname = info.get("fullname", "Неизвестно")
        status_code = info.get("status_code", "UNKNOWN")
        status_text = info.get("status_text", "Неизвестно")
        kg_status_short = info.get("kg_status_short", "")
        kg_color = info.get("kg_color", "NONE")

        # Следующая свободная строка
        row = ws.max_row + 1
        ws.append([fullname, imei, status_text, kg_status_short])

        # Цвет по статусу регистрации (колонка C)
        if status_code == "REGISTERED":
            fill_reg = green_fill
        elif status_code == "UNREGISTERED":
            fill_reg = red_fill
        else:
            fill_reg = blue_fill
        ws[f"C{row}"].fill = fill_reg

        # Цвет по KG-статусу (колонка D)
        if kg_color == "RED":
            fill_kg = red_fill
        elif kg_color == "BLUE":
            fill_kg = blue_fill
        elif kg_color == "GREEN":
            fill_kg = green_fill
        else:
            fill_kg = None

        if fill_kg:
            ws[f"D{row}"].fill = fill_kg

    file_path = "imei_result.xlsx"
    wb.save(file_path)
    return file_path


# --------- Handlers ---------

@bot.message_handler(commands=["start"])
def start_handler(message):
    text = (
        "Здравствуйте! 👋\n\n"
        "Этот бот проверяет IMEI по базе imei.kg.\n"
        "Показывает модель, статус регистрации и KG-статус (SIM/срок регистрации).\n\n"
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
