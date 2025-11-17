import os
import requests
import telebot
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# TOKENni Render / Environment Variables ichida TOKEN nomi bilan bergansiz
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Environmentda TOKEN topilmadi!")

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"


# ---------- API bilan ishlash ----------

def call_api(imei: str) -> dict:
    """
    Bitta IMEI bo'yicha API dan ma'lumot olib keladi.
    Natija har doim shu formatda:
    - ok = True  bo'lsa: fullname, status_text, status_code bor
    - ok = False bo'lsa: error matni bor
    """
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=15)
        data = r.json()
    except Exception:
        return {
            "ok": False,
            "imei": imei,
            "error": "Сервер ошибка"
        }

    if data.get("status") != "SUCCESS":
        msg = data.get("message") or "IMEI не найден"
        return {
            "ok": False,
            "imei": imei,
            "error": msg
        }

    d = data.get("data") or {}
    reg = d.get("registrationStatus") or {}
    status_code = (reg.get("status") or "").upper()

    if status_code == "REGISTERED":
        status_text = "Зарегистрирован"
        short_code = "REGISTERED"
    elif status_code == "UNREGISTERED":
        status_text = "Не зарегистрирован"
        short_code = "UNREGISTERED"
    else:
        status_text = "Неизвестно"
        short_code = "UNKNOWN"

    fullname = (
        d.get("standardisedFullName")
        or d.get("gsmaMarketingName")
        or "Неизвестно"
    )

    return {
        "ok": True,
        "imei": imei,
        "fullname": fullname,
        "status_text": status_text,
        "status_code": short_code
    }


# ---------- Telegram javobi (bitta IMEI) ----------

def format_text_answer(res: dict) -> str:
    """
    Bitta IMEI uchun chiroyli matnli javob.
    """
    if not res["ok"]:
        return (
            f"❌ IMEI: <code>{res['imei']}</code>\n"
            f"Статус: {res['error']}"
        )

    if res["status_code"] == "REGISTERED":
        emoji = "✅"
    elif res["status_code"] == "UNREGISTERED":
        emoji = "❌"
    else:
        emoji = "ℹ️"

    return (
        f"📱 IMEI: <code>{res['imei']}</code>\n\n"
        f"Полное название: {res['fullname']}\n"
        f"Статус: {emoji} {res['status_text']}"
    )


# ---------- Excel (.xlsx) yaratish ----------

def create_excel(results: list) -> str:
    """
    results — call_api dan qaytgan obyektlar ro'yxati.
    Excel fayl yaratib, fayl nomini qaytaradi.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "IMEI"

    # Sarlavhalar
    ws.append(["Полное название", "IMEI", "Статус"])

    # Ranglar (faqat status ustuniga)
    green_fill = PatternFill("solid", fgColor="C6EFCE")  # registered
    red_fill = PatternFill("solid", fgColor="FFC7CE")    # unregistered / error
    grey_fill = PatternFill("solid", fgColor="D9D9D9")   # unknown

    for res in results:
        if not res["ok"]:
            row = ["—", res["imei"], f"Ошибка: {res['error']}"]
            fill = red_fill
        else:
            row = [res["fullname"], res["imei"], res["status_text"]]
            if res["status_code"] == "REGISTERED":
                fill = green_fill
            elif res["status_code"] == "UNREGISTERED":
                fill = red_fill
            else:
                fill = grey_fill

        ws.append(row)
        # Oxirgi qo'shilgan qatordagi 3-ustun (Статус)ni bo'yash
        status_cell = ws.cell(row=ws.max_row, column=3)
        status_cell.fill = fill

    filename = "imei_results.xlsx"
    wb.save(filename)
    return filename


# ---------- /start va IMEI handlerlar ----------

@bot.message_handler(commands=['start'])
def start(message):
    text = (
        "Assalomu alaykum! 👋\n\n"
        "IMEI tekshirish uchun IMEI raqamini yuboring.\n"
        "Bir nechta IMEI bo'yicha Excel olish uchun:\n"
        "<code>/excel 354058249103607 353548756116480 ...</code>\n"
        "yoki\n"
        "<code>/excel</code> deb yozib, keyingi qatorda IMEI larni tashlang."
    )
    bot.reply_to(message, text)


def extract_imeis(text: str) -> list:
    """
    Matndan faqat raqamli IMEIlarni ajratib oladi.
    /excel 123 456
    yoki
    /excel\n123\n456
    ham ishlaydi.
    """
    parts = text.replace("\n", " ").split()
    imeis = [p for p in parts if p.isdigit()]
    return imeis


@bot.message_handler(commands=['excel'])
def handle_excel(message):
    imeis = extract_imeis(message.text)

    if not imeis:
        bot.reply_to(
            message,
            "IMEI larni ham yozing, masalan:\n"
            "<code>/excel 354058249103607 353548756116480</code>\n"
            "yoki:\n"
            "<code>/excel</code>\n"
            "354058249103607\n353548756116480"
        )
        return

    results = []
    for imei in imeis:
        res = call_api(imei)
        results.append(res)

    file_name = create_excel(results)

    with open(file_name, "rb") as f:
        bot.send_document(message.chat.id, f)


@bot.message_handler(content_types=['text'])
def handle_imei(message):
    imei = message.text.strip()

    if not imei.isdigit():
        bot.reply_to(message, "Iltimos, faqat IMEI raqamini yuboring (faqat raqamlar).")
        return

    res = call_api(imei)
    text = format_text_answer(res)
    bot.reply_to(message, text)


print("Bot ishga tushdi...")
bot.infinity_polling(skip_pending=True)
