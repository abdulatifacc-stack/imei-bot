import telebot
import requests
import csv
import io
import os

TOKEN = os.getenv("TOKEN")
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"

# ---------- IMEI TEKSHIRISH FUNKSIYASI ----------
def check_imei(imei):
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=10)
        data = r.json()

        d = data.get("data", {}) or {}

        model = d.get("gsmaMarketingName") or d.get("standardisedFullName") or "Неизвестно"
        fullname = d.get("standardisedFullName") or model

        reg = d.get("registrationStatus", {}) or {}
        status_code = (reg.get("status") or "").upper()

        if status_code == "REGISTERED":
            status = "Зарегистрирован"
        elif status_code == "UNREGISTERED":
            status = "Не зарегистрирован"
        else:
            status = "Неизвестно"

        return [imei, model, fullname, status]

    except:
        return [imei, "Ошибка", "Ошибка", "Ошибка"]

# ---------- /excel KOMANDA ----------
@bot.message_handler(commands=["excel"])
def excel_handler(message):
    lines = message.text.split()
    imeis = [line.strip() for line in lines[1:] if line.strip().isdigit()]

    if not imeis:
        bot.reply_to(message, "❌ IMEI топилмади!")
        return

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["IMEI", "Модель", "Полное название", "Статус"])

    for imei in imeis:
        writer.writerow(check_imei(imei))

    output.seek(0)
    file = io.BytesIO(output.getvalue().encode("utf-8"))
    bot.send_document(message.chat.id, file, caption="📄 Ваш Excel (CSV) файл готов!")

# ---------- DONALIK IMEI TEKSHIRISH ----------
@bot.message_handler(func=lambda msg: msg.text.isdigit())
def single_imei(message):
    imei = message.text.strip()
    data = check_imei(imei)

    text = f"""
📱 <b>IMEI:</b> <code>{data[0]}</code>

<b>Модель:</b> {data[1]}
<b>Полное название:</b> {data[2]}

<b>Статус:</b> {data[3]}
"""
    bot.reply_to(message, text)

# ---------- BOT ISHGA TUSHADI ----------
bot.infinity_polling(skip_pending=True)
