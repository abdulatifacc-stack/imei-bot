# imei_bot_single_column.py
import os
import re
import json
import requests
import telebot

# Google Sheets
import gspread
from google.oauth2.service_account import Credentials

# ---------- ADMIN VA FAYLLAR ----------
ADMIN_ID = 357556285
USERS_FILE = "users.json"                 # foydalanuvchilar ro'yxati
CHECKED_FILE = "checked_imeis.json"       # oldin tekshirilgan IMEI lar
MODEL_CONF_FILE = "model_config.json"     # {"header": "iPhone 16 PM"}

# Google Sheets: ID yoki nom
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")         # tavsiya: ID
SHEET_NAME     = os.getenv("SHEET_NAME", "IMEI_Table")    # ID bo'lmasa, nom

# Service Account JSON (siz yuklagan fayl)
SERVICE_ACCOUNT_FILE = "/mnt/data/pristine-disk-479021-t5-5a15549085f2.json"

# Telegram token
TOKEN = os.getenv("TOKEN")
if not TOKEN:
    raise RuntimeError("Переменная окружения TOKEN не задана!")
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

# IMEI API
API_URL = "https://www.imei.kg/api/phys/imei/status?imei={imei}"

# ---------- YORDAMCHILAR (LOCAL STORAGE) ----------
def _load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default

def _save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _load_checked():
    # { "356....": "Apple iPhone 16 Pro Max ...", ... }
    return _load_json(CHECKED_FILE, {})

def _save_checked(db: dict):
    _save_json(CHECKED_FILE, db)

def record_checked_imei(imei: str, fullname: str):
    db = _load_checked()
    db[str(imei)] = fullname or ""
    _save_checked(db)

def _get_model_header():
    cfg = _load_json(MODEL_CONF_FILE, {})
    return cfg.get("header", "").strip()

def _set_model_header(header: str):
    _save_json(MODEL_CONF_FILE, {"header": header.strip()})

# ---------- USER/ADMIN LOG ----------
def save_user(message):
    u = message.from_user
    uid = str(u.id)
    data = _load_json(USERS_FILE, {})
    data[uid] = {
        "username": u.username or "",
        "first_name": u.first_name or "",
        "last_name": u.last_name or "",
    }
    _save_json(USERS_FILE, data)

def notify_admin_imeis(message, imeis):
    if not ADMIN_ID:
        return
    u = message.from_user
    username = f"@{u.username}" if u.username else "—"
    full_name = " ".join(x for x in [u.first_name, u.last_name] if x) or "—"
    imei_text = "\n".join(f"• {i}" for i in imeis)
    text = f"""
🟢 IMEI botdan foydalanildi

👤 User ID: <code>{u.id}</code>
🔗 Username: {username}
👤 Ism: {full_name}
💬 Chat ID: <code>{message.chat.id}</code>

📱 IMEI(lar):
{imei_text}
"""
    try:
        bot.send_message(ADMIN_ID, text.strip())
    except Exception:
        pass

# ---------- API va PARSING ----------
def fetch_imei_data(imei: str) -> dict:
    try:
        r = requests.get(API_URL.format(imei=imei), timeout=10)
        return r.json()
    except Exception:
        return {"status": "ERROR", "message": "Request failed"}

def parse_imei(data: dict) -> dict:
    if data.get("status") != "SUCCESS":
        return {
            "model": "Неизвестно",
            "fullname": "Неизвестно",
            "status_code": "UNKNOWN",
            "status_text": "Ошибка или неверный IMEI",
            "kg_chat": "🇰🇬 KG: ошибка при запросе. ⚪️",
        }
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
        status_code = "UNKNOWN"
        status_text = "Неизвестно"

    # kg_chat faqat ma'lumot uchun
    kg_chat = "🇰🇬 KG: информация получена."  # qisqa, batafsil kerak bo'lmasa
    return {
        "model": model,
        "fullname": fullname,
        "status_code": status_code,
        "status_text": status_text,
        "kg_chat": kg_chat,
    }

def format_text_answer(imei: str, info: dict) -> str:
    status_code = info["status_code"]
    emoji = "✅" if status_code == "REGISTERED" else ("❌" if status_code == "UNREGISTERED" else "ℹ️")
    return f"""
📱 {imei}

<pre><code>Полное название: {info['fullname']}
IMEI: {imei}
Статус: {emoji} {info['status_text']}
{info['kg_chat']}</code></pre>
""".strip()

# ---------- Google Sheets (bitta ustun rejimi) ----------
def _get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    return gspread.authorize(creds)

def _open_worksheet():
    gc = _get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID) if SPREADSHEET_ID else gc.open(SHEET_NAME)
    return sh.sheet1

def _ensure_header(ws, header_text: str):
    """
    A1 bo'sh bo'lsa — header yozamiz.
    Agar A1 bor bo'lsa va boshqacha bo'lsa, A1 ni yangilaymiz.
    """
    a1 = ws.acell("A1").value
    if (a1 or "").strip() != header_text.strip():
        ws.update("A1", header_text.strip())

def _get_existing_imeis(ws):
    """
    A2 dan pastga hozirgacha yozilgan IMEIlarni o‘qib, to‘plam qaytaradi.
    """
    all_vals = ws.col_values(1)  # A ustun
    # birinchi qator — header, shuning uchun 2-qatorlardan boshlab IMEIlar
    imeis = [v.strip() for v in all_vals[1:] if str(v).strip()]
    return set(imeis)

def _collect_new_imeis_for_current_model(header_text: str):
    """
    checked_imeis.json dagi IMEIlarni O‘SHA BIR MODELDAN kolleksiya qiladi.
    Talab: 'faqat o'tgan IMEIlar bazaga' — demak, faqat tekshirilganlardan olamiz.
    Modelni moslashtirish: soddalashtirib — header ichidagi '15'/'16' ga qarab filtrlash.
    xohlasangiz yanada nozik filtrlash qo‘shamiz.
    """
    db = _load_checked()  # { imei: fullname }
    want_16 = "16" in header_text
    want_15 = "15" in header_text

    out = []
    for imei, fullname in db.items():
        name = (fullname or "").casefold()
        if want_16 and ("iphone" in name and "16" in name and ("pro max" in name or "promax" in name)):
            out.append(imei)
        elif want_15 and ("iphone" in name and "15" in name and ("pro max" in name or "promax" in name)):
            out.append(imei)
        # Agar headerga raqam yozilmagan bo‘lsa (erkin nom), HAMMASINI yozish ham mumkin:
        elif not want_16 and not want_15:
            out.append(imei)

    # tartib saqlab, dublikatni olib tashlaymiz
    seen, uniq = set(), []
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq

def sync_single_column():
    """
    A ustunda bitta model uchun (A1 = header) IMEIlar APPEND qilib boriladi.
    Hech qachon tozalash yo'q. FAQAT yangilar qo'shiladi (sheet dagi mavjudlar va ro'yxatdagi dublikatlar chiqarib tashlanadi).
    """
    header = _get_model_header()
    if not header:
        raise RuntimeError("Model sarlavhasi o‘rnatilmagan. Avval /setmodel ni ishlating, masalan: /setmodel iPhone 16 PM")

    ws = _open_worksheet()
    _ensure_header(ws, header)

    existing = _get_existing_imeis(ws)            # sheetdagi mavjud IMEIlar
    candidates = _collect_new_imeis_for_current_model(header)  # tekshirilganlardan moslari

    # faqat yangilar
    to_add = [i for i in candidates if i not in existing]
    if not to_add:
        return 0

    # qaysi qatorga yozishni topamiz: oxirgi qator + 1
    last_row = len(existing) + 1  # A1 header, existing = A2..A(1+len)
    # batch yozish
    rows = [[i] for i in to_add]
    start = last_row + 1
    end = last_row + len(to_add)
    ws.update(f"A{start}:A{end}", rows)
    return len(to_add)

# ---------- HANDLERLAR ----------
@bot.message_handler(commands=["start"])
def start_handler(message):
    save_user(message)
    header = _get_model_header() or "— o‘rnatilmagan —"
    bot.reply_to(
        message,
        "Здравствуйте! 👋\n\n"
        "Отправьте IMEI (один или несколько) — bot tekshiradi va javob beradi.\n"
        "Admin:\n"
        "• /setmodel <nom> — A1 sarlavha (masalan: iPhone 16 PM)\n"
        "• /sync — bitta ustunga (A) faqat yangi IMEIlarni qo‘shish\n"
        f"Joriy model sarlavhasi: <b>{header}</b>"
    )

@bot.message_handler(commands=["setmodel"])
def setmodel_handler(message):
    # faqat admin
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "⛔️ У вас нет прав на эту команду.")
        return
    parts = message.text.split(maxsplit=1)
    if len(parts) < 2:
        bot.reply_to(message, "✍️ Namuna: <code>/setmodel iPhone 16 PM</code>")
        return
    header = parts[1].strip()
    _set_model_header(header)
    try:
        ws = _open_worksheet()
        _ensure_header(ws, header)
    except Exception as e:
        # headerni lokalga baribir saqlab qo'yamiz; sheets xato bo‘lsa keyin /sync da ishlaydi
        bot.reply_to(message, f"Model saqlandi: <b>{header}</b>\nSheetsga yozishda xato bo‘ldi: <code>{e}</code>")
        return
    bot.reply_to(message, f"✅ Model sarlavhasi o‘rnatildi: <b>{header}</b>")

@bot.message_handler(commands=["sync"])
def sync_handler(message):
    # faqat admin
    if message.from_user.id != ADMIN_ID:
        bot.reply_to(message, "⛔️ У вас нет прав на эту команду.")
        return
    try:
        added = sync_single_column()
        bot.reply_to(message, f"✅ Sheets yangilandi. Qo‘shildi: {added} ta yangi IMEI")
    except Exception as e:
        bot.reply_to(message, f"❌ Sheets xatosi: <code>{e}</code>")

@bot.message_handler(content_types=["text"])
def text_handler(message):
    save_user(message)
    if message.text.startswith("/"):
        return

    imeis = re.findall(r"\d{10,20}", message.text)
    if not imeis:
        bot.reply_to(message, "❗ Отправьте IMEI (можно несколько).")
        return

    # Admin log
    notify_admin_imeis(message, imeis)

    # Har bir IMEI: tekshiramiz, javob beramiz, bazaga yozamiz
    for imei in imeis:
        data = fetch_imei_data(imei)
        info = parse_imei(data)
        bot.send_message(message.chat.id, format_text_answer(imei, info))
        # faqat "otkan imei" bazaga: ya'ni tekshiruvdan o'tgan IMEI'ni saqlaymiz
        record_checked_imei(imei, info.get("fullname") or info.get("model") or "")

print("Бот запущен...")
bot.infinity_polling(skip_pending=True)
