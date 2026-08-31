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
Aiogram 3 + Gemini 3.6 Flash + PPTXProcessor + FontManager.
"""

import os
import sys
import re
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
from backend.main import translate_clean_filename, UPLOADS_DIR, EXPORTS_DIR

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Token configuration
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8899026470:AAFa4WG85YKIEw0q2LPKirDeKrPWdC7kpqE").strip()

# In-memory user settings storage
USER_SETTINGS: Dict[int, Dict[str, Any]] = {}

def get_user_settings(user_id: int) -> Dict[str, Any]:
    if user_id not in USER_SETTINGS:
        USER_SETTINGS[user_id] = {
            "target_script": "latin", # 'latin' or 'cyrillic'
            "domain": "IT & Dasturlash",
            "auto_fit": True,
            "api_key": None
        }
    return USER_SETTINGS[user_id]

# Keyboard Builders
def get_main_keyboard(user_id: int) -> InlineKeyboardMarkup:
    st = get_user_settings(user_id)
    script_label = "🔤 Yozuv: Lotin" if st["target_script"] == "latin" else "🔤 Ёзув: Кирилл"
    domain_label = f"🏢 Soha: {st['domain']}"
    autofit_label = "⚡ Auto-fit: Yoqilgan" if st["auto_fit"] else "⚡ Auto-fit: O'chirilgan"

    kb = [
        [
            InlineKeyboardButton(text=script_label, callback_data="toggle_script"),
            InlineKeyboardButton(text=autofit_label, callback_data="toggle_autofit")
        ],
        [
            InlineKeyboardButton(text=domain_label, callback_data="change_domain")
        ],
        [
            InlineKeyboardButton(text="🔤 Shriftlar bazasi (130+)", callback_data="show_fonts"),
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

# Initialize Dispatcher
dp = Dispatcher()

@dp.message(CommandStart())
async def cmd_start(msg: Message):
    user_id = msg.from_user.id
    settings = get_user_settings(user_id)
    welcome_text = (
        f"👋 <b>Assalomu alaykum, {msg.from_user.first_name}!</b>\n\n"
        f"🚀 <b>SlideTranslate AI Botiga xush kelibsiz!</b>\n"
        f"Men PowerPoint (<b>.pptx</b>) taqdimotlaringizni dizayni, shakllari, jadvallari va "
        f"<b>130+ shriftlarini</b> buzmagan holda O'zbek tiliga professional tarjima qilib beraman.\n\n"
        f"📁 <b>Ishni boshlash uchun menga istalgan .pptx fayl yuboring!</b>"
    )
    await msg.answer(welcome_text, parse_mode="HTML", reply_markup=get_main_keyboard(user_id))

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
        "1. Menga <b>.pptx</b> formatidagi taqdimot faylini yuboring.\n"
        "2. Bot barcha slaydlarni tahlil qiladi va Gemini 3.6 Flash yordamida tarjima qiladi.\n"
        "3. Tayyor bo'lgan faylni yuklab oling!\n\n"
        "⚡ <b>Xususiyatlar:</b>\n"
        "• 100% Dizayn va ranglar saqlanadi\n"
        "• Reklama va Slide Master logotiplari tozalanadi\n"
        "• So'zlar qutilardan toshib ketmaydi (Anti-overflow)\n"
        "• Lotin va Kirill yozuvlarini qo'llab-quvvatlaydi"
    )
    await msg.answer(help_text, parse_mode="HTML")

# Callbacks
@dp.callback_query(F.data == "toggle_script")
async def cb_toggle_script(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    st["target_script"] = "cyrillic" if st["target_script"] == "latin" else "latin"
    await cb.message.edit_reply_markup(reply_markup=get_main_keyboard(user_id))
    await cb.answer(f"Yozuv o'zgartirildi: {'Kirill' if st['target_script'] == 'cyrillic' else 'Lotin'}")

@dp.callback_query(F.data == "toggle_autofit")
async def cb_toggle_autofit(cb: CallbackQuery):
    user_id = cb.from_user.id
    st = get_user_settings(user_id)
    st["auto_fit"] = not st["auto_fit"]
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
        "Dasturchi: Google Antigravity AI Engine"
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
            f"⚡ <b>Gemini 3.6 Flash orqali tarjima qilinmoqda...</b>",
            parse_mode="HTML"
        )

        # 3. Translate with Gemini Translator
        translator = GeminiTranslator(api_key=st["api_key"])
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

        # 6. Edit status and send translated document
        await status_msg.edit_text("✅ <b>Tarjima tayyor! Fayl yuborilmoqda...</b>", parse_mode="HTML")

        caption = (
            f"🎉 <b>Taqdimotingiz muvaffaqiyatli tarjima qilindi!</b>\n\n"
            f"📊 <b>Natijalar:</b>\n"
            f"• Slaydlar: <b>{slides_count} ta</b>\n"
            f"• Matn bloklari: <b>{total_items} ta</b>\n"
            f"• Yozuv: <b>{'Lotin' if target_script == 'latin' else 'Кирилл'}</b>\n"
            f"• Shriftlar va dizayn: <b>100% Saqlangan</b>\n\n"
            f"<i>SlideTranslate AI — Gemini 3.6 Flash</i>"
        )

        input_file = FSInputFile(path=out_path, filename=out_filename)
        await msg.reply_document(document=input_file, caption=caption, parse_mode="HTML")
        await status_msg.delete()

    except Exception as e:
        logger.error(f"Xatolik yuz berdi: {e}", exc_info=True)
        await status_msg.edit_text(f"❌ <b>Xatolik yuz berdi:</b> {str(e)}", parse_mode="HTML")

async def run_bot():
    global BOT_TOKEN
    if not BOT_TOKEN:
        print("=" * 65)
        print("  ⚠️ TELEGRAM_BOT_TOKEN o'rnatilmagan!")
        print("  Iltimos, Telegram Bot tokeningizni kiriting (@BotFather dan olingan):")
        print("=" * 65)
        try:
            BOT_TOKEN = input("Bot Token: ").strip()
        except EOFError:
            BOT_TOKEN = ""

    if not BOT_TOKEN:
        print("Xatolik: Bot Token kiritilmadi. Bot to'xtatildi.")
        return

    bot = Bot(token=BOT_TOKEN)
    print("=" * 65)
    print("      🤖 SlideTranslate AI — Telegram Boti Ishga Tushdi!")
    print("      Foydalanuvchilar buyruqlari va PPTX fayllari kutilmoqda...")
    print("=" * 65)
    await dp.start_polling(bot)

def main():
    try:
        asyncio.run(run_bot())
    except (KeyboardInterrupt, SystemExit):
        print("\nBot to'xtatildi.")

if __name__ == "__main__":
    main()
