import os
filepath = r"C:\Users\user\Desktop\SlideTranslate_AI\bot.py"
with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

old_func = """async def run_bot():
    await bot.delete_webhook(drop_pending_updates=True)
    print("=" * 65)
    print("      🤖 SlideTranslate AI — Telegram Boti Ishga Tushdi!")
    print("      Har bir foydalanuvchi o'zining Gemini kalitidan foydalanadi.")
    print("=" * 65)
    await dp.start_polling(bot)"""

new_func = """def get_bot_mode() -> str:
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
    await dp.start_polling(bot)"""

if old_func in content:
    content = content.replace(old_func, new_func)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print("Replacement successful")
else:
    print("Old function not found exactly as string. Trying regex...")
    import re
    # Match from async def run_bot(): down to await dp.start_polling(...)
    pattern = re.compile(r'async def run_bot\(\):.*?await dp\.start_polling\(bot\)', re.DOTALL)
    if pattern.search(content):
        content = pattern.sub(new_func, content)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        print("Replacement via regex successful")
    else:
        print("Regex also failed to match.")