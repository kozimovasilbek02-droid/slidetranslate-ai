import os, uuid, re, urllib.parse, asyncio
from typing import Dict, Any, List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

from contextlib import asynccontextmanager

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
EXPORTS_DIR = os.path.join(BASE_DIR, "exports")

def get_bot_mode() -> str:
    mode = os.getenv("BOT_MODE", "").lower().strip()
    if mode:
        return mode
    if os.getenv("USE_WEBHOOK", "false").lower() == "true":
        return "webhook"
    return "polling"

async def bot_worker():
    bot_mode = get_bot_mode()
    if bot_mode in ["webhook", "none", "disabled", "off"]:
        print(f"ℹ️ Telegram bot mode is '{bot_mode}'. Background polling worker skipped.")
        return
        
    await asyncio.sleep(2)
    try:
        from bot import dp, bot as telegram_bot
        print("🤖 SlideTranslate Telegram Bot 24/7 doimiy polling rejimida ishga tushmoqda...")
        while True:
            try:
                await telegram_bot.delete_webhook(drop_pending_updates=True)
                print("✅ SlideTranslate Telegram Bot polling faol va xabarlarni qabul qilmoqda!")
                await dp.start_polling(telegram_bot)
            except asyncio.CancelledError:
                break
            except Exception as e:
                print(f"⚠️ Telegram botda xatolik: {e}, 5 soniyada qayta urinmoqda...")
                await asyncio.sleep(5)
    except Exception as e:
        print(f"Telegram botni ishga tushirishda xatolik: {e}")

async def keep_alive_pinger():
    await asyncio.sleep(30)
    url = os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL") or "https://slidetranslate-ai.onrender.com"
    health_url = f"{url.rstrip('/')}/health"
    while True:
        try:
            import httpx
            async with httpx.AsyncClient() as client:
                await client.get(health_url, timeout=15)
        except Exception:
            pass
        await asyncio.sleep(600)

@asynccontextmanager
async def lifespan(app: FastAPI):
    task1 = asyncio.create_task(bot_worker())
    task2 = asyncio.create_task(keep_alive_pinger())
    yield
    task1.cancel()
    task2.cancel()
    try:
        from bot import bot as telegram_bot
        await telegram_bot.session.close()
    except Exception:
        pass

app = FastAPI(title="SlideTranslate AI", version="2.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Translation-Warning", "X-Font-Warning"],
)

@app.api_route("/", methods=["GET", "HEAD"])
@app.api_route("/health", methods=["GET", "HEAD"])
@app.api_route("/api/v1/health", methods=["GET", "HEAD"])
async def health_check():
    return {"status": "ok", "service": "SlideTranslate AI", "version": "2.5.0", "uptime": "24/7 active"}

@app.post("/webhook")
async def telegram_webhook(update: dict):
    bot_mode = get_bot_mode()
    if bot_mode != "webhook":
        return {"ok": False, "error": f"Webhook is disabled (current bot mode: '{bot_mode}')"}
        
    try:
        from aiogram.types import Update as TgUpdate
        from bot import dp, bot as telegram_bot
        telegram_update = TgUpdate.model_validate(update, context={"bot": telegram_bot})
        await dp.feed_update(telegram_bot, telegram_update)
    except Exception as e:
        print(f"Error handling webhook update: {e}")
    return {"ok": True}

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(EXPORTS_DIR, exist_ok=True)

SESSIONS: Dict[str, Dict[str, Any]] = {}

DICT_TITLES = {
    "人力资源部": "Inson Resurslari Departamenti",
    "人力资源": "Inson Resurslari",
    "工作总结": "Ish Xulosasi",
    "工作汇报": "Ish Hisoboti",
    "总结": "Xulosa",
    "汇报": "Hisobot",
    "PPT模板": "Taqdimot Shabloni",
    "模板": "Shabloni",
    "年度": "Yillik",
    "招聘": "Ishga Qabul",
    "培训": "Trening",
    "方案": "Rejasi"
}

def translate_clean_filename(raw_filename: str, translator: GeminiTranslator, target_script: str = "latin") -> str:
    base, _ = os.path.splitext(raw_filename)
    base = re.sub(r"[\(\[\{]\d+[\)\]\}]", "", base)
    base = re.sub(r"[-_]?\s*(tarjima(si)?|ozbekcha|uz|translated|translation)", "", base, flags=re.IGNORECASE).strip()
    
    translated_ok = False
    try:
        translated = translator.translate_single_text(base, target_script=target_script)
        if translated and translated.strip() and translated.strip() != base.strip():
            base = translated.strip()
            translated_ok = True
    except Exception:
        pass
        
    if not translated_ok:
        for k, v in DICT_TITLES.items():
            base = base.replace(k, v)
        
    base = re.sub(r'[/\\:*?"<>|_]', " ", base)
    base = re.sub(r"(\s*[-_]?\s*[Tt]arjima(si)?|\s*[-_]?\s*[Oo][’'`]?zbekcha)$", "", base, flags=re.IGNORECASE)
    base = re.sub(r"\s+", " ", base).strip()
    return base if base else "Taqdimot"

class TranslateRequest(BaseModel):
    session_id: str
    api_key: Optional[str] = None
    target_script: str = "latin"
    domain: str = "general"
    glossary: Optional[Dict[str, str]] = None

class ExportRequest(BaseModel):
    session_id: str
    auto_fit: bool = True
    target_script: str = "latin"
    api_key: Optional[str] = None

@app.post("/api/upload")
async def upload_pptx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".pptx", ".potx")):
        raise HTTPException(status_code=400, detail="Faqat .pptx formatidagi fayllar qabul qilinadi.")
    session_id = str(uuid.uuid4())
    save_path = os.path.join(UPLOADS_DIR, f"{session_id}_{file.filename}")
    with open(save_path, "wb") as f:
        f.write(await file.read())
    data = PPTXProcessor.extract_presentation_data(save_path)
    SESSIONS[session_id] = {
        "session_id": session_id,
        "filename": file.filename,
        "pptx_path": save_path,
        "slides": data["slides"],
        "slides_count": data["slides_count"],
        "total_items": data["total_items"]
    }
    return {
        "session_id": session_id,
        "filename": file.filename,
        "slides_count": data["slides_count"],
        "total_items": data["total_items"],
        "slides": data["slides"]
    }

@app.post("/api/translate")
async def translate_presentation(req: TranslateRequest):
    session = SESSIONS.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessiya topilmadi")
    translator = GeminiTranslator(api_key=req.api_key)
    all_items = [it for s in session.get("slides", []) for it in s.get("items", [])]
    if not all_items:
        return {"status": "success", "slides": session["slides"]}
    batch_size = 75
    translations_map = {}
    for i in range(0, len(all_items), batch_size):
        batch = all_items[i:i + batch_size]
        res = translator.translate_items_batch(batch, target_script=req.target_script, domain=req.domain, glossary=req.glossary)
        for r in res:
            translations_map[r["id"]] = r["translated_text"]
    for slide in session["slides"]:
        for item in slide["items"]:
            if item["id"] in translations_map:
                item["translated_text"] = translations_map[item["id"]]
    return {
        "status": "success",
        "session": session,
        "slides": session["slides"]
    }

@app.post("/api/export")
async def export_presentation(req: ExportRequest):
    session = SESSIONS.get(req.session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessiya topilmadi")
    orig_pptx = session["pptx_path"]
    orig_name = session.get("filename", "presentation.pptx")
    
    translator = GeminiTranslator(api_key=req.api_key)
    clean_title = translate_clean_filename(orig_name, translator, req.target_script)
    if not clean_title.lower().endswith(".pptx"):
        out_filename = f"{clean_title}.pptx"
    else:
        out_filename = clean_title
        
    out_path = os.path.join(EXPORTS_DIR, f"{req.session_id}_{out_filename}")
    
    trans_map = {it["id"]: it.get("translated_text", "") for s in session.get("slides", []) for it in s.get("items", [])}
    PPTXProcessor.apply_translations_and_export(orig_pptx, trans_map, out_path, auto_fit=req.auto_fit, target_script=req.target_script)
    
    return {
        "status": "success",
        "download_url": f"/api/download/{req.session_id}/{urllib.parse.quote(out_filename)}",
        "download_filename": out_filename
    }

@app.get("/api/download/{session_id}/{filename}")
async def download_file(session_id: str, filename: str):
    filename = urllib.parse.unquote(filename)
    if not filename.lower().endswith(".pptx"):
        filename += ".pptx"
        
    fpath = os.path.join(EXPORTS_DIR, f"{session_id}_{filename}")
    if not os.path.exists(fpath):
        for f in os.listdir(EXPORTS_DIR):
            if f.startswith(session_id):
                fpath = os.path.join(EXPORTS_DIR, f)
                break
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="Fayl topilmadi.")
        
    safe_ascii = re.sub(r'[^\w\s.-]', '', filename)
    if not safe_ascii.lower().endswith(".pptx"):
        safe_ascii += ".pptx"
    enc = urllib.parse.quote(filename)
    headers = {
        "Content-Disposition": f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{enc}',
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }
    return FileResponse(fpath, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers=headers, filename=safe_ascii)

@app.get("/api/fonts")
async def get_fonts():
    from backend.core.font_manager import font_manager, FONTS_DIR
    return {
        "status": "success",
        "total_fonts": len(font_manager.registry),
        "fonts": font_manager.registry
    }

@app.post("/api/fonts/preload")
async def preload_fonts():
    from backend.core.font_manager import font_manager
    res = font_manager.preload_top_100_fonts(max_workers=8)
    return {"status": "success", "result": res}

@app.get("/api/fonts/css")
async def get_fonts_css():
    from backend.core.font_manager import font_manager
    from fastapi.responses import Response
    css_content = font_manager.generate_font_css()
    return Response(content=css_content, media_type="text/css")

@app.get("/api/fonts/file/{filename}")
async def get_font_file(filename: str):
    from backend.core.font_manager import FONTS_DIR
    safe_name = os.path.basename(filename)
    font_path = os.path.join(FONTS_DIR, safe_name)
    if not os.path.exists(font_path):
        raise HTTPException(status_code=404, detail="Shrift fayli topilmadi")
    return FileResponse(font_path, media_type="font/ttf")

@app.post("/api/v1/translate")
async def translate_v1_direct(
    file: UploadFile = File(...),
    target_lang: str = Form("uz-Latn"),
    api_key: Optional[str] = Form(None),
    domain: str = Form("general"),
    auto_fit: bool = Form(True)
):
    if not file.filename or not file.filename.lower().endswith((".pptx", ".potx")):
        raise HTTPException(status_code=400, detail="Faqat .pptx formatidagi fayllar qabul qilinadi.")

    target_script = "cyrillic" if target_lang.lower() in ["uz-cyrl", "cyrillic", "kirill"] else "latin"
    session_id = str(uuid.uuid4())
    in_path = os.path.join(UPLOADS_DIR, f"{session_id}_{file.filename}")
    with open(in_path, "wb") as f:
        f.write(await file.read())

    # Extract presentation data
    extracted = PPTXProcessor.extract_presentation_data(in_path)
    all_items = [it for s in extracted.get("slides", []) for it in s.get("items", [])]

    translator = GeminiTranslator(api_key=api_key)
    stats: Dict[str, Any] = {}
    if all_items:
        translated_results = translator.translate_items_batch(
            items=all_items,
            target_script=target_script,
            domain=domain,
            stats=stats
        )
        trans_map = {r["id"]: r["translated_text"] for r in translated_results}
    else:
        trans_map = {}

    clean_title = translate_clean_filename(file.filename, translator, target_script)
    if not clean_title.lower().endswith(".pptx"):
        out_filename = f"{clean_title}.pptx"
    else:
        out_filename = clean_title

    out_path = os.path.join(EXPORTS_DIR, f"{session_id}_{out_filename}")
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=in_path,
        translations_map=trans_map,
        output_pptx_path=out_path,
        auto_fit=auto_fit,
        target_script=target_script
    )

    safe_ascii = re.sub(r'[^\w\s.-]', '', out_filename)
    if not safe_ascii.lower().endswith(".pptx"):
        safe_ascii += ".pptx"
    enc = urllib.parse.quote(out_filename)

    headers = {
        "Content-Disposition": f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{enc}',
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }
    if stats.get("failed", 0) > 0:
        headers["X-Translation-Warning"] = f"{stats['failed']}/{stats.get('total', 0)} ta matn tarjima qilinmadi"

    return FileResponse(out_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers=headers, filename=safe_ascii)

dist = os.path.join(BASE_DIR, "frontend", "dist")
if os.path.exists(dist):
    app.mount("/", StaticFiles(directory=dist, html=True), name="frontend")