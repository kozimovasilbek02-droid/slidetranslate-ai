# -*- coding: utf-8 -*-
import sys, os, time, json
sys.stdout.reconfigure(encoding='utf-8')

# Ensure path
sys.path.insert(0, r'C:\Users\user\Desktop\SlideTranslate_AI')

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

source_pptx = r"C:\Users\user\Desktop\Powepoint\SlidesCarnival\Prezentatsiyalar\Porter's Value Chain Analysis Infographics\Porter's Value Chain Analysis Infographics.pptx"
out_dir = r"C:\Users\user\Desktop\SlideTranslate_AI\Ozbekcha_Taqdimotlar"
os.makedirs(out_dir, exist_ok=True)
out_pptx = os.path.join(out_dir, "Porter_Qiymat_Zanjiri_Tahlili_Infografikasi.pptx")

print("="*65)
print("🚀 SOFF.UZ UCHUN SLAYDNI TARJIMA QILISH BOSHLANDI")
print(f"📄 Manba: {source_pptx}")
print("="*65)

# 1. Matnlarni ajratish
extracted = PPTXProcessor.extract_presentation_data(source_pptx)
print(f"✓ Ajratib olindi: {extracted['slides_count']} ta slayd, {extracted['total_items']} ta matn bloki.")

all_items = []
for s in extracted['slides']:
    for item in s['items']:
        all_items.append(item)

glossary = {
    "Porter's Value Chain": "Porter qiymat zanjiri",
    "Value Chain Analysis": "Qiymat zanjiri tahlili",
    "Primary Activities": "Asosiy faoliyat turlari",
    "Support Activities": "Yordamchi faoliyat turlari",
    "Inbound Logistics": "Kiruvchi logistika va ta'minot",
    "Outbound Logistics": "Chiquvchi logistika va yetkazib berish",
    "Operations": "Ishlab chiqarish va operatsiyalar",
    "Marketing & Sales": "Marketing va sotuv",
    "Service": "Servis va mijozlarga xizmat ko'rsatish",
    "Infrastructure": "Kompaniya infratuzilmasi",
    "Human Resource Management": "Inson resurslarini boshqarish (HR)",
    "Technology Development": "Texnologik rivojlanish va IT",
    "Procurement": "Xaridlar va moddiy ta'minot",
    "Margin": "Foyda marjasi",
    "Competitive Advantage": "Raqobat ustunligi"
}

# 2. Gemini orqali tarjima qilish
translator = GeminiTranslator()
print("Gemini AI orqali professional biznes va iqtisodiy uslubda tarjima qilinmoqda...")
translated_results = translator.translate_items_batch(
    items=all_items,
    target_script='latin',
    domain='Biznes strategiya, korxona boshqaruvi va Porter qiymat zanjiri tahlili',
    glossary=glossary
)

trans_map = {r['id']: r['translated_text'] for r in translated_results}

# 3. Tarjimani yangi PPTX faylga joylash
PPTXProcessor.apply_translations_and_export(
    original_pptx_path=source_pptx,
    translations_map=trans_map,
    output_pptx_path=out_pptx,
    auto_fit=True,
    target_script='latin'
)

print(f"\n✅ Muvaffaqiyatli tarjima qilindi va saqlandi:")
print(f"📁 Fayl manzili: {out_pptx}")
print(f"📊 Hajmi: {os.path.getsize(out_pptx)/(1024):.1f} KB")
