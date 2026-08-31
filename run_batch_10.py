# -*- coding: utf-8 -*-
import sys, os, time, json
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

source_dir = r'C:\Users\user\Desktop\Powepoint\SlidesCarnival\Prezentatsiyalar'
out_dir = r'C:\Users\user\Desktop\SlideTranslate_AI\batch_outputs'
os.makedirs(out_dir, exist_ok=True)

test_files = [
    (r'Basic Investment Decision Tree Infographics\Basic Investment Decision Tree Infographics.pptx', 'Investitsiyalar, bank ishi va moliya'),
    (r'Catering Invoice Template\Catering Invoice Template.pptx', 'Restoran, katering va to\'lov hisob-fakturalari (Invoices)'),
    (r'Counting Shapes Math Worksheet\Counting Shapes Math Worksheet.pptx', 'Boshlang\'ich matematika va geometriya darsligi'),
    (r'HR Value Chain Slides\HR Value Chain Slides.pptx', 'Inson resurslari (HR) va kadrlar boshqaruvi'),
    (r'Idea SWOT Analysis Infographic Template\Idea SWOT Analysis Infographic Template.pptx', 'Biznes strategiya, SWOT tahlil va marketing'),
    (r'Memory Hierarchy Infographic\Memory Hierarchy Infographic.pptx', 'Kompyuter arxitekturasi va axborot texnologiyalari (IT)'),
    (r'Oil And Gas Value Chain Slides\Oil And Gas Value Chain Slides.pptx', 'Neft-gaz sanoati va energetika qiymat zanjiri'),
    (r'Startup Executive Summary Slides\Startup Executive Summary Slides.pptx', 'Startaplar, venchur investitsiyalar va biznes reja'),
    (r'Classroom Pledge Poster\Classroom Pledge Poster.pptx', 'Ta\'lim, maktab qoidalari va posterlar'),
    (r'Career Choice Decision Tree Infographics\Career Choice Decision Tree Infographics.pptx', 'Kasb tanlash, karyera va infografik qarorlar daraxti')
]

translator = GeminiTranslator()
results = []

print('=' * 75)
print('🚀 10 TA TAQDIMOTNI TO\'LIQ AVTOMATIK TARJIMA VA TAHLIL QILISH')
print('=' * 75)

for idx, (rel_p, domain) in enumerate(test_files, 1):
    full_src = os.path.join(source_dir, rel_p)
    base_name = os.path.splitext(os.path.basename(full_src))[0]
    out_path = os.path.join(out_dir, f'{base_name}_Tarjima_UZ.pptx')
    
    if os.path.exists(out_path) and os.path.getsize(out_path) > 1000:
        print(f'[{idx}/10] Oldindan mavjud: {base_name}')
        prs = Presentation(out_path)
        extracted = PPTXProcessor.extract_presentation_data(out_path)
        results.append({
            'index': idx,
            'name': base_name,
            'status': 'SUCCESS',
            'slides': len(prs.slides),
            'items': extracted['total_items'],
            'domain': domain,
            'out_path': out_path,
            'size_bytes': os.path.getsize(out_path)
        })
        continue

    print(f'[{idx}/10] Qayta ishlanmoqda: {base_name}')
    start_t = time.time()
    try:
        extracted = PPTXProcessor.extract_presentation_data(full_src)
        all_items = [it for s in extracted['slides'] for it in s['items']]
        translated = translator.translate_items_batch(items=all_items, target_script='latin', domain=domain)
        trans_map = {r['id']: r['translated_text'] for r in translated}
        PPTXProcessor.apply_translations_and_export(full_src, trans_map, out_path, auto_fit=True, target_script='latin')
        elapsed = time.time() - start_t
        results.append({
            'index': idx,
            'name': base_name,
            'status': 'SUCCESS',
            'slides': extracted['slides_count'],
            'items': len(all_items),
            'time_sec': round(elapsed, 2),
            'domain': domain,
            'out_path': out_path,
            'size_bytes': os.path.getsize(out_path)
        })
        print(f'  ✓ Tayyorlandi: {os.path.getsize(out_path)} bayt ({elapsed:.2f}s)')
    except Exception as e:
        print(f'  ✗ Xatolik: {e}')
        results.append({'index': idx, 'name': base_name, 'status': 'FAIL', 'error': str(e)})

print('\n' + '=' * 75)
print('🎉 10 TA TAQDIMOT NATIJALARI HISOBOTI:')
print('=' * 75)
print(json.dumps(results, ensure_ascii=False, indent=2))
