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
    r'Accountant CV Resume\Accountant CV Resume.pptx',
    r'Basic Investment Decision Tree Infographics\Basic Investment Decision Tree Infographics.pptx',
    r'Catering Invoice Template\Catering Invoice Template.pptx',
    r'Counting Shapes Math Worksheet\Counting Shapes Math Worksheet.pptx',
    r'HR Value Chain Slides\HR Value Chain Slides.pptx',
    r'Idea SWOT Analysis Infographic Template\Idea SWOT Analysis Infographic Template.pptx',
    r'Memory Hierarchy Infographic\Memory Hierarchy Infographic.pptx',
    r'Oil And Gas Value Chain Slides\Oil And Gas Value Chain Slides.pptx',
    r'Porter\'s Value Chain Analysis Infographics\Porter\'s Value Chain Analysis Infographics.pptx',
    r'Startup Executive Summary Slides\Startup Executive Summary Slides.pptx'
]

domain_map = {
    'Accountant': 'Buxgalteriya, moliya va rezyume (CV)',
    'Investment': 'Investitsiyalar, bank ishi va moliya',
    'Catering': 'Restoran, katering va to\'lov hisob-fakturalari (Invoices)',
    'Math': 'Boshlang\'ich matematika va geometriya darsligi',
    'HR': 'Inson resurslari (HR) va kadrlar boshqaruvi',
    'SWOT': 'Biznes strategiya, SWOT tahlil va marketing',
    'Memory': 'Kompyuter arxitekturasi va axborot texnologiyalari (IT)',
    'Oil': 'Neft-gaz sanoati va energetika qiymat zanjiri',
    'Porter': 'Strategik menejment va Porter qiymat zanjiri tahlili',
    'Startup': 'Startaplar, venchur investitsiyalar va biznes reja'
}

translator = GeminiTranslator()
results_summary = []

print('=' * 75)
print('🚀 10 TA TAQDIMOTNI AVTOMATIK TARJIMA VA ANALIZ QILISH BOSHLANDI')
print('=' * 75)

for idx, rel_path in enumerate(test_files, 1):
    full_src = os.path.join(source_dir, rel_path)
    base_name = os.path.splitext(os.path.basename(full_src))[0]
    out_path = os.path.join(out_dir, f'{base_name}_Tarjima_UZ.pptx')
    
    print(f'\n[{idx}/10] Qayta ishlanmoqda: {base_name}')
    start_t = time.time()
    
    # 1. Extraction
    try:
        extracted = PPTXProcessor.extract_presentation_data(full_src)
        slides_count = extracted['slides_count']
        total_items = extracted['total_items']
        print(f'  ✓ Ajratib olindi: {slides_count} ta slayd, {total_items} ta matn bloki')
    except Exception as e:
        print(f'  ✗ Extraction xatoligi: {e}')
        results_summary.append({'name': base_name, 'status': 'EXTRACT_FAIL', 'error': str(e)})
        continue

    # Flatten items
    all_items = []
    for s in extracted['slides']:
        for item in s['items']:
            all_items.append(item)

    if not all_items:
        print('  ! Matn topilmadi')
        continue

    # Detect domain
    domain = 'Biznes va taqdimotlar'
    for key, dom in domain_map.items():
        if key.lower() in base_name.lower():
            domain = dom
            break

    # 2. Translation via Gemini
    try:
        translated_results = translator.translate_items_batch(
            items=all_items,
            target_script='latin',
            domain=domain
        )
        trans_map = {r['id']: r['translated_text'] for r in translated_results}
        print(f'  ✓ Tarjima qilindi: {len(trans_map)}/{len(all_items)} ta element (Soha: {domain})')
    except Exception as e:
        print(f'  ✗ Translation xatoligi: {e}')
        results_summary.append({'name': base_name, 'status': 'TRANSLATE_FAIL', 'error': str(e)})
        continue

    # 3. Apply & Export
    try:
        PPTXProcessor.apply_translations_and_export(
            original_pptx_path=full_src,
            translations_map=trans_map,
            output_pptx_path=out_path,
            auto_fit=True,
            target_script='latin'
        )
        elapsed = time.time() - start_t
        out_size = os.path.getsize(out_path)
        print(f'  ✓ Yangi PPTX saqlandi: {out_size} bayt ({elapsed:.2f}s)')
        results_summary.append({
            'name': base_name,
            'status': 'SUCCESS',
            'slides': slides_count,
            'items': total_items,
            'time_sec': round(elapsed, 2),
            'domain': domain,
            'out_path': out_path
        })
    except Exception as e:
        print(f'  ✗ Export xatoligi: {e}')
        results_summary.append({'name': base_name, 'status': 'EXPORT_FAIL', 'error': str(e)})

print('\n' + '=' * 75)
print('📊 UMUMIY NATIJALAR HISOBOTI:')
print('=' * 75)
print(json.dumps(results_summary, ensure_ascii=False, indent=2))
