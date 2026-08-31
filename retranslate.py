# -*- coding: utf-8 -*-
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

orig_path = r'C:\Users\user\Desktop\SlideTranslate_AI\temp_sessions\2abdc0f9-c72d-4b1a-a869-9cd0d9591731\original.pptx'
out_path = r'C:\Users\user\Downloads\Agriculture Value Chain Slides_Tarjima.pptx'
out_path_v2 = r'C:\Users\user\Downloads\Agriculture Value Chain Slides_Tarjima_Mukammal.pptx'

extracted = PPTXProcessor.extract_presentation_data(orig_path)
print(f'Extracted {extracted["slides_count"]} slides, {extracted["total_items"]} items.')

all_items = []
for s in extracted['slides']:
    for item in s['items']:
        all_items.append(item)

glossary = {
    'Agriculture': "Qishloq xo'jaligi",
    'Value Chain': "Qiymat zanjiri",
    'Pre-production': "Xomashyo va tayyorgarlik bosqichi",
    'Post-production': "Hosilni yig'ish va saqlash",
    'Production': "Yetishtirish va ishlab chiqarish",
    'Processing': "Sanoatda qayta ishlash",
    'Point 01': "01-ko'rsatkich",
    'Point 02': "02-ko'rsatkich",
    'Point 03': "03-ko'rsatkich",
    'Credits': "Mualliflik huquqlari va minnatdorchilik"
}

translator = GeminiTranslator()
print('Translating with enhanced Gemini translator...')
translated_results = translator.translate_items_batch(
    items=all_items,
    target_script='latin',
    domain='Qishloq xo\'jaligi, agrosanoat va biznes taqdimotlari',
    glossary=glossary
)

trans_map = {r['id']: r['translated_text'] for r in translated_results}

print('\n--- Sample Results ---')
for item_id, text in list(trans_map.items())[:12]:
    print(f'  {item_id}: {text}')

PPTXProcessor.apply_translations_and_export(
    original_pptx_path=orig_path,
    translations_map=trans_map,
    output_pptx_path=out_path,
    auto_fit=True,
    target_script='latin'
)
PPTXProcessor.apply_translations_and_export(
    original_pptx_path=orig_path,
    translations_map=trans_map,
    output_pptx_path=out_path_v2,
    auto_fit=True,
    target_script='latin'
)

print(f'\n[SUCCESS] Generated perfected presentations at:\n  1. {out_path}\n  2. {out_path_v2}')
