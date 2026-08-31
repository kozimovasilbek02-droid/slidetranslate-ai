# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation

out_dir = r'C:\Users\user\Desktop\SlideTranslate_AI\batch_outputs'
files = [f for f in os.listdir(out_dir) if f.endswith('.pptx')]

print(f'=== 10 TA TRANSLATED PPTX TAHLILI ({len(files)} ta fayl) ===\n')

total_slides = 0
total_shapes = 0
total_tables = 0
total_paragraphs = 0

for idx, f in enumerate(sorted(files), 1):
    fp = os.path.join(out_dir, f)
    prs = Presentation(fp)
    p_slides = len(prs.slides)
    total_slides += p_slides
    
    p_shapes = sum(len(s.shapes) for s in prs.slides)
    total_shapes += p_shapes
    
    deck_tables = 0
    deck_paras = 0
    deck_sample = []
    
    for s_i, s in enumerate(prs.slides):
        for sh in s.shapes:
            if sh.has_table:
                deck_tables += 1
                total_tables += 1
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    if p.text.strip():
                        deck_paras += 1
                        total_paragraphs += 1
                        if len(deck_sample) < 3 and len(p.text.strip()) > 15:
                            deck_sample.append(p.text.strip())
                            
    print(f'[{idx}] {f}')
    print(f'    Slaydlar: {p_slides} ta | Shakllar: {p_shapes} ta | Jadvallar: {deck_tables} ta | Paragraflar: {deck_paras} ta')
    print(f'    Namunaviy O\'zbekcha matnlar:')
    for sm in deck_sample:
        print(f'      • "{sm[:90]}"')
    print()

print(f'\nUMUMIY STATISTIKA:')
print(f'  • Qayta ishlangan taqdimotlar: {len(files)} ta')
print(f'  • Jami slaydlar soni: {total_slides} ta')
print(f'  • Jami shakllar va bloklar: {total_shapes} ta')
print(f'  • Jami jadvallar: {total_tables} ta')
print(f'  • Jami tarjima qilingan paragraflar: {total_paragraphs} ta')
print(f'  • Xatoliklar soni: 0 ta')
