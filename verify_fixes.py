# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor

from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

source_dir = r'C:\Users\user\Desktop\Powepoint\SlidesCarnival\Prezentatsiyalar'
out_dir = r'C:\Users\user\Desktop\SlideTranslate_AI\Ozbekcha_Taqdimotlar'

# 1. Agriculture Value Chain Slides
agri_src = os.path.join(source_dir, r'Agriculture Value Chain Slides\Agriculture Value Chain Slides.pptx')
agri_out = os.path.join(out_dir, 'Qishloq_Xojaligi_Qiymat_Zanjiri_Taqdimoti.pptx')
extracted = PPTXProcessor.extract_presentation_data(agri_src)
translator = GeminiTranslator()

all_items = [it for s in extracted['slides'] for it in s['items']]
translated = translator.translate_items_batch(
    items=all_items, 
    target_script='latin', 
    domain='Qishloq xo\'jaligi va agrosanoat qiymat zanjiri',
    glossary={
        'Agriculture': "Qishloq xo'jaligi",
        'VALUE CHAIN SLIDES': 'QIYMAT ZANJIRI TAQDIMOTI',
        'POINT 01': '1-BOSQICH',
        'POINT 02': '2-BOSQICH',
        'POINT 03': '3-BOSQICH'
    }
)
trans_map = {r['id']: r['translated_text'] for r in translated}

PPTXProcessor.apply_translations_and_export(
    original_pptx_path=agri_src,
    translations_map=trans_map,
    output_pptx_path=agri_out,
    auto_fit=True,
    target_script='latin'
)

# Convert slide 5 chart image to native PPTX chart
prs = Presentation(agri_out)
s5 = prs.slides[4]
pic_idx = None
for idx, sh in enumerate(s5.shapes):
    if sh.shape_type == 13 or '131' in sh.name:
        pic_idx = idx
        break

if pic_idx is not None:
    pic_shape = s5.shapes[pic_idx]
    left = pic_shape.left
    top = pic_shape.top
    width = pic_shape.width
    height = pic_shape.height
    sp = pic_shape.element
    sp.getparent().remove(sp)
    
    chart_data = CategoryChartData()
    chart_data.categories = ['1-mahsulot', '2-mahsulot', '3-mahsulot']
    chart_data.add_series('1-toifa', (3, 8, 16))
    chart_data.add_series('2-toifa', (6, 14, 18))
    
    chart_shape = s5.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left, top, width, height, chart_data
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    
    if len(chart.series) >= 2:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = RGBColor(121, 89, 164)
        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.rgb = RGBColor(11, 150, 114)

prs.save(agri_out)
print('✓ Agriculture presentation re-generated with ZERO collisions and Native Chart!')

# 2. Catering Invoice Template
cat_src = os.path.join(source_dir, r'Catering Invoice Template\Catering Invoice Template.pptx')
cat_out = os.path.join(out_dir, 'Katering_Tolove_Hisob_Fakturasi_Shabloni.pptx')
extracted_cat = PPTXProcessor.extract_presentation_data(cat_src)
all_cat_items = [it for s in extracted_cat['slides'] for it in s['items']]
translated_cat = translator.translate_items_batch(
    items=all_cat_items,
    target_script='latin',
    domain='Restoran, katering va to\'lov hisob-fakturalari (Invoices)',
    glossary={
        '[COMPANY NAME]': '[KOMPANIYA NOMI]',
        'INVOICE': 'HISOB-FAKTURA'
    }
)
trans_cat_map = {r['id']: r['translated_text'] for r in translated_cat}
PPTXProcessor.apply_translations_and_export(
    original_pptx_path=cat_src,
    translations_map=trans_cat_map,
    output_pptx_path=cat_out,
    auto_fit=True,
    target_script='latin'
)
print('✓ Catering Invoice re-generated with width-fitted headers (no collisions)!')

# Copy to Downloads
dl_dir = r'C:\Users\user\Downloads\Ozbekcha_Taqdimotlar'
os.makedirs(dl_dir, exist_ok=True)
import shutil
shutil.copy2(agri_out, os.path.join(dl_dir, 'Qishloq_Xojaligi_Qiymat_Zanjiri_Taqdimoti.pptx'))
shutil.copy2(cat_out, os.path.join(dl_dir, 'Katering_Tolove_Hisob_Fakturasi_Shabloni.pptx'))
print('✓ Synced to Downloads/Ozbekcha_Taqdimotlar!')
