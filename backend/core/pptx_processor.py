# -*- coding: utf-8 -*-
import os
import re
from typing import Dict, Any, List, Optional
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE

from backend.core.transliteration import ensure_script
from backend.core.font_manager import font_manager

def sanitize_control_chars(text: str) -> str:
    if not text:
        return ""
    t = re.sub(r'_x[0-9a-fA-F]{4}_', ' ', text)
    t = re.sub(r'[\u0001-\u0008\u000b\u000c\u000e-\u001f\u007f]', ' ', t)
    t = re.sub(r'[ \t]+', ' ', t)
    return t.strip()

class PPTXProcessor:
    WATERMARK_PATTERNS = [
        r"presentationgo", r"ypppt", r"slidescarnival", r"slidesmania",
        r"slidesgo", r"poweredtemplate", r"slidemodel", r"designed with",
        r"free templates?", r"questions or need help", r"visit our faq",
        r"更多精品", r"ppt模板", r"ppt背景", r"by:\s*", r"\.com",
        r"51ppt", r"优品ppt"
    ]

    @staticmethod
    def _is_real_text(text: str) -> bool:
        if not text:
            return False
        t = sanitize_control_chars(text)
        if not t:
            return False
        if len(t) == 1 and ord(t[0]) >= 0xE000:
            return False
        return True

    @staticmethod
    def _is_watermark_text(txt: str) -> bool:
        if not txt:
            return False
        t = sanitize_control_chars(txt).lower()
        return any(re.search(pat, t) for pat in PPTXProcessor.WATERMARK_PATTERNS)

    @staticmethod
    def _is_watermark_recursive(shape) -> bool:
        if shape.has_text_frame and PPTXProcessor._is_watermark_text(shape.text_frame.text):
            return True
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                for sub in shape.shapes:
                    if PPTXProcessor._is_watermark_recursive(sub):
                        return True
            except Exception:
                pass
        return False

    @staticmethod
    def clean_presentation_watermarks(prs: Presentation) -> int:
        removed = 0
        
        # 1. Clean Slide Masters and Layouts
        for master in prs.slide_masters:
            for sh in list(master.shapes):
                if PPTXProcessor._is_watermark_recursive(sh):
                    try:
                        sh._element.getparent().remove(sh._element)
                        removed += 1
                    except Exception:
                        pass
                elif sh.shape_type != MSO_SHAPE_TYPE.PLACEHOLDER:
                    if sh.left is not None and sh.top is not None:
                        if (sh.left > 8000000 and sh.top < 500000) or sh.top > 6000000 or (sh.left < 500000 and sh.top < 500000):
                            try:
                                sh._element.getparent().remove(sh._element)
                                removed += 1
                            except Exception:
                                pass

            for layout in master.slide_layouts:
                for sh in list(layout.shapes):
                    if PPTXProcessor._is_watermark_recursive(sh):
                        try:
                            sh._element.getparent().remove(sh._element)
                            removed += 1
                        except Exception:
                            pass
                    elif sh.shape_type != MSO_SHAPE_TYPE.PLACEHOLDER:
                        if sh.left is not None and sh.top is not None:
                            if (sh.left > 8000000 and sh.top < 500000) or sh.top > 6000000 or (sh.left < 500000 and sh.top < 500000):
                                try:
                                    sh._element.getparent().remove(sh._element)
                                    removed += 1
                                except Exception:
                                    pass

        # 2. Clean Slide Level Watermarks
        for slide in prs.slides:
            for sh in list(slide.shapes):
                if PPTXProcessor._is_watermark_recursive(sh):
                    try:
                        sh._element.getparent().remove(sh._element)
                        removed += 1
                    except Exception:
                        pass

        return removed

    @staticmethod
    def _extract_font_meta(paragraph) -> Dict[str, Any]:
        font_name = "Calibri"
        font_size_pt = 16.0
        is_bold = False
        is_italic = False
        font_color = "#333333"

        if paragraph.runs:
            first_r = paragraph.runs[0]
            if first_r.font:
                if first_r.font.name:
                    font_name = first_r.font.name
                if first_r.font.size and first_r.font.size.pt:
                    font_size_pt = first_r.font.size.pt
                is_bold = bool(first_r.font.bold)
                is_italic = bool(first_r.font.italic)
                try:
                    if first_r.font.color and first_r.font.color.rgb:
                        font_color = f"#{first_r.font.color.rgb}"
                except Exception:
                    pass
        elif paragraph.font:
            if paragraph.font.name:
                font_name = paragraph.font.name
            if paragraph.font.size and paragraph.font.size.pt:
                font_size_pt = paragraph.font.size.pt
            is_bold = bool(paragraph.font.bold)
            is_italic = bool(paragraph.font.italic)

        try:
            font_manager.ensure_font_available(font_name)
        except Exception:
            pass

        return {
            "font_name": font_name,
            "font_size_pt": round(font_size_pt, 1),
            "is_bold": is_bold,
            "is_italic": is_italic,
            "font_color": font_color
        }

    @staticmethod
    def extract_presentation_data(pptx_path: str) -> Dict[str, Any]:
        if not os.path.exists(pptx_path):
            raise FileNotFoundError(f"PPTX topilmadi: {pptx_path}")
        prs = Presentation(pptx_path)
        slides_data = []
        total_items_count = 0
        slide_width = prs.slide_width
        slide_height = prs.slide_height

        for s_idx, slide in enumerate(prs.slides, start=1):
            slide_items = []
            PPTXProcessor._extract_shapes_recursive(
                shapes=slide.shapes,
                slide_index=s_idx,
                slide_width=slide_width,
                slide_height=slide_height,
                items_list=slide_items
            )
            total_items_count += len(slide_items)
            slides_data.append({
                "slide_index": s_idx,
                "slide_id": f"slide_{s_idx}",
                "items_count": len(slide_items),
                "items": slide_items
            })

        return {
            "slide_width_pt": slide_width.pt if slide_width else 960,
            "slide_height_pt": slide_height.pt if slide_height else 540,
            "slides_count": len(slides_data),
            "total_items": total_items_count,
            "slides": slides_data
        }

    @staticmethod
    def _extract_shapes_recursive(shapes, slide_index: int, slide_width, slide_height, items_list: list, prefix: str = ""):
        for sh_idx, shape in enumerate(shapes):
            sh_id_str = f"{prefix}sh{getattr(shape, 'shape_id', sh_idx)}"
            box_info = {"left": 5, "top": 5, "width": 90, "height": 20}
            try:
                if shape.left is not None and shape.top is not None and slide_width and slide_height:
                    box_info = {
                        "left": round((shape.left / slide_width) * 100, 2),
                        "top": round((shape.top / slide_height) * 100, 2),
                        "width": round((shape.width / slide_width) * 100, 2),
                        "height": round((shape.height / slide_height) * 100, 2)
                    }
            except Exception:
                pass

            # 1. Group shapes
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    PPTXProcessor._extract_shapes_recursive(shape.shapes, slide_index, slide_width, slide_height, items_list, prefix=f"{sh_id_str}_g")
                except Exception:
                    pass
                continue

            # 2. Table
            if shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if cell.text_frame:
                            for p_idx, p in enumerate(cell.text_frame.paragraphs):
                                p_text = sanitize_control_chars(p.text)
                                if PPTXProcessor._is_real_text(p_text) and not PPTXProcessor._is_watermark_text(p_text):
                                    f_meta = PPTXProcessor._extract_font_meta(p)
                                    item_id = f"s{slide_index}_{sh_id_str}_tbl_r{r_idx}_c{c_idx}_p{p_idx}"
                                    items_list.append({
                                        "id": item_id,
                                        "slide_index": slide_index,
                                        "shape_name": f"Jadval [Qator {r_idx+1}, Ustun {c_idx+1}]",
                                        "item_type": "table_cell",
                                        "original_text": p_text,
                                        "translated_text": p_text,
                                        "font_name": f_meta["font_name"],
                                        "font_size_pt": f_meta["font_size_pt"],
                                        "is_bold": f_meta["is_bold"],
                                        "is_italic": f_meta["is_italic"],
                                        "font_color": f_meta["font_color"],
                                        "box": box_info
                                    })
                continue

            # 3. Regular Shape / TextBox
            if shape.has_text_frame:
                tf = shape.text_frame
                for p_idx, p in enumerate(tf.paragraphs):
                    p_text = sanitize_control_chars(p.text)
                    if PPTXProcessor._is_real_text(p_text) and not PPTXProcessor._is_watermark_text(p_text):
                        f_meta = PPTXProcessor._extract_font_meta(p)
                        item_id = f"s{slide_index}_{sh_id_str}_p{p_idx}"
                        items_list.append({
                            "id": item_id,
                            "slide_index": slide_index,
                            "shape_name": shape.name or f"Shakl {sh_idx+1}",
                            "item_type": "title" if p_idx == 0 and len(p_text) < 40 else "body",
                            "original_text": p_text,
                            "translated_text": p_text,
                            "font_name": f_meta["font_name"],
                            "font_size_pt": f_meta["font_size_pt"],
                            "is_bold": f_meta["is_bold"],
                            "is_italic": f_meta["is_italic"],
                            "font_color": f_meta["font_color"],
                            "box": box_info
                        })

            # 4. Native PowerPoint Charts
            if shape.has_chart:
                try:
                    c_elem = shape.chart._element
                    for t_idx, t_node in enumerate(c_elem.xpath('.//c:title//a:t')):
                        t_text = sanitize_control_chars(t_node.text) if t_node.text else ""
                        if PPTXProcessor._is_real_text(t_text) and not PPTXProcessor._is_watermark_text(t_text):
                            item_id = f"s{slide_index}_{sh_id_str}_chtitle_{t_idx}"
                            items_list.append({
                                "id": item_id,
                                "slide_index": slide_index,
                                "shape_name": f"Diagramma Sarlavhasi ({shape.name or f'Shakl {sh_idx+1}'})",
                                "item_type": "title",
                                "original_text": t_text,
                                "translated_text": t_text,
                                "font_name": "Calibri",
                                "font_size_pt": 18.0,
                                "is_bold": True,
                                "is_italic": False,
                                "font_color": "#333333",
                                "box": box_info
                            })
                    for cat_idx, v_node in enumerate(c_elem.xpath('.//c:cat//c:pt//c:v')):
                        c_text = sanitize_control_chars(v_node.text) if v_node.text else ""
                        if PPTXProcessor._is_real_text(c_text) and not PPTXProcessor._is_watermark_text(c_text):
                            item_id = f"s{slide_index}_{sh_id_str}_chcat_{cat_idx}"
                            items_list.append({
                                "id": item_id,
                                "slide_index": slide_index,
                                "shape_name": f"Diagramma Toifasi [{c_text}]",
                                "item_type": "body",
                                "original_text": c_text,
                                "translated_text": c_text,
                                "font_name": "Calibri",
                                "font_size_pt": 14.0,
                                "is_bold": False,
                                "is_italic": False,
                                "font_color": "#333333",
                                "box": box_info
                            })
                    for ser_idx, v_node in enumerate(c_elem.xpath('.//c:ser//c:tx//c:v')):
                        s_text = sanitize_control_chars(v_node.text) if v_node.text else ""
                        if PPTXProcessor._is_real_text(s_text) and not PPTXProcessor._is_watermark_text(s_text):
                            item_id = f"s{slide_index}_{sh_id_str}_chser_{ser_idx}"
                            items_list.append({
                                "id": item_id,
                                "slide_index": slide_index,
                                "shape_name": f"Diagramma Seriyasi [{s_text}]",
                                "item_type": "body",
                                "original_text": s_text,
                                "translated_text": s_text,
                                "font_name": "Calibri",
                                "font_size_pt": 14.0,
                                "is_bold": False,
                                "is_italic": False,
                                "font_color": "#333333",
                                "box": box_info
                            })
                except Exception:
                    pass

    @staticmethod
    def apply_translations_and_export(
        original_pptx_path: str,
        translations_map: Dict[str, str],
        output_pptx_path: str,
        auto_fit: bool = True,
        target_script: str = "latin",
        clean_watermarks: bool = True
    ) -> str:
        if not os.path.exists(original_pptx_path):
            raise FileNotFoundError(f"Original PPTX topilmadi: {original_pptx_path}")

        prs = Presentation(original_pptx_path)

        if clean_watermarks:
            PPTXProcessor.clean_presentation_watermarks(prs)

        for s_idx, slide in enumerate(prs.slides, start=1):
            PPTXProcessor._apply_to_shapes_recursive(
                shapes=slide.shapes,
                slide_index=s_idx,
                translations_map=translations_map,
                auto_fit=auto_fit,
                target_script=target_script
            )

        os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
        prs.save(output_pptx_path)
        return output_pptx_path

    @staticmethod
    def _apply_to_shapes_recursive(shapes, slide_index: int, translations_map: Dict[str, str], auto_fit: bool, target_script: str, prefix: str = ""):
        for sh_idx, shape in enumerate(shapes):
            sh_id_str = f"{prefix}sh{getattr(shape, 'shape_id', sh_idx)}"

            # 1. Group
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    PPTXProcessor._apply_to_shapes_recursive(shape.shapes, slide_index, translations_map, auto_fit, target_script, prefix=f"{sh_id_str}_g")
                except Exception:
                    pass
                continue

            # 2. Table
            if shape.has_table:
                table = shape.table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if cell.text_frame:
                            cell.text_frame.word_wrap = True
                            cell.text_frame.margin_left = Inches(0.01)
                            cell.text_frame.margin_right = Inches(0.01)
                            cell.text_frame.margin_top = Inches(0.01)
                            cell.text_frame.margin_bottom = Inches(0.01)
                            for p_idx, p in enumerate(cell.text_frame.paragraphs):
                                p_orig = sanitize_control_chars(p.text)
                                if PPTXProcessor._is_real_text(p_orig) and not PPTXProcessor._is_watermark_text(p_orig):
                                    item_id = f"s{slide_index}_{sh_id_str}_tbl_r{r_idx}_c{c_idx}_p{p_idx}"
                                    trans = translations_map.get(item_id)
                                    if trans:
                                        PPTXProcessor._set_paragraph_text_safe(p, ensure_script(sanitize_control_chars(trans), target_script), auto_fit=auto_fit, shape=shape)
                continue

            # 3. Regular Shape / TextBox
            if shape.has_text_frame:
                tf = shape.text_frame
                tf.word_wrap = True
                tf.margin_left = Inches(0.01)
                tf.margin_right = Inches(0.01)
                tf.margin_top = Inches(0.01)
                tf.margin_bottom = Inches(0.01)
                for p_idx, p in enumerate(tf.paragraphs):
                    p_orig = sanitize_control_chars(p.text)
                    if PPTXProcessor._is_real_text(p_orig) and not PPTXProcessor._is_watermark_text(p_orig):
                        item_id = f"s{slide_index}_{sh_id_str}_p{p_idx}"
                        trans = translations_map.get(item_id)
                        if trans:
                            PPTXProcessor._set_paragraph_text_safe(p, ensure_script(sanitize_control_chars(trans), target_script), auto_fit=auto_fit, shape=shape)

            # 4. Native Charts
            if shape.has_chart:
                try:
                    c_elem = shape.chart._element
                    for t_idx, t_node in enumerate(c_elem.xpath('.//c:title//a:t')):
                        item_id = f"s{slide_index}_{sh_id_str}_chtitle_{t_idx}"
                        trans = translations_map.get(item_id)
                        if trans:
                            t_node.text = ensure_script(sanitize_control_chars(trans), target_script)
                    for cat_idx, v_node in enumerate(c_elem.xpath('.//c:cat//c:pt//c:v')):
                        item_id = f"s{slide_index}_{sh_id_str}_chcat_{cat_idx}"
                        trans = translations_map.get(item_id)
                        if trans:
                            v_node.text = ensure_script(sanitize_control_chars(trans), target_script)
                    for ser_idx, v_node in enumerate(c_elem.xpath('.//c:ser//c:tx//c:v')):
                        item_id = f"s{slide_index}_{sh_id_str}_chser_{ser_idx}"
                        trans = translations_map.get(item_id)
                        if trans:
                            v_node.text = ensure_script(sanitize_control_chars(trans), target_script)
                except Exception:
                    pass

    @staticmethod
    def _set_paragraph_text_safe(paragraph, new_text: str, auto_fit: bool = True, shape = None):
        # Remove any lingering <a:br> elements that cause control character _x000B_ or trailing line jumps
        try:
            for br in list(paragraph._p.xpath('.//a:br')):
                br.getparent().remove(br)
        except Exception:
            pass

        if not paragraph.runs:
            paragraph.text = sanitize_control_chars(new_text)
            return

        orig_text = sanitize_control_chars("".join(r.text for r in paragraph.runs))
        first_run = paragraph.runs[0]

        orig_font_name = first_run.font.name if (first_run.font and first_run.font.name) else None

        clean_val = sanitize_control_chars(new_text)
        safe_text = re.sub(r"([A-Za-zА-Яа-яЎўҒғҚқҲҳ])['\‘\`\ʼ]([A-Za-zА-Яа-яЎўҒғҚқҲҳ])", lambda m: m.group(1) + "ʻ\u2060" + m.group(2), clean_val)

        orig_len = float(len(orig_text))
        new_len = float(len(safe_text))

        box_width_pt = None
        if shape and hasattr(shape, "width") and shape.width:
            try:
                box_width_pt = shape.width.pt
            except Exception:
                pass

        for r in paragraph.runs:
            current_pt = 16.0
            if r.font and r.font.size and r.font.size.pt:
                current_pt = r.font.size.pt
            else:
                if orig_len <= 6:
                    current_pt = 22.0
                elif orig_len <= 15:
                    current_pt = 18.0
                elif orig_len <= 30:
                    current_pt = 14.0
                else:
                    current_pt = 12.0

            try:
                new_pt = current_pt
                
                # A. Yakka so'z sarlavhalar (masalan: MUNDARIJA, BO'LIM, 01)
                if " " not in safe_text.strip():
                    if new_len > orig_len:
                        new_pt = max(10.0, current_pt * (orig_len / new_len) * 1.1)
                    if box_width_pt:
                        max_w_pt = (box_width_pt - 4) / (new_len * 0.58)
                        new_pt = min(new_pt, max(9.0, max_w_pt))
                
                # B. Subtitle / Qo'shimcha sarlavha (masalan: Work report, CONTENTS)
                elif orig_text.lower() in ["work report", "contents", "business plan", "company report"]:
                    new_pt = min(14.0, current_pt)
                
                # C. Katta va uzun sarlavhalar (masalan: Kompaniyamizning so'nggi SWOT tahlili hisoboti)
                elif current_pt >= 24.0:
                    if new_len >= 35:
                        new_pt = max(14.0, current_pt * 0.72)
                    elif new_len >= 20:
                        new_pt = max(16.0, current_pt * 0.82)
                    else:
                        new_pt = max(18.0, current_pt * 0.90)
                
                # D. O'rta sarlavhalar (>= 18pt)
                elif current_pt >= 18.0:
                    ratio = (orig_len / new_len) if new_len > orig_len else 1.0
                    new_pt = max(11.0, current_pt * (ratio ** 0.85))
                
                # E. Uzun izoh va matnlar (new_len >= 80)
                elif new_len >= 80:
                    ratio = (orig_len / new_len) if new_len > orig_len else 1.0
                    new_pt = max(8.5, min(12.0, current_pt * (ratio ** 0.80)))
                
                # F. Standart matnlar
                elif current_pt >= 12.0:
                    ratio = (orig_len / new_len) if new_len > orig_len else 1.0
                    new_pt = max(8.0, current_pt * (ratio ** 0.78))
                else:
                    ratio = (orig_len / new_len) if new_len > orig_len else 1.0
                    new_pt = max(6.5, current_pt * (ratio ** 0.72))
                
                r.font.size = Pt(new_pt)
                if orig_font_name and r.font:
                    r.font.name = orig_font_name
            except Exception:
                pass

        first_run.text = safe_text
        if len(paragraph.runs) > 1:
            for r in paragraph.runs[1:]:
                r.text = ""
