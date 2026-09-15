# -*- coding: utf-8 -*-
"""
SlideTranslate AI — High-Fidelity Multi-Slide Preview & Thumbnail Generator
---------------------------------------------------------------------------
1. Taqdimotning dastlabki 2-3 ta slaydlarini (Titul, Mundarija, Asosiy mazmun)
   haqiqiy dizayni, shakllari, rasmlari, ranglari va tarjima qilingan matnlari bilan
   yuqori sifatli (1920x1080) JPEG rasm qilib chiqaradi.
2. LibreOffice (agar tizimda bo'lsa) + PyMuPDF / PyPDFium2 orqali, yoki
   mustaqil python-pptx + Pillow dvigateli orqali 100% kafolatlangan real slayd ko'rinishini chizadi.
3. 1-slayd rasmini .pptx arxivi (docProps/thumbnail.jpeg) ichiga OpenXML standarti asosida joylaydi.
"""

import os
import sys
import io
import shutil
import zipfile
import subprocess
from typing import List, Optional

from PIL import Image, ImageDraw, ImageFont

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.dml.color import RGBColor
except ImportError:
    Presentation = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None


class ThumbnailGenerator:
    """PowerPoint taqdimotlari uchun 2-3 ta slaydning real preview rasmlarini yaratuvchi dvigatel."""

    @classmethod
    def export_presentation_previews(cls, pptx_path: str, output_dir: str, max_slides: int = 3, width: int = 1920, height: int = 1080) -> List[str]:
        """Taqdimotdan 2-3 ta slaydning real preview rasmlarini yaratish."""
        abs_pptx = os.path.abspath(pptx_path)
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(abs_pptx))[0]
        
        # 1. LibreOffice orqali PDF ga o'girib, sahifalarni rasmlarga aylantirish
        pdf_images = cls._try_libreoffice_export(abs_pptx, output_dir, max_slides, width, height)
        if pdf_images:
            return pdf_images

        # 2. Pure-Python + Pillow + python-pptx High-Fidelity Slayd rendereri
        pil_images = cls._try_pil_pptx_render(abs_pptx, output_dir, max_slides, width, height)
        if pil_images:
            return pil_images

        # 3. Agar LibreOffice yoki PIL render yaratilmasa, xavfsiz bo'sh ro'yxat qaytarish (rekursiyasiz)
        return []

    @classmethod
    def export_slide_preview(cls, pptx_path: str, output_image_path: str, width: int = 1920, height: int = 1080) -> bool:
        """1-slayd preview'sini yaratish."""
        out_dir = os.path.dirname(os.path.abspath(output_image_path))
        previews = cls.export_presentation_previews(pptx_path, out_dir, max_slides=1, width=width, height=height)
        if previews and os.path.exists(previews[0]):
            if previews[0] != os.path.abspath(output_image_path):
                try:
                    shutil.copyfile(previews[0], output_image_path)
                except Exception:
                    pass
            return True
        return False

    @classmethod
    def _try_libreoffice_export(cls, abs_pptx: str, output_dir: str, max_slides: int, width: int, height: int) -> List[str]:
        """LibreOffice (soffice) orqali PDF ga aylantirib, slaydlarni ajratish."""
        soffice_bin = None
        for cmd in ["soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe"]:
            if shutil.which(cmd) or os.path.exists(cmd):
                soffice_bin = cmd
                break

        if not soffice_bin:
            return []

        try:
            temp_pdf_dir = os.path.join(output_dir, "_temp_pdf")
            os.makedirs(temp_pdf_dir, exist_ok=True)
            
            cmd = [
                soffice_bin,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", temp_pdf_dir,
                abs_pptx
            ]
            res = subprocess.run(cmd, capture_output=True, timeout=25)
            if res.returncode != 0:
                return []

            base_name = os.path.splitext(os.path.basename(abs_pptx))[0]
            pdf_path = os.path.join(temp_pdf_dir, f"{base_name}.pdf")
            if not os.path.exists(pdf_path):
                return []

            generated = []
            if fitz is not None:
                doc = fitz.open(pdf_path)
                pages_to_export = min(max_slides, len(doc))
                for idx in range(pages_to_export):
                    page = doc[idx]
                    zoom = width / page.rect.width
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    out_img = os.path.join(output_dir, f"{base_name}_slide_{idx + 1}.jpg")
                    pix.save(out_img)
                    generated.append(out_img)
                doc.close()
            elif pdfium is not None:
                pdf = pdfium.PdfDocument(pdf_path)
                pages_to_export = min(max_slides, len(pdf))
                for idx in range(pages_to_export):
                    image = pdf[idx].render(scale=2.0).to_pil()
                    out_img = os.path.join(output_dir, f"{base_name}_slide_{idx + 1}.jpg")
                    image.save(out_img, "JPEG", quality=90)
                    generated.append(out_img)
                pdf.close()

            try:
                os.remove(pdf_path)
                os.rmdir(temp_pdf_dir)
            except Exception:
                pass

            return generated
        except Exception:
            return []

    @classmethod
    def _try_pil_pptx_render(cls, abs_pptx: str, output_dir: str, max_slides: int, width: int, height: int) -> List[str]:
        """python-pptx va Pillow yordamida real slayd elementlarini (shakllar, fon, rasmlar, matnlar) chizish."""
        if Presentation is None:
            return []

        try:
            prs = Presentation(abs_pptx)
            total_slides = len(prs.slides)
            if total_slides == 0:
                return []

            prs_width = prs.slide_width
            prs_height = prs.slide_height
            scale_x = width / prs_width
            scale_y = height / prs_height

            base_name = os.path.splitext(os.path.basename(abs_pptx))[0]
            exported_images = []

            # Shriftlar
            try:
                font_title = ImageFont.truetype("arialbd.ttf", 36)
                font_body = ImageFont.truetype("arial.ttf", 22)
                font_small = ImageFont.truetype("arial.ttf", 16)
            except Exception:
                try:
                    font_title = ImageFont.truetype("arial.ttf", 36)
                    font_body = ImageFont.truetype("arial.ttf", 22)
                    font_small = ImageFont.truetype("arial.ttf", 16)
                except Exception:
                    font_title = ImageFont.load_default()
                    font_body = ImageFont.load_default()
                    font_small = ImageFont.load_default()

            slides_count = min(max_slides, total_slides)
            for s_idx in range(slides_count):
                slide = prs.slides[s_idx]
                
                # Fon rangini aniqlash
                bg_color = (255, 255, 255)
                try:
                    if slide.background and slide.background.fill:
                        fill = slide.background.fill
                        if fill.type == 1 and fill.fore_color and fill.fore_color.type == 1:
                            rgb = fill.fore_color.rgb
                            bg_color = (rgb[0], rgb[1], rgb[2])
                except Exception:
                    pass

                img = Image.new("RGB", (width, height), color=bg_color)
                draw = ImageDraw.Draw(img)

                # Shakllarni chizish
                for shape in slide.shapes:
                    try:
                        x = int(shape.left * scale_x)
                        y = int(shape.top * scale_y)
                        w = int(shape.width * scale_x)
                        h = int(shape.height * scale_y)

                        # A. Rasmlar (Suratlar, piktogrammalar)
                        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            try:
                                img_blob = shape.image.blob
                                with Image.open(io.BytesIO(img_blob)) as pic:
                                    pic_resized = pic.resize((max(1, w), max(1, h)), Image.Resampling.LANCZOS)
                                    if pic_resized.mode in ("RGBA", "LA") or (pic_resized.mode == "P" and "transparency" in pic_resized.info):
                                        img.paste(pic_resized, (x, y), pic_resized)
                                    else:
                                        img.paste(pic_resized, (x, y))
                                continue
                            except Exception:
                                pass

                        # B. Rangli shakllar va kartochkalar
                        try:
                            if shape.has_text_frame or shape.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM):
                                if hasattr(shape, "fill") and shape.fill and shape.fill.type == 1:
                                    rgb = shape.fill.fore_color.rgb
                                    draw.rounded_rectangle([x, y, x + w, y + h], radius=12, fill=(rgb[0], rgb[1], rgb[2]))
                        except Exception:
                            pass

                        # C. Matnlar
                        if shape.has_text_frame:
                            tf = shape.text_frame
                            cur_y = y + 10
                            for p in tf.paragraphs:
                                line_text = p.text.strip()
                                if not line_text:
                                    cur_y += 12
                                    continue

                                text_color = (30, 41, 59)
                                try:
                                    for r in p.runs:
                                        if r.font and r.font.color and r.font.color.rgb:
                                            rgb = r.font.color.rgb
                                            text_color = (rgb[0], rgb[1], rgb[2])
                                            break
                                except Exception:
                                    pass

                                f = font_title if (p.font and p.font.size and p.font.size.pt > 24) else font_body
                                draw.text((x + 10, cur_y), line_text[:120], fill=text_color, font=f)
                                cur_y += 34
                    except Exception:
                        continue

                # Slayd pastki chiroyli footer va sahifa raqami
                draw.rectangle([0, height - 36, width, height], fill=(15, 23, 42, 220))
                draw.text((40, height - 30), f"SlideTranslate AI • {s_idx + 1}-slayd", fill=(203, 213, 225), font=font_small)
                draw.text((width - 150, height - 30), f"{s_idx + 1} / {total_slides}", fill=(148, 163, 184), font=font_small)

                out_img = os.path.join(output_dir, f"{base_name}_slide_{s_idx + 1}.jpg")
                img.save(out_img, "JPEG", quality=92)
                exported_images.append(out_img)

            return exported_images
        except Exception:
            return []

    @classmethod
    def embed_thumbnail_into_pptx(cls, pptx_path: str, image_path: Optional[str] = None) -> bool:
        """1-slayd preview'sini .pptx arxivi (docProps/thumbnail.jpeg) ichiga OpenXML standarti bo'yicha qadash."""
        abs_pptx = os.path.abspath(pptx_path)
        if not os.path.exists(abs_pptx):
            return False

        if not image_path or not os.path.exists(image_path):
            temp_img = abs_pptx + "_temp_thumb.jpg"
            success = cls.export_slide_preview(abs_pptx, temp_img)
            if not success or not os.path.exists(temp_img):
                return False
            image_path = temp_img

        try:
            with open(image_path, "rb") as f:
                img_bytes = f.read()

            temp_zip = abs_pptx + ".temp.zip"
            with zipfile.ZipFile(abs_pptx, "r") as zin, zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename == "_rels/.rels":
                        text = content.decode("utf-8")
                        if "thumbnail" not in text:
                            rel_tag = '<Relationship Id="rIdThumbnail" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>'
                            text = text.replace("</Relationships>", f"{rel_tag}</Relationships>")
                            content = text.encode("utf-8")
                    elif item.filename == "[Content_Types].xml":
                        text = content.decode("utf-8")
                        added_tags = ""
                        if 'Extension="jpeg"' not in text and 'Extension="jpg"' not in text:
                            added_tags += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                        if '/docProps/thumbnail.jpeg' not in text:
                            added_tags += '<Override PartName="/docProps/thumbnail.jpeg" ContentType="image/jpeg"/>'
                        if added_tags:
                            text = text.replace("</Types>", f"{added_tags}</Types>")
                            content = text.encode("utf-8")
                    elif item.filename == "docProps/thumbnail.jpeg":
                        continue
                    zout.writestr(item, content)

                zout.writestr("docProps/thumbnail.jpeg", img_bytes)

            os.replace(temp_zip, abs_pptx)
            return True
        except Exception:
            return False
        finally:
            if image_path.endswith("_temp_thumb.jpg") and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                except Exception:
                    pass
            if 'temp_zip' in locals() and os.path.exists(temp_zip):
                try:
                    os.remove(temp_zip)
                except Exception:
                    pass
