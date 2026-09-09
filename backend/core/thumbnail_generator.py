# -*- coding: utf-8 -*-
"""
SlideTranslate AI — Thumbnail & Preview Generator
--------------------------------------------------
Har bir tarjima qilingan taqdimot uchun:
1. 1-slaydning yuqori sifatli (Ultra HD / 1920x1080) rasm preview'sini (.jpg) yaratadi.
2. Slayd preview rasmini to'g'ridan-to'g'ri .pptx arxividagi 'docProps/thumbnail.jpeg'
   ichiga OpenXML standartlari bo'yicha integratsiya qiladi.
   Bu orqali Windows, macOS hamda Soff.uz va boshqa platformalar fayl preview'sini
   darhol avtomatik ko'rsatadi!
"""

import os
import sys
import zipfile
import shutil
from typing import Optional

try:
    import win32com.client
except ImportError:
    win32com = None

from PIL import Image, ImageDraw, ImageFont


class ThumbnailGenerator:
    """PowerPoint taqdimotlari uchun preview rasm yaratish va PPTX ga o'rnatish vositasi."""

    @staticmethod
    def export_slide_preview(pptx_path: str, output_image_path: str, width: int = 1920, height: int = 1080) -> bool:
        """PowerPoint COM orqali 1-slaydni yuqori sifatli JPG rasm qilib saqlash."""
        abs_pptx = os.path.abspath(pptx_path)
        abs_img = os.path.abspath(output_image_path)
        os.makedirs(os.path.dirname(abs_img), exist_ok=True)

        # 1. PowerPoint COM orqali urinish
        if win32com is not None:
            try:
                ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
                # ReadOnly=1 (True), Untitled=0 (False), WithWindow=0 (False)
                presentation = ppt_app.Presentations.Open(abs_pptx, 1, 0, 0)
                if presentation.Slides.Count > 0:
                    presentation.Slides(1).Export(abs_img, "JPG", width, height)
                presentation.Close()
                if os.path.exists(abs_img) and os.path.getsize(abs_img) > 1000:
                    return True
            except Exception as e:
                print(f"[ThumbnailGenerator] PowerPoint COM xatoligi: {e}")

        # 2. Agar PowerPoint COM ishlamasa, fallback PIL orqali chiroyli muqova yaratish
        try:
            img = Image.new("RGB", (width, height), color=(245, 247, 250))
            draw = ImageDraw.Draw(img)
            
            # Sarlavha chizish
            title = os.path.splitext(os.path.basename(pptx_path))[0].replace("_", " ")
            draw.rectangle([(60, 60), (width - 60, height - 60)], outline=(30, 64, 175), width=6)
            draw.text((120, 200), title[:60], fill=(15, 23, 42))
            draw.text((120, 300), "SlideTranslate AI — O'zbekcha Taqdimot", fill=(100, 116, 139))
            
            img.save(abs_img, "JPEG", quality=95)
            return True
        except Exception:
            return False

    @staticmethod
    def embed_thumbnail_into_pptx(pptx_path: str, image_path: Optional[str] = None) -> bool:
        """
        Preview rasmini .pptx arxivi ichiga (docProps/thumbnail.jpeg) OpenXML standarti
        asosida to'liq joylash.
        """
        abs_pptx = os.path.abspath(pptx_path)
        if not os.path.exists(abs_pptx):
            return False

        # Agar rasm berilmagan bo'lsa, avtomatik yaratish
        if not image_path or not os.path.exists(image_path):
            temp_img = abs_pptx + "_preview.jpg"
            success = ThumbnailGenerator.export_slide_preview(abs_pptx, temp_img)
            if not success or not os.path.exists(temp_img):
                return False
            image_path = temp_img

        try:
            with open(image_path, "rb") as f:
                img_bytes = f.read()

            temp_zip_path = abs_pptx + ".temp.zip"
            with zipfile.ZipFile(abs_pptx, "r") as zin, zipfile.ZipFile(temp_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
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

            os.replace(temp_zip_path, abs_pptx)
            return True
        except Exception as e:
            print(f"[ThumbnailGenerator] Thumbnail embed xatoligi: {e}")
            return False
        finally:
            # Temp image ni tozalash (agar alohida temp bo'lsa)
            if image_path.endswith("_preview.jpg") and os.path.exists(image_path):
                try:
                    os.remove(image_path)
                except Exception:
                    pass
