# -*- coding: utf-8 -*-
"""
backend.core.gemini_translator
------------------------------
Google Gemini orqali slayd matnlarini yuqori aniqlikda, sohaviy kontekst,
anti-overflow va parallel batch tarzda tarjima qiluvchi modul.
"""
from __future__ import annotations

import os
import json
import re
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Optional

from google import genai
from google.genai import types

from backend.core.transliteration import ensure_script, latin_to_cyrillic

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gemini-3.1-flash-lite"
FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
    "gemini-3.6-flash",
]

def sanitize_text(text: str) -> str:
    """Boshqaruv belgilari va OpenXML artefaktlarini tozalaydi, probellarni saqlaydi."""
    if not text:
        return ""
    # _x000B_ va boshqaruv belgilarini probel bilan almashtiramiz (so'zlar yopishib qolmasligi uchun)
    t = re.sub(r'_x[0-9a-fA-F]{4}_', ' ', text)
    t = re.sub(r'[\u0001-\u0008\u000b\u000c\u000e-\u001f\u007f]', ' ', t)
    t = re.sub(r'[ \t]+', ' ', t)
    return t.strip()


class TranslationError(RuntimeError):
    """Tarjima xizmati xatolik qaytarganda."""


class GeminiTranslator:
    """
    Taqdimot matnlarini kontekst, terminologiya va slayd tuzilmasini saqlagan holda
    tarjima qiluvchi yuqori aniqlikdagi AI tarjimon dvigateli.
    """

    def __init__(self, api_key: Optional[str] = None, model_name: Optional[str] = None):
        self.api_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY", "")
        )
        self.client = genai.Client(api_key=self.api_key) if self.api_key else genai.Client()
        self.model_candidates = [model_name] if model_name else list(FALLBACK_MODELS)
        self.model_name = self.model_candidates[0]

    def translate_items_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None,
        stats: Optional[Dict[str, Any]] = None,
        batch_size: int = 35,
        max_workers: int = 4,
    ) -> List[Dict[str, Any]]:
        """
        Slayd elementlari ro'yxatini parallel batch tarzda tarjima qiladi.
        items: [{"id": "s1_sh0_p0", "text": "Strategic Vision", ...}, ...]
        """
        if not items:
            if stats is not None:
                stats.update(total=0, failed=0, failed_ids=[])
            return []

        batches = [items[i:i + batch_size] for i in range(0, len(items), batch_size)]

        if len(batches) == 1:
            results = self._translate_single_batch(batches[0], target_script, domain, glossary)
        else:
            results = []
            workers = min(max_workers, len(batches))
            with ThreadPoolExecutor(max_workers=workers) as executor:
                future_to_batch = {
                    executor.submit(self._translate_single_batch, b, target_script, domain, glossary): b
                    for b in batches
                }
                for future in as_completed(future_to_batch):
                    b = future_to_batch[future]
                    try:
                        res = future.result()
                        results.extend(res)
                    except Exception as exc:
                        logger.error("Batch tarjima xatosi: %s", exc)
                        for it in b:
                            orig = it.get("original_text") or it.get("text", "")
                            results.append({"id": it["id"], "translated_text": sanitize_text(orig)})

        order_map = {it["id"]: it for it in results}
        ordered_results = []
        failed_ids = []

        for it in items:
            orig = sanitize_text(it.get("original_text") or it.get("text", ""))
            matched = order_map.get(it["id"])
            if matched:
                tr_text = matched.get("translated_text", "")
                if not tr_text or tr_text.strip() == orig.strip():
                    # Matn tarjima qilinmagan bo'lishi mumkin (masalan qisqa nomlar bundan mustasno)
                    pass
                ordered_results.append(matched)
            else:
                ordered_results.append({"id": it["id"], "translated_text": orig})
                failed_ids.append(it["id"])

        if stats is not None:
            stats.update(total=len(items), failed=len(failed_ids), failed_ids=failed_ids)

        return ordered_results

    def _translate_single_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None,
        max_retries: int = 5,
    ) -> List[Dict[str, Any]]:
        is_cyrillic = target_script.lower() in ["cyrillic", "kirill", "uz-cyrl", "ўзбекча"]
        script_name = "O'zbek tili (Lotin yozuvi)" if not is_cyrillic else "Ўзбек тили (Кирилл ёзуви)"

        glossary_instructions = ""
        if glossary:
            terms = "; ".join([f"'{k}' => '{v}'" for k, v in glossary.items()])
            glossary_instructions = f"\nMUHIM QOIDA - Maxsus atamalar lug'atiga qat'iy amal qiling:\n{terms}\n"

        system_instruction = f"""Siz professional xalqaro PowerPoint taqdimotlari bo'yicha ekspert AI tarjimonsiz.
Vazifangiz berilgan slayd matnlarini {script_name}ga professional, ravon va slayd ramkalariga sig'adigan darajada IXCHAM tarjima qilishdir.

Soha / Kontekst: {domain}
{glossary_instructions}

QAT'IY QOIDALAR:
1. HAR BIR INGLIZCHA MATNNI O'ZBEK TILIGA O'GIRING:
   - "Renaissance Genius" => "Uyg'onish davri dahosi"
   - "The Last Supper" => "So'nggi kecha"
   - "Mona Lisa" => "Mona Liza"
   - "Codex Leicester" => "Lester kodeksi"
   - "Passing on the Torch" => "Merosni davom ettirish" (ma'nosiga qarab)
   - "Vitruvian Man" => "Vitruviy odami"
   - "Executive Summary" => "Rahbarlik uchun xulosa"
   - "Contents" / "Table of contents" => "Mundarija"
   - "Work Report" => "Ish hisoboti"
2. IXCHAMLIK (ANTI-OVERFLOW):
   - Slayd bloklaridan toshib ketmasligi uchun cho'zilgan jumlalardan qoching.
   - Sarlavhalarni lo'nda va ixcham saqlang.
3. SONLAR, FOIZLAR, FORMULALAR:
   - Raqamlar, yillar (masalan: 1452-1519), foizlar va maxsus belgilarni o'zgartirmang.
4. TOZALASH:
   - Boshqaruv belgilarini (_x000B_, \\v, \\r) va ortiqcha probellarni tozalang.
5. JAVOB FORMATI:
   - Kiruvchi JSON massividagi har bir element uchun 'id' va 'translated' kalitlari bilan JSON massiv qaytaring.
6. SHABLON SARLAVHALARINI TO'G'RI TARJIMA QILISH:
   - "Agenda Style" => "Kun tartibi" ("uslubi" so'zini qo'shmang)
   - "Our Team Style" / "Team Style" => "Bizning jamoa"
   - "Infographic Style" => "Infografika"
   - "Portfolio Style" => "Portfolio"
   - "Real Estate" => "Ko'chmas mulk"
   - "Content Here" / "Contents Title" => "Mundarija"
   - Shablonlardagi "Style", "Layout" kabi sun'iy so'zlarni sarlavhaga qo'shmang, lo'nda va tabiiy nomlang.
7. LOREM IPSUM VA SHABLON MATNLARINI TO'LIQ O'ZBEKCHALASHTIRING:
   - "Lorem ipsum dolor sit amet..." kabi har qanday soxta lotincha/inglizcha matnlarni hech qachon shundayligicha qoldirmang!
   - Qisqa sarlavhalar uchun: "Mavzu bo'yicha qisqacha izoh" yoki "Taqdimotning qisqacha mazmuni";
   - Uzun matnlar uchun: "Ushbu bo'limda taqdimot mavzusi yuzasidan batafsil ma'lumotlar, asosiy ko'rsatkichlar va tahliliy xulosalar keltiriladi."
"""

        prompt_payload = [
            {"id": it["id"], "text": sanitize_text(it.get("original_text") or it.get("text", ""))}
            for it in items
        ]
        user_content = json.dumps(prompt_payload, ensure_ascii=False, indent=2)

        result_map: Dict[str, str] = {}
        last_exception = None

        for attempt in range(max_retries):
            cur_model = self.model_candidates[attempt % len(self.model_candidates)]
            try:
                response = self.client.models.generate_content(
                    model=cur_model,
                    contents=user_content,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        temperature=0.2,
                        response_mime_type="application/json"
                    )
                )

                raw_text = response.text.strip()
                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                raw_text = re.sub(r"\s*```$", "", raw_text)

                parsed = json.loads(raw_text)
                if isinstance(parsed, list):
                    for row in parsed:
                        if isinstance(row, dict) and "id" in row:
                            t_val = row.get("translated") or row.get("translated_text") or row.get("text") or ""
                            clean_t = sanitize_text(t_val)
                            # Kirill kerak bo'lsa transliteratsiya bilan mustahkamlash
                            if is_cyrillic:
                                clean_t = ensure_script(clean_t, "cyrillic")
                            result_map[row["id"]] = clean_t
                if result_map:
                    break
            except Exception as e:
                last_exception = e
                wait_sec = 1.5 * (2 ** attempt)
                logger.warning("Gemini chaqiruvi muvaffaqiyatsiz (urinish %d/%d): %s. Kutish: %.1fs",
                               attempt + 1, max_retries, e, wait_sec)
                time.sleep(wait_sec)

        # Muvaffaqiyatsiz bo'lgan elementlar uchun zaxira (fallback)
        results = []
        for it in items:
            item_id = it["id"]
            orig = sanitize_text(it.get("original_text") or it.get("text", ""))
            if item_id in result_map and result_map[item_id].strip():
                results.append({"id": item_id, "translated_text": result_map[item_id]})
            else:
                # Zaxira: deep_translator yoki asl matn
                tr = ""
                try:
                    from deep_translator import GoogleTranslator
                    gt = GoogleTranslator(source="auto", target="uz")
                    tr = gt.translate(orig) if orig else ""
                except Exception:
                    pass
                val = tr or orig
                if is_cyrillic:
                    val = ensure_script(val, "cyrillic")
                results.append({"id": item_id, "translated_text": sanitize_text(val)})

        return results

    def translate_single_text(self, text: str, target_script: str = "latin") -> str:
        """Yagona satrni (masalan fayl sarlavhasini) toza o'zbek tiliga o'giradi."""
        if not text or not text.strip():
            return "Taqdimot"
        is_cyrillic = target_script.lower() in ["cyrillic", "kirill", "uz-cyrl", "ўзбекча"]
        script_name = "O'zbek tili (Lotin yozuvi)" if not is_cyrillic else "Ўзбек тили (Кирилл ёзуви)"
        clean_in = sanitize_text(text)

        prompt = f"""Fayl sarlavhasini {script_name}ga qisqa, toza va professional tarjima qiling.
QOIDALAR:
1. FAQAT tarjima qilingan nomni qaytaring.
2. Oxiriga '_Tarjima', '_Ozbekcha' yoki qavslar qo'shmang.
3. '_' (pastki chiziq) ishlatmang, so'zlar orasida probel bo'lsin.

Matn: "{clean_in}\""""

        for model in self.model_candidates:
            try:
                res = self.client.models.generate_content(
                    model=model,
                    contents=prompt
                )
                out = res.text.strip().strip('\'"').strip()
                out = re.sub(r'(\s*[-_]?\s*(tarjima|ozbekcha|ўзбекча)[a-z]*)$', '', out, flags=re.IGNORECASE)
                out = re.sub(r'[/\\:*?"<>|_]', ' ', out)
                out = re.sub(r'\s+', ' ', out).strip()
                if out:
                    return ensure_script(out, target_script) if is_cyrillic else out
            except Exception:
                continue

        # Zaxira (fallback)
        try:
            from deep_translator import GoogleTranslator
            gt = GoogleTranslator(source='auto', target='uz')
            tr = gt.translate(clean_in)
            out = re.sub(r'(\s*[-_]?\s*(tarjima|ozbekcha|ўзбекча)[a-z]*)$', '', tr, flags=re.IGNORECASE)
            out = re.sub(r'[/\\:*?"<>|_]', ' ', out)
            out = re.sub(r'\s+', ' ', out).strip()
            return ensure_script(out, target_script) if (out and is_cyrillic) else (out or clean_in)
        except Exception:
            return clean_in
