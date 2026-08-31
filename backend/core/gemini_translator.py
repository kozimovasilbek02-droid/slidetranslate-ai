# -*- coding: utf-8 -*-
import os
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Optional
from google import genai
from google.genai import types

from backend.core.transliteration import ensure_script

def sanitize_text(text: str) -> str:
    if not text:
        return ""
    t = re.sub(r'_x[0-9a-fA-F]{4}_', ' ', text)
    t = re.sub(r'[\u0001-\u0008\u000b\u000c\u000e-\u001f\u007f]', ' ', t)
    t = re.sub(r'[ \t]+', ' ', t)
    return t.strip()

class GeminiTranslator:
    def __init__(self, api_key: Optional[str] = None, model_name: Optional[str] = None):
        self.api_key = api_key or os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
        self.client = genai.Client(api_key=self.api_key) if self.api_key else genai.Client()
        self.model_candidates = [
            "gemini-3.6-flash",
            "gemini-3.5-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.1-pro-preview"
        ]
        self.model_name = model_name or os.environ.get("GEMINI_MODEL") or self.model_candidates[0]

    def translate_items_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None
    ) -> List[Dict[str, Any]]:
        if not items:
            return []

        sub_batch_size = 35
        batches = [items[i:i + sub_batch_size] for i in range(0, len(items), sub_batch_size)]
        
        if len(batches) == 1:
            return self._translate_single_batch(batches[0], target_script, domain, glossary)
        
        results = []
        with ThreadPoolExecutor(max_workers=min(4, len(batches))) as executor:
            future_to_batch = {
                executor.submit(self._translate_single_batch, b, target_script, domain, glossary): b
                for b in batches
            }
            for future in as_completed(future_to_batch):
                try:
                    res = future.result()
                    results.extend(res)
                except Exception:
                    b = future_to_batch[future]
                    results.extend([{"id": it["id"], "translated_text": sanitize_text(it["original_text"])} for it in b])

        order_map = {it["id"]: it for it in results}
        return [order_map.get(it["id"], {"id": it["id"], "translated_text": sanitize_text(it["original_text"])}) for it in items]

    def _translate_single_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None
    ) -> List[Dict[str, Any]]:
        script_name = "O'zbek tili (Lotin yozuvi)" if target_script.lower() == "latin" else "Ўзбек тили (Кирилл ёзуви)"
        
        glossary_instructions = ""
        if glossary:
            terms = ", ".join([f"{k} => {v}" for k, v in glossary.items()])
            glossary_instructions = f"Maxsus lug'at atamalari: {terms}"

        system_instruction = f"""Siz professional PowerPoint taqdimotlari tarjimonisiz.
Vazifangiz taqdimot matnlarini {script_name}ga professional, ravon va slayd ramkalariga sig'adigan darajada IXCHAM tarjima qilishdir.

QAT'IY QOIDALAR:
1. DIAGRAMMA VA SHAKL SARLAVHALARI (O'TA IXCHAM BO'LSIN):
   - '添加标题文本' / '添加标题' / '在此添加标题' => 'Sarlavha' (yoki 'Mavzu') - hech qachon 'Sarlavha matnini qo'shing' deb 3 qatorga cho'zmang!
   - '根据自己的需要添加适当的文字...' => 'Bu yerga qisqacha tavsif matni kiritiladi.'
   - '目录' / 'CONTENTS' => 'Mundarija'
   - 'Work report' / '汇报完毕' => 'Ish hisoboti' / 'E\'tiboringiz uchun rahmat'
2. LOREM IPSUM VA SHABLON MATNLAR:
   - 'Lorem Ipsum' => '1-bosqich' (yoki 'Namuna sarlavhasi')
   - 'Lorem ipsum dolor sit amet...' => 'Bu yerga loyihangizning qisqacha tavsifi, asosiy vazifalar yoki maqsadlar yoziladi.'
3. IXCHAMLIK (ANTI-OVERFLOW): Matnlar slayd shakllaridan toshib ketmasligi uchun cho'zilgan so'zlardan qoching.
4. Boshqaruv belgilari (_x000B_, \\v, \\r) va ortiqcha probellarni tozalang.
5. Kiruvchi JSON massividagi har bir element uchun 'id' va 'translated' kalitlari bilan JSON massiv qaytaring.
{glossary_instructions}
"""

        prompt_payload = [{"id": it["id"], "text": sanitize_text(it["original_text"])} for it in items]
        user_content = json.dumps(prompt_payload, ensure_ascii=False, indent=2)

        for attempt in range(len(self.model_candidates)):
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
                result_map = {}
                if isinstance(parsed, list):
                    for row in parsed:
                        if isinstance(row, dict) and "id" in row:
                            t_val = row.get("translated") or row.get("translated_text") or row.get("text") or row.get("uzbek") or ""
                            result_map[row["id"]] = ensure_script(sanitize_text(t_val), target_script)

                return [{"id": it["id"], "translated_text": result_map.get(it["id"], sanitize_text(it["original_text"]))} for it in items]

            except Exception as e:
                time.sleep(1.0 * (attempt + 1))
                if attempt == len(self.model_candidates) - 1:
                    return [{"id": it["id"], "translated_text": sanitize_text(it["original_text"])} for it in items]

    def translate_single_text(self, text: str, target_script: str = "latin") -> str:
        if not text or not text.strip():
            return "Taqdimot"
        script_name = "O'zbek tili (Lotin yozuvi)" if target_script.lower() == "latin" else "Ўзбек тили (Кирилл ёзуви)"
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
                    return out
            except Exception:
                continue
        return clean_in
