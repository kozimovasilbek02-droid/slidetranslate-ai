# -*- coding: utf-8 -*-
"""
backend.core.transliteration
----------------------------
O'zbek lotin va kirill yozuvlari o'rtasida yuqori aniqlikdagi deterministik transliteratsiya dvigateli.
"""
from __future__ import annotations
import re

APOSTROPHES = {"'", "ʻ", "`", "ʼ", "’", "‘"}
_LATIN_VOWELS = set("aeiou")

_PLAIN_DIGRAPHS = {
    "sh": "ш",
    "ch": "ч",
    "ng": "нг",
    "ts": "ц",
}

_Y_VOWEL_DIGRAPHS = {
    "a": "я",
    "o": "ё",
    "u": "ю",
    "e": "е",
}

_SINGLE_LETTERS = {
    "a": "а", "b": "б", "d": "д", "f": "ф", "g": "г",
    "h": "ҳ", "i": "и", "j": "ж", "k": "к", "l": "л", "m": "м",
    "n": "н", "o": "о", "p": "п", "q": "қ", "r": "р", "s": "с",
    "t": "т", "u": "у", "v": "в", "x": "х", "y": "й", "z": "з",
}

CYRILLIC_TO_LATIN = {
    "ш": "sh", "Ш": "Sh", "ч": "ch", "Ч": "Ch",
    "ё": "yo", "Ё": "Yo", "ю": "yu", "Ю": "Yu", "я": "ya", "Я": "Ya",
    "ў": "oʻ", "Ў": "Oʻ", "ғ": "gʻ", "Ғ": "Gʻ",
    "а": "a", "А": "A", "б": "b", "Б": "B", "в": "v", "В": "V", "г": "g", "Г": "G",
    "д": "d", "Д": "D", "е": "e", "Е": "E", "ж": "j", "Ж": "J", "з": "z", "Z": "Z",
    "и": "i", "И": "I", "й": "y", "Й": "Y", "к": "k", "К": "K", "л": "l", "Л": "L",
    "м": "m", "М": "M", "н": "n", "Н": "N", "о": "o", "О": "O", "п": "p", "П": "P",
    "р": "r", "Р": "R", "с": "s", "С": "S", "т": "t", "Т": "T", "у": "u", "У": "U",
    "ф": "f", "Ф": "F", "х": "x", "Х": "X", "ц": "ts", "Ц": "Ts", "щ": "sh", "Щ": "Sh",
    "ъ": "'", "Ъ": "'", "ь": "", "Ь": "", "э": "e", "Э": "E", "қ": "q", "Қ": "Q",
    "ҳ": "h", "Ҳ": "H"
}

def _match_case(source: str, target: str) -> str:
    """Kirill natijasining katta/kichik harfini lotincha manbaga moslashtiradi."""
    if source.isupper():
        return target.upper()
    if source[:1].isupper():
        return target[0].upper() + target[1:]
    return target

def latin_to_cyrillic(text: str) -> str:
    """
    O'zbek lotin matnini qat'iy fonetik va imlo qoidalari asosida kirill yozuviga aylantiradi:
    1. gʻ / oʻ (har qanday apostrof varianti bilan) - eng yuqori ustuvorlik ('yo'q' -> 'йўқ' bo'lishi uchun).
    2. Yolg'iz tutuq belgisi -> ъ
    3. sh, ch, ng, ts digraflari
    4. ya, yo, yu, ye iotlangan unlilari
    5. 'e' - so'z boshi va unlilardan keyin 'э', aks holda 'е'
    6. Qolgan yagona harflar
    """
    if not text:
        return ""
    result: list[str] = []
    i = 0
    n = len(text)
    at_vowel_or_boundary = True

    while i < n:
        ch = text[i]
        lower = ch.lower()

        # 1) gʻ / oʻ -- eng yuqori ustuvorlik
        if lower in ("g", "o") and i + 1 < n and text[i + 1] in APOSTROPHES:
            cyr = "ғ" if lower == "g" else "ў"
            result.append(_match_case(ch, cyr))
            at_vowel_or_boundary = (lower == "o")
            i += 2
            continue

        # 2) Yolg'iz tutuq belgisi (o'/g' dan tashqari) -> ъ
        if ch in APOSTROPHES:
            result.append("ъ")
            at_vowel_or_boundary = False
            i += 1
            continue

        if not lower.isalpha():
            result.append(ch)
            at_vowel_or_boundary = True
            i += 1
            continue

        two = text[i:i + 2].lower()

        # 3) sh / ch / ng / ts
        if two in _PLAIN_DIGRAPHS:
            result.append(_match_case(text[i:i + 2], _PLAIN_DIGRAPHS[two]))
            at_vowel_or_boundary = False
            i += 2
            continue

        # 4) y + unli (lekin 'yoʻ' holatida 'oʻ' ga xalaqit bermaslik)
        if lower == "y" and i + 1 < n:
            next_lower = text[i + 1].lower()
            if next_lower in _Y_VOWEL_DIGRAPHS:
                is_yo_apostrophe_clash = (
                    next_lower == "o" and i + 2 < n and text[i + 2] in APOSTROPHES
                )
                if not is_yo_apostrophe_clash:
                    result.append(_match_case(text[i:i + 2], _Y_VOWEL_DIGRAPHS[next_lower]))
                    at_vowel_or_boundary = True
                    i += 2
                    continue

        # 5) 'e' -- kontekstga qarab э yoki е
        if lower == "e":
            cyr = "э" if at_vowel_or_boundary else "е"
            result.append(_match_case(ch, cyr))
            at_vowel_or_boundary = True
            i += 1
            continue

        # 6) qolgan yagona harflar
        if lower in _SINGLE_LETTERS:
            result.append(_match_case(ch, _SINGLE_LETTERS[lower]))
            at_vowel_or_boundary = lower in _LATIN_VOWELS
            i += 1
            continue

        result.append(ch)
        at_vowel_or_boundary = True
        i += 1

    return "".join(result)

def cyrillic_to_latin(text: str) -> str:
    """Kirill matnini lotin yozuviga o'giradi."""
    if not text:
        return ""
    result = text
    for k, v in CYRILLIC_TO_LATIN.items():
        result = result.replace(k, v)
    return result

def ensure_script(text: str, target_script: str = "latin") -> str:
    """Matnni ko'rsatilgan maqsadli alifboga (latin yoki cyrillic) o'giradi."""
    if not text:
        return ""
    is_cyrillic_target = target_script.lower() in ["cyrillic", "kirill", "uz-cyrl", "ўзбекча"]
    has_cyrillic = bool(re.search(r"[а-яА-ЯёЁўЎғҒқҚҳҲ]", text))
    if is_cyrillic_target and not has_cyrillic:
        return latin_to_cyrillic(text)
    elif not is_cyrillic_target and has_cyrillic:
        return cyrillic_to_latin(text)
    return text
