# -*- coding: utf-8 -*-
import re

LATIN_TO_CYRILLIC = {
    "sh": "ш", "Sh": "Ш", "SH": "Ш",
    "ch": "ч", "Ch": "Ч", "CH": "Ч",
    "yo": "ё", "Yo": "Ё", "YO": "Ё",
    "yu": "ю", "Yu": "Ю", "YU": "Ю",
    "ya": "я", "Ya": "Я", "YA": "Я",
    "ye": "е", "Ye": "Е", "YE": "Е",
    "oʻ": "ў", "Oʻ": "Ў", "o‘": "ў", "O‘": "Ў", "o'": "ў", "O'": "Ў", "o`": "ў", "O`": "Ў",
    "gʻ": "ғ", "Gʻ": "Ғ", "g‘": "ғ", "G‘": "Ғ", "g'": "ғ", "G'": "Ғ", "g`": "ғ", "G`": "Ғ",
    "a": "а", "A": "А", "b": "б", "B": "Б", "d": "д", "D": "Д", "e": "е", "E": "Е",
    "f": "ф", "F": "Ф", "g": "г", "G": "Г", "h": "ҳ", "H": "Ҳ", "i": "и", "I": "И",
    "j": "ж", "J": "Ж", "k": "к", "K": "К", "l": "л", "L": "Л", "m": "м", "M": "М",
    "n": "н", "N": "Н", "o": "о", "O": "О", "p": "п", "P": "П", "q": "қ", "Q": "Қ",
    "r": "р", "R": "Р", "s": "с", "S": "С", "t": "т", "T": "Т", "u": "у", "U": "У",
    "v": "в", "V": "В", "x": "х", "X": "Х", "y": "й", "Y": "Й", "z": "з", "Z": "З",
    "'": "ъ", "’": "ъ", "`": "ъ", "ʼ": "ъ"
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

def latin_to_cyrillic(text: str) -> str:
    if not text:
        return ""
    result = text
    two_letter = ["sh", "Sh", "SH", "ch", "Ch", "CH", "yo", "Yo", "YO", "yu", "Yu", "YU", "ya", "Ya", "YA", "ye", "Ye", "YE", "oʻ", "Oʻ", "o‘", "O‘", "o'", "O'", "o`", "O`", "gʻ", "Gʻ", "g‘", "G‘", "g'", "G'", "g`", "G`"]
    for k in two_letter:
        if k in result:
            result = result.replace(k, LATIN_TO_CYRILLIC[k])
    for k, v in LATIN_TO_CYRILLIC.items():
        if len(k) == 1:
            result = result.replace(k, v)
    return result

def cyrillic_to_latin(text: str) -> str:
    if not text:
        return ""
    result = text
    for k, v in CYRILLIC_TO_LATIN.items():
        result = result.replace(k, v)
    return result

def ensure_script(text: str, target_script: str = "latin") -> str:
    if not text:
        return ""
    is_cyrillic_target = target_script.lower() in ["cyrillic", "kirill", "ўзбекча"]
    has_cyrillic = bool(re.search(r"[а-яА-ЯёЁўЎғҒқҚҳҲ]", text))
    if is_cyrillic_target and not has_cyrillic:
        return latin_to_cyrillic(text)
    elif not is_cyrillic_target and has_cyrillic:
        return cyrillic_to_latin(text)
    return text
