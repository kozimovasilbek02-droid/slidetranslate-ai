# -*- coding: utf-8 -*-
"""
tests/test_transliteration.py
O'zbek lotin va kirill transliteratsiya qoidalarini tekshiruvchi testlar to'plami.
"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from backend.core.transliteration import latin_to_cyrillic, cyrillic_to_latin, ensure_script

class TestSingleLetters:
    def test_simple_words(self):
        assert latin_to_cyrillic("kitob") == "китоб"
        assert latin_to_cyrillic("maktab") == "мактаб"
        assert latin_to_cyrillic("pul") == "пул"
        assert latin_to_cyrillic("vatan") == "ватан"

class TestDigraphs:
    def test_sh(self):
        assert latin_to_cyrillic("shahar") == "шаҳар"

    def test_ch(self):
        assert latin_to_cyrillic("chiroyli") == "чиройли"

    def test_ng(self):
        assert latin_to_cyrillic("rang") == "ранг"

    def test_oq_gq_apostrophes(self):
        assert latin_to_cyrillic("o'quvchi") == "ўқувчи"
        assert latin_to_cyrillic("g'alaba") == "ғалаба"
        assert latin_to_cyrillic("oʻzbek") == "ўзбек"
        assert latin_to_cyrillic("o`zbek") == "ўзбек"

class TestYoApostropheClash:
    def test_yoq(self):
        assert latin_to_cyrillic("yo'q") == "йўқ"

    def test_yomon(self):
        assert latin_to_cyrillic("yomon") == "ёмон"

class TestYVowelDigraphs:
    def test_ya(self):
        assert latin_to_cyrillic("yaxshi") == "яхши"

    def test_yu(self):
        assert latin_to_cyrillic("yulduz") == "юлдуз"

    def test_yer(self):
        assert latin_to_cyrillic("yer") == "ер"

class TestTutuqBelgisi:
    def test_sanat(self):
        assert latin_to_cyrillic("san'at") == "санъат"

    def test_talim(self):
        assert latin_to_cyrillic("ta'lim") == "таълим"

class TestEContextRule:
    def test_e_at_start(self):
        assert latin_to_cyrillic("eshik") == "эшик"
        assert latin_to_cyrillic("e'tibor") == "эътибор"

    def test_e_after_consonant(self):
        assert latin_to_cyrillic("kelajak") == "келажак"

class TestCasePreservation:
    def test_all_caps(self):
        assert latin_to_cyrillic("KITOB") == "КИТОБ"
        assert latin_to_cyrillic("SHAHAR") == "ШАҲАР"

    def test_title_case(self):
        assert latin_to_cyrillic("O'zbekiston") == "Ўзбекистон"
        assert latin_to_cyrillic("Shahar") == "Шаҳар"

class TestEnsureScript:
    def test_ensure_script_latin_to_cyrillic(self):
        assert ensure_script("O'zbekiston", "cyrillic") == "Ўзбекистон"

    def test_ensure_script_cyrillic_to_latin(self):
        assert ensure_script("Ўзбекистон", "latin") == "Oʻzbekiston"
