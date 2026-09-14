# -*- coding: utf-8 -*-
"""
tests/test_gemini_translator.py
GeminiTranslator modulining tozalash, guruhlash va xatoliklarni qayta ishlash testlari.
"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from unittest.mock import MagicMock
from backend.core.gemini_translator import sanitize_text, GeminiTranslator

class TestSanitizeText:
    def test_empty(self):
        assert sanitize_text("") == ""
        assert sanitize_text(None) == ""

    def test_soft_break_char(self):
        # \x0b or _x000B_ should be replaced with space so words don't stick
        assert sanitize_text("Leonardo\x0bDa Vinci") == "Leonardo Da Vinci"
        assert sanitize_text("Leonardo_x000B_Da Vinci") == "Leonardo Da Vinci"

    def test_extra_spaces_collapsed(self):
        assert sanitize_text("  Hello    World  ") == "Hello World"

class TestGeminiTranslatorMocked:
    def test_mock_batch_translation(self):
        translator = GeminiTranslator(api_key="test_dummy_key")
        
        # Mocking genai response
        mock_response = MagicMock()
        mock_response.text = '[{"id": "item1", "translated": "Renessans dahosi"}]'
        translator.client.models.generate_content = MagicMock(return_value=mock_response)
        
        items = [{"id": "item1", "original_text": "Renaissance Genius"}]
        stats = {}
        res = translator.translate_items_batch(items, target_script="latin", stats=stats)
        
        assert len(res) == 1
        assert res[0]["id"] == "item1"
        assert res[0]["translated_text"] == "Renessans dahosi"
        assert stats["total"] == 1
        assert stats["failed"] == 0

    def test_empty_items(self):
        translator = GeminiTranslator(api_key="test_dummy_key")
        stats = {}
        res = translator.translate_items_batch([], target_script="latin", stats=stats)
        assert res == []
        assert stats["total"] == 0
