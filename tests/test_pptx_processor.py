# -*- coding: utf-8 -*-
"""
tests/test_pptx_processor.py
PPTXProcessor modulining reklama slaydlarini aniqlash, suvbelgilar va tozalash testlari.
"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from backend.core.pptx_processor import sanitize_control_chars, PPTXProcessor

class TestControlChars:
    def test_sanitize_spaces(self):
        assert sanitize_control_chars("A\x0bB") == "A B"
        assert sanitize_control_chars("Test_x000B_Text") == "Test Text"

class TestWatermarksAndAds:
    def test_watermark_detection(self):
        assert PPTXProcessor._is_watermark_text("Visit slidescarnival.com for more") is True
        assert PPTXProcessor._is_watermark_text("Designed by presentationgo") is True
        assert PPTXProcessor._is_watermark_text("Leonardo Da Vinci Biography") is False

    def test_real_text(self):
        assert PPTXProcessor._is_real_text("") is False
        assert PPTXProcessor._is_real_text("   ") is False
        assert PPTXProcessor._is_real_text("Mavzu") is True
