# -*- coding: utf-8 -*-
"""
run_tests.py
SlideTranslate AI uchun sodda va ishonchli test ishga tushiruvchi skript.
Hech qanday tashqi kutubxona (pytest) talab qilmaydi.
"""
import sys, os, time, traceback

# Encoding
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

def run():
    print("=" * 65)
    print("🧪 SLIDETRANSLATE AI — TEST SUITE")
    print("=" * 65)
    
    passed = 0
    failed = 0
    
    # 1. Transliteration tests
    print("\n[1/4] Test: Transliteration (Lotin <-> Kirill)...")
    try:
        from backend.core.transliteration import latin_to_cyrillic, cyrillic_to_latin, ensure_script
        
        assert latin_to_cyrillic("kitob") == "китоб"
        assert latin_to_cyrillic("maktab") == "мактаб"
        assert latin_to_cyrillic("shahar") == "шаҳар"
        assert latin_to_cyrillic("chiroyli") == "чиройли"
        assert latin_to_cyrillic("rang") == "ранг"
        assert latin_to_cyrillic("o'quvchi") == "ўқувчи"
        assert latin_to_cyrillic("g'alaba") == "ғалаба"
        assert latin_to_cyrillic("oʻzbek") == "ўзбек"
        assert latin_to_cyrillic("yo'q") == "йўқ"
        assert latin_to_cyrillic("yomon") == "ёмон"
        assert latin_to_cyrillic("yaxshi") == "яхши"
        assert latin_to_cyrillic("yer") == "ер"
        assert latin_to_cyrillic("san'at") == "санъат"
        assert latin_to_cyrillic("eshik") == "эшик"
        assert latin_to_cyrillic("kelajak") == "келажак"
        assert latin_to_cyrillic("O'zbekiston") == "Ўзбекистон"
        assert ensure_script("O'zbekiston", "cyrillic") == "Ўзбекистон"
        assert ensure_script("Ўзбекистон", "latin") == "Oʻzbekiston"
        
        print("  ✓ Barcha 18 ta transliteratsiya tekshiruvlari muvaffaqiyatli!")
        passed += 1
    except Exception as e:
        print(f"  ✗ Xato: {e}")
        traceback.print_exc()
        failed += 1

    # 2. Gemini Translator tests
    print("\n[2/4] Test: Gemini Translator (Sanitization & Mock)...")
    try:
        from unittest.mock import MagicMock
        from backend.core.gemini_translator import sanitize_text, GeminiTranslator
        
        assert sanitize_text("") == ""
        assert sanitize_text("Leonardo\x0bDa Vinci") == "Leonardo Da Vinci"
        assert sanitize_text("Leonardo_x000B_Da Vinci") == "Leonardo Da Vinci"
        assert sanitize_text("  Hello    World  ") == "Hello World"
        
        translator = GeminiTranslator(api_key="test_dummy_key")
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
        
        print("  ✓ Sanitization va mock batch tarjima muvaffaqiyatli!")
        passed += 1
    except Exception as e:
        print(f"  ✗ Xato: {e}")
        traceback.print_exc()
        failed += 1

    # 3. PPTX Processor tests
    print("\n[3/4] Test: PPTX Processor (Ads & Watermarks)...")
    try:
        from backend.core.pptx_processor import sanitize_control_chars, PPTXProcessor
        
        assert sanitize_control_chars("A\x0bB") == "A B"
        assert sanitize_control_chars("Test_x000B_Text") == "Test Text"
        assert PPTXProcessor._is_watermark_text("Visit slidescarnival.com for more") is True
        assert PPTXProcessor._is_watermark_text("Designed by presentationgo") is True
        assert PPTXProcessor._is_watermark_text("Leonardo Da Vinci Biography") is False
        assert PPTXProcessor._is_real_text("") is False
        assert PPTXProcessor._is_real_text("Mavzu") is True
        
        print("  ✓ Suvbelgi va boshqaruv belgilarini tozalash muvaffaqiyatli!")
        passed += 1
    except Exception as e:
        print(f"  ✗ Xato: {e}")
        traceback.print_exc()
        failed += 1

    # 4. API Endpoints
    print("\n[4/4] Test: FastAPI App & Direct /api/v1/translate Route...")
    try:
        from backend.main import app
        route_paths = [r.path for r in app.routes]
        assert "/health" in route_paths
        assert "/api/v1/health" in route_paths
        assert "/api/v1/translate" in route_paths
        assert "/api/upload" in route_paths
        assert "/api/translate" in route_paths
        assert "/api/export" in route_paths
        
        print(f"  ✓ Barcha zaruriy marshrutlar ({len(route_paths)} ta) mavjud va to'g'ri sozlangan!")
        passed += 1
    except Exception as e:
        print(f"  ✗ Xato: {e}")
        traceback.print_exc()
        failed += 1

    print("\n" + "=" * 65)
    print(f"NATIJA: {passed} ta o'tdi, {failed} ta muvaffaqiyatsiz.")
    print("=" * 65)
    if failed > 0:
        sys.exit(1)

if __name__ == "__main__":
    run()
