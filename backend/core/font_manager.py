# -*- coding: utf-8 -*-
import os
import re
import json
import requests
from typing import Dict, Any, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONTS_DIR = os.path.join(BASE_DIR, "fonts")
REGISTRY_FILE = os.path.join(FONTS_DIR, "fonts_registry.json")

SYSTEM_FONTS = [
    "Calibri", "Arial", "Times New Roman", "Segoe UI", "Aptos", "Tahoma", "Verdana",
    "Georgia", "Trebuchet MS", "Impact", "Comic Sans MS", "Consolas", "Courier New",
    "Microsoft YaHei", "SimSun", "SimHei", "KaiTi", "FangSong", "Malgun Gothic", "Meiryo"
]

TOP_100_FONTS = [
    # Modern Sans-Serif (40)
    "Roboto", "Open Sans", "Montserrat", "Poppins", "Lato", "Inter", "Oswald", "Raleway",
    "Nunito", "Rubik", "Ubuntu", "Work Sans", "Fira Sans", "Quicksand", "DM Sans",
    "Plus Jakarta Sans", "Outfit", "Jost", "Cabin", "Arimo", "Heebo", "Mulish",
    "Titillium Web", "Manrope", "Josefin Sans", "Dosis", "Exo 2", "Archivo", "Overpass",
    "Asap", "Questrial", "Catamaran", "Signika", "Karla", "Chivo", "League Spartan",
    "Lexend", "Urbanist", "Space Grotesk", "Syne",

    # Elegant Serif & Editorial (20)
    "Playfair Display", "Merriweather", "PT Sans", "PT Serif", "Lora", "Bitter",
    "Cormorant Garamond", "Libre Baskerville", "Cinzel", "Prata", "Bodoni Moda",
    "Spectral", "Domine", "Cardo", "Faustina", "EB Garamond", "Vollkorn", "Alegreya",
    "Marcellus", "Cormorant",

    # Display, Bold Headings & Artistic (25)
    "Bebas Neue", "Anton", "Lobster", "Pacifico", "Dancing Script", "Caveat",
    "Comfortaa", "Shadows Into Light", "Indie Flower", "Satisfy", "Amatic SC",
    "Righteous", "Great Vibes", "Sacramento", "Courgette", "Tangerine", "Parisienne",
    "Alex Brush", "Yellowtail", "Bad Script", "Marck Script", "Allura", "Archivo Black",
    "Teko", "Fjalla One",

    # Tech, Monospace & Minimalist (15)
    "Inconsolata", "Source Code Pro", "Fira Code", "Space Mono", "JetBrains Mono",
    "Roboto Mono", "IBM Plex Mono", "Ubuntu Mono", "Source Sans 3", "Source Serif 4",
    "Pathway Gothic One", "Yanone Kaffeesatz", "League Gothic", "Antonio", "Saira"
]

class FontManager:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(FontManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        os.makedirs(FONTS_DIR, exist_ok=True)
        self.registry: Dict[str, Dict[str, Any]] = self._load_registry()
        self._initialized = True

    def _load_registry(self) -> Dict[str, Dict[str, Any]]:
        if os.path.exists(REGISTRY_FILE):
            try:
                with open(REGISTRY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        
        initial = {}
        for sf in SYSTEM_FONTS:
            initial[sf.lower()] = {
                "name": sf,
                "type": "system",
                "file": None,
                "category": "sans-serif" if "sans" in sf.lower() or sf in ["Calibri", "Arial", "Segoe UI", "Tahoma", "Verdana"] else "serif"
            }
        self._save_registry(initial)
        return initial

    def _save_registry(self, reg: Optional[Dict[str, Any]] = None):
        data = reg or self.registry
        try:
            with open(REGISTRY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving font registry: {e}")

    def get_font_info(self, font_name: str) -> Optional[Dict[str, Any]]:
        if not font_name:
            return None
        return self.registry.get(font_name.lower())

    def fetch_font_from_web(self, font_name: str) -> Optional[Dict[str, Any]]:
        if not font_name or not font_name.strip():
            return None
        
        clean_name = font_name.strip()
        key = clean_name.lower()

        if key in self.registry and self.registry[key].get("file"):
            local_path = os.path.join(FONTS_DIR, self.registry[key]["file"])
            if os.path.exists(local_path):
                return self.registry[key]

        headers = {
            'User-Agent': 'Mozilla/5.0 (Linux; U; Android 2.2; en-us; Nexus One Build/FRF91) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1'
        }
        url = f"https://fonts.googleapis.com/css?family={clean_name.replace(' ', '+')}:400,700"

        try:
            r = requests.get(url, headers=headers, timeout=6)
            if r.status_code == 200:
                matches = re.findall(r'url\((https://[^\)]+)\)', r.text)
                if matches:
                    font_url = matches[0]
                    font_resp = requests.get(font_url, timeout=10)
                    if font_resp.status_code == 200:
                        safe_filename = re.sub(r'[^a-zA-Z0-9_\-]', '_', clean_name) + ".ttf"
                        file_path = os.path.join(FONTS_DIR, safe_filename)
                        with open(file_path, "wb") as f:
                            f.write(font_resp.content)

                        info = {
                            "name": clean_name,
                            "type": "custom",
                            "file": safe_filename,
                            "url": font_url,
                            "size_kb": round(len(font_resp.content) / 1024, 2)
                        }
                        self.registry[key] = info
                        self._save_registry()
                        print(f"✅ Downloaded & Cached font: {clean_name} ({info['size_kb']} KB)")
                        return info
        except Exception as e:
            print(f"⚠️ Could not fetch font '{clean_name}': {e}")

        fallback = {
            "name": clean_name,
            "type": "system",
            "file": None
        }
        self.registry[key] = fallback
        self._save_registry()
        return fallback

    def ensure_font_available(self, font_name: str) -> Dict[str, Any]:
        if not font_name:
            return {"name": "Calibri", "type": "system"}
        
        key = font_name.strip().lower()
        if key in self.registry:
            info = self.registry[key]
            if info.get("file"):
                fp = os.path.join(FONTS_DIR, info["file"])
                if os.path.exists(fp):
                    return info
            else:
                return info

        return self.fetch_font_from_web(font_name) or {"name": font_name, "type": "system"}

    def preload_top_100_fonts(self, max_workers: int = 8) -> Dict[str, Any]:
        print(f"🚀 Starting Top {len(TOP_100_FONTS)} fonts preload...")
        success_count = 0
        failed_count = 0

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_font = {executor.submit(self.fetch_font_from_web, fn): fn for fn in TOP_100_FONTS}
            for future in as_completed(future_to_font):
                fn = future_to_font[future]
                try:
                    res = future.result()
                    if res and res.get("file"):
                        success_count += 1
                    else:
                        failed_count += 1
                except Exception:
                    failed_count += 1

        print(f"🎉 Preload finished! Downloaded: {success_count}, Existing/System: {failed_count}")
        return {
            "total_top_fonts": len(TOP_100_FONTS),
            "downloaded": success_count,
            "system_or_existing": failed_count,
            "registry_total": len(self.registry)
        }

    def generate_font_css(self) -> str:
        css_rules = []
        for key, item in self.registry.items():
            if item.get("file"):
                fname = item["name"]
                ffile = item["file"]
                rule = f"@font-face {{\n  font-family: '{fname}';\n  src: url('/api/fonts/file/{ffile}') format('truetype');\n  font-weight: normal;\n  font-style: normal;\n  font-display: swap;\n}}"
                css_rules.append(rule)
        return "\n\n".join(css_rules)

font_manager = FontManager()
