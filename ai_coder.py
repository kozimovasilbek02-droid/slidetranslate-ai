# -*- coding: utf-8 -*-
"""
Frontier AI Coder & Proxy Engine
---------------------------------
Antigravity va mahalliy loyihalarda eng so'nggi Frontier modellari orqali kod yozish vositasi:
- Anthropic Claude Fable 5.1 (claude-fable-5-1 / claude-3-7-sonnet fallback)
- OpenAI Codex GPT-6 Astra (gpt-6-astra / gpt-4o fallback)
- OpenRouter Unified Endpoint
"""

import os
import sys
import json
import argparse
from typing import List, Optional, Dict, Any

# Ensure UTF-8 output
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

KEYS_FILE = os.path.expanduser("~/.frontier_ai_keys.json")

def load_keys() -> Dict[str, str]:
    if os.path.exists(KEYS_FILE):
        try:
            with open(KEYS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_key(provider: str, key: str):
    keys = load_keys()
    keys[provider.lower()] = key.strip()
    with open(KEYS_FILE, "w", encoding="utf-8") as f:
        json.dump(keys, f, indent=2)
    print(f"✅ {provider.upper()} API kaliti xavfsiz saqlandi: {KEYS_FILE}")

def get_key(provider: str) -> Optional[str]:
    keys = load_keys()
    prov = provider.lower()
    if prov in keys:
        return keys[prov]
    if prov == "anthropic":
        return os.environ.get("ANTHROPIC_API_KEY")
    elif prov == "openai":
        return os.environ.get("OPENAI_API_KEY")
    elif prov == "openrouter":
        return os.environ.get("OPENROUTER_API_KEY")
    return None

def query_claude_fable(prompt: str, context_files: Optional[List[str]] = None, model_name: str = "claude-fable-5-1") -> str:
    import anthropic
    api_key = get_key("anthropic") or get_key("openrouter")
    if not api_key:
        raise ValueError("Anthropic API kaliti topilmadi. O'rnatish: python ai_coder.py --set-key anthropic <KALIT>")

    client = anthropic.Anthropic(api_key=api_key)

    file_context = ""
    if context_files:
        for fpath in context_files:
            if os.path.exists(fpath):
                with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                    file_context += f"\n--- FILE: {os.path.basename(fpath)} ---\n{f.read()}\n"

    system_instruction = (
        "You are Claude Fable 5.1, the highest-tier software engineer from Anthropic. "
        "Write clean, production-grade, bug-free code. Return only the requested code and concise explanations."
    )

    full_user_content = f"{file_context}\n\nTask: {prompt}" if file_context else prompt

    candidate_models = [model_name, "claude-3-7-sonnet-20250219", "claude-3-5-sonnet-20241022"]
    for m in candidate_models:
        try:
            print(f"🤖 Claude Fable / Anthropic modeliga so'rov yuborilmoqda ({m})...")
            response = client.messages.create(
                model=m,
                max_tokens=8192,
                system=system_instruction,
                messages=[{"role": "user", "content": full_user_content}]
            )
            return response.content[0].text
        except Exception as e:
            err = str(e).lower()
            if "model" in err or "not_found" in err or "permission" in err:
                print(f"⚠️ {m} modeli tekshirildi, keyingi modelga o'tilmoqda...")
                continue
            raise e
    raise RuntimeError("Anthropic modellari bilan bog'lanishda xatolik yuz berdi.")

def query_gpt_astra(prompt: str, context_files: Optional[List[str]] = None, model_name: str = "gpt-6-astra") -> str:
    from openai import OpenAI
    api_key = get_key("openai") or get_key("openrouter")
    if not api_key:
        raise ValueError("OpenAI API kaliti topilmadi. O'rnatish: python ai_coder.py --set-key openai <KALIT>")

    client = OpenAI(api_key=api_key)

    file_context = ""
    if context_files:
        for fpath in context_files:
            if os.path.exists(fpath):
                with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                    file_context += f"\n--- FILE: {os.path.basename(fpath)} ---\n{f.read()}\n"

    system_instruction = (
        "You are Codex GPT-6 Astra, OpenAI's premier autonomous agentic coding model. "
        "Implement exact, high-performance, resilient solutions. Deliver complete, working code."
    )

    full_user_content = f"{file_context}\n\nTask: {prompt}" if file_context else prompt

    candidate_models = [model_name, "gpt-4o", "chatgpt-4o-latest", "o3-mini"]
    for m in candidate_models:
        try:
            print(f"🤖 OpenAI GPT-6 Astra modeliga so'rov yuborilmoqda ({m})...")
            response = client.chat.completions.create(
                model=m,
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": full_user_content}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            err = str(e).lower()
            if "model" in err or "not_found" in err or "permission" in err:
                print(f"⚠️ {m} modeli tekshirildi, keyingi modelga o'tilmoqda...")
                continue
            raise e
    raise RuntimeError("OpenAI modellari bilan bog'lanishda xatolik yuz berdi.")

def main():
    parser = argparse.ArgumentParser(description="Frontier AI Coder (Claude Fable 5.1 & GPT-6 Astra)")
    parser.add_argument("--model", choices=["fable", "astra", "claude", "gpt"], default="fable", help="Tanlangan model")
    parser.add_argument("--prompt", type=str, help="Kod yozish topshirig'i")
    parser.add_argument("--files", nargs="*", help="Kontekst uchun kiritiladigan fayllar")
    parser.add_argument("--output", type=str, help="Natijani saqlash uchun fayl yo'li")
    parser.add_argument("--set-key", nargs=2, metavar=("PROVIDER", "KEY"), help="API kalitni saqlash (anthropic|openai|openrouter)")
    parser.add_argument("--status", action="store_true", help="Saqlangan kalitlar holatini ko'rish")

    args = parser.parse_args()

    if args.set_key:
        prov, key = args.set_key
        save_key(prov, key)
        return

    if args.status:
        keys = load_keys()
        print("🔑 Saqlangan Frontier Kalitlar Holati:")
        for p in ["anthropic", "openai", "openrouter"]:
            val = keys.get(p) or os.environ.get(f"{p.upper()}_API_KEY")
            if val:
                masked = val[:7] + "..." + val[-4:]
                print(f" • {p.upper()}: ✅ Sozlangan ({masked})")
            else:
                print(f" • {p.upper()}: ❌ Kiritilmagan")
        return

    if not args.prompt:
        print("⚠️ Iltimos, topshiriq bering. Masalan:")
        print("   python ai_coder.py --model fable --prompt 'FastAPI uchun yangi modul yoz'")
        print("   python ai_coder.py --model astra --prompt 'Docker faylni optimallashtir'")
        print("\nKalitlarni kiritish uchun:")
        print("   python ai_coder.py --set-key anthropic <SIZNING_ANTHROPIC_KALITINGIZ>")
        print("   python ai_coder.py --set-key openai <SIZNING_OPENAI_KALITINGIZ>")
        return

    model_type = args.model.lower()
    if model_type in ["fable", "claude"]:
        result = query_claude_fable(args.prompt, args.files)
    else:
        result = query_gpt_astra(args.prompt, args.files)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"\n💾 Natija muvaffaqiyatli saqlandi: {args.output}")
    else:
        print("\n" + "=" * 60)
        print("           🚀 FRONTIER MODEL JAVOBI")
        print("=" * 60 + "\n")
        print(result)

if __name__ == "__main__":
    main()
