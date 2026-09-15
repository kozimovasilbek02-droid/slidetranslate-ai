# -*- coding: utf-8 -*-
"""
SlideTranslate AI — Continuous Background Runner
------------------------------------------------
Automated runner to launch and maintain the Telegram bot process.
Auto-restarts on network disconnections or unhandled exits.
"""

import sys
import os
import time
import subprocess

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr.encoding != 'utf-8':
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BOT_SCRIPT = os.path.join(SCRIPT_DIR, "bot.py")

def main():
    print("=" * 65)
    print("      🚀 SlideTranslate AI — Background Bot Runner")
    print("=" * 65)
    print(f"📁 Working Directory: {SCRIPT_DIR}")
    print(f"🤖 Target Script: {BOT_SCRIPT}")

    restart_count = 0
    max_restarts_per_minute = 10
    recent_starts = []

    while True:
        now = time.time()
        recent_starts = [t for t in recent_starts if now - t < 60]
        if len(recent_starts) >= max_restarts_per_minute:
            print("⚠️ Reached max restart frequency limit (10/min). Sleeping 30s...")
            time.sleep(30)
            recent_starts.clear()

        recent_starts.append(now)
        restart_count += 1
        print(f"\n[Runner] Launching bot.py process (Attempt #{restart_count})...")

        try:
            p = subprocess.Popen([sys.executable, "-u", BOT_SCRIPT], cwd=SCRIPT_DIR)
            code = p.wait()
            print(f"[Runner] Bot process exited with returncode {code}.")
        except KeyboardInterrupt:
            print("\n[Runner] Stopped by user (Ctrl+C). Exiting runner.")
            break
        except Exception as e:
            print(f"[Runner] Error starting bot process: {e}")

        time.sleep(2)

if __name__ == "__main__":
    main()
