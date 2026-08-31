# -*- coding: utf-8 -*-
"""
SlideTranslate AI — One-click Complete Launcher
Starts FastAPI Backend (Port 8000) and Vite React Frontend (Port 5173).
"""

import os
import sys
import time
import subprocess
import webbrowser

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def free_port(port):
    try:
        output = subprocess.check_output(f"netstat -ano | findstr :{port}", shell=True).decode()
        for line in output.strip().split("\n"):
            parts = line.strip().split()
            if len(parts) >= 5 and f":{port}" in parts[1]:
                pid = int(parts[-1])
                if pid > 4 and pid != os.getpid():
                    subprocess.run(f"taskkill /F /PID {pid}", shell=True, capture_output=True)
    except Exception:
        pass

def print_banner():
    print("=" * 65)
    print("      🚀 SlideTranslate AI — Professional Slayd Tarjimoni")
    print("      FastAPI (Port 8000) + React / Vite (Port 5173) + Gemini AI")
    print("=" * 65)

def main():
    print_banner()
    root_dir = os.path.abspath(os.path.dirname(__file__))
    frontend_dir = os.path.join(root_dir, "frontend")

    print("[1/3] Portlar tozalanmoqda (8000, 5173)...")
    free_port(8000)
    free_port(5173)
    time.sleep(0.5)

    print("[2/3] Backend (FastAPI) ishga tushirilmoqda...")
    backend_cmd = [sys.executable, "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000", "--reload"]
    backend_proc = subprocess.Popen(backend_cmd, cwd=root_dir)

    print("[3/3] Frontend (Vite) ishga tushirilmoqda...")
    npm_cmd = "npm.cmd" if os.name == "nt" else "npm"
    frontend_proc = subprocess.Popen([npm_cmd, "run", "dev", "--", "--host", "--port", "5173"], cwd=frontend_dir)

    time.sleep(2)
    url = "http://localhost:5173"
    print(f"\n🎉 Barcha xizmatlar muvaffaqiyatli ishga tushdi!")
    print(f"      🌐 Veb-ilova manzili: {url}")
    print(f"      📡 Backend API: http://localhost:8000/docs")

    print("\n[Dasturni to'xtatish uchun Ctrl+C tugmalarini bosing]\n")
    try:
        backend_proc.wait()
        frontend_proc.wait()
    except KeyboardInterrupt:
        print("\nTo'xtatilmoqda...")
        backend_proc.terminate()
        frontend_proc.terminate()
        print("Xizmatlar to'xtatildi.")

if __name__ == "__main__":
    main()
