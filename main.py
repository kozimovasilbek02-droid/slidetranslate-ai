import os, sys, uvicorn
from backend.main import app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    bot_mode = os.environ.get('BOT_MODE', '').lower() or ('webhook' if os.environ.get('USE_WEBHOOK', 'false').lower() == 'true' else 'polling')
    print(f"🚀 SlideTranslate AI server starting on port {port} [BOT_MODE={bot_mode}]")
    uvicorn.run('backend.main:app', host='0.0.0.0', port=port)