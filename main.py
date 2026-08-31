import os, sys, uvicorn
from backend.main import app

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    uvicorn.run('backend.main:app', host='0.0.0.0', port=port)
