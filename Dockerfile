# Base Python Image
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies & font utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    fontconfig \
    libfreetype6-dev \
    libxml2-dev \
    libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

# Install Python requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend and frontend files
COPY backend/ ./backend/
COPY frontend/dist/ ./frontend/dist/
COPY bot.py .
COPY start_app.py .

# Create uploads, exports, fonts directories
RUN mkdir -p uploads exports backend/fonts temp_sessions

# Expose FastAPI port
EXPOSE 8000

CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
