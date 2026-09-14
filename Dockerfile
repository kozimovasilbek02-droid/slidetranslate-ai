# Stage 1: Build React Frontend
FROM node:20-alpine AS frontend-builder
WORKDIR /frontend
COPY frontend/package*.json ./
RUN npm install
COPY frontend/ ./
RUN npm run build

# Stage 2: Python Backend & Telegram Bot Worker
FROM python:3.11-slim
WORKDIR /app

# Install system dependencies & font utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    fontconfig \
    libfreetype6-dev \
    libxml2-dev \
    libxslt1-dev \
    libreoffice-impress-nogui \
    fonts-dejavu-core \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Install Python requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend, bot and config
COPY backend/ ./backend/
COPY bot.py .
COPY start_app.py .

# Copy built frontend from Stage 1
COPY --from=frontend-builder /frontend/dist/ ./frontend/dist/

# Create runtime directories
RUN mkdir -p uploads exports backend/fonts temp_sessions

# Expose Web port
EXPOSE 10000

# Start FastAPI Web Server with integrated Telegram Bot Worker
CMD ["sh", "-c", "uvicorn backend.main:app --host 0.0.0.0 --port ${PORT:-10000}"]
