#!/usr/bin/env bash
# SlideTranslate AI — 1-Click VPS Deployment Script (Ubuntu / Debian)
set -e

echo "=========================================================="
echo "   🚀 SlideTranslate AI — VPS O'rnatish va Ishga Tushirish"
echo "=========================================================="

# 1. Update system & install Docker if missing
if ! command -v docker &> /dev/null; then
    echo "[1/4] Docker o'rnatilmoqda..."
    apt-get update -y
    apt-get install -y ca-certificates curl gnupg lsb-release
    mkdir -p /etc/apt/keyrings
    curl -fsSL https://download.docker.com/linux/ubuntu/gpg | gpg --dearmor -o /etc/apt/keyrings/docker.gpg
    echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | tee /etc/apt/sources.list.d/docker.list > /dev/null
    apt-get update -y
    apt-get install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin docker-compose
fi

# 2. Build & Launch Docker containers
echo "[2/4] Loyiha konteynerlari tayyorlanmoqda (Web + Bot)..."
docker-compose down || true
docker-compose build
docker-compose up -d

echo "[3/4] Konteynerlar holati tekshirilmoqda..."
docker-compose ps

echo "=========================================================="
echo "🎉 TABRIKLAYMIZ! Loyihangiz VPS serverda 24/7 ishga tushdi!"
echo "   🌐 Veb-ilova: http://$(curl -s ifconfig.me):8000"
echo "   🤖 Telegram Bot: @slayd_tarjimabot faol holatda!"
echo "=========================================================="
