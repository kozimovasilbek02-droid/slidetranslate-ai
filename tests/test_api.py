# -*- coding: utf-8 -*-
"""
tests/test_api.py
FastAPI endpointlarining ishlashini tekshiruvchi testlar to'plami.
"""
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from fastapi.testclient import TestClient
from backend.main import app

client = TestClient(app)

class TestHealthEndpoints:
    def test_root_health(self):
        resp = client.get("/health")
        assert resp.status_code == 200
        data = resp.json()
        assert data["status"] == "ok"
        assert "SlideTranslate AI" in data["service"]

    def test_v1_health(self):
        resp = client.get("/api/v1/health")
        assert resp.status_code == 200
        data = resp.json()
        assert data["status"] == "ok"

    def test_invalid_translate_file_extension(self):
        resp = client.post(
            "/api/v1/translate",
            files={"file": ("test.txt", b"dummy content", "text/plain")}
        )
        assert resp.status_code == 400
