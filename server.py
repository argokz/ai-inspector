#!/usr/bin/env python3
"""
HTTP Backend сервер для веб-интерфейса AI-Inspector (на чистом Python без зависимостей)
"""

import os
import sys
import json
import tempfile
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

from report_validator import (
    ReportValidator,
    extract_from_file,
    DocxLayoutInfo,
    load_env
)

PORT = 8000
PUBLIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "public")

class AIInspectorHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PUBLIC_DIR, **kwargs)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/status":
            env = load_env()
            has_gemini = bool(env.get("GEMINI_API_KEY", "").strip())
            response = {
                "status": "ok",
                "gemini_configured": has_gemini,
                "version": "1.0.0"
            }
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(response).encode("utf-8"))
            return

        # Иначе раздавать статику
        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/validate":
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))
            post_body = self.rfile.read(content_length)

            env = load_env()
            gemini_key = env.get("GEMINI_API_KEY", "").strip()

            text_content = ""
            layout = DocxLayoutInfo()
            use_ai = False

            if "multipart/form-data" in content_type:
                # Извлекаем границу (boundary)
                boundary = content_type.split("boundary=")[-1].strip().encode("utf-8")
                parts = post_body.split(b"--" + boundary)
                
                filename = "uploaded_file"
                file_bytes = b""
                
                for part in parts:
                    if b'Content-Disposition: form-data;' in part:
                        headers_part, _, body_part = part.partition(b"\r\n\r\n")
                        headers_str = headers_part.decode("utf-8", errors="ignore")
                        body_part = body_part.rstrip(b"\r\n")
                        
                        if 'name="file"' in headers_str:
                            if 'filename="' in headers_str:
                                filename = headers_str.split('filename="')[1].split('"')[0]
                            file_bytes = body_part
                        elif 'name="use_ai"' in headers_str:
                            use_ai = (body_part.decode("utf-8", errors="ignore").lower() == "true")

                if file_bytes:
                    suffix = os.path.splitext(filename)[1] or ".docx"
                    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tf:
                        tf.write(file_bytes)
                        temp_path = tf.name

                    try:
                        text_content, layout = extract_from_file(temp_path)
                    finally:
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
            else:
                try:
                    data = json.loads(post_body.decode("utf-8"))
                    text_content = data.get("text", "")
                    use_ai = data.get("use_ai", False)
                except Exception:
                    text_content = post_body.decode("utf-8", errors="ignore")

            if not text_content.strip():
                self.send_response(400)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Пустой отчет или файл не распознан"}).encode("utf-8"))
                return

            ai_key_to_use = gemini_key if use_ai else ""
            validator = ReportValidator(text_content, layout)
            result = validator.run_all(gemini_key=ai_key_to_use)

            # Сериализуемый ответ
            response_data = {
                "total_score": result["total_score"],
                "max_total": result["max_total"],
                "percentage": result["percentage"],
                "grade_letter": result["grade_letter"],
                "gpa": result["gpa"],
                "layout": {
                    "is_docx": layout.is_docx,
                    "has_page_borders": layout.has_page_borders,
                    "borders_first_page_only": layout.borders_first_page_only,
                    "has_title_page_setting": layout.has_title_page_setting,
                    "left_margin_cm": layout.left_margin_cm,
                    "right_margin_cm": layout.right_margin_cm,
                },
                "criteria": {
                    k: {
                        "name": v.name,
                        "score": round(v.score, 1),
                        "max_score": v.max_score,
                        "passed": v.passed,
                        "comments": v.comments,
                        "warnings": v.warnings,
                        "penalties": v.penalties
                    } for k, v in result["details"].items()
                },
                "gemini_review": result.get("gemini_review")
            }

            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(response_data, ensure_ascii=False).encode("utf-8"))
            return

        self.send_response(404)
        self.end_headers()

def run_server():
    os.makedirs(PUBLIC_DIR, exist_ok=True)
    server = HTTPServer(("0.0.0.0", PORT), AIInspectorHandler)
    print(f"==================================================")
    print(f" AI-Inspector Web Server запущен на http://localhost:{PORT}")
    print(f"==================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен.")

if __name__ == "__main__":
    run_server()
