from io import BytesIO
from pathlib import Path
import threading
import time
import uuid

import msoffcrypto
from flask import Flask, jsonify, render_template, request, send_file


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB


@app.after_request
def _disable_cache(response):
    # 로컬 앱이므로 브라우저 캐시 완전 비활성.
    # 업데이트 후 이전 버전 UI가 캐시로 뜨는 문제 방지.
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
DEFAULT_PASSWORD = "minhwa6331^^!!"
CLIENT_TIMEOUT_SECONDS = 8

active_clients = {}
client_lock = threading.Lock()


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def build_output_name(filename: str) -> str:
    file_path = Path(filename)
    return f"{file_path.stem}(암호화 제거){file_path.suffix}"


def register_client() -> str:
    client_id = uuid.uuid4().hex
    with client_lock:
        active_clients[client_id] = time.time()
    return client_id


def touch_client(client_id: str) -> bool:
    with client_lock:
        if client_id not in active_clients:
            return False
        active_clients[client_id] = time.time()
        return True


def unregister_client(client_id: str) -> None:
    with client_lock:
        active_clients.pop(client_id, None)


def has_active_clients() -> bool:
    now = time.time()
    with client_lock:
        expired = [client_id for client_id, last_seen in active_clients.items() if now - last_seen > CLIENT_TIMEOUT_SECONDS]
        for client_id in expired:
            active_clients.pop(client_id, None)
        return bool(active_clients)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        client_id = register_client()
        return render_template("index.html", client_id=client_id)

    upload = request.files.get("excel_file")

    if not upload or not upload.filename:
        return render_template("index.html", error="엑셀 파일을 선택해주세요.")

    if not is_allowed_file(upload.filename):
        return render_template("index.html", error="지원 파일 형식은 xls, xlsx, xlsm, xlsb 입니다.")
    
    output_name = build_output_name(upload.filename)

    try:
        encrypted_stream = BytesIO(upload.read())
        decrypted_stream = BytesIO()

        office_file = msoffcrypto.OfficeFile(encrypted_stream)
        office_file.load_key(password=DEFAULT_PASSWORD)
        office_file.decrypt(decrypted_stream)

        decrypted_stream.seek(0)

        return send_file(
            decrypted_stream,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/octet-stream",
        )
    except Exception as exc:
        return render_template(
            "index.html",
            client_id=register_client(),
            error=f"복호화에 실패했습니다. 고정 비밀번호로 열 수 없는 파일이거나 지원하지 않는 암호화 방식일 수 있습니다. ({exc})",
        )


@app.route("/ping", methods=["POST"])
def ping():
    payload = request.get_json(silent=True) or {}
    client_id = payload.get("clientId", "")
    ok = touch_client(client_id)
    return jsonify({"ok": ok})


@app.route("/disconnect", methods=["POST"])
def disconnect():
    payload = request.get_json(silent=True) or {}
    client_id = payload.get("clientId", "")
    unregister_client(client_id)
    return ("", 204)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
