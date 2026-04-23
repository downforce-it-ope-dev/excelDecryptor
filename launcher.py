import json
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path
from tkinter import Tk, messagebox
from wsgiref.simple_server import make_server

from app import app, has_active_clients
from app_version import APP_VERSION


HOST = "127.0.0.1"
PORT = 5000
UPDATE_CONFIG_PATH = Path(__file__).with_name("update_config.json")


def is_port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) == 0


def open_browser() -> None:
    url = f"http://{HOST}:{PORT}"
    for _ in range(30):
        if is_port_open(HOST, PORT):
            webbrowser.open(url)
            return
        time.sleep(0.2)


def stop_when_no_browser(server) -> None:
    had_client = False

    while True:
        if has_active_clients():
            had_client = True
        elif had_client:
            time.sleep(2)
            if not has_active_clients():
                server.shutdown()
                return

        time.sleep(1)


def load_manifest_url() -> str:
    try:
        with open(UPDATE_CONFIG_PATH, "r", encoding="utf-8") as file_obj:
            config = json.load(file_obj)
        return str(config.get("manifest_url", "")).strip()
    except Exception:
        return ""


def parse_version(version_text: str) -> tuple[int, ...]:
    parts = []
    for piece in version_text.strip().split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def load_manifest(manifest_url: str) -> dict:
    with urllib.request.urlopen(manifest_url, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


def download_update(download_url: str) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="excel_decryptor_update_"))
    target_path = temp_dir / "dftsExcelDecryptor.exe"

    with urllib.request.urlopen(download_url, timeout=60) as response, open(target_path, "wb") as file_obj:
        file_obj.write(response.read())

    return target_path


def ask_update(latest_version: str) -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno(
        "업데이트 확인",
        f"새 버전({latest_version})이 있습니다.\n업데이트 후 프로그램을 다시 실행할까요?",
        parent=root,
    )
    root.destroy()
    return result


def write_update_script(current_exe: Path, downloaded_exe: Path) -> Path:
    script_path = downloaded_exe.with_name("apply_update.cmd")
    script_content = "\r\n".join(
        [
            "@echo off",
            "setlocal enableextensions",
            f'set "SRC={downloaded_exe}"',
            f'set "DST={current_exe}"',
            "",
            "for /L %%i in (1,1,30) do (",
            '  taskkill /F /IM "%~nx0" > nul 2>&1',
            '  copy /Y "%SRC%" "%DST%" > nul 2>&1',
            "  if not errorlevel 1 goto success",
            "  timeout /t 1 /nobreak > nul",
            ")",
            "",
            'msg * "업데이트 파일 교체에 실패했습니다. 프로그램을 완전히 종료한 뒤 다시 시도해주세요."',
            "goto end",
            "",
            ":success",
            'start "" "%DST%"',
            'del "%SRC%" > nul 2>&1',
            "",
            ":end",
            'del "%~f0" > nul 2>&1',
        ]
    )
    script_path.write_text(script_content, encoding="utf-8")
    return script_path


def run_updater_if_needed() -> bool:
    manifest_url = load_manifest_url()
    executable_path = Path(sys.executable)

    if not manifest_url:
        return False

    if executable_path.suffix.lower() != ".exe":
        return False

    try:
        manifest = load_manifest(manifest_url)
        latest_version = str(manifest.get("version", "")).strip()
        download_url = str(manifest.get("download_url", "")).strip()

        if not latest_version or not download_url:
            return False

        if parse_version(latest_version) <= parse_version(APP_VERSION):
            return False

        if not ask_update(latest_version):
            return False

        downloaded_exe = download_update(download_url)
        update_script = write_update_script(executable_path, downloaded_exe)

        subprocess.Popen(
            ["cmd", "/c", str(update_script)],
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return True
    except Exception:
        return False


def main() -> None:
    if run_updater_if_needed():
        return

    with make_server(HOST, PORT, app) as server:
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()

        watcher_thread = threading.Thread(target=stop_when_no_browser, args=(server,), daemon=True)
        watcher_thread.start()

        server.serve_forever()


if __name__ == "__main__":
    main()
