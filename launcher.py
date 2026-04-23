import json
import os
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
UPDATE_LOG_PATH = Path.home() / "dftsExcelDecryptor_update.log"


def log_update(message: str) -> None:
    try:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(UPDATE_LOG_PATH, "a", encoding="utf-8") as file_obj:
            file_obj.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


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
    except Exception as exc:
        log_update(f"manifest url read failed: {exc}")
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


def write_update_script(current_exe: Path, downloaded_exe: Path, current_pid: int) -> Path:
    script_path = downloaded_exe.with_name("apply_update.cmd")
    log_path = UPDATE_LOG_PATH
    script_content = "\r\n".join(
        [
            "@echo off",
            "setlocal enableextensions",
            f'set "SRC={downloaded_exe}"',
            f'set "DST={current_exe}"',
            f'set "PID={current_pid}"',
            f'set "LOG={log_path}"',
            'echo [%date% %time%] updater script started >> "%LOG%"',
            "",
            "for /L %%i in (1,1,60) do (",
            '  tasklist /FI "PID eq %PID%" | find "%PID%" > nul',
            "  if errorlevel 1 goto replace",
            '  echo [%date% %time%] waiting for process %PID% to exit >> "%LOG%"',
            "  timeout /t 1 /nobreak > nul",
            ")",
            'echo [%date% %time%] timeout waiting for old process exit >> "%LOG%"',
            "goto end",
            "",
            ":replace",
            "for /L %%i in (1,1,30) do (",
            '  copy /Y "%SRC%" "%DST%" > nul 2>&1',
            "  if not errorlevel 1 goto success",
            '  echo [%date% %time%] copy retry %%i failed >> "%LOG%"',
            "  timeout /t 1 /nobreak > nul",
            ")",
            'echo [%date% %time%] failed to replace executable >> "%LOG%"',
            "goto end",
            "",
            ":success",
            'echo [%date% %time%] executable replaced successfully >> "%LOG%"',
            'start "" "%DST%"',
            'del "%SRC%" > nul 2>&1',
            'echo [%date% %time%] restarted updated executable >> "%LOG%"',
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
        log_update("manifest url missing")
        return False

    if executable_path.suffix.lower() != ".exe":
        log_update(f"skip updater in non-exe mode: {executable_path}")
        return False

    try:
        log_update(f"current version={APP_VERSION}")
        log_update(f"manifest url={manifest_url}")

        manifest = load_manifest(manifest_url)
        latest_version = str(manifest.get("version", "")).strip()
        download_url = str(manifest.get("download_url", "")).strip()

        log_update(f"latest version={latest_version}")
        log_update(f"download url={download_url}")

        if not latest_version or not download_url:
            log_update("manifest missing version or download url")
            return False

        if parse_version(latest_version) <= parse_version(APP_VERSION):
            log_update("already latest version")
            return False

        if not ask_update(latest_version):
            log_update("user cancelled update")
            return False

        downloaded_exe = download_update(download_url)
        log_update(f"downloaded update to {downloaded_exe}")

        update_script = write_update_script(executable_path, downloaded_exe, os.getpid())
        log_update(f"created update script {update_script}")

        subprocess.Popen(
            ["cmd", "/c", str(update_script)],
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        log_update("spawned update script and exiting current app")
        return True
    except Exception as exc:
        log_update(f"updater failed: {exc}")
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
