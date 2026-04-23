import json
import os
import shutil
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
UPDATER_EXE_NAME = "dftsExcelDecryptorUpdater.exe"


def log_update(message: str) -> None:
    try:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(UPDATE_LOG_PATH, "a", encoding="utf-8") as file_obj:
            file_obj.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


def open_browser() -> None:
    url = f"http://{HOST}:{PORT}"
    webbrowser.open(url)


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
    with urllib.request.urlopen(manifest_url, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))


def download_update(download_url: str, work_dir: Path) -> Path:
    target_path = work_dir / "dftsExcelDecryptor.exe"
    with urllib.request.urlopen(download_url, timeout=120) as response, open(target_path, "wb") as file_obj:
        shutil.copyfileobj(response, file_obj)
    return target_path


def ask_update(latest_version: str) -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno(
        "업데이트 확인",
        f"새 버전({latest_version})이 있습니다.\n업데이트를 진행할까요?",
        parent=root,
    )
    if result:
        messagebox.showinfo(
            "업데이트 진행",
            "업데이트 파일을 다운로드하고 있습니다.\n잠시 후 프로그램이 자동으로 다시 실행됩니다.",
            parent=root,
        )
    root.destroy()
    return result


def get_bundled_updater_path() -> Path | None:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    updater_path = base_dir / UPDATER_EXE_NAME
    if updater_path.exists():
        return updater_path
    return None


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

        updater_source = get_bundled_updater_path()
        if updater_source is None:
            log_update("bundled updater exe not found")
            return False

        work_dir = Path(tempfile.mkdtemp(prefix="dfts_excel_decryptor_update_"))
        downloaded_exe = download_update(download_url, work_dir)
        updater_target = work_dir / UPDATER_EXE_NAME
        shutil.copy2(updater_source, updater_target)

        log_update(f"downloaded update to {downloaded_exe}")
        log_update(f"copied updater helper to {updater_target}")

        subprocess.Popen(
            [str(updater_target), str(downloaded_exe), str(executable_path), str(os.getpid()), str(UPDATE_LOG_PATH)],
            close_fds=True,
        )
        log_update("spawned updater helper and exiting current app")
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
