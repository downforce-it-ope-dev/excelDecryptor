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
from tkinter import Tk, ttk, StringVar, messagebox
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
    # 브라우저가 한 번이라도 뜬 적이 있었는지.
    had_client = False
    # 브라우저가 아예 안 열리는 비정상 상황 대비용 grace period.
    # 시작 후 이 시간까지도 client가 없으면 그냥 종료한다.
    startup_grace_deadline = time.time() + 30

    while True:
        if has_active_clients():
            had_client = True
        else:
            should_check_shutdown = had_client or time.time() > startup_grace_deadline
            if should_check_shutdown:
                # 짧은 재확인 (네트워크 블립/ping 경합 방지)
                time.sleep(0.5)
                if not has_active_clients():
                    server.shutdown()
                    return

        time.sleep(0.5)


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
    root.destroy()
    return result


class _UpdateProgressUI:
    """다운로드 중 표시할 진행바 창. 스레드에서 다운로드하는 동안 pump()로 UI를 돌린다."""

    def __init__(self, initial_text: str) -> None:
        self.root = Tk()
        self.root.title("dftsExcelDecryptor 업데이트")
        self.root.geometry("420x140")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.status_var = StringVar(value=initial_text)

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="업데이트 진행 중", font=("Malgun Gothic", 12, "bold")).pack(anchor="w")
        ttk.Label(frame, textvariable=self.status_var, wraplength=380).pack(anchor="w", pady=(12, 12))

        self.bar = ttk.Progressbar(frame, mode="indeterminate")
        self.bar.pack(fill="x")
        self.bar.start(12)

        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        self.pump()

    def set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.pump()

    def pump(self) -> None:
        try:
            self.root.update()
        except Exception:
            pass

    def close(self) -> None:
        try:
            self.bar.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass


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

        # 사용자의 "예" 클릭 직후 곧바로 진행바 창을 띄운다.
        ui = _UpdateProgressUI("업데이트 파일을 다운로드하는 중입니다...")
        try:
            # 다운로드는 워커 스레드에서 수행하고, 메인 스레드는 UI를 pump 한다.
            download_result: dict = {}

            def _download_worker() -> None:
                try:
                    download_result["path"] = download_update(download_url, work_dir)
                except Exception as exc:
                    download_result["error"] = exc

            worker_thread = threading.Thread(target=_download_worker, daemon=True)
            worker_thread.start()
            while worker_thread.is_alive():
                ui.pump()
                time.sleep(0.05)
            ui.pump()

            if "error" in download_result:
                raise download_result["error"]

            downloaded_exe = download_result["path"]
            updater_target = work_dir / UPDATER_EXE_NAME
            shutil.copy2(updater_source, updater_target)

            log_update(f"downloaded update to {downloaded_exe}")
            log_update(f"copied updater helper to {updater_target}")

            ui.set_status("업데이트 도구를 시작하는 중입니다...")

            subprocess.Popen(
                [str(updater_target), str(downloaded_exe), str(executable_path), str(os.getpid()), str(UPDATE_LOG_PATH)],
                close_fds=True,
            )
            log_update("spawned updater helper and exiting current app")

            # helper 창이 뜰 잠깐의 시간을 주고 본 앱 쪽 창을 닫는다 (깜빡임 최소화).
            for _ in range(6):
                ui.pump()
                time.sleep(0.1)
        finally:
            ui.close()
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

    # wsgiref/PyInstaller bootloader 잔류로 태스크 매니저에 프로세스가
    # 남는 것을 막기 위해 명시적으로 프로세스를 종료한다.
    os._exit(0)


if __name__ == "__main__":
    main()
