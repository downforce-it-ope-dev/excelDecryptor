import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from tkinter import Tk, ttk, StringVar, messagebox


SYNCHRONIZE = 0x00100000
WAIT_OBJECT_0 = 0x00000000
WAIT_TIMEOUT = 0x00000102
INFINITE = 0xFFFFFFFF

kernel32 = ctypes.windll.kernel32


def open_process(pid: int):
    return kernel32.OpenProcess(SYNCHRONIZE, False, pid)


def wait_for_process_exit(pid: int, timeout_ms: int) -> bool:
    handle = open_process(pid)
    if not handle:
        return True
    try:
        result = kernel32.WaitForSingleObject(handle, timeout_ms)
        return result == WAIT_OBJECT_0
    finally:
        kernel32.CloseHandle(handle)


def log_message(log_path: Path, message: str) -> None:
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as file_obj:
        file_obj.write(f"[{timestamp}] {message}\n")


class UpdateWindow:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("dftsExcelDecryptor 업데이트")
        self.root.geometry("420x140")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.status = StringVar(value="업데이트를 준비하고 있습니다...")

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="업데이트 진행 중", font=("Malgun Gothic", 12, "bold")).pack(anchor="w")
        ttk.Label(frame, textvariable=self.status, wraplength=380).pack(anchor="w", pady=(12, 12))

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x")
        self.progress.start(12)

        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

    def set_status(self, text: str) -> None:
        self.status.set(text)
        self.root.update_idletasks()

    def close(self) -> None:
        self.progress.stop()
        self.root.destroy()


class UpdateError(Exception):
    pass


def replace_executable(source: Path, destination: Path, log_path: Path) -> None:
    backup = destination.with_suffix(destination.suffix + ".bak")

    if backup.exists():
        backup.unlink()

    last_error = None
    for attempt in range(1, 31):
        try:
            if backup.exists():
                backup.unlink()
            if destination.exists():
                destination.replace(backup)
            os.replace(source, destination)
            if backup.exists():
                backup.unlink()
            log_message(log_path, f"replace succeeded on attempt {attempt}")
            return
        except Exception as exc:
            last_error = exc
            log_message(log_path, f"replace attempt {attempt} failed: {exc}")
            time.sleep(1)

    if backup.exists() and not destination.exists():
        try:
            backup.replace(destination)
        except Exception as exc:
            log_message(log_path, f"backup restore failed: {exc}")

    raise UpdateError(f"실행 파일 교체 실패: {last_error}")


def main() -> int:
    if len(sys.argv) != 5:
        return 1

    source = Path(sys.argv[1])
    destination = Path(sys.argv[2])
    old_pid = int(sys.argv[3])
    log_path = Path(sys.argv[4])

    window = UpdateWindow()

    try:
        log_message(log_path, f"updater helper started source={source} destination={destination} pid={old_pid}")
        window.set_status("기존 프로그램 종료를 기다리는 중입니다...")

        if not wait_for_process_exit(old_pid, 90000):
            raise UpdateError("기존 프로그램이 종료되지 않았습니다.")

        log_message(log_path, "old process exited")
        window.set_status("업데이트 파일을 적용하는 중입니다...")
        replace_executable(source, destination, log_path)

        window.set_status("업데이트 완료. 프로그램을 다시 실행합니다...")
        log_message(log_path, "launching updated executable")
        subprocess.Popen([str(destination)], close_fds=True)
        time.sleep(1)
        window.close()
        return 0
    except Exception as exc:
        log_message(log_path, f"updater helper failed: {exc}")
        window.close()
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror("업데이트 실패", f"업데이트 중 오류가 발생했습니다.\n{exc}")
        root.destroy()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
