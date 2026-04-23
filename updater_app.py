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
    # .bak을 destination 옆(바탕화면 등)이 아니라 시스템 temp 디렉터리에 둔다.
    # OneDrive 동기화, 제어된 폴더 액세스(Controlled Folder Access),
    # 일부 백신의 랜섬웨어 휴리스틱이 보호 폴더에서의 .bak 생성을 차단하는 문제 회피.
    backup_dir = Path(tempfile.mkdtemp(prefix="dfts_excel_decryptor_backup_"))
    backup = backup_dir / f"{destination.name}.bak.{int(time.time())}"
    log_message(log_path, f"backup target={backup}")

    last_error = None
    for attempt in range(1, 31):
        try:
            if destination.exists():
                # 1차: temp로 백업 이동
                try:
                    os.replace(str(destination), str(backup))
                except Exception as backup_exc:
                    # 백업 이동이 실패하면 폴백: 기존 파일을 바로 삭제
                    log_message(
                        log_path,
                        f"attempt {attempt} backup move failed, falling back to unlink: {backup_exc}",
                    )
                    destination.unlink()

            # 2차: 새 exe를 제자리에 배치
            os.replace(str(source), str(destination))
            log_message(log_path, f"replace succeeded on attempt {attempt}")

            # 성공했으면 백업 및 temp 디렉터리 정리
            try:
                if backup.exists():
                    backup.unlink()
                shutil.rmtree(backup_dir, ignore_errors=True)
            except Exception as cleanup_exc:
                log_message(log_path, f"backup cleanup warning: {cleanup_exc}")
            return
        except Exception as exc:
            last_error = exc
            log_message(log_path, f"replace attempt {attempt} failed: {exc}")
            time.sleep(1)

    # 모든 시도 실패: 가능하면 원상복구
    if backup.exists() and not destination.exists():
        try:
            os.replace(str(backup), str(destination))
            log_message(log_path, "restored destination from backup")
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
