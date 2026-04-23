import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path


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


def download_file(url: str) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="excel_decryptor_update_"))
    target = temp_dir / "update.exe"
    with urllib.request.urlopen(url, timeout=60) as response, open(target, "wb") as file_obj:
        shutil.copyfileobj(response, file_obj)
    return target


def replace_executable(current_exe: Path, new_exe: Path) -> None:
    backup_path = current_exe.with_suffix(current_exe.suffix + ".bak")

    if backup_path.exists():
        backup_path.unlink()

    shutil.move(str(current_exe), str(backup_path))
    shutil.move(str(new_exe), str(current_exe))

    if backup_path.exists():
        backup_path.unlink()


def restart_app(executable_path: Path) -> None:
    subprocess.Popen([str(executable_path)], close_fds=True)


def main() -> int:
    if len(sys.argv) < 4:
        return 1

    current_exe = Path(sys.argv[1])
    current_version = sys.argv[2]
    manifest_url = sys.argv[3]

    try:
        manifest = load_manifest(manifest_url)
        latest_version = str(manifest.get("version", "")).strip()
        download_url = str(manifest.get("download_url", "")).strip()

        if not latest_version or not download_url:
            return 0

        if parse_version(latest_version) <= parse_version(current_version):
            return 0

        time.sleep(2)
        downloaded_file = download_file(download_url)
        replace_executable(current_exe, downloaded_file)
        restart_app(current_exe)
        return 2
    except Exception:
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
