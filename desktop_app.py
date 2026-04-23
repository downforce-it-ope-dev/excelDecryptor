from io import BytesIO
from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import msoffcrypto


ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}


class ExcelDecryptApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("엑셀 암호 해제기")
        self.root.geometry("560x360")
        self.root.resizable(False, False)

        self.file_path_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.status_var = tk.StringVar(value="암호화된 엑셀 파일을 선택해주세요.")

        self._build_ui()

    def _build_ui(self) -> None:
        wrapper = ttk.Frame(self.root, padding=24)
        wrapper.pack(fill="both", expand=True)

        title = ttk.Label(wrapper, text="엑셀 암호 해제기", font=("Malgun Gothic", 18, "bold"))
        title.pack(anchor="w")

        desc = ttk.Label(
            wrapper,
            text="암호가 걸린 엑셀 파일을 선택하고 비밀번호를 입력하면\n복호화된 파일을 새 이름으로 저장합니다.",
            font=("Malgun Gothic", 10),
            justify="left",
        )
        desc.pack(anchor="w", pady=(8, 20))

        file_label = ttk.Label(wrapper, text="엑셀 파일", font=("Malgun Gothic", 10, "bold"))
        file_label.pack(anchor="w")

        file_row = ttk.Frame(wrapper)
        file_row.pack(fill="x", pady=(8, 16))

        self.file_entry = ttk.Entry(file_row, textvariable=self.file_path_var, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True)

        self.file_button = ttk.Button(file_row, text="파일 선택", command=self.choose_file)
        self.file_button.pack(side="left", padx=(8, 0))

        password_label = ttk.Label(wrapper, text="비밀번호", font=("Malgun Gothic", 10, "bold"))
        password_label.pack(anchor="w")

        self.password_entry = ttk.Entry(wrapper, textvariable=self.password_var, show="*")
        self.password_entry.pack(fill="x", pady=(8, 20))

        self.decrypt_button = ttk.Button(wrapper, text="암호 해제 후 저장", command=self.decrypt_file)
        self.decrypt_button.pack(fill="x")

        self.progress = ttk.Progressbar(wrapper, mode="indeterminate")
        self.progress.pack(fill="x", pady=(18, 10))

        status = ttk.Label(wrapper, textvariable=self.status_var, font=("Malgun Gothic", 9))
        status.pack(anchor="w")

        hint = ttk.Label(
            wrapper,
            text="지원 형식: xls, xlsx, xlsm, xlsb",
            font=("Malgun Gothic", 9),
        )
        hint.pack(anchor="w", pady=(8, 0))

    def choose_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="암호화된 엑셀 파일 선택",
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm *.xlsb"),
                ("All files", "*.*"),
            ],
        )

        if not file_path:
            return

        extension = Path(file_path).suffix.lower()
        if extension not in ALLOWED_EXTENSIONS:
            messagebox.showerror("오류", "지원 파일 형식은 xls, xlsx, xlsm, xlsb 입니다.")
            return

        self.file_path_var.set(file_path)
        self.status_var.set("파일이 선택되었습니다. 비밀번호를 입력해주세요.")

    def decrypt_file(self) -> None:
        file_path = self.file_path_var.get().strip()
        password = self.password_var.get()

        if not file_path:
            messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
            return

        if not password:
            messagebox.showerror("오류", "비밀번호를 입력해주세요.")
            return

        extension = Path(file_path).suffix.lower()
        default_name = f"decrypted_{Path(file_path).name}"

        save_path = filedialog.asksaveasfilename(
            title="복호화된 파일 저장",
            defaultextension=extension,
            initialfile=default_name,
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm *.xlsb"),
                ("All files", "*.*"),
            ],
        )

        if not save_path:
            return

        self._set_busy(True)
        self.status_var.set("복호화 중입니다. 잠시만 기다려주세요.")

        thread = threading.Thread(
            target=self._decrypt_worker,
            args=(file_path, password, save_path),
            daemon=True,
        )
        thread.start()

    def _decrypt_worker(self, file_path: str, password: str, save_path: str) -> None:
        try:
            with open(file_path, "rb") as source:
                encrypted_stream = BytesIO(source.read())

            office_file = msoffcrypto.OfficeFile(encrypted_stream)
            office_file.load_key(password=password)

            decrypted_stream = BytesIO()
            office_file.decrypt(decrypted_stream)
            decrypted_stream.seek(0)

            with open(save_path, "wb") as target:
                target.write(decrypted_stream.read())

            self.root.after(0, self._on_success, save_path)
        except Exception as exc:
            self.root.after(0, self._on_error, str(exc))

    def _on_success(self, save_path: str) -> None:
        self._set_busy(False)
        self.status_var.set("복호화가 완료되었습니다.")
        messagebox.showinfo("완료", f"복호화가 완료되었습니다.\n저장 위치:\n{save_path}")

    def _on_error(self, error_message: str) -> None:
        self._set_busy(False)
        self.status_var.set("복호화에 실패했습니다.")
        messagebox.showerror(
            "실패",
            "복호화에 실패했습니다.\n비밀번호가 틀렸거나 지원하지 않는 암호화 방식일 수 있습니다.\n\n"
            f"상세: {error_message}",
        )

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.file_button.config(state=state)
        self.decrypt_button.config(state=state)
        self.password_entry.config(state=state)

        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = ExcelDecryptApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
