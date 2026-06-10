"""Графический интерфейс VKR Checker."""
from __future__ import annotations

from pathlib import Path
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .folder_checker import check_folder, find_docx_files


class VKRCheckerApp(tk.Tk):
    """Окно приложения для проверки всех DOCX-файлов в выбранной папке."""

    def __init__(self) -> None:
        super().__init__()
        self.title("VKR Checker")
        self.geometry("760x520")
        self.minsize(720, 480)

        self.selected_folder = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="Выберите папку с DOCX-файлами")
        self.progress_text = tk.StringVar(value="0 / 0")
        self.last_summary_report: Path | None = None
        self.last_reports_dir: Path | None = None
        self.is_running = False

        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)

        title = ttk.Label(
            self,
            text="Проверка оформления ВКР",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=18, pady=(16, 8))

        folder_frame = ttk.LabelFrame(self, text="Папка для проверки")
        folder_frame.grid(row=1, column=0, sticky="ew", padx=18, pady=8)
        folder_frame.columnconfigure(0, weight=1)

        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.selected_folder)
        self.folder_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        self.choose_button = ttk.Button(
            folder_frame,
            text="Выбрать папку",
            command=self.choose_folder,
        )
        self.choose_button.grid(row=0, column=1, padx=(0, 10), pady=10)

        buttons_frame = ttk.Frame(self)
        buttons_frame.grid(row=2, column=0, sticky="ew", padx=18, pady=8)

        self.start_button = ttk.Button(
            buttons_frame,
            text="Начать проверку",
            command=self.start_check,
        )
        self.start_button.pack(side="left")

        self.open_report_button = ttk.Button(
            buttons_frame,
            text="Открыть сводный отчёт",
            command=self.open_summary_report,
            state="disabled",
        )
        self.open_report_button.pack(side="left", padx=8)

        self.open_folder_button = ttk.Button(
            buttons_frame,
            text="Открыть папку отчётов",
            command=self.open_reports_folder,
            state="disabled",
        )
        self.open_folder_button.pack(side="left")

        progress_frame = ttk.Frame(self)
        progress_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=8)
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")

        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_text, width=10)
        self.progress_label.grid(row=0, column=1, padx=(10, 0))

        log_frame = ttk.LabelFrame(self, text="Ход проверки")
        log_frame.grid(row=4, column=0, sticky="nsew", padx=18, pady=(8, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log = tk.Text(log_frame, wrap="word", height=14, state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)
        self.log.configure(yscrollcommand=scrollbar.set)

        status_bar = ttk.Label(self, textvariable=self.status_text, anchor="w")
        status_bar.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 10))

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="Выберите папку с DOCX-файлами")
        if not folder:
            return
        self.selected_folder.set(folder)
        count = len(find_docx_files(folder))
        self.status_text.set(f"Найдено DOCX-файлов: {count}")
        self._append_log(f"Выбрана папка: {folder}")
        self._append_log(f"Найдено DOCX-файлов: {count}")

    def start_check(self) -> None:
        if self.is_running:
            return

        folder = self.selected_folder.get().strip()
        if not folder:
            messagebox.showwarning("Папка не выбрана", "Сначала выберите папку с DOCX-файлами.")
            return

        folder_path = Path(folder)
        if not folder_path.exists() or not folder_path.is_dir():
            messagebox.showerror("Ошибка", "Выбранная папка не найдена.")
            return

        docx_count = len(find_docx_files(folder_path))
        if docx_count == 0:
            messagebox.showinfo("DOCX не найдены", "В выбранной папке нет файлов .docx.")
            return

        self.is_running = True
        self.last_summary_report = None
        self.last_reports_dir = None
        self.start_button.configure(state="disabled")
        self.choose_button.configure(state="disabled")
        self.open_report_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=docx_count)
        self.progress_text.set(f"0 / {docx_count}")
        self.status_text.set("Проверка выполняется...")
        self._append_log("\nНачата проверка.")

        thread = threading.Thread(target=self._run_check_thread, args=(folder_path,), daemon=True)
        thread.start()

    def _run_check_thread(self, folder_path: Path) -> None:
        try:
            summaries, summary_report, error_log = check_folder(
                folder_path,
                progress_callback=self._on_progress_threadsafe,
            )
            self.after(0, self._finish_success, summaries, summary_report, error_log)
        except Exception as exc:  # noqa: BLE001
            self.after(0, self._finish_error, exc)

    def _on_progress_threadsafe(self, current: int, total: int, file_path: Path, status: str) -> None:
        self.after(0, self._update_progress, current, total, file_path, status)

    def _update_progress(self, current: int, total: int, file_path: Path, status: str) -> None:
        self.progress.configure(value=current)
        self.progress_text.set(f"{current} / {total}")
        self.status_text.set(f"{status}: {file_path.name}")
        self._append_log(f"[{current}/{total}] {status}: {file_path.name}")

    def _finish_success(self, summaries, summary_report: Path, error_log: Path) -> None:
        self.is_running = False
        self.last_summary_report = summary_report
        self.last_reports_dir = summary_report.parent

        total = len(summaries)
        total_errors = sum(s.errors or 0 for s in summaries)
        total_warnings = sum(s.warnings or 0 for s in summaries)
        failed = sum(1 for s in summaries if s.status == "failed")

        self.status_text.set(
            f"Готово. Файлов: {total}, ошибок: {total_errors}, предупреждений: {total_warnings}, технических сбоев: {failed}"
        )
        self._append_log("Проверка завершена.")
        self._append_log(f"Сводный отчёт: {summary_report}")
        self._append_log(f"Лог ошибок: {error_log}")

        self.start_button.configure(state="normal")
        self.choose_button.configure(state="normal")
        self.open_report_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")

        messagebox.showinfo(
            "Проверка завершена",
            f"Проверено файлов: {total}\nОшибок: {total_errors}\nПредупреждений: {total_warnings}\nСводный отчёт сохранён в папке vkr_reports.",
        )

    def _finish_error(self, exc: Exception) -> None:
        self.is_running = False
        self.status_text.set(f"Ошибка: {exc}")
        self._append_log(f"Критическая ошибка: {exc}")
        self.start_button.configure(state="normal")
        self.choose_button.configure(state="normal")
        messagebox.showerror("Ошибка", str(exc))

    def open_summary_report(self) -> None:
        if self.last_summary_report and self.last_summary_report.exists():
            open_path(self.last_summary_report)
        else:
            messagebox.showwarning("Отчёт не найден", "Сначала выполните проверку.")

    def open_reports_folder(self) -> None:
        if self.last_reports_dir and self.last_reports_dir.exists():
            open_path(self.last_reports_dir)
        else:
            messagebox.showwarning("Папка не найдена", "Сначала выполните проверку.")

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def open_path(path: Path) -> None:
    """Открывает файл или папку стандартным способом для текущей ОС."""
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def run_gui() -> None:
    app = VKRCheckerApp()
    app.mainloop()
