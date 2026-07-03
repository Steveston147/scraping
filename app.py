"""Simple Windows desktop app for the inbound programme scraper MVP."""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from dotenv import load_dotenv

from programme_scraper import DEFAULT_OPENAI_MODEL, OUTPUT_FILE, run_scraper


class ProgrammeScraperApp:
    """Small tkinter GUI so non-engineers can run the scraper by double-clicking."""

    def __init__(self, root: tk.Tk) -> None:
        load_dotenv()
        self.root = root
        self.root.title("Inbound Programme Scraper MVP")
        self.root.geometry("760x560")
        self.message_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.input_file_var = tk.StringVar()
        self.output_folder_var = tk.StringVar(value=str(Path.cwd()))
        self.api_key_var = tk.StringVar(value=os.getenv("OPENAI_API_KEY", ""))
        self.model_var = tk.StringVar(value=os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL))

        self._build_ui()
        self.root.after(200, self._process_messages)

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 6}
        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(frame, text="Input Excel file (input_urls.xlsx)").grid(row=0, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.input_file_var, width=72).grid(row=1, column=0, sticky="ew", **padding)
        ttk.Button(frame, text="Browse...", command=self._select_input_file).grid(row=1, column=1, **padding)

        ttk.Label(frame, text="Output folder").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.output_folder_var, width=72).grid(row=3, column=0, sticky="ew", **padding)
        ttk.Button(frame, text="Browse...", command=self._select_output_folder).grid(row=3, column=1, **padding)

        ttk.Label(frame, text="OpenAI API key (not saved by this app)").grid(row=4, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.api_key_var, width=72, show="*").grid(row=5, column=0, sticky="ew", **padding)

        ttk.Label(frame, text="OpenAI model").grid(row=6, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.model_var, width=30).grid(row=7, column=0, sticky="w", **padding)

        self.run_button = ttk.Button(frame, text="Run", command=self._start_run)
        self.run_button.grid(row=8, column=0, sticky="w", **padding)

        ttk.Label(frame, text="Progress messages").grid(row=9, column=0, sticky="w", **padding)
        self.log_box = scrolledtext.ScrolledText(frame, height=16, wrap="word", state="disabled")
        self.log_box.grid(row=10, column=0, columnspan=2, sticky="nsew", **padding)

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(10, weight=1)

    def _select_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input_urls.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.input_file_var.set(path)

    def _select_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder_var.set(folder)

    def _start_run(self) -> None:
        input_file = self.input_file_var.get().strip()
        output_folder = self.output_folder_var.get().strip()
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get().strip() or DEFAULT_OPENAI_MODEL

        if not input_file:
            messagebox.showerror("Missing input file", "Please select input_urls.xlsx.")
            return
        if not Path(input_file).is_file():
            messagebox.showerror("Input file not found", "The selected input file does not exist.")
            return
        if not output_folder:
            messagebox.showerror("Missing output folder", "Please select an output folder.")
            return
        if not Path(output_folder).is_dir():
            messagebox.showerror("Output folder not found", "The selected output folder does not exist.")
            return
        if not api_key:
            messagebox.showerror("Missing OpenAI API key", "Please enter your OpenAI API key.")
            return

        self.run_button.configure(state="disabled")
        self._clear_log()
        self._append_log("Starting scraper...")
        output_path = str(Path(output_folder) / OUTPUT_FILE)

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            args=(input_file, output_path, api_key, model),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self, input_file: str, output_path: str, api_key: str, model: str) -> None:
        try:
            created_file = run_scraper(
                input_path=input_file,
                output_path=output_path,
                api_key=api_key,
                model=model,
                progress_callback=self._queue_progress,
            )
            self.message_queue.put(("done", created_file))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def _queue_progress(self, message: str) -> None:
        self.message_queue.put(("progress", message))

    def _process_messages(self) -> None:
        while True:
            try:
                kind, message = self.message_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                self._append_log(message)
            elif kind == "done":
                self._append_log(f"Completed. Output created: {message}")
                self.run_button.configure(state="normal")
                messagebox.showinfo("Completed", f"Created output file:\n{message}")
            elif kind == "error":
                self._append_log(f"Error: {message}")
                self.run_button.configure(state="normal")
                messagebox.showerror("Error", message)
        self.root.after(200, self._process_messages)

    def _append_log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    ProgrammeScraperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
