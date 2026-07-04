"""Windows-friendly Tkinter GUI for the inbound programme scraper."""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from dotenv import load_dotenv

from programme_scraper import OUTPUT_FILE, run_scraper


class ScraperApp(tk.Tk):
    """Small desktop app that runs the scraper without requiring command-line use."""

    def __init__(self) -> None:
        super().__init__()
        load_dotenv()
        self.title("Inbound Programme Scraper")
        self.geometry("760x520")
        self.resizable(True, True)
        self.messages: "queue.Queue[str]" = queue.Queue()
        self.worker: threading.Thread | None = None

        self.input_file = tk.StringVar()
        self.output_folder = tk.StringVar(value=str(Path.cwd()))
        self.api_key = tk.StringVar(value=os.getenv("OPENAI_API_KEY", ""))
        self.model = tk.StringVar(value=os.getenv("OPENAI_MODEL", "gpt-4o-mini"))

        self._build_ui()
        self.after(200, self._poll_messages)

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 6}
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Input Excel file").grid(row=0, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.input_file).grid(row=0, column=1, sticky="ew", **padding)
        ttk.Button(frame, text="Browse...", command=self._pick_input).grid(row=0, column=2, **padding)

        ttk.Label(frame, text="Output folder").grid(row=1, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.output_folder).grid(row=1, column=1, sticky="ew", **padding)
        ttk.Button(frame, text="Choose...", command=self._pick_output_folder).grid(row=1, column=2, **padding)

        ttk.Label(frame, text="OpenAI API key").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.api_key, show="*").grid(row=2, column=1, sticky="ew", **padding)

        ttk.Label(frame, text="Model").grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.model).grid(row=3, column=1, sticky="ew", **padding)

        self.run_button = ttk.Button(frame, text="Run", command=self._run)
        self.run_button.grid(row=4, column=1, sticky="e", **padding)

        ttk.Label(frame, text="Progress messages").grid(row=5, column=0, columnspan=3, sticky="w", **padding)
        self.log = tk.Text(frame, height=18, wrap="word", state="disabled")
        self.log.grid(row=6, column=0, columnspan=3, sticky="nsew", **padding)
        frame.rowconfigure(6, weight=1)

    def _pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select input_urls.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.input_file.set(path)

    def _pick_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_folder.set(path)

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self) -> None:
        input_path = self.input_file.get().strip()
        output_folder = self.output_folder.get().strip()
        api_key = self.api_key.get().strip()
        model = self.model.get().strip() or "gpt-4o-mini"

        if not input_path:
            messagebox.showerror("Missing input", "Please choose input_urls.xlsx.")
            return
        if not output_folder:
            messagebox.showerror("Missing output folder", "Please choose an output folder.")
            return
        if not api_key:
            messagebox.showerror("Missing API key", "Please enter your OpenAI API key.")
            return

        output_path = str(Path(output_folder) / OUTPUT_FILE)
        self.run_button.configure(state="disabled")
        self._append_log("Starting scraper...")

        def worker() -> None:
            try:
                run_scraper(input_path=input_path, output_path=output_path, api_key=api_key, model=model, progress_callback=self.messages.put)
                self.messages.put(f"DONE:{output_path}")
            except Exception as exc:
                self.messages.put(f"ERROR:{exc}")

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _poll_messages(self) -> None:
        while True:
            try:
                message = self.messages.get_nowait()
            except queue.Empty:
                break
            if message.startswith("DONE:"):
                output_path = message.removeprefix("DONE:")
                self._append_log(f"Completed. Output saved to {output_path}")
                messagebox.showinfo("Complete", f"Scraper finished.\n\nOutput saved to:\n{output_path}")
                self.run_button.configure(state="normal")
            elif message.startswith("ERROR:"):
                error = message.removeprefix("ERROR:")
                self._append_log(f"Error: {error}")
                messagebox.showerror("Error", error)
                self.run_button.configure(state="normal")
            else:
                self._append_log(message)
        self.after(200, self._poll_messages)


if __name__ == "__main__":
    ScraperApp().mainloop()
