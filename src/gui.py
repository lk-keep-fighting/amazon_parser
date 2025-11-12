"""Tkinter-based interface for the Amazon Product Parser.

This module provides a small graphical frontend that mirrors the capabilities of
``python -m src.amazon_product_parser``.  Users can parse individual Amazon
product pages or run the Excel batch processor without working with the command
line.
"""

from __future__ import annotations

import json
import os
import platform
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .amazon_product_parser import AmazonProductParser, process_excel


class AmazonParserGUI:
    """Render a window that exposes parser operations through a GUI."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Amazon Product Parser")
        self.root.minsize(720, 520)

        self.parser = AmazonProductParser()
        self._last_excel_destination: str | None = None

        style = ttk.Style()
        style.configure(
            "Large.TButton",
            padding=(16, 12),
            font=("TkDefaultFont", 11, "bold"),
        )

        container = ttk.Frame(root, padding=12)
        container.pack(fill="both", expand=True)

        heading = ttk.Label(
            container,
            text=(
                "Use the tabs below to parse a single Amazon URL or process an Excel "
                "workbook. These actions mirror the command line interface described "
                "in the README."
            ),
            wraplength=680,
            justify="left",
        )
        heading.pack(anchor="w", pady=(0, 10))

        self.notebook = ttk.Notebook(container)
        self.notebook.pack(fill="both", expand=True)

        self._build_single_product_tab()
        self._build_excel_tab()

    # ------------------------------------------------------------------
    # UI builders
    # ------------------------------------------------------------------
    def _build_single_product_tab(self) -> None:
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Single Product")

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

        self.url_var = tk.StringVar()
        self.pretty_var = tk.BooleanVar(value=True)
        self.single_status_var = tk.StringVar(value="")

        url_label = ttk.Label(frame, text="Amazon product URL:")
        url_label.grid(row=0, column=0, sticky="w")

        url_entry = ttk.Entry(frame, textvariable=self.url_var)
        url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=(0, 6))
        url_entry.focus_set()

        pretty_check = ttk.Checkbutton(
            frame,
            text="Pretty-print JSON output",
            variable=self.pretty_var,
        )
        pretty_check.grid(row=1, column=0, sticky="w")

        self.single_button = ttk.Button(frame, text="Parse product", command=self._on_parse_single)
        self.single_button.grid(row=1, column=2, sticky="e")

        status_label = ttk.Label(frame, textvariable=self.single_status_var, foreground="gray")
        status_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 4))

        self.single_output = ScrolledText(frame, wrap="word", height=18)
        self.single_output.configure(state="disabled", font=("Courier New", 10))
        self.single_output.grid(row=3, column=0, columnspan=3, sticky="nsew")

    def _build_excel_tab(self) -> None:
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Excel Batch")

        frame.columnconfigure(1, weight=1)

        self.excel_input_var = tk.StringVar()
        self.excel_output_var = tk.StringVar()
        self.excel_sheet_var = tk.StringVar()
        self.excel_column_var = tk.StringVar(value="1")
        self.excel_status_var = tk.StringVar(value="")

        input_label = ttk.Label(frame, text="Input workbook:")
        input_label.grid(row=0, column=0, sticky="w")

        input_entry = ttk.Entry(frame, textvariable=self.excel_input_var)
        input_entry.grid(row=0, column=1, sticky="ew", pady=(0, 4))

        input_button = ttk.Button(frame, text="Browse…", command=self._browse_input_file)
        input_button.grid(row=0, column=2, sticky="ew", padx=(6, 0), pady=(0, 4))

        output_label = ttk.Label(frame, text="Output path (optional):")
        output_label.grid(row=1, column=0, sticky="w")

        output_entry = ttk.Entry(frame, textvariable=self.excel_output_var)
        output_entry.grid(row=1, column=1, sticky="ew", pady=(0, 4))

        output_button = ttk.Button(frame, text="Browse…", command=self._browse_output_file)
        output_button.grid(row=1, column=2, sticky="ew", padx=(6, 0), pady=(0, 4))

        sheet_label = ttk.Label(frame, text="Sheet name (optional):")
        sheet_label.grid(row=2, column=0, sticky="w")

        sheet_entry = ttk.Entry(frame, textvariable=self.excel_sheet_var)
        sheet_entry.grid(row=2, column=1, sticky="ew", pady=(0, 4))

        column_label = ttk.Label(frame, text="URL column (1-based):")
        column_label.grid(row=3, column=0, sticky="w")

        column_entry = ttk.Entry(frame, textvariable=self.excel_column_var, width=8)
        column_entry.grid(row=3, column=1, sticky="w", pady=(0, 8))

        hint = ttk.Label(
            frame,
            text=(
                "Each Amazon URL or ASIN in the selected column will be fetched and parsed. "
                "Results are written to adjacent columns, mirroring the CLI behaviour."
            ),
            wraplength=660,
            justify="left",
        )
        hint.grid(row=4, column=0, columnspan=3, sticky="w")

        self.excel_button = ttk.Button(
            frame,
            text="Process workbook",
            command=self._on_process_excel,
            style="Large.TButton",
        )
        self.excel_button.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 6))

        self.excel_open_button = ttk.Button(
            frame,
            text="Open output workbook",
            command=self._on_open_excel_destination,
            state="disabled",
        )
        self.excel_open_button.grid(row=6, column=0, columnspan=3, sticky="ew")

        excel_status = ttk.Label(frame, textvariable=self.excel_status_var, foreground="gray")
        excel_status.grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------
    def _on_parse_single(self) -> None:
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter an Amazon product URL to parse.")
            return

        self.single_status_var.set("Parsing product…")
        self.single_button.config(state="disabled")
        self._set_text(self.single_output, "")

        def worker() -> None:
            try:
                product = self.parser.parse(url)
            except Exception as exc:  # pragma: no cover - network/UI runtime path
                message = str(exc)

                def notify_error() -> None:
                    self.single_status_var.set("Failed to parse product")
                    messagebox.showerror("Parsing failed", message)
                    self.single_button.config(state="normal")

                self.root.after(0, notify_error)
                return

            json_payload = product.to_dict()
            pretty = self.pretty_var.get()
            output_text = json.dumps(json_payload, ensure_ascii=False, indent=2 if pretty else None)

            def show_result() -> None:
                self._set_text(self.single_output, output_text)
                if not pretty:
                    # Maintain a trailing newline for single-line output.
                    self.single_output.config(state="normal")
                    self.single_output.insert(tk.END, "\n")
                    self.single_output.config(state="disabled")
                self.single_status_var.set("Product parsed successfully")
                self.single_button.config(state="normal")

            self.root.after(0, show_result)

        threading.Thread(target=worker, daemon=True).start()

    def _on_process_excel(self) -> None:
        input_path = self.excel_input_var.get().strip()
        output_path = self.excel_output_var.get().strip() or None
        sheet_name = self.excel_sheet_var.get().strip() or None
        column_text = self.excel_column_var.get().strip()

        if not input_path:
            messagebox.showwarning("Missing workbook", "Please select the Excel workbook to process.")
            return

        if column_text:
            try:
                column_index = int(column_text)
            except ValueError:
                messagebox.showwarning("Invalid column", "The URL column must be a positive integer.")
                return
        else:
            column_index = 1

        if column_index < 1:
            messagebox.showwarning("Invalid column", "The URL column must be a positive integer.")
            return

        self.excel_status_var.set("Processing workbook…")
        self.excel_button.config(state="disabled")
        self.excel_open_button.config(state="disabled")
        self._last_excel_destination = None

        def worker() -> None:
            try:
                destination = process_excel(
                    input_path,
                    output_path=output_path,
                    sheet_name=sheet_name,
                    column=column_index,
                    parser=self.parser,
                )
            except Exception as exc:  # pragma: no cover - runtime path

                def notify_error() -> None:
                    self.excel_status_var.set("Failed to process workbook")
                    messagebox.showerror("Excel processing failed", str(exc))
                    self.excel_button.config(state="normal")
                    self.excel_open_button.config(state="disabled")

                self.root.after(0, notify_error)
                return

            def notify_success() -> None:
                self._last_excel_destination = destination
                self.excel_status_var.set(f"Workbook updated: {destination}")
                messagebox.showinfo("Processing complete", f"Excel updated: {destination}")
                self.excel_button.config(state="normal")
                self.excel_open_button.config(state="normal")

            self.root.after(0, notify_success)

        threading.Thread(target=worker, daemon=True).start()

    def _on_open_excel_destination(self) -> None:
        if not self._last_excel_destination:
            messagebox.showinfo(
                "No workbook",
                "Run the batch processing to generate an output workbook before opening it.",
            )
            return

        path = os.path.abspath(self._last_excel_destination)
        if not os.path.exists(path):
            self.excel_open_button.config(state="disabled")
            messagebox.showwarning("File not found", f"Could not find the workbook at:\n{path}")
            return

        try:
            system = platform.system()
            if system == "Windows":
                opener = getattr(os, "startfile", None)
                if opener:
                    opener(path)
                else:
                    raise OSError("Opening files is not supported on this platform.")
            elif system == "Darwin":
                subprocess.Popen(
                    ["open", path],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            else:
                subprocess.Popen(
                    ["xdg-open", path],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception as exc:
            messagebox.showerror("Open failed", f"Failed to open the workbook: {exc}")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _browse_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self.excel_input_var.set(file_path)

    def _browse_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Select output workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self.excel_output_var.set(file_path)

    def _set_text(self, widget: ScrolledText, text: str) -> None:
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, text)
        widget.config(state="disabled")


def launch() -> None:
    """Start the GUI application."""

    root = tk.Tk()
    AmazonParserGUI(root)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - UI entry point
    launch()
