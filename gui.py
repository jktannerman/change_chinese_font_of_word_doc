"""Tkinter GUI wrapper for change_chinese_font.py.

Double-click launch.bat (or run this file directly) to open the window.
Drag a .docx onto launch.bat to pre-fill the file path.

Usage:
    py -3.13 gui.py [input.docx]
"""

from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from change_chinese_font import DEFAULT_FONT, process_document
from docx import Document

PRESET_FONTS: list[str] = [
    "FangSong",
    "SimSun",
    "SimHei",
    "KaiTi",
    "Microsoft YaHei",
    "Other...",
]

_PADDING = {"padx": 8, "pady": 5}


class ChineseFontApp(tk.Tk):
    """Main application window."""

    def __init__(self, initial_file: str | None = None) -> None:
        """Initialise the window and build all widgets.

        Args:
            initial_file: Optional path to pre-fill the file field (from drag-and-drop).
        """
        super().__init__()
        self.title("Chinese Font Changer")
        self.resizable(False, False)

        self._file_var = tk.StringVar()
        self._font_var = tk.StringVar(value=DEFAULT_FONT)
        self._custom_font_var = tk.StringVar()
        self._status_var = tk.StringVar(value="Ready")
        self._last_output: Path | None = None

        self._build_widgets()
        self._center_window()

        if initial_file:
            self._file_var.set(initial_file)
            self._update_convert_button()

    # ------------------------------------------------------------------
    # Widget construction
    # ------------------------------------------------------------------

    def _build_widgets(self) -> None:
        """Create and lay out all widgets."""
        outer = ttk.Frame(self, padding=12)
        outer.grid(row=0, column=0, sticky="nsew")

        # --- File row ---
        ttk.Label(outer, text="File:").grid(row=0, column=0, sticky="w", **_PADDING)
        file_entry = ttk.Entry(outer, textvariable=self._file_var, width=48, state="readonly")
        file_entry.grid(row=0, column=1, sticky="ew", **_PADDING)
        ttk.Button(outer, text="Browse…", command=self._browse).grid(
            row=0, column=2, sticky="w", **_PADDING
        )

        # --- Font row ---
        ttk.Label(outer, text="Font:").grid(row=1, column=0, sticky="w", **_PADDING)
        self._font_combo = ttk.Combobox(
            outer,
            textvariable=self._font_var,
            values=PRESET_FONTS,
            state="readonly",
            width=24,
        )
        self._font_combo.grid(row=1, column=1, sticky="w", **_PADDING)
        self._font_combo.bind("<<ComboboxSelected>>", self._on_font_changed)

        # Custom font entry (hidden until "Other..." is chosen)
        self._custom_entry = ttk.Entry(outer, textvariable=self._custom_font_var, width=24)
        self._custom_label = ttk.Label(outer, text="Font name:")

        # --- Convert button ---
        self._convert_btn = ttk.Button(
            outer, text="Convert", command=self._convert, state="disabled"
        )
        self._convert_btn.grid(row=3, column=1, sticky="w", **_PADDING)

        # --- Status label ---
        status_label = ttk.Label(
            outer, textvariable=self._status_var, foreground="gray", wraplength=380
        )
        status_label.grid(row=4, column=0, columnspan=3, sticky="w", **_PADDING)

        # Expand middle column
        outer.columnconfigure(1, weight=1)

        # Trace file var to enable/disable Convert
        self._file_var.trace_add("write", lambda *_: self._update_convert_button())

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_font_changed(self, _event: object = None) -> None:
        """Show or hide the custom font entry based on combobox selection."""
        if self._font_var.get() == "Other...":
            self._custom_label.grid(row=2, column=0, sticky="w", **_PADDING)
            self._custom_entry.grid(row=2, column=1, sticky="w", **_PADDING)
            self._custom_entry.focus_set()
        else:
            self._custom_label.grid_remove()
            self._custom_entry.grid_remove()

    def _browse(self) -> None:
        """Open a file dialog and populate the file field."""
        path = filedialog.askopenfilename(
            title="Select a Word document",
            filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self._file_var.set(path)

    def _convert(self) -> None:
        """Resolve the font, run the conversion, and notify the user."""
        input_path = Path(self._file_var.get())

        # Resolve font name
        selected = self._font_var.get()
        if selected == "Other...":
            font_name = self._custom_font_var.get().strip()
            if not font_name:
                messagebox.showerror("Font required", "Please enter a custom font name.")
                return
        else:
            font_name = selected

        if not input_path.exists():
            messagebox.showerror("File not found", f"Cannot find:\n{input_path}")
            return

        output_path = input_path.with_name(f"{input_path.stem}_modified.docx")

        self._status_var.set("Processing…")
        self._convert_btn.configure(state="disabled")
        self.update_idletasks()

        try:
            doc = Document(str(input_path))
            modified = process_document(doc, font_name)
            doc.save(str(output_path))
        except Exception as exc:  # noqa: BLE001
            self._status_var.set("Error — see dialog.")
            messagebox.showerror("Conversion failed", str(exc))
            self._convert_btn.configure(state="normal")
            return

        self._last_output = output_path
        self._status_var.set(f"Saved: {output_path}  ({modified} runs changed)")
        self._convert_btn.configure(state="normal")

        answer = messagebox.askquestion(
            "Done",
            f"Saved as:\n{output_path}\n\nOpen the folder in Explorer?",
            icon="info",
        )
        if answer == "yes":
            self._open_folder(output_path)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _update_convert_button(self) -> None:
        """Enable the Convert button only when a file path is present."""
        state = "normal" if self._file_var.get().strip() else "disabled"
        self._convert_btn.configure(state=state)

    def _open_folder(self, path: Path) -> None:
        """Open Windows Explorer with *path* selected.

        Args:
            path: The file to highlight in Explorer.
        """
        subprocess.Popen(["explorer", "/select,", str(path)])

    def _center_window(self) -> None:
        """Center the window on screen after geometry is calculated."""
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


def main() -> None:
    """Entry point: parse optional drag-and-drop argument and launch the GUI."""
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    app = ChineseFontApp(initial_file=initial_file)
    app.mainloop()


if __name__ == "__main__":
    main()
