from __future__ import annotations

import json
import queue
import subprocess
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import PhotoImage, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from . import APP_NAME
from .config import DEFAULT_TEMPLATE_PATH, WINDOW_ICON_RELATIVE_PATH
from .generator import (
    GenerationError,
    default_output_dir,
    default_year,
    ensure_excel_available,
    generate_year_files,
)
from .runtime import resource_path, state_file_path


BG_APP = "#dbe1e8"
BG_HEADER = "#17324d"
BG_HEADER_ALT = "#214261"
BG_CARD = "#f6f8fb"
BG_CARD_ALT = "#eef2f6"
BG_SURFACE = "#ffffff"
TEXT_DARK = "#18324b"
TEXT_MUTED = "#66798d"
TEXT_LIGHT = "#eef4fb"
ACCENT_BLUE = "#205493"
ACCENT_BLUE_HOVER = "#183f73"
ACCENT_RED = "#c53a47"
ACCENT_RED_DEEP = "#a62835"
BORDER = "#c7d0da"
LOG_BG = "#0f1a28"
LOG_FG = "#dbe8f4"
LOG_INFO = "#8bb8ff"
LOG_DONE = "#7bd389"
LOG_ERROR = "#ff8b98"


@dataclass
class AppState:
    template_path: str = ""
    output_dir: str = ""
    year: str = str(default_year())


def load_state() -> AppState:
    path = state_file_path()
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return AppState(
                template_path=str(data.get("template_path", "")),
                output_dir=str(data.get("output_dir", "")),
                year=str(data.get("year", default_year())),
            )
        except Exception:
            pass

    template = str(DEFAULT_TEMPLATE_PATH) if DEFAULT_TEMPLATE_PATH.exists() else ""
    output_dir = ""
    if template:
        output_dir = str(default_output_dir(Path(template), default_year()))
    return AppState(template_path=template, output_dir=output_dir)


def save_state(state: AppState) -> None:
    state_file_path().write_text(
        json.dumps(
            {
                "template_path": state.template_path,
                "output_dir": state.output_dir,
                "year": state.year,
            },
            indent=2,
            ensure_ascii=True,
        ),
        encoding="utf-8",
    )


def open_folder(path: Path) -> None:
    subprocess.Popen(["explorer.exe", str(path)])


class PMPlanApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1180x760")
        self.root.minsize(1040, 700)
        self.root.configure(bg=BG_APP)

        self.icon_preview: PhotoImage | None = None
        self._configure_styles()
        self._apply_icon()

        state = load_state()
        self.template_var = StringVar(value=state.template_path)
        self.year_var = StringVar(value=state.year)
        self.output_var = StringVar(value=state.output_dir)
        self.status_var = StringVar(value="Ready for production.")
        self.year_badge_var = StringVar()
        self.generated_dir: Path | None = None

        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.generate_button: ttk.Button | None = None
        self.open_button: ttk.Button | None = None
        self.progress: ttk.Progressbar | None = None
        self.log_text: ScrolledText | None = None
        self.status_value_label: tk.Label | None = None

        self.template_var.trace_add("write", self._on_input_change)
        self.year_var.trace_add("write", self._on_input_change)
        self.output_var.trace_add("write", self._on_input_change)

        self._build_ui()
        self._refresh_header_badge()
        self.root.after(150, self._poll_queue)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("App.TFrame", background=BG_APP)
        style.configure("Card.TFrame", background=BG_SURFACE)
        style.configure(
            "Field.TLabel",
            background=BG_SURFACE,
            foreground=TEXT_DARK,
            font=("Segoe UI Semibold", 10),
        )
        style.configure(
            "Hint.TLabel",
            background=BG_SURFACE,
            foreground=TEXT_MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "PanelTitle.TLabel",
            background=BG_SURFACE,
            foreground=TEXT_DARK,
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "Surface.TEntry",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground=TEXT_DARK,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=8,
        )
        style.map(
            "Surface.TEntry",
            bordercolor=[("focus", ACCENT_BLUE)],
            lightcolor=[("focus", ACCENT_BLUE)],
            darkcolor=[("focus", ACCENT_BLUE)],
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10),
            padding=(18, 12),
            background=ACCENT_BLUE,
            foreground="#ffffff",
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", ACCENT_BLUE_HOVER), ("disabled", "#97b4d8")],
            foreground=[("disabled", "#eef4fb")],
        )
        style.configure(
            "Neutral.TButton",
            font=("Segoe UI", 10),
            padding=(16, 11),
            background="#edf1f5",
            foreground=TEXT_DARK,
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Neutral.TButton",
            background=[("active", "#e1e7ee"), ("disabled", "#f4f6f8")],
            foreground=[("disabled", "#95a4b3")],
        )
        style.configure(
            "Progress.Horizontal.TProgressbar",
            troughcolor="#e3e8ee",
            background=ACCENT_RED,
            bordercolor="#e3e8ee",
            lightcolor=ACCENT_RED,
            darkcolor=ACCENT_RED,
            thickness=10,
        )

    def _apply_icon(self) -> None:
        icon_path = resource_path(*WINDOW_ICON_RELATIVE_PATH.parts)
        if icon_path.exists():
            try:
                self.root.iconbitmap(default=str(icon_path))
            except Exception:
                pass

        preview_path = resource_path("assets", "app_icon.png")
        if preview_path.exists():
            try:
                self.icon_preview = PhotoImage(file=str(preview_path)).subsample(8, 8)
            except Exception:
                self.icon_preview = None

    def _make_card(self, parent, bg: str = BG_SURFACE, padx: int = 18, pady: int = 18) -> tk.Frame:
        card = tk.Frame(
            parent,
            bg=bg,
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
            padx=padx,
            pady=pady,
        )
        return card

    def _build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=BG_APP, padx=20, pady=20)
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(2, weight=1)

        self._build_header(outer).grid(row=0, column=0, sticky="ew")
        tk.Frame(outer, bg=BG_APP, height=14).grid(row=1, column=0, sticky="ew")
        body = tk.Frame(outer, bg=BG_APP)
        body.grid(row=2, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=7)
        body.grid_columnconfigure(1, weight=4)
        body.grid_rowconfigure(0, weight=1)

        self._build_setup_panel(body).grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        self._build_status_panel(body).grid(row=0, column=1, sticky="nsew")

    def _build_header(self, parent) -> tk.Frame:
        header = tk.Frame(
            parent,
            bg=BG_HEADER,
            bd=0,
            highlightthickness=1,
            highlightbackground="#244a6d",
            padx=24,
            pady=22,
        )
        header.grid_columnconfigure(1, weight=1)

        left_badge = tk.Frame(header, bg=BG_HEADER_ALT, padx=14, pady=14)
        left_badge.grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 18))
        if self.icon_preview is not None:
            tk.Label(left_badge, image=self.icon_preview, bg=BG_HEADER_ALT).pack()
        else:
            tk.Label(
                left_badge,
                text="PM",
                bg=BG_HEADER_ALT,
                fg=TEXT_LIGHT,
                font=("Segoe UI Semibold", 18),
            ).pack()

        tk.Label(
            header,
            text="Mass Production PM Schedule Generator",
            bg=BG_HEADER,
            fg=TEXT_LIGHT,
            font=("Segoe UI Semibold", 20),
        ).grid(row=0, column=1, sticky="w")

        tk.Label(
            header,
            text=(
                "Operational desktop tool for planning, generating, and exporting Jan-Dec "
                "ALL BACKLINE PM workbooks with controlled PM PLAN and DE-DROSS scheduling."
            ),
            bg=BG_HEADER,
            fg="#bfd0df",
            font=("Segoe UI", 10),
            wraplength=660,
            justify="left",
        ).grid(row=1, column=1, sticky="w", pady=(6, 0))

        header_pill = tk.Frame(header, bg=ACCENT_RED, padx=14, pady=12)
        header_pill.grid(row=0, column=2, rowspan=2, sticky="ne")
        tk.Label(
            header_pill,
            textvariable=self.year_badge_var,
            bg=ACCENT_RED,
            fg="#ffffff",
            font=("Segoe UI Semibold", 10),
        ).pack(anchor="e")

        accent_line = tk.Frame(header, bg=ACCENT_RED, height=4)
        accent_line.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(18, 0))

        return header

    def _build_setup_panel(self, parent) -> tk.Frame:
        panel = self._make_card(parent)
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            panel,
            text="Production Setup",
            bg=BG_SURFACE,
            fg=TEXT_DARK,
            font=("Segoe UI Semibold", 15),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            panel,
            text="Select the template, choose the production year, and define the output workspace.",
            bg=BG_SURFACE,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="w", pady=(6, 18))

        self._build_field(
            panel,
            row=2,
            label="Template workbook",
            hint="Primary source workbook used as the base structure for all generated monthly files.",
            variable=self.template_var,
            button_text="Browse...",
            button_command=self._pick_template,
        )
        self._build_field(
            panel,
            row=4,
            label="Output workspace",
            hint="Target folder for the generated Jan-Dec PM plan workbooks.",
            variable=self.output_var,
            button_text="Browse...",
            button_command=self._pick_output,
        )
        self._build_year_field(panel, row=6)

        ttk.Separator(panel, orient="horizontal").grid(row=8, column=0, sticky="ew", pady=18)

        action_row = tk.Frame(panel, bg=BG_SURFACE)
        action_row.grid(row=9, column=0, sticky="w")
        action_row.grid_columnconfigure(2, weight=1)

        self.generate_button = ttk.Button(
            action_row,
            text="Generate Jan-Dec files",
            command=self._start_generation,
            style="Primary.TButton",
        )
        self.generate_button.grid(row=0, column=0, sticky="w")

        self.open_button = ttk.Button(
            action_row,
            text="Open output folder",
            command=self._open_generated_folder,
            state="disabled",
            style="Neutral.TButton",
        )
        self.open_button.grid(row=0, column=1, sticky="w", padx=(12, 0))

        return panel

    def _build_field(
        self,
        parent,
        row: int,
        label: str,
        hint: str,
        variable: StringVar,
        button_text: str,
        button_command,
    ) -> None:
        block = tk.Frame(parent, bg=BG_SURFACE)
        block.grid(row=row, column=0, sticky="ew", pady=(0, 18))
        block.grid_columnconfigure(0, weight=1)

        ttk.Label(block, text=label, style="Field.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))

        input_row = tk.Frame(block, bg=BG_SURFACE)
        input_row.grid(row=1, column=0, sticky="ew")
        input_row.grid_columnconfigure(0, weight=1)

        ttk.Entry(input_row, textvariable=variable, style="Surface.TEntry").grid(
            row=0,
            column=0,
            sticky="ew",
        )
        ttk.Button(
            input_row,
            text=button_text,
            command=button_command,
            style="Neutral.TButton",
            width=16,
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(14, 0),
        )
        ttk.Label(block, text=hint, style="Hint.TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))

    def _build_year_field(self, parent, row: int) -> None:
        block = tk.Frame(parent, bg=BG_SURFACE)
        block.grid(row=row, column=0, sticky="ew", pady=(0, 18))
        block.grid_columnconfigure(0, weight=1)

        ttk.Label(block, text="Production year", style="Field.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))

        year_wrap = tk.Frame(block, bg=BG_SURFACE)
        year_wrap.grid(row=1, column=0, sticky="ew")
        year_wrap.grid_columnconfigure(1, weight=1)

        ttk.Entry(year_wrap, textvariable=self.year_var, width=16, style="Surface.TEntry").grid(row=0, column=0, sticky="w")
        ttk.Button(
            year_wrap,
            text="Use current year",
            command=self._use_current_year,
            style="Neutral.TButton",
            width=16,
        ).grid(row=0, column=1, sticky="w", padx=(14, 0))
        ttk.Label(
            block,
            text="Date columns automatically follow the real machine calendar for the selected year.",
            style="Hint.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

    def _build_status_panel(self, parent) -> tk.Frame:
        panel = self._make_card(parent, bg=BG_HEADER, padx=18, pady=18)
        panel.configure(highlightbackground="#284866")
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            panel,
            text="Run Status",
            bg=BG_HEADER,
            fg="#d8e4ef",
            font=("Segoe UI Semibold", 14),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            panel,
            text="Live state for the current generation session.",
            bg=BG_HEADER,
            fg="#9db3c8",
            font=("Segoe UI", 9),
        ).grid(row=1, column=0, sticky="w", pady=(4, 14))

        self.status_value_label = tk.Label(
            panel,
            textvariable=self.status_var,
            bg=BG_HEADER,
            fg="#ffffff",
            font=("Segoe UI Semibold", 12),
            justify="left",
            wraplength=330,
        )
        self.status_value_label.grid(row=2, column=0, sticky="w")

        self.progress = ttk.Progressbar(panel, mode="indeterminate", style="Progress.Horizontal.TProgressbar")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(14, 14))

        overview = tk.Frame(panel, bg=BG_HEADER_ALT, padx=12, pady=12)
        overview.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        tk.Label(
            overview,
            text="Schedule scope",
            bg=BG_HEADER_ALT,
            fg="#ffccd2",
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w")
        tk.Label(
            overview,
            text=(
                "PM PLAN anchor is applied to BT01-BT09, A12, and A13. "
                "PM PLAN runs every 28 days and DE-DROSS runs every 7 days around that anchor."
            ),
            bg=BG_HEADER_ALT,
            fg="#d0dceb",
            font=("Segoe UI", 9),
            wraplength=300,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        note = tk.Frame(panel, bg=BG_HEADER_ALT, padx=12, pady=10)
        note.grid(row=5, column=0, sticky="ew")
        tk.Label(
            note,
            text="System requirement",
            bg=BG_HEADER_ALT,
            fg="#ffccd2",
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w")
        tk.Label(
            note,
            text="Microsoft Excel must be installed and available through COM automation.",
            bg=BG_HEADER_ALT,
            fg="#d0dceb",
            font=("Segoe UI", 9),
            wraplength=300,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        log_wrap = tk.Frame(panel, bg=BG_HEADER, pady=0)
        log_wrap.grid(row=6, column=0, sticky="nsew", pady=(12, 0))
        log_wrap.grid_columnconfigure(0, weight=1)
        log_wrap.grid_rowconfigure(1, weight=1)

        tk.Label(
            log_wrap,
            text="Process log",
            bg=BG_HEADER,
            fg="#d8e4ef",
            font=("Segoe UI Semibold", 10),
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.log_text = ScrolledText(
            log_wrap,
            height=9,
            wrap="word",
            font=("Consolas", 9),
            bg=LOG_BG,
            fg=LOG_FG,
            insertbackground=LOG_FG,
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=10,
            spacing1=3,
            spacing3=3,
        )
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")
        self.log_text.tag_configure("log", foreground=LOG_FG)
        self.log_text.tag_configure("status", foreground=LOG_INFO)
        self.log_text.tag_configure("done", foreground=LOG_DONE)
        self.log_text.tag_configure("error", foreground=LOG_ERROR)
        return panel

    def _append_log(self, message: str, kind: str = "log") -> None:
        if self.log_text is None:
            return
        timestamp = datetime.now().strftime("%H:%M:%S")
        tag = kind if kind in {"log", "status", "done", "error"} else "log"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_input_change(self, *_args) -> None:
        self._refresh_header_badge()

    def _refresh_header_badge(self) -> None:
        year_text = self.year_var.get().strip()

        if year_text.isdigit():
            self.year_badge_var.set(f"Target year {year_text}")
        else:
            self.year_badge_var.set("Target year not set")

    def _pick_template(self) -> None:
        selected = filedialog.askopenfilename(
            title="Choose template Excel file",
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            self.template_var.set(selected)
            self._refresh_default_output()

    def _pick_output(self) -> None:
        selected = filedialog.askdirectory(title="Choose output folder")
        if selected:
            self.output_var.set(selected)

    def _use_current_year(self) -> None:
        self.year_var.set(str(default_year()))
        self._refresh_default_output()

    def _refresh_default_output(self) -> None:
        template_text = self.template_var.get().strip()
        year_text = self.year_var.get().strip()
        if not template_text or not year_text.isdigit():
            return

        current = self.output_var.get().strip()
        template_path = Path(template_text)
        suggested = default_output_dir(template_path, int(year_text))
        if not current or current == str(default_output_dir(template_path, default_year())):
            self.output_var.set(str(suggested))

    def _validate_inputs(self) -> tuple[Path, Path, int] | None:
        template_text = self.template_var.get().strip()
        output_text = self.output_var.get().strip()
        year_text = self.year_var.get().strip()

        if not template_text:
            messagebox.showerror("Missing template", "Please choose a template Excel file.")
            return None
        if not output_text:
            messagebox.showerror("Missing output folder", "Please choose an output folder.")
            return None
        if not year_text.isdigit():
            messagebox.showerror("Invalid year", "Please enter a numeric year.")
            return None

        template = Path(template_text)
        output_dir = Path(output_text)
        if not template.exists():
            messagebox.showerror("Template not found", f"File was not found:\n{template}")
            return None

        return template, output_dir, int(year_text)

    def _set_running(self, running: bool) -> None:
        if self.generate_button is None or self.progress is None:
            return
        if running:
            self.generate_button.configure(state="disabled")
            self.progress.start(10)
        else:
            self.generate_button.configure(state="normal")
            self.progress.stop()

    def _start_generation(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        validated = self._validate_inputs()
        if validated is None:
            return

        template, output_dir, year = validated
        save_state(
            AppState(
                template_path=str(template),
                output_dir=str(output_dir),
                year=str(year),
            )
        )

        existing_files = sorted(output_dir.glob("*.xls*")) if output_dir.exists() else []
        if existing_files:
            overwrite = messagebox.askyesno(
                "Overwrite files?",
                (
                    "The output folder already contains Excel files.\n"
                    "The generator may overwrite files with the same names.\n\n"
                    "Continue?"
                ),
            )
            if not overwrite:
                return

        self.generated_dir = output_dir
        if self.open_button is not None:
            self.open_button.configure(state="disabled")
        self.status_var.set("Checking Microsoft Excel...")
        self._append_log("Starting generation...", "status")
        self._set_running(True)

        def worker() -> None:
            try:
                self.log_queue.put(("status", "Checking Microsoft Excel..."))
                ensure_excel_available()
                self.log_queue.put(("status", "Generating Jan-Dec workbooks..."))
                results = generate_year_files(
                    template_path=template,
                    output_dir=output_dir,
                    year=year,
                    log=lambda message: self.log_queue.put(("log", message)),
                )
                self.log_queue.put(
                    ("done", f"Generated {len(results)} file(s) in {output_dir}")
                )
            except GenerationError as exc:
                self.log_queue.put(("error", str(exc)))
            except Exception as exc:  # pragma: no cover
                self.log_queue.put(("error", f"Unexpected error: {exc}"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _poll_queue(self) -> None:
        while True:
            try:
                kind, message = self.log_queue.get_nowait()
            except queue.Empty:
                break

            self._append_log(message, kind)
            if kind == "status":
                self.status_var.set(message)
            elif kind == "done":
                self.status_var.set(message)
                self._set_running(False)
                if self.open_button is not None:
                    self.open_button.configure(state="normal")
                messagebox.showinfo("Finished", message)
            elif kind == "error":
                self.status_var.set("Failed.")
                self._set_running(False)
                messagebox.showerror("Generation failed", message)

        self.root.after(150, self._poll_queue)

    def _open_generated_folder(self) -> None:
        if self.generated_dir is None:
            return
        self.generated_dir.mkdir(parents=True, exist_ok=True)
        open_folder(self.generated_dir)


def run_ui() -> None:
    root = Tk()
    PMPlanApp(root)
    root.mainloop()
