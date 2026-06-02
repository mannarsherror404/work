from pathlib import Path
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from transformer import transform_forward, transform_reverse
from validator import ValidationError


APP_NAME = "Data Mask Transformer"


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()


def get_default_output_dir() -> Path:
    output_dir = APP_DIR / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def resource_path(relative_path: str) -> Path:
    if getattr(sys, "frozen", False):
        base_path = Path(getattr(sys, "_MEIPASS"))
    else:
        base_path = APP_DIR
    return base_path / relative_path


BG_COLOR = "#edf3f7"
CARD_COLOR = "#ffffff"
HEADER_COLOR = "#131f33"
HEADER_ACCENT = "#d9dcff"
ACCENT_COLOR = "#3f46d8"
ACCENT_DARK = "#27309f"
TEXT_COLOR = "#1d2938"
MUTED_COLOR = "#667085"
SUCCESS_COLOR = "#168464"
FIELD_BG = "#f8fafc"
FIELD_BORDER = "#d8dee8"
BUTTON_SOFT = "#eef0ff"
FONT_FAMILY = "Segoe UI" if sys.platform.startswith("win") else "Helvetica"


class DataTransformerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title(APP_NAME)
        self.geometry("980x700")
        self.minsize(920, 660)
        self.configure(background=BG_COLOR)

        self.input_path = tk.StringVar()
        self.operation = tk.StringVar(value="forward")
        self.output_path = tk.StringVar(value="No output created yet.")
        self.status_text = tk.StringVar(value="Ready")

        self._configure_style()
        self._set_window_icon()
        self._build_layout()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("App.TFrame", background=BG_COLOR)
        style.configure("Panel.TFrame", background=CARD_COLOR, relief="flat")
        style.configure("Header.TFrame", background=HEADER_COLOR)
        style.configure(
            "Title.TLabel",
            background=HEADER_COLOR,
            foreground="#ffffff",
            font=(FONT_FAMILY, 25, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=HEADER_COLOR,
            foreground=HEADER_ACCENT,
            font=(FONT_FAMILY, 11),
        )
        style.configure(
            "Body.TLabel",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            font=(FONT_FAMILY, 12, "bold"),
        )
        style.configure(
            "Muted.TLabel",
            background=CARD_COLOR,
            foreground=MUTED_COLOR,
            font=(FONT_FAMILY, 10),
        )
        style.configure(
            "Status.TLabel",
            background=CARD_COLOR,
            foreground=SUCCESS_COLOR,
            font=(FONT_FAMILY, 12, "bold"),
        )

    def _set_window_icon(self) -> None:
        icon_path = resource_path("assets/DataMaskTransformer.png")
        if icon_path.exists():
            try:
                self.icon_image = tk.PhotoImage(file=icon_path)
                self.iconphoto(True, self.icon_image)
            except tk.TclError:
                pass

    def _build_layout(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=28, style="App.TFrame")
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        header_card = RoundedCard(main, radius=28, fill=HEADER_COLOR, shadow="#d3dde8")
        header_card.grid(row=0, column=0, sticky="ew")
        header_card.inner.columnconfigure(0, weight=1)

        header = ttk.Frame(header_card.inner, padding=(34, 30, 34, 32), style="Header.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Create vendor-safe Excel files and restore them using the saved mapping.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        panel_card = RoundedCard(main, radius=28, fill=CARD_COLOR, shadow="#d9e2ec")
        panel_card.grid(row=1, column=0, sticky="nsew", pady=(22, 0))
        panel_card.inner.columnconfigure(0, weight=1)
        panel_card.inner.rowconfigure(0, weight=1)

        panel = ttk.Frame(panel_card.inner, padding=34, style="Panel.TFrame")
        panel.grid(row=0, column=0, sticky="nsew")
        panel.columnconfigure(1, weight=1)

        ttk.Label(panel, text="Input file", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        RoundedTextField(panel, textvariable=self.input_path).grid(
            row=0, column=1, sticky="ew", padx=(18, 12), pady=12
        )
        RoundedButton(
            panel,
            text="Browse",
            command=self._choose_input_file,
            fill=BUTTON_SOFT,
            hover="#dfe3ff",
            foreground=ACCENT_DARK,
            width=130,
        ).grid(row=0, column=2, sticky="ew", pady=12)

        ttk.Label(panel, text="Operation", style="Body.TLabel").grid(row=1, column=0, sticky="w")
        operation_frame = ttk.Frame(panel, style="Panel.TFrame")
        operation_frame.grid(row=1, column=1, sticky="w", padx=(18, 0), pady=16)

        self.forward_option = RoundedToggle(
            operation_frame,
            text="Export vendor file",
            value="forward",
            variable=self.operation,
            command=self._sync_operation_buttons,
            selected=True,
            width=185,
        )
        self.forward_option.grid(row=0, column=0, sticky="w", padx=(0, 12))

        self.reverse_option = RoundedToggle(
            operation_frame,
            text="Reverse vendor file",
            value="reverse",
            variable=self.operation,
            command=self._sync_operation_buttons,
            width=195,
        )
        self.reverse_option.grid(row=0, column=1, sticky="w")

        actions = ttk.Frame(panel, style="Panel.TFrame")
        actions.grid(row=2, column=1, sticky="w", padx=(18, 0), pady=(22, 22))

        self.run_button = RoundedButton(
            actions,
            text="Run transformation",
            command=self._run_clicked,
            fill=ACCENT_COLOR,
            hover=ACCENT_DARK,
            foreground="#ffffff",
            width=225,
            font_weight="bold",
        )
        self.run_button.grid(row=0, column=0, sticky="w")

        RoundedButton(
            actions,
            text="Open output folder",
            command=self._open_output_folder,
            fill=BUTTON_SOFT,
            hover="#dfe3ff",
            foreground=ACCENT_DARK,
            width=195,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))

        ttk.Separator(panel).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 24))

        ttk.Label(panel, text="Status", style="Body.TLabel").grid(row=4, column=0, sticky="w")
        status_box = ttk.Frame(panel, style="Panel.TFrame")
        status_box.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(18, 0), pady=10)
        status_box.columnconfigure(0, weight=1)
        ttk.Label(status_box, textvariable=self.status_text, style="Status.TLabel").grid(
            row=0, column=0, sticky="ew"
        )

        ttk.Label(panel, text="Output", style="Body.TLabel").grid(row=5, column=0, sticky="w")
        RoundedTextField(panel, textvariable=self.output_path, readonly=True).grid(
            row=5, column=1, columnspan=2, sticky="ew", padx=(18, 0), pady=10
        )

        notes = ttk.Label(
            panel,
            text="Forward creates a saved mapping automatically. Reverse uses the mapping saved beside the vendor file.",
            style="Muted.TLabel",
        )
        notes.grid(row=6, column=0, columnspan=3, sticky="w", pady=(24, 0))

    def _sync_operation_buttons(self) -> None:
        self.forward_option.set_selected(self.operation.get() == "forward")
        self.reverse_option.set_selected(self.operation.get() == "reverse")

    def _choose_input_file(self) -> None:
        initial_dir = APP_DIR / "input"
        if getattr(sys, "frozen", False) or not initial_dir.exists():
            initial_dir = Path.home() / "Desktop"

        path = filedialog.askopenfilename(
            title="Select input file",
            initialdir=str(initial_dir),
            filetypes=[
                ("Excel or CSV files", "*.xlsx *.xls *.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.input_path.set(path)

    def _run_clicked(self) -> None:
        input_file = self.input_path.get().strip()

        if not input_file:
            messagebox.showerror("Missing input file", "Please select an input file.")
            return

        self.run_button.set_enabled(False)
        self.status_text.set("Running transformation...")
        self.output_path.set("Working...")

        worker = threading.Thread(
            target=self._run_transform,
            args=(input_file, self.operation.get()),
            daemon=True,
        )
        worker.start()

    def _run_transform(self, input_file: str, operation: str) -> None:
        try:
            output_dir = get_default_output_dir()

            if operation == "forward":
                output = transform_forward(input_file, output_dir=output_dir)
            else:
                output = transform_reverse(input_file, output_dir=output_dir)

        except ValidationError as exc:
            self.after(0, self._show_error, "Validation error", str(exc))
        except Exception as exc:
            self.after(0, self._show_error, "Unexpected error", str(exc))
        else:
            self.after(0, self._show_success, output)

    def _show_success(self, output: Path) -> None:
        self.run_button.set_enabled(True)
        self.status_text.set("Done")
        self.output_path.set(str(output))
        messagebox.showinfo("Transformation complete", f"Output created:\n{output}")

    def _show_error(self, title: str, message: str) -> None:
        self.run_button.set_enabled(True)
        self.status_text.set("Error")
        self.output_path.set("No output created.")
        messagebox.showerror(title, message)

    def _open_output_folder(self) -> None:
        output_dir = get_default_output_dir()

        if sys.platform.startswith("win"):
            os.startfile(output_dir)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(output_dir)], check=False)
        else:
            subprocess.run(["xdg-open", str(output_dir)], check=False)


class RoundedCard(tk.Canvas):
    def __init__(
        self,
        parent: tk.Misc,
        radius: int = 20,
        fill: str = CARD_COLOR,
        shadow: str = "#d8e0e8",
    ) -> None:
        super().__init__(
            parent,
            background=BG_COLOR,
            highlightthickness=0,
            borderwidth=0,
            relief="flat",
        )
        self.radius = radius
        self.fill = fill
        self.shadow = shadow
        self.inner = tk.Frame(self, background=fill, borderwidth=0, highlightthickness=0)
        self.window_id = self.create_window(0, 0, window=self.inner, anchor="nw")
        self.bind("<Configure>", self._draw)

    def _draw(self, event: tk.Event) -> None:
        width = event.width
        height = event.height
        self.delete("card")
        self._rounded_rect(7, 9, width - 1, height - 1, self.radius, fill=self.shadow, tags="card")
        self._rounded_rect(0, 0, width - 10, height - 12, self.radius, fill=self.fill, tags="card")
        self.coords(self.window_id, 0, 0)
        self.itemconfigure(self.window_id, width=width - 10, height=height - 12)
        self.tag_lower("card")

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: object) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self.create_polygon(points, smooth=True, splinesteps=18, outline="", **kwargs)


class RoundedButton(tk.Canvas):
    def __init__(
        self,
        parent: tk.Misc,
        text: str,
        command: Callable[[], None],
        fill: str,
        hover: str,
        foreground: str,
        width: int = 180,
        height: int = 48,
        radius: int = 18,
        font_weight: str = "normal",
    ) -> None:
        super().__init__(
            parent,
            width=width,
            height=height,
            background=CARD_COLOR,
            highlightthickness=0,
            borderwidth=0,
        )
        self.command = command
        self.fill = fill
        self.hover = hover
        self.foreground = foreground
        self.radius = radius
        self.enabled = True
        self.text = text
        self.font = (FONT_FAMILY, 11, font_weight)
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _event: self._draw(self.hover))
        self.bind("<Leave>", lambda _event: self._draw(self.fill))
        self._draw(self.fill)

    def set_enabled(self, enabled: bool) -> None:
        self.enabled = enabled
        self._draw("#d7dce5" if not enabled else self.fill)

    def _click(self, _event: tk.Event) -> None:
        if self.enabled:
            self.command()

    def _draw(self, fill: str) -> None:
        self.delete("all")
        self._rounded_rect(0, 0, int(self["width"]), int(self["height"]), self.radius, fill=fill)
        self.create_text(
            int(self["width"]) // 2,
            int(self["height"]) // 2,
            text=self.text,
            fill=self.foreground if self.enabled else MUTED_COLOR,
            font=self.font,
        )

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: object) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self.create_polygon(points, smooth=True, splinesteps=18, outline="", **kwargs)


class RoundedToggle(RoundedButton):
    def __init__(
        self,
        parent: tk.Misc,
        text: str,
        value: str,
        variable: tk.StringVar,
        command: Callable[[], None],
        selected: bool = False,
        width: int = 180,
    ) -> None:
        self.value = value
        self.variable = variable
        self.external_command = command
        super().__init__(
            parent,
            text=text,
            command=self._select,
            fill=ACCENT_COLOR if selected else FIELD_BG,
            hover=ACCENT_DARK if selected else "#edf0ff",
            foreground="#ffffff" if selected else TEXT_COLOR,
            width=width,
            height=44,
            radius=18,
            font_weight="bold" if selected else "normal",
        )

    def _select(self) -> None:
        self.variable.set(self.value)
        self.external_command()

    def set_selected(self, selected: bool) -> None:
        self.fill = ACCENT_COLOR if selected else FIELD_BG
        self.hover = ACCENT_DARK if selected else "#edf0ff"
        self.foreground = "#ffffff" if selected else TEXT_COLOR
        self.font = (FONT_FAMILY, 11, "bold" if selected else "normal")
        self._draw(self.fill)


class RoundedTextField(tk.Canvas):
    def __init__(self, parent: tk.Misc, textvariable: tk.StringVar, readonly: bool = False) -> None:
        super().__init__(
            parent,
            height=48,
            background=CARD_COLOR,
            highlightthickness=0,
            borderwidth=0,
        )
        self.textvariable = textvariable
        self.readonly = readonly
        self.entry = tk.Entry(
            self,
            textvariable=textvariable,
            relief="flat",
            borderwidth=0,
            background=FIELD_BG,
            foreground=TEXT_COLOR,
            font=(FONT_FAMILY, 12),
            readonlybackground=FIELD_BG,
        )
        if readonly:
            self.entry.configure(state="readonly")
        self.window_id = self.create_window(16, 24, window=self.entry, anchor="w")
        self.bind("<Configure>", self._draw)

    def _draw(self, event: tk.Event) -> None:
        width = event.width
        self.delete("field")
        self._rounded_rect(0, 0, width, 48, 18, fill=FIELD_BG, outline=FIELD_BORDER, tags="field")
        self.coords(self.window_id, 16, 24)
        self.itemconfigure(self.window_id, width=max(40, width - 32), height=26)
        self.tag_lower("field")

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: object) -> None:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        self.create_polygon(points, smooth=True, splinesteps=18, **kwargs)


def run_app() -> None:
    app = DataTransformerApp()
    app.mainloop()
