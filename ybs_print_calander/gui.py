"""Tkinter GUI for the YBS Print Calander application."""

from __future__ import annotations

import calendar
import datetime as dt
import json
import queue
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import ttk
from typing import Dict, Iterable, List, Tuple

from .client import AuthenticationError, NetworkError, OrderRecord, YBSClient

BACKGROUND_COLOR = "#0b1d3a"
ACCENT_COLOR = "#1f3a63"
TEXT_COLOR = "#f8f9fa"
SUCCESS_COLOR = "#28a745"
FAIL_COLOR = "#dc3545"
PENDING_COLOR = "#f0ad4e"
DAY_CELL_BACKGROUND = "#102a54"
DAY_CELL_HOVER_VALID = "#25497a"
DAY_CELL_HOVER_INVALID = "#5a1f1f"
TODAY_BORDER_COLOR = "#f6c343"
TODAY_HEADER_BACKGROUND = "#1d4e89"
ASSIGNMENT_HEADER_BACKGROUND = "#2561a8"
ASSIGNMENT_LIST_BACKGROUND = "#17406d"
NOTES_TEXT_BACKGROUND = "#0f3460"
ORDERS_LIST_BACKGROUND = "#0d274a"

DRAG_THRESHOLD = 5

DateKey = Tuple[int, int, int]


STATE_PATH = Path.home() / ".ybs_print_calander" / "state.json"


@dataclass
class DayCell:
    """Container for widgets that make up a calendar day cell."""

    frame: tk.Frame
    header_label: tk.Label
    notes_text: tk.Text
    orders_list: tk.Listbox
    default_bg: str
    border_color: str = ACCENT_COLOR
    border_thickness: int = 1
    is_today: bool = False


class YBSApp:
    """Encapsulates the Tkinter application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("YBS Print Calander")
        self.root.configure(background=BACKGROUND_COLOR)
        self.root.geometry("720x480")

        self.client = YBSClient()
        self._queue: "queue.Queue[tuple[str, object]]" = queue.Queue()

        today = dt.date.today()
        self._current_year = today.year
        self._current_month = today.month

        self._day_cells: Dict[DateKey, DayCell] = {}
        self._calendar_notes: Dict[DateKey, str] = {}
        self._calendar_assignments: Dict[DateKey, List[Tuple[str, str]]] = {}
        self._calendar_hover: DateKey | None = None
        self._drag_data: dict[str, object] = {}
        self._state_path: Path = STATE_PATH
        self._state_path.parent.mkdir(parents=True, exist_ok=True)
        self._state_save_after_id: str | None = None

        self._load_state()
        self._reset_drag_state()

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.last_refresh_var = tk.StringVar(value="")
        self._order_filter_var = tk.StringVar()
        self.month_label_var = tk.StringVar(value=today.strftime("%B %Y"))
        self._all_orders: list[OrderRecord] = []

        self._configure_style()
        self._build_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_queue()

    @staticmethod
    def _serialize_date_key(date_key: DateKey) -> str:
        try:
            year, month, day = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            return f"{date_key[0]}-{date_key[1]}-{date_key[2]}"
        return f"{year:04d}-{month:02d}-{day:02d}"

    @staticmethod
    def _deserialize_date_key(key: str) -> DateKey | None:
        if not isinstance(key, str):
            return None
        parts = key.split("-")
        if len(parts) != 3:
            return None
        try:
            return (int(parts[0]), int(parts[1]), int(parts[2]))
        except (TypeError, ValueError):
            return None

    def _load_state(self) -> None:
        notes: Dict[DateKey, str] = {}
        assignments: Dict[DateKey, List[Tuple[str, str]]] = {}

        data: object | None = None
        try:
            with self._state_path.open("r", encoding="utf-8") as state_file:
                data = json.load(state_file)
        except FileNotFoundError:
            data = None
        except (OSError, json.JSONDecodeError):
            data = None

        if isinstance(data, dict):
            raw_notes = data.get("notes", {})
            if isinstance(raw_notes, dict):
                for key_str, value in raw_notes.items():
                    date_key = self._deserialize_date_key(key_str)
                    if date_key is None or not isinstance(value, str):
                        continue
                    notes[date_key] = value

            raw_assignments = data.get("assignments", {})
            if isinstance(raw_assignments, dict):
                for key_str, value in raw_assignments.items():
                    date_key = self._deserialize_date_key(key_str)
                    if date_key is None or not isinstance(value, list):
                        continue

                    normalized_assignments: List[Tuple[str, str]] = []
                    for entry in value:
                        if isinstance(entry, (list, tuple)):
                            first = str(entry[0]) if len(entry) > 0 else ""
                            second = str(entry[1]) if len(entry) > 1 else ""
                            normalized_assignments.append((first, second))
                        elif isinstance(entry, str):
                            normalized_assignments.append((entry, ""))

                    if normalized_assignments:
                        assignments[date_key] = normalized_assignments

        self._calendar_notes = notes
        self._calendar_assignments = assignments

    def _save_state(self) -> None:
        self._state_save_after_id = None

        notes = {
            self._serialize_date_key(key): value
            for key, value in self._calendar_notes.items()
        }
        assignments = {
            self._serialize_date_key(key): [list(item) for item in value]
            for key, value in self._calendar_assignments.items()
        }

        state = {"notes": notes, "assignments": assignments}

        try:
            self._state_path.parent.mkdir(parents=True, exist_ok=True)
            with self._state_path.open("w", encoding="utf-8") as state_file:
                json.dump(state, state_file, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _schedule_state_save(self) -> None:
        if self._state_save_after_id is not None:
            try:
                self.root.after_cancel(self._state_save_after_id)
            except tk.TclError:
                pass

        try:
            self._state_save_after_id = self.root.after(1000, self._save_state)
        except tk.TclError:
            self._state_save_after_id = None

    def _on_close(self) -> None:
        if self._state_save_after_id is not None:
            try:
                self.root.after_cancel(self._state_save_after_id)
            except tk.TclError:
                pass
            self._state_save_after_id = None

        self._save_state()

        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _configure_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:  # pragma: no cover - fallback path
            pass

        style.configure("Dark.TFrame", background=BACKGROUND_COLOR)
        style.configure("Dark.TLabel", background=BACKGROUND_COLOR, foreground=TEXT_COLOR)
        style.configure(
            "Dark.TLabelframe",
            background=BACKGROUND_COLOR,
            foreground=TEXT_COLOR,
            bordercolor=ACCENT_COLOR,
            borderwidth=1,
        )
        style.configure(
            "Dark.TLabelframe.Label",
            background=BACKGROUND_COLOR,
            foreground=TEXT_COLOR,
        )
        style.configure("Dark.TPanedwindow", background=BACKGROUND_COLOR)
        style.configure(
            "Dark.TButton",
            background=ACCENT_COLOR,
            foreground=TEXT_COLOR,
            borderwidth=0,
            focusthickness=3,
            focuscolor=ACCENT_COLOR,
            padding=6,
        )
        style.map("Dark.TButton", background=[("active", "#25497a")])

        style.configure(
            "Dark.Treeview",
            background="#102a54",
            foreground=TEXT_COLOR,
            fieldbackground="#102a54",
            bordercolor=ACCENT_COLOR,
            borderwidth=1,
            rowheight=26,
        )
        style.map("Dark.Treeview", background=[("selected", "#1e90ff")])
        style.configure(
            "Dark.Treeview.Heading",
            background=ACCENT_COLOR,
            foreground=TEXT_COLOR,
            bordercolor=ACCENT_COLOR,
            relief="flat",
            padding=6,
        )
        style.map("Dark.Treeview.Heading", background=[("active", "#25497a")])

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, style="Dark.TFrame")
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        login_frame = ttk.Frame(container, style="Dark.TFrame")
        login_frame.pack(fill=tk.X, pady=(0, 20))

        username_label = ttk.Label(login_frame, text="Username", style="Dark.TLabel")
        username_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        username_entry = ttk.Entry(login_frame, textvariable=self.username_var, width=30)
        username_entry.grid(row=0, column=1, sticky=tk.W)
        username_entry.focus_set()

        password_label = ttk.Label(login_frame, text="Password", style="Dark.TLabel")
        password_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        password_entry = ttk.Entry(login_frame, textvariable=self.password_var, show="*", width=30)
        password_entry.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        username_entry.bind("<Return>", self._on_enter_pressed)
        password_entry.bind("<Return>", self._on_enter_pressed)

        button_frame = ttk.Frame(login_frame, style="Dark.TFrame")
        button_frame.grid(row=0, column=2, rowspan=2, padx=(20, 0), sticky=tk.N)

        self.status_canvas = tk.Canvas(
            button_frame,
            width=20,
            height=20,
            highlightthickness=0,
            bg=BACKGROUND_COLOR,
            bd=0,
        )
        self.status_canvas.pack(pady=(0, 8))
        self.status_light = self.status_canvas.create_oval(2, 2, 18, 18, fill="#555555", outline="")

        self.login_button = ttk.Button(
            button_frame,
            text="Login",
            style="Dark.TButton",
            command=self._on_login_clicked,
        )
        self.login_button.pack()

        self.refresh_button = ttk.Button(
            button_frame,
            text="Refresh",
            style="Dark.TButton",
            command=self._on_refresh_clicked,
            state=tk.DISABLED,
        )
        self.refresh_button.pack(pady=(8, 0))

        self.status_message = ttk.Label(login_frame, text="", style="Dark.TLabel")
        self.status_message.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))

        self.last_refresh_label = ttk.Label(
            login_frame,
            textvariable=self.last_refresh_var,
            style="Dark.TLabel",
            font=("TkDefaultFont", 8),
        )
        self.last_refresh_label.grid(
            row=3, column=0, columnspan=3, sticky=tk.W, pady=(2, 0)
        )

        content_paned = ttk.Panedwindow(container, orient=tk.HORIZONTAL, style="Dark.TPanedwindow")
        content_paned.pack(fill=tk.BOTH, expand=True)

        table_frame = ttk.LabelFrame(
            content_paned,
            text="Orders",
            style="Dark.TLabelframe",
            padding=10,
        )
        table_frame.configure(labelanchor="n")

        filter_label = ttk.Label(table_frame, text="Filter", style="Dark.TLabel")
        filter_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        filter_entry = ttk.Entry(table_frame, textvariable=self._order_filter_var)
        filter_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        filter_entry.bind("<KeyRelease>", self._apply_order_filter)

        # Treeview and scrollbar
        self.tree = ttk.Treeview(
            table_frame,
            columns=("order", "company"),
            show="headings",
            style="Dark.Treeview",
        )
        self.tree.heading("order", text="Order#", anchor="center")
        self.tree.heading("company", text="Company", anchor="center")
        self.tree.column("order", anchor="center", width=120, stretch=False)
        self.tree.column("company", anchor="center", width=400)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=2, column=0, sticky="nsew")
        scrollbar.grid(row=2, column=1, sticky="ns")

        self.tree.bind("<ButtonPress-1>", self._on_order_press)
        self.tree.bind("<B1-Motion>", self._on_order_drag)
        self.tree.bind("<ButtonRelease-1>", self._on_order_release)

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(2, weight=1)

        content_paned.add(table_frame, weight=3)

        calendar_frame = ttk.LabelFrame(
            content_paned,
            text="Calendar",
            style="Dark.TLabelframe",
            padding=10,
        )
        calendar_frame.configure(labelanchor="n")

        header_frame = ttk.Frame(calendar_frame, style="Dark.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header_frame.columnconfigure(2, weight=1)

        previous_button = ttk.Button(
            header_frame,
            text="Previous",
            style="Dark.TButton",
            command=lambda: self._change_month(-1),
        )
        previous_button.grid(row=0, column=0, padx=(0, 10))

        today_button = ttk.Button(
            header_frame,
            text="Today",
            style="Dark.TButton",
            command=self._go_to_today,
        )
        today_button.grid(row=0, column=1, padx=(0, 10))

        self.month_label = ttk.Label(
            header_frame,
            textvariable=self.month_label_var,
            style="Dark.TLabel",
        )
        self.month_label.grid(row=0, column=2, sticky="ew")

        next_button = ttk.Button(
            header_frame,
            text="Next",
            style="Dark.TButton",
            command=lambda: self._change_month(1),
        )
        next_button.grid(row=0, column=3, padx=(10, 0))

        self.calendar_grid = ttk.Frame(calendar_frame, style="Dark.TFrame")
        self.calendar_grid.grid(row=1, column=0, sticky="nsew")

        for column_index in range(7):
            self.calendar_grid.columnconfigure(column_index, weight=1, uniform="calendar")
        self.calendar_grid.rowconfigure(0, weight=0)

        calendar_frame.columnconfigure(0, weight=1)
        calendar_frame.rowconfigure(1, weight=1)

        self._render_calendar()

        content_paned.add(calendar_frame, weight=2)

    def _render_calendar(self) -> None:
        year = self._current_year
        month = self._current_month
        today = dt.date.today()
        today_key = (today.year, today.month, today.day)
        first_of_month = dt.date(year, month, 1)
        self.month_label_var.set(first_of_month.strftime("%B %Y"))

        self._remove_calendar_hover()

        for date_key, day_cell in list(self._day_cells.items()):
            self._save_day_notes(date_key)
            day_cell.frame.destroy()

        self._day_cells.clear()

        for child in self.calendar_grid.winfo_children():
            child.destroy()

        month_structure = calendar.Calendar().monthdayscalendar(year, month)

        for column_index in range(7):
            header_label = ttk.Label(
                self.calendar_grid,
                text=calendar.day_abbr[column_index],
                style="Dark.TLabel",
                anchor="center",
            )
            header_label.grid(row=0, column=column_index, sticky="nsew", padx=2, pady=(0, 6))

        for row_index, week in enumerate(month_structure, start=1):
            self.calendar_grid.rowconfigure(row_index, weight=1, uniform="calendar_rows")
            for column_index, day in enumerate(week):
                if day == 0:
                    placeholder = tk.Frame(
                        self.calendar_grid,
                        bg=BACKGROUND_COLOR,
                        bd=0,
                        highlightthickness=0,
                    )
                    placeholder.grid(
                        row=row_index, column=column_index, sticky="nsew", padx=2, pady=2
                    )
                    continue

                date_key = (year, month, day)
                is_today = date_key == today_key
                border_color = TODAY_BORDER_COLOR if is_today else ACCENT_COLOR
                border_thickness = 2 if is_today else 1
                cell_frame = tk.Frame(
                    self.calendar_grid,
                    bg=DAY_CELL_BACKGROUND,
                    highlightbackground=border_color,
                    highlightcolor=border_color,
                    highlightthickness=border_thickness,
                    bd=0,
                )
                cell_frame.grid(row=row_index, column=column_index, sticky="nsew", padx=2, pady=2)
                cell_frame.grid_propagate(False)
                cell_frame.configure(width=110, height=110)
                cell_frame.columnconfigure(0, weight=1)
                cell_frame.rowconfigure(1, weight=1)
                cell_frame.rowconfigure(2, weight=1)

                header_label = tk.Label(
                    cell_frame,
                    text=str(day),
                    anchor="nw",
                    bg=DAY_CELL_BACKGROUND,
                    fg=TEXT_COLOR,
                    font=("TkDefaultFont", 10, "bold"),
                    padx=4,
                    pady=2,
                )
                header_label.grid(row=0, column=0, sticky="ew")

                notes_text = tk.Text(
                    cell_frame,
                    height=3,
                    wrap=tk.WORD,
                    bg=NOTES_TEXT_BACKGROUND,
                    fg=TEXT_COLOR,
                    insertbackground=TEXT_COLOR,
                    relief="flat",
                    bd=0,
                )
                notes_text.grid(row=1, column=0, sticky="nsew", padx=4, pady=(2, 2))

                orders_list = tk.Listbox(
                    cell_frame,
                    height=3,
                    activestyle="none",
                    exportselection=False,
                )
                orders_list.configure(
                    bg=ORDERS_LIST_BACKGROUND,
                    fg=TEXT_COLOR,
                    highlightbackground=ACCENT_COLOR,
                    highlightcolor=ACCENT_COLOR,
                    selectbackground="#1e90ff",
                    selectforeground=TEXT_COLOR,
                    relief="flat",
                    bd=0,
                )
                orders_list.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 4))

                day_cell = DayCell(
                    frame=cell_frame,
                    header_label=header_label,
                    notes_text=notes_text,
                    orders_list=orders_list,
                    default_bg=DAY_CELL_BACKGROUND,
                )
                day_cell.border_color = border_color
                day_cell.border_thickness = border_thickness
                day_cell.is_today = is_today
                self._day_cells[date_key] = day_cell

                existing_notes = self._calendar_notes.get(date_key, "")
                if existing_notes:
                    notes_text.insert("1.0", existing_notes)

                notes_text.bind(
                    "<FocusOut>",
                    lambda event, key=date_key: self._save_day_notes(key),
                )

                cell_frame.bind(
                    "<Double-Button-1>",
                    lambda event, key=date_key: self._open_day_details(key),
                )
                notes_text.bind(
                    "<Double-Button-1>",
                    lambda event, key=date_key: self._open_day_details(key),
                )
                orders_list.bind(
                    "<Double-Button-1>",
                    lambda event, key=date_key: self._open_day_details(key),
                )
                orders_list.bind(
                    "<Delete>",
                    lambda event, key=date_key: self._on_day_order_delete(event, key),
                )
                orders_list.bind(
                    "<ButtonPress-1>",
                    lambda event, key=date_key: self._on_day_order_press(event, key),
                )
                orders_list.bind(
                    "<B1-Motion>",
                    lambda event, key=date_key: self._on_day_order_drag(event, key),
                )
                orders_list.bind(
                    "<ButtonRelease-1>",
                    lambda event, key=date_key: self._on_day_order_release(event, key),
                )

                self._update_day_cell_display(date_key)

        self._calendar_hover = None

    def _apply_day_cell_base_style(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        border_color = getattr(day_cell, "border_color", ACCENT_COLOR)
        border_thickness = getattr(day_cell, "border_thickness", 1)
        day_cell.frame.configure(
            bg=day_cell.default_bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=border_thickness,
        )

        assignments = self._calendar_assignments.get(date_key, [])
        has_assignments = bool(assignments)

        if has_assignments:
            header_bg = ASSIGNMENT_HEADER_BACKGROUND
            orders_bg = ASSIGNMENT_LIST_BACKGROUND
        else:
            header_bg = (
                TODAY_HEADER_BACKGROUND
                if getattr(day_cell, "is_today", False)
                else day_cell.default_bg
            )
            orders_bg = ORDERS_LIST_BACKGROUND

        day_cell.header_label.configure(bg=header_bg, fg=TEXT_COLOR)
        day_cell.orders_list.configure(bg=orders_bg, fg=TEXT_COLOR)

    def _save_day_notes(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        text_value = day_cell.notes_text.get("1.0", "end-1c")
        changed = False
        if text_value.strip():
            if self._calendar_notes.get(date_key) != text_value:
                self._calendar_notes[date_key] = text_value
                changed = True
        elif date_key in self._calendar_notes:
            self._calendar_notes.pop(date_key, None)
            changed = True

        if changed:
            self._schedule_state_save()

    def _on_day_order_delete(
        self, event: tk.Event | None, date_key: DateKey
    ) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        orders_list = day_cell.orders_list
        selection = orders_list.curselection()
        assignments = self._calendar_assignments.get(date_key, [])

        if not assignments:
            orders_list.selection_clear(0, tk.END)
            self._set_status(FAIL_COLOR, "No orders scheduled for this day.")
            return

        if not selection:
            self._set_status(FAIL_COLOR, "Please select an order to remove.")
            return

        index = selection[0]
        if index < 0 or index >= len(assignments):
            self._set_status(
                FAIL_COLOR, "Unable to determine which order to remove."
            )
            return

        removed_assignment = assignments.pop(index)
        if assignments:
            self._calendar_assignments[date_key] = assignments
        else:
            self._calendar_assignments.pop(date_key, None)

        orders_list.selection_clear(0, tk.END)
        self._update_day_cell_display(date_key)
        self._schedule_state_save()

        message = self._format_removal_message(date_key, removed_assignment)
        self._set_status(SUCCESS_COLOR, message)

    def _open_day_details(self, date_key: DateKey) -> None:
        window = tk.Toplevel(self.root)
        window.title(self._format_date_label(date_key))
        window.configure(bg=BACKGROUND_COLOR)
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, style="Dark.TFrame", padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        date_label = ttk.Label(
            frame,
            text=f"Orders for {self._format_date_label(date_key)}",
            style="Dark.TLabel",
        )
        date_label.grid(row=0, column=0, columnspan=2, sticky="w")

        info_var = tk.StringVar(value="")
        info_label = ttk.Label(frame, textvariable=info_var, style="Dark.TLabel")
        info_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 10))

        listbox = tk.Listbox(
            frame,
            height=8,
            selectmode=tk.SINGLE,
            activestyle="none",
            exportselection=False,
        )
        listbox.configure(
            bg="#102a54",
            fg=TEXT_COLOR,
            highlightbackground=ACCENT_COLOR,
            highlightcolor=ACCENT_COLOR,
            selectbackground="#1e90ff",
            relief="flat",
        )
        listbox.grid(row=2, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=2, column=1, sticky="ns")

        button_frame = ttk.Frame(frame, style="Dark.TFrame")
        button_frame.grid(row=3, column=0, columnspan=2, sticky="e", pady=(15, 0))

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        def update_button_states(*_: object) -> None:
            assignments = self._calendar_assignments.get(date_key, [])
            has_assignments = bool(assignments)
            info_var.set("" if has_assignments else "No orders scheduled for this day.")
            clear_state = tk.NORMAL if has_assignments else tk.DISABLED
            clear_button.config(state=clear_state)
            if has_assignments and listbox.curselection():
                remove_button.config(state=tk.NORMAL)
            else:
                remove_button.config(state=tk.DISABLED)

        def refresh_list(select_index: int | None = None) -> None:
            assignments = self._calendar_assignments.get(date_key, [])
            listbox.delete(0, tk.END)
            for assignment in assignments:
                listbox.insert(tk.END, self._format_assignment_label(assignment))
            listbox.selection_clear(0, tk.END)
            if (
                select_index is not None
                and 0 <= select_index < len(assignments)
                and assignments
            ):
                listbox.selection_set(select_index)
            update_button_states()

        def remove_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                self._set_status(FAIL_COLOR, "Please select an order to remove.")
                return

            index = selection[0]
            assignments = self._calendar_assignments.get(date_key, [])
            if not assignments:
                self._set_status(FAIL_COLOR, "No orders scheduled for this day.")
                refresh_list()
                return

            if index < 0 or index >= len(assignments):
                self._set_status(FAIL_COLOR, "Unable to determine which order to remove.")
                update_button_states()
                return

            removed_assignment = assignments.pop(index)
            if assignments:
                self._calendar_assignments[date_key] = assignments
            else:
                self._calendar_assignments.pop(date_key, None)

            self._update_day_cell_display(date_key)
            self._schedule_state_save()
            next_index = min(index, len(assignments) - 1)
            refresh_list(select_index=next_index if assignments else None)

            message = self._format_removal_message(date_key, removed_assignment)
            self._set_status(SUCCESS_COLOR, message)

        def clear_day() -> None:
            assignments = self._calendar_assignments.get(date_key, [])
            if not assignments:
                self._set_status(FAIL_COLOR, "No orders scheduled for this day.")
                refresh_list()
                return

            removed_count = len(assignments)
            self._calendar_assignments.pop(date_key, None)
            self._update_day_cell_display(date_key)
            self._schedule_state_save()
            refresh_list()

            date_label_text = self._format_date_label(date_key)
            plural = "s" if removed_count != 1 else ""
            message = f"Cleared {removed_count} order{plural} from {date_label_text}."
            self._set_status(SUCCESS_COLOR, message)

        def close_dialog() -> None:
            window.destroy()

        remove_button = ttk.Button(
            button_frame,
            text="Remove Selected",
            style="Dark.TButton",
            command=remove_selected,
        )
        remove_button.grid(row=0, column=0, padx=(0, 10))

        clear_button = ttk.Button(
            button_frame,
            text="Clear Day",
            style="Dark.TButton",
            command=clear_day,
        )
        clear_button.grid(row=0, column=1, padx=(0, 10))

        close_button = ttk.Button(
            button_frame,
            text="Close",
            style="Dark.TButton",
            command=close_dialog,
        )
        close_button.grid(row=0, column=2)

        listbox.bind("<<ListboxSelect>>", update_button_states)
        window.protocol("WM_DELETE_WINDOW", close_dialog)

        refresh_list()
        window.focus_set()

    def _change_month(self, delta_months: int) -> None:
        year = self._current_year
        month = self._current_month + delta_months

        while month < 1:
            month += 12
            year -= 1
        while month > 12:
            month -= 12
            year += 1

        self._current_year = year
        self._current_month = month
        self._remove_calendar_hover()
        self._render_calendar()

    def _go_to_today(self) -> None:
        today = dt.date.today()
        self._current_year = today.year
        self._current_month = today.month
        self._remove_calendar_hover()
        self._render_calendar()

    def _reset_drag_state(self) -> None:
        self._drag_data = {
            "item": None,
            "values": (),
            "start_x": 0,
            "start_y": 0,
            "widget": None,
            "active": False,
            "source": None,
            "source_date_key": None,
            "source_index": None,
            "source_assignment": None,
        }

    def _on_order_press(self, event: tk.Event) -> None:
        self._end_drag()
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            self.tree.selection_remove(self.tree.selection())
            return

        self.tree.selection_set(item_id)
        values = self.tree.item(item_id, "values") or ()
        if not isinstance(values, tuple):
            values = tuple(values)

        order_values = tuple(str(value) for value in values)

        self._drag_data.update(
            {
                "item": item_id,
                "values": order_values,
                "start_x": event.x_root,
                "start_y": event.y_root,
                "widget": None,
                "active": False,
                "source": "tree",
                "source_date_key": None,
                "source_index": None,
                "source_assignment": order_values,
            }
        )

    def _on_order_drag(self, event: tk.Event) -> None:
        item_id = self._drag_data.get("item")
        if not item_id:
            return

        if not self._drag_data.get("active"):
            start_x = int(self._drag_data.get("start_x", event.x_root))
            start_y = int(self._drag_data.get("start_y", event.y_root))
            if (
                abs(event.x_root - start_x) >= DRAG_THRESHOLD
                or abs(event.y_root - start_y) >= DRAG_THRESHOLD
            ):
                self._begin_drag()

        if not self._drag_data.get("active"):
            return

        self._position_drag_window(event.x_root, event.y_root)
        target_info = self._detect_calendar_target(event.x_root, event.y_root)
        self._update_calendar_hover(target_info)

    def _on_order_release(self, event: tk.Event) -> None:
        item_id = self._drag_data.get("item")
        if not item_id:
            return

        if not self._drag_data.get("active"):
            self._end_drag()
            return

        target_info = self._detect_calendar_target(event.x_root, event.y_root)
        normalized_key: DateKey | None = None
        if target_info:
            raw_key = target_info.get("date_key")
            if isinstance(raw_key, tuple) and len(raw_key) == 3:
                try:
                    normalized_key = (int(raw_key[0]), int(raw_key[1]), int(raw_key[2]))
                except (TypeError, ValueError):
                    normalized_key = None

        if normalized_key is not None:
            year, month, day_value = normalized_key
            values = self._drag_data.get("values", ())
            if not isinstance(values, (tuple, list)) or not values:
                self._queue.put(
                    (
                        "calendar_drop",
                        False,
                        "Unable to determine which order was dragged.",
                        None,
                    )
                )
            else:
                order_values = tuple(str(value) for value in values)
                order_number = order_values[0]
                company = order_values[1] if len(order_values) > 1 else ""
                message = f"Assigned order {order_number}"
                if company:
                    message += f" ({company})"
                try:
                    display_date = dt.date(year, month, day_value)
                except ValueError:
                    display_date = None
                if display_date is not None:
                    message += f" to {display_date.strftime('%B %d, %Y')}."
                else:
                    message += f" to day {day_value}."
                self._queue.put(
                    (
                        "calendar_drop",
                        True,
                        message,
                        {"date_key": normalized_key, "values": order_values},
                    )
                )
        else:
            self._queue.put(
                (
                    "calendar_drop",
                    False,
                    "Please drop orders onto a valid calendar day.",
                    None,
                )
            )

        self._end_drag()

    def _on_day_order_press(self, event: tk.Event, date_key: DateKey) -> None:
        self._end_drag()

        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        orders_list = day_cell.orders_list
        assignments = self._calendar_assignments.get(date_key, [])
        if not assignments:
            orders_list.selection_clear(0, tk.END)
            return

        try:
            index = int(orders_list.nearest(event.y))
        except (tk.TclError, ValueError):
            return

        if index < 0 or index >= len(assignments):
            orders_list.selection_clear(0, tk.END)
            return

        try:
            bbox = orders_list.bbox(index)
        except tk.TclError:
            bbox = None

        if not bbox or not (bbox[1] <= event.y <= bbox[1] + bbox[3]):
            orders_list.selection_clear(0, tk.END)
            return

        orders_list.selection_clear(0, tk.END)
        orders_list.selection_set(index)
        orders_list.activate(index)

        assignment = assignments[index]
        normalized_assignment = tuple(str(value) for value in assignment)
        if len(normalized_assignment) < 2:
            normalized_assignment = normalized_assignment + ("",) * (2 - len(normalized_assignment))
        assignment_values = normalized_assignment[:2]

        try:
            normalized_date_key = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            normalized_date_key = date_key

        self._drag_data.update(
            {
                "item": ("calendar", normalized_date_key, index),
                "values": assignment_values,
                "start_x": event.x_root,
                "start_y": event.y_root,
                "widget": None,
                "active": False,
                "source": "calendar",
                "source_date_key": normalized_date_key,
                "source_index": index,
                "source_assignment": assignment_values,
            }
        )

    def _on_day_order_drag(self, event: tk.Event, date_key: DateKey) -> None:
        if self._drag_data.get("source") != "calendar":
            return

        try:
            normalized_date_key = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            normalized_date_key = date_key

        if self._drag_data.get("source_date_key") != normalized_date_key:
            return

        item_id = self._drag_data.get("item")
        if not item_id:
            return

        if not self._drag_data.get("active"):
            start_x = int(self._drag_data.get("start_x", event.x_root))
            start_y = int(self._drag_data.get("start_y", event.y_root))
            if (
                abs(event.x_root - start_x) >= DRAG_THRESHOLD
                or abs(event.y_root - start_y) >= DRAG_THRESHOLD
            ):
                self._begin_drag()

        if not self._drag_data.get("active"):
            return

        self._position_drag_window(event.x_root, event.y_root)
        target_info = self._detect_calendar_target(event.x_root, event.y_root)
        self._update_calendar_hover(target_info)

    def _on_day_order_release(self, event: tk.Event, date_key: DateKey) -> None:
        if self._drag_data.get("source") != "calendar":
            self._end_drag()
            return

        try:
            normalized_date_key = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            normalized_date_key = date_key

        if self._drag_data.get("source_date_key") != normalized_date_key:
            self._end_drag()
            return

        item_id = self._drag_data.get("item")
        if not item_id:
            self._end_drag()
            return

        if not self._drag_data.get("active"):
            self._end_drag()
            return

        target_info = self._detect_calendar_target(event.x_root, event.y_root)
        normalized_key: DateKey | None = None
        if target_info:
            raw_key = target_info.get("date_key")
            if isinstance(raw_key, (tuple, list)) and len(raw_key) == 3:
                try:
                    normalized_key = (
                        int(raw_key[0]),
                        int(raw_key[1]),
                        int(raw_key[2]),
                    )
                except (TypeError, ValueError):
                    normalized_key = None

        if normalized_key is not None:
            values = self._drag_data.get("values", ())
            if not isinstance(values, (tuple, list)) or not values:
                self._queue.put(
                    (
                        "calendar_drop",
                        False,
                        "Unable to determine which order was dragged.",
                        None,
                    )
                )
            else:
                order_values = tuple(str(value) for value in values)
                order_number = order_values[0] if len(order_values) > 0 else ""
                company = order_values[1] if len(order_values) > 1 else ""

                raw_source_key = self._drag_data.get("source_date_key")
                normalized_source: DateKey | None = None
                if isinstance(raw_source_key, (tuple, list)) and len(raw_source_key) == 3:
                    try:
                        normalized_source = (
                            int(raw_source_key[0]),
                            int(raw_source_key[1]),
                            int(raw_source_key[2]),
                        )
                    except (TypeError, ValueError):
                        normalized_source = None

                target_label = self._format_date_label(normalized_key)
                source_label = (
                    self._format_date_label(normalized_source)
                    if normalized_source is not None
                    else ""
                )

                order_label = "order"
                if order_number:
                    order_label += f" {order_number}"
                if company:
                    order_label += f" ({company})"

                if normalized_source is not None and normalized_source != normalized_key:
                    message = f"Moved {order_label} from {source_label} to {target_label}."
                elif normalized_source is not None and normalized_source == normalized_key:
                    capitalized = order_label[0].upper() + order_label[1:]
                    message = f"{capitalized} remains scheduled for {target_label}."
                else:
                    message = f"Assigned {order_label} to {target_label}."

                payload: dict[str, object] = {
                    "date_key": normalized_key,
                    "values": order_values,
                }

                if normalized_source is not None:
                    payload["source_date_key"] = normalized_source
                    source_index = self._drag_data.get("source_index")
                    try:
                        payload["source_index"] = int(source_index)
                    except (TypeError, ValueError):
                        pass

                    source_assignment = self._drag_data.get("source_assignment")
                    if isinstance(source_assignment, (tuple, list)):
                        payload["source_assignment"] = tuple(
                            str(value) for value in source_assignment
                        )
                    else:
                        payload["source_assignment"] = order_values

                self._queue.put(("calendar_drop", True, message, payload))
        else:
            self._queue.put(
                (
                    "calendar_drop",
                    False,
                    "Please drop orders onto a valid calendar day.",
                    None,
                )
            )

        self._end_drag()

    def _begin_drag(self) -> None:
        item_id = self._drag_data.get("item")
        values = self._drag_data.get("values", ())
        if not item_id or not isinstance(values, (tuple, list)):
            return

        order_values = tuple(str(value) for value in values)
        text_parts = [part for part in order_values if part]
        label_text = " - ".join(text_parts) if text_parts else ""

        drag_window = tk.Toplevel(self.root)
        drag_window.overrideredirect(True)
        try:  # pragma: no cover - platform dependent feature
            drag_window.attributes("-topmost", True)
        except tk.TclError:
            pass
        drag_window.configure(bg=ACCENT_COLOR)

        label = tk.Label(
            drag_window,
            text=label_text,
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            padx=8,
            pady=4,
            bd=0,
        )
        label.pack()

        self._drag_data["widget"] = drag_window
        self._drag_data["values"] = order_values
        self._drag_data["active"] = True

        start_x = int(self._drag_data.get("start_x", 0))
        start_y = int(self._drag_data.get("start_y", 0))
        self._position_drag_window(start_x, start_y)

    def _position_drag_window(self, x_root: int, y_root: int) -> None:
        widget = self._drag_data.get("widget")
        if widget is None:
            return
        widget.geometry(f"+{x_root + 16}+{y_root + 16}")

    def _detect_calendar_target(self, x_root: int, y_root: int) -> Dict[str, object] | None:
        try:
            grid_x = self.calendar_grid.winfo_rootx()
            grid_y = self.calendar_grid.winfo_rooty()
            grid_width = self.calendar_grid.winfo_width()
            grid_height = self.calendar_grid.winfo_height()
        except tk.TclError:
            return None

        if not (
            grid_x <= x_root <= grid_x + grid_width
            and grid_y <= y_root <= grid_y + grid_height
        ):
            return None

        for date_key, day_cell in self._day_cells.items():
            frame = day_cell.frame
            if not frame.winfo_viewable():
                continue

            cell_x = frame.winfo_rootx()
            cell_y = frame.winfo_rooty()
            cell_width = frame.winfo_width()
            cell_height = frame.winfo_height()

            if (
                cell_x <= x_root <= cell_x + cell_width
                and cell_y <= y_root <= cell_y + cell_height
            ):
                return {
                    "date_key": date_key,
                    "day": date_key[2],
                    "frame": frame,
                }

        return {"date_key": None, "day": None, "frame": None}

    def _update_calendar_hover(self, target_info: Dict[str, object] | None) -> None:
        if not target_info:
            self._remove_calendar_hover()
            return

        date_key = target_info.get("date_key")
        if date_key:
            self._apply_calendar_hover(date_key, True)
        else:
            self._remove_calendar_hover()

    def _apply_calendar_hover(self, date_key: DateKey, is_valid: bool) -> None:
        if not is_valid:
            self._remove_calendar_hover()
            return

        if self._calendar_hover == date_key:
            return

        self._remove_calendar_hover()

        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        hover_color = DAY_CELL_HOVER_VALID if is_valid else DAY_CELL_HOVER_INVALID
        border_color = "#1e90ff" if is_valid else FAIL_COLOR

        day_cell.frame.configure(
            bg=hover_color,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=day_cell.border_thickness + 1,
        )
        day_cell.header_label.configure(bg=hover_color)

        self._calendar_hover = date_key

    def _remove_calendar_hover(self) -> None:
        if self._calendar_hover is None:
            return

        date_key = self._calendar_hover
        self._apply_day_cell_base_style(date_key)

        self._calendar_hover = None

    def _end_drag(self) -> None:
        widget = self._drag_data.get("widget")
        if widget is not None:
            try:  # pragma: no cover - defensive cleanup
                widget.destroy()
            except tk.TclError:
                pass
        self._remove_calendar_hover()
        self._reset_drag_state()

    def _handle_calendar_drop(self, success: bool, message: str, payload: object | None) -> None:
        color = SUCCESS_COLOR if success else FAIL_COLOR
        self._set_status(color, message)
        if not success:
            return

        if not isinstance(payload, dict):
            return

        date_key = payload.get("date_key")
        values = payload.get("values")
        if (
            not isinstance(date_key, (tuple, list))
            or len(date_key) != 3
            or not isinstance(values, (tuple, list))
        ):
            return

        try:
            normalized_key = (int(date_key[0]), int(date_key[1]), int(date_key[2]))
        except (TypeError, ValueError):
            return

        order_values = tuple(str(value) for value in values)
        raw_source_key = payload.get("source_date_key")
        normalized_source: DateKey | None = None
        if isinstance(raw_source_key, (tuple, list)) and len(raw_source_key) == 3:
            try:
                normalized_source = (
                    int(raw_source_key[0]),
                    int(raw_source_key[1]),
                    int(raw_source_key[2]),
                )
            except (TypeError, ValueError):
                normalized_source = None

        source_assignment_raw = payload.get("source_assignment")
        source_assignment: Tuple[str, ...] | None = None
        if isinstance(source_assignment_raw, (tuple, list)):
            source_assignment = tuple(str(value) for value in source_assignment_raw)

        source_index_raw = payload.get("source_index")
        if isinstance(source_index_raw, int):
            source_index = source_index_raw
        else:
            try:
                source_index = int(source_index_raw)
            except (TypeError, ValueError):
                source_index = None

        removed_from_source = False
        if normalized_source is not None:
            assignments = self._calendar_assignments.get(normalized_source, [])
            expected_assignment = source_assignment or order_values

            removal_index: int | None = None
            if (
                isinstance(source_index, int)
                and 0 <= source_index < len(assignments)
            ):
                candidate = assignments[source_index]
                candidate_tuple = tuple(str(value) for value in candidate)
                if candidate_tuple == expected_assignment:
                    removal_index = source_index

            if removal_index is None:
                for idx, assignment in enumerate(assignments):
                    if tuple(str(value) for value in assignment) == expected_assignment:
                        removal_index = idx
                        break

            if removal_index is not None:
                assignments.pop(removal_index)
                if assignments:
                    self._calendar_assignments[normalized_source] = assignments
                else:
                    self._calendar_assignments.pop(normalized_source, None)
                self._update_day_cell_display(normalized_source)
                removed_from_source = True

        added_to_target = self._assign_order_to_day(normalized_key, order_values)
        if removed_from_source and not added_to_target:
            self._schedule_state_save()

    def _assign_order_to_day(
        self, date_key: DateKey, order_values: Tuple[str, ...]
    ) -> bool:
        order_number = order_values[0] if len(order_values) > 0 else ""
        company = order_values[1] if len(order_values) > 1 else ""
        normalized: Tuple[str, str] = (str(order_number), str(company))

        assignments = self._calendar_assignments.setdefault(date_key, [])
        changed = False
        if normalized not in assignments:
            assignments.append(normalized)
            changed = True

        self._update_day_cell_display(date_key)

        if changed:
            self._schedule_state_save()

        return changed

    def _update_day_cell_display(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        assignments = self._calendar_assignments.get(date_key, [])
        orders_list = day_cell.orders_list
        orders_list.delete(0, tk.END)
        for assignment in assignments:
            orders_list.insert(tk.END, self._format_assignment_label(assignment))

        try:
            day_value = int(date_key[2])
        except (TypeError, ValueError):
            day_text = str(date_key[2])
        else:
            if assignments:
                day_text = f"{day_value} ({len(assignments)})"
            else:
                day_text = str(day_value)

        day_cell.header_label.configure(text=day_text)
        self._apply_day_cell_base_style(date_key)

    def _format_assignment_label(self, assignment: Tuple[str, str]) -> str:
        order_number = assignment[0].strip()
        company = assignment[1].strip()
        if order_number and company:
            return f"{order_number} - {company}"
        if order_number:
            return order_number
        if company:
            return company
        return "Unnamed order"

    def _format_date_label(self, date_key: DateKey) -> str:
        try:
            year, month, day = (int(date_key[0]), int(date_key[1]), int(date_key[2]))
            display_date = dt.date(year, month, day)
        except (TypeError, ValueError):
            year, month, day = date_key
            return f"{month:02d}/{day:02d}/{year}"
        else:
            return display_date.strftime("%B %d, %Y")

    def _format_removal_message(
        self, date_key: DateKey, assignment: Tuple[str, str]
    ) -> str:
        order_number = assignment[0].strip()
        company = assignment[1].strip()

        if order_number:
            message = f"Removed order {order_number}"
        else:
            message = "Removed order"

        if company:
            if order_number:
                message += f" ({company})"
            else:
                message += f" for {company}"

        message += f" from {self._format_date_label(date_key)}."
        return message

    def _on_login_clicked(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get()

        if not username or not password:
            self._set_status(FAIL_COLOR, "Please enter both a username and password.")
            return

        self.login_button.config(state=tk.DISABLED)
        self.refresh_button.config(state=tk.DISABLED)
        self._set_status(PENDING_COLOR, "Attempting login...")

        thread = threading.Thread(
            target=self._perform_login,
            args=(username, password),
            daemon=True,
        )
        thread.start()

    def _on_enter_pressed(self, event: object | None) -> None:
        self._on_login_clicked()

    def _on_refresh_clicked(self) -> None:
        self.login_button.config(state=tk.DISABLED)
        self.refresh_button.config(state=tk.DISABLED)
        self._set_status(PENDING_COLOR, "Refreshing orders...")

        thread = threading.Thread(target=self._perform_refresh, daemon=True)
        thread.start()

    def _set_status(self, color: str, message: str) -> None:
        self.status_canvas.itemconfigure(self.status_light, fill=color)
        self.status_message.config(text=message)

    def _update_last_refresh(self, success: bool) -> None:
        if success:
            timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.last_refresh_var.set(f"Last updated: {timestamp}")
            return

        current_value = self.last_refresh_var.get().strip()
        if current_value:
            base_value = current_value.split(" (stale)", 1)[0]
            self.last_refresh_var.set(f"{base_value} (stale)")
        else:
            self.last_refresh_var.set("")

    def _format_status_with_last_refresh(self, message: str) -> str:
        last_refresh_text = self.last_refresh_var.get().strip()
        if message and last_refresh_text:
            return f"{message} ({last_refresh_text})"
        return message

    def _perform_login(self, username: str, password: str) -> None:
        try:
            self.client.login(username, password)
            orders = self.client.fetch_orders()
        except (AuthenticationError, NetworkError) as exc:
            self._queue.put(("login_result", False, str(exc), [], "login"))
        except Exception as exc:  # pragma: no cover - defensive
            self._queue.put(("login_result", False, f"Unexpected error: {exc}", [], "login"))
        else:
            self._queue.put(("login_result", True, "Login successful.", orders, "login"))

    def _perform_refresh(self) -> None:
        try:
            orders = self.client.fetch_orders()
        except (AuthenticationError, NetworkError) as exc:
            self._queue.put(("login_result", False, str(exc), [], "refresh"))
        except Exception as exc:  # pragma: no cover - defensive
            self._queue.put(("login_result", False, f"Unexpected error: {exc}", [], "refresh"))
        else:
            self._queue.put(("login_result", True, "Orders refreshed.", orders, "refresh"))

    def _poll_queue(self) -> None:
        try:
            while True:
                event = self._queue.get_nowait()
                if not event:
                    continue

                event_type = event[0]
                if event_type == "login_result":
                    success = bool(event[1])
                    message = str(event[2])
                    payload = event[3] if len(event) > 3 else []
                    extra = event[4] if len(event) > 4 else "login"
                    operation = extra if isinstance(extra, str) else "login"

                    orders: List[OrderRecord] = []
                    if isinstance(payload, (list, tuple)):
                        orders = list(payload)

                    self._handle_login_result(success, message, orders, operation)
                    continue

                if event_type == "calendar_drop":
                    success = bool(event[1]) if len(event) > 1 else False
                    message = str(event[2]) if len(event) > 2 else ""
                    payload = event[3] if len(event) > 3 else None
                    self._handle_calendar_drop(success, message, payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_queue)

    def _handle_login_result(
        self, success: bool, message: str, orders: List[OrderRecord], operation: str = "login"
    ) -> None:
        operation_key = operation.lower() if isinstance(operation, str) else "login"

        self._update_last_refresh(success)

        color = SUCCESS_COLOR if success else FAIL_COLOR
        status_message = self._format_status_with_last_refresh(message)
        self._set_status(color, status_message)
        self.login_button.config(state=tk.NORMAL)

        if success:
            self.refresh_button.config(state=tk.NORMAL)
            self._populate_orders(orders)
            if operation_key == "login" and not orders:
                empty_message = "Login successful, but no orders were found."
                formatted_message = self._format_status_with_last_refresh(empty_message)
                self._set_status(color, formatted_message)
        else:
            self.refresh_button.config(state=tk.DISABLED)
            if operation_key == "login":
                self.password_var.set("")

    def _populate_orders(self, orders: Iterable[OrderRecord]) -> None:
        self._all_orders = list(orders)
        self._apply_order_filter()

    def _apply_order_filter(self, event: object | None = None) -> None:
        filter_text = self._order_filter_var.get().strip().lower()

        for item in self.tree.get_children():
            self.tree.delete(item)

        for order in self._all_orders:
            order_number = str(getattr(order, "order_number", ""))
            company = str(getattr(order, "company", ""))
            if filter_text and not (
                filter_text in order_number.lower() or filter_text in company.lower()
            ):
                continue
            self.tree.insert("", tk.END, values=(order_number, company))


def launch_app() -> None:
    root = tk.Tk()
    app = YBSApp(root)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - manual usage
    launch_app()
