"""Tkinter GUI for the YBS Print Calander application."""

from __future__ import annotations

import calendar
import datetime as dt
import queue
import threading
import tkinter as tk
from tkinter import ttk
from typing import Dict, Iterable, List, Tuple

from .client import AuthenticationError, NetworkError, OrderRecord, YBSClient

BACKGROUND_COLOR = "#0b1d3a"
ACCENT_COLOR = "#1f3a63"
TEXT_COLOR = "#f8f9fa"
SUCCESS_COLOR = "#28a745"
FAIL_COLOR = "#dc3545"
PENDING_COLOR = "#f0ad4e"

DRAG_THRESHOLD = 5


class YBSApp:
    """Encapsulates the Tkinter application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("YBS Print Calander")
        self.root.configure(background=BACKGROUND_COLOR)
        self.root.geometry("720x480")

        self.client = YBSClient()
        self._queue: "queue.Queue[tuple[str, object]]" = queue.Queue()

        self._calendar_cells: Dict[int, Tuple[str, str]] = {}
        self._calendar_cell_lookup: Dict[Tuple[str, str], int] = {}
        self._calendar_assignments: Dict[int, List[Tuple[str, str]]] = {}
        self._calendar_hover: tuple[str, str] | None = None
        self._drag_data: dict[str, object] = {}
        self._reset_drag_state()

        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()

        self._configure_style()
        self._build_layout()
        self._poll_queue()

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

        self.status_message = ttk.Label(login_frame, text="", style="Dark.TLabel")
        self.status_message.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))

        content_paned = ttk.Panedwindow(container, orient=tk.HORIZONTAL, style="Dark.TPanedwindow")
        content_paned.pack(fill=tk.BOTH, expand=True)

        table_frame = ttk.LabelFrame(
            content_paned,
            text="Orders",
            style="Dark.TLabelframe",
            padding=10,
        )
        table_frame.configure(labelanchor="n")

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

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<ButtonPress-1>", self._on_order_press)
        self.tree.bind("<B1-Motion>", self._on_order_drag)
        self.tree.bind("<ButtonRelease-1>", self._on_order_release)

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        content_paned.add(table_frame, weight=3)

        calendar_frame = ttk.LabelFrame(
            content_paned,
            text="Calendar",
            style="Dark.TLabelframe",
            padding=10,
        )
        calendar_frame.configure(labelanchor="n")

        today = dt.date.today()
        month_name = today.strftime("%B %Y")
        month_label = ttk.Label(calendar_frame, text=month_name, style="Dark.TLabel")
        month_label.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        columns = [f"day_{i}" for i in range(7)]
        self.calendar_tree = ttk.Treeview(
            calendar_frame,
            columns=columns,
            show="headings",
            style="Dark.Treeview",
            height=7,
            selectmode="none",
        )

        self.calendar_tree.tag_configure("hover_valid", background="#25497a")
        self.calendar_tree.tag_configure("hover_invalid", background="#5a1f1f")

        for index, day_name in enumerate(calendar.day_abbr):
            column_id = columns[index]
            self.calendar_tree.heading(column_id, text=day_name, anchor="center")
            self.calendar_tree.column(column_id, anchor="center", width=50, stretch=True)

        month_structure = calendar.Calendar().monthdayscalendar(today.year, today.month)
        self._calendar_cells.clear()
        self._calendar_cell_lookup.clear()
        self._calendar_assignments.clear()
        self._calendar_hover = None

        for week in month_structure:
            formatted_week = [str(day) if day != 0 else "" for day in week]
            item_id = self.calendar_tree.insert("", tk.END, values=formatted_week)
            for index, day in enumerate(week):
                if day == 0:
                    continue
                column_name = columns[index]
                self._calendar_cells[day] = (item_id, column_name)
                self._calendar_cell_lookup[(item_id, column_name)] = day

        self.calendar_tree.state(["disabled"])
        self.calendar_tree.grid(row=1, column=0, sticky="nsew")

        calendar_frame.columnconfigure(0, weight=1)
        calendar_frame.rowconfigure(1, weight=1)

        content_paned.add(calendar_frame, weight=2)

    def _reset_drag_state(self) -> None:
        self._drag_data = {
            "item": None,
            "values": (),
            "start_x": 0,
            "start_y": 0,
            "widget": None,
            "active": False,
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

        self._drag_data.update(
            {
                "item": item_id,
                "values": values,
                "start_x": event.x_root,
                "start_y": event.y_root,
                "widget": None,
                "active": False,
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
        if target_info and target_info.get("day") is not None:
            day_value = int(target_info["day"])
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
                message += f" to day {day_value}."
                self._queue.put(
                    (
                        "calendar_drop",
                        True,
                        message,
                        {"day": day_value, "values": order_values},
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
        calendar_x = self.calendar_tree.winfo_rootx()
        calendar_y = self.calendar_tree.winfo_rooty()
        width = self.calendar_tree.winfo_width()
        height = self.calendar_tree.winfo_height()

        if not (
            calendar_x <= x_root <= calendar_x + width
            and calendar_y <= y_root <= calendar_y + height
        ):
            return None

        relative_x = x_root - calendar_x
        relative_y = y_root - calendar_y
        row_id = self.calendar_tree.identify_row(relative_y)
        column = self.calendar_tree.identify_column(relative_x)

        if not row_id or not column:
            return {"item": row_id, "column": column, "day": None}

        try:
            column_index = int(column.lstrip("#")) - 1
        except ValueError:
            return {"item": row_id, "column": column, "day": None}

        columns = self.calendar_tree["columns"]
        if column_index < 0 or column_index >= len(columns):
            return {"item": row_id, "column": column, "day": None}

        column_name = columns[column_index]
        day_value = self._calendar_cell_lookup.get((row_id, column_name))
        return {"item": row_id, "column": column_name, "day": day_value}

    def _update_calendar_hover(self, target_info: Dict[str, object] | None) -> None:
        if not target_info or not target_info.get("item"):
            self._remove_calendar_hover()
            return

        item_id = str(target_info["item"])
        day_value = target_info.get("day")
        is_valid = day_value is not None
        self._apply_calendar_hover(item_id, is_valid)

    def _apply_calendar_hover(self, item_id: str, is_valid: bool) -> None:
        tag_to_apply = "hover_valid" if is_valid else "hover_invalid"

        if self._calendar_hover and self._calendar_hover[0] != item_id:
            self._remove_calendar_hover()

        if self._calendar_hover and self._calendar_hover == (item_id, tag_to_apply):
            return

        current_tags = set(self.calendar_tree.item(item_id, "tags") or ())
        current_tags.discard("hover_valid")
        current_tags.discard("hover_invalid")
        current_tags.add(tag_to_apply)
        self.calendar_tree.item(item_id, tags=tuple(current_tags))
        self._calendar_hover = (item_id, tag_to_apply)

    def _remove_calendar_hover(self) -> None:
        if not self._calendar_hover:
            return

        item_id, tag_name = self._calendar_hover
        current_tags = set(self.calendar_tree.item(item_id, "tags") or ())
        if tag_name in current_tags:
            current_tags.remove(tag_name)
            self.calendar_tree.item(item_id, tags=tuple(current_tags))
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

        day_value = payload.get("day")
        values = payload.get("values")
        if not isinstance(day_value, int) or not isinstance(values, (tuple, list)):
            return

        order_values = tuple(str(value) for value in values)
        self._assign_order_to_day(day_value, order_values)

    def _assign_order_to_day(self, day: int, order_values: Tuple[str, ...]) -> None:
        target = self._calendar_cells.get(day)
        if not target:
            return

        item_id, column_name = target
        order_number = order_values[0] if len(order_values) > 0 else ""
        company = order_values[1] if len(order_values) > 1 else ""
        normalized: Tuple[str, str] = (str(order_number), str(company))

        assignments = self._calendar_assignments.setdefault(day, [])
        if normalized not in assignments:
            assignments.append(normalized)

        display_text = self._format_day_cell(day, assignments)
        self.calendar_tree.set(item_id, column_name, display_text)

    def _format_day_cell(self, day: int, assignments: List[Tuple[str, str]]) -> str:
        if not assignments:
            return str(day)

        labels: List[str] = []
        for order_number, company in assignments:
            label = order_number.strip()
            if not label and company:
                label = company.strip()
            if label:
                labels.append(label)

        if not labels:
            return str(day)

        return f"{day} ({', '.join(labels)})"

    def _on_login_clicked(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get()

        if not username or not password:
            self._set_status(FAIL_COLOR, "Please enter both a username and password.")
            return

        self.login_button.config(state=tk.DISABLED)
        self._set_status(PENDING_COLOR, "Attempting login...")

        thread = threading.Thread(
            target=self._perform_login,
            args=(username, password),
            daemon=True,
        )
        thread.start()

    def _on_enter_pressed(self, event: object | None) -> None:
        self._on_login_clicked()

    def _set_status(self, color: str, message: str) -> None:
        self.status_canvas.itemconfigure(self.status_light, fill=color)
        self.status_message.config(text=message)

    def _perform_login(self, username: str, password: str) -> None:
        try:
            self.client.login(username, password)
            orders = self.client.fetch_orders()
        except (AuthenticationError, NetworkError) as exc:
            self._queue.put(("login_result", False, str(exc), []))
        except Exception as exc:  # pragma: no cover - defensive
            self._queue.put(("login_result", False, f"Unexpected error: {exc}", []))
        else:
            self._queue.put(("login_result", True, "Login successful.", orders))

    def _poll_queue(self) -> None:
        try:
            while True:
                event_type, success, message, payload = self._queue.get_nowait()
                if event_type == "login_result":
                    self._handle_login_result(bool(success), str(message), list(payload))
                elif event_type == "calendar_drop":
                    self._handle_calendar_drop(bool(success), str(message), payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_queue)

    def _handle_login_result(self, success: bool, message: str, orders: List[OrderRecord]) -> None:
        color = SUCCESS_COLOR if success else FAIL_COLOR
        self._set_status(color, message)
        self.login_button.config(state=tk.NORMAL)
        if success:
            self._populate_orders(orders)
            if not orders:
                self.status_message.config(text="Login successful, but no orders were found.")

    def _populate_orders(self, orders: Iterable[OrderRecord]) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for order in orders:
            self.tree.insert("", tk.END, values=(order.order_number, order.company))


def launch_app() -> None:
    root = tk.Tk()
    app = YBSApp(root)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - manual usage
    launch_app()
