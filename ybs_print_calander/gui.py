"""Tkinter GUI for the YBS Print Calander application."""

from __future__ import annotations

import calendar
import datetime as dt
import queue
import threading
import tkinter as tk
from tkinter import ttk
from typing import Iterable, List

from .client import AuthenticationError, NetworkError, OrderRecord, YBSClient

BACKGROUND_COLOR = "#0b1d3a"
ACCENT_COLOR = "#1f3a63"
TEXT_COLOR = "#f8f9fa"
SUCCESS_COLOR = "#28a745"
FAIL_COLOR = "#dc3545"
PENDING_COLOR = "#f0ad4e"


class YBSApp:
    """Encapsulates the Tkinter application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("YBS Print Calander")
        self.root.configure(background=BACKGROUND_COLOR)
        self.root.geometry("720x480")

        self.client = YBSClient()
        self._queue: "queue.Queue[tuple[str, object]]" = queue.Queue()

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

        for index, day_name in enumerate(calendar.day_abbr):
            column_id = columns[index]
            self.calendar_tree.heading(column_id, text=day_name, anchor="center")
            self.calendar_tree.column(column_id, anchor="center", width=50, stretch=True)

        month_structure = calendar.Calendar().monthdayscalendar(today.year, today.month)
        for week in month_structure:
            formatted_week = [day if day != 0 else "" for day in week]
            self.calendar_tree.insert("", tk.END, values=formatted_week)

        self.calendar_tree.state(["disabled"])
        self.calendar_tree.grid(row=1, column=0, sticky="nsew")

        calendar_frame.columnconfigure(0, weight=1)
        calendar_frame.rowconfigure(1, weight=1)

        content_paned.add(calendar_frame, weight=2)

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
