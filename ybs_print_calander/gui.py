"""Tkinter GUI for the YBS Print Calander application."""

from __future__ import annotations

import calendar
import datetime as dt
import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Dict, Iterable, List, Tuple

from . import __version__
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
ACTIVE_DAY_BORDER_COLOR = "#1e90ff"
ACTIVE_DAY_HEADER_BACKGROUND = "#337ab7"
ASSIGNMENT_HEADER_BACKGROUND = "#2561a8"
ASSIGNMENT_LIST_BACKGROUND = "#17406d"
NOTES_TEXT_BACKGROUND = "#0f3460"
ORDERS_LIST_BACKGROUND = "#0d274a"
ADJACENT_MONTH_DAY_CELL_BACKGROUND = "#0b1f3d"
ADJACENT_MONTH_TEXT_COLOR = "#a5b3c8"
ADJACENT_MONTH_NOTES_BACKGROUND = "#0d2749"
ADJACENT_MONTH_ORDERS_BACKGROUND = "#0a1c36"

APP_NAME = "YBS Print Calander"
APP_TITLE = f"{APP_NAME} v{__version__}"
ABOUT_TITLE = f"About {APP_NAME}"
APP_DESCRIPTION = "Manage print orders and assignments."
ABOUT_MESSAGE = (
    f"{APP_NAME}\n"
    f"Version {__version__}\n\n"
    f"{APP_DESCRIPTION}"
)

DRAG_THRESHOLD = 5

DateKey = Tuple[int, int, int]


XRANDR_MONITOR_PATTERN = re.compile(
    r"^\s*\S+\s+connected(?:\s+primary)?\s+(?P<w>\d+)x(?P<h>\d+)\+(?P<x>-?\d+)\+(?P<y>-?\d+)",
    re.IGNORECASE,
)


STATE_PATH = Path.home() / ".ybs_print_calander" / "state.json"


class HoverTooltip:
    """Display contextual hover text for a widget after a small delay."""

    def __init__(self, widget: tk.Widget, text: str, *, delay: int = 500) -> None:
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id: str | None = None
        self._pointer: tuple[int, int] | None = None
        self._visible = False
        self._widget_destroyed = False

        tooltip: tk.Toplevel | None
        try:
            tooltip = tk.Toplevel(widget.winfo_toplevel())
        except tk.TclError:
            tooltip = None

        self._tooltip = tooltip

        if tooltip is not None:
            tooltip.withdraw()
            try:
                tooltip.wm_overrideredirect(True)
            except tk.TclError:
                pass
            try:
                tooltip.transient(widget.winfo_toplevel())
            except tk.TclError:
                pass
            try:
                tooltip.attributes("-topmost", True)
            except tk.TclError:
                pass
            tooltip.configure(background=ACCENT_COLOR, padx=1, pady=1)

            label = tk.Label(
                tooltip,
                text=text,
                background=ACCENT_COLOR,
                foreground=TEXT_COLOR,
                borderwidth=0,
                highlightthickness=0,
                justify="left",
                wraplength=280,
                padx=10,
                pady=6,
            )
            label.pack()

        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<Motion>", self._on_motion, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")
        widget.bind("<Destroy>", self._on_widget_destroy, add="+")

    def _on_enter(self, event: tk.Event) -> None:
        self._pointer = self._event_pointer(event)
        self._schedule_show()

    def _on_motion(self, event: tk.Event) -> None:
        self._pointer = self._event_pointer(event)
        if self._visible:
            self._hide()
        self._schedule_show()

    def _on_leave(self, _: tk.Event | None) -> None:
        self._hide()

    def _on_widget_destroy(self, _: tk.Event | None) -> None:
        self._widget_destroyed = True
        self._hide()
        tooltip = self._tooltip
        if tooltip is not None:
            try:
                tooltip.destroy()
            except tk.TclError:
                pass
            self._tooltip = None

    def _event_pointer(self, event: tk.Event) -> tuple[int, int]:
        x_root = getattr(event, "x_root", 0)
        y_root = getattr(event, "y_root", 0)
        try:
            return (int(x_root), int(y_root))
        except (TypeError, ValueError):
            return (0, 0)

    def _schedule_show(self) -> None:
        if self._widget_destroyed:
            return
        self._cancel_scheduled_show()
        try:
            self._after_id = self.widget.after(self.delay, self._show)
        except tk.TclError:
            self._after_id = None

    def _cancel_scheduled_show(self) -> None:
        if self._after_id is None:
            return
        try:
            self.widget.after_cancel(self._after_id)
        except tk.TclError:
            pass
        self._after_id = None

    def _show(self) -> None:
        self._after_id = None
        if self._visible or self._tooltip is None or self._widget_destroyed:
            return
        if not self.widget.winfo_exists():
            return
        try:
            self._tooltip.deiconify()
            self._tooltip.lift()
        except tk.TclError:
            return
        self._visible = True
        self._reposition()

    def _hide(self) -> None:
        self._cancel_scheduled_show()
        if not self._visible or self._tooltip is None:
            return
        try:
            self._tooltip.withdraw()
        except tk.TclError:
            pass
        self._visible = False

    def _reposition(self) -> None:
        if not self._visible or self._tooltip is None:
            return
        if self._pointer is None:
            return
        x, y = self._pointer
        try:
            self._tooltip.geometry(f"+{x + 16}+{y + 16}")
        except tk.TclError:
            pass


@dataclass
class DayCell:
    """Container for widgets that make up a calendar day cell."""

    frame: tk.Frame
    header_label: tk.Label
    notes_text: tk.Text
    orders_list: tk.Listbox
    default_bg: str
    header_fg: str = TEXT_COLOR
    notes_bg: str = NOTES_TEXT_BACKGROUND
    notes_fg: str = TEXT_COLOR
    orders_bg: str = ORDERS_LIST_BACKGROUND
    orders_fg: str = TEXT_COLOR
    border_color: str = ACCENT_COLOR
    border_thickness: int = 1
    is_today: bool = False
    in_current_month: bool = True


class YBSApp:
    """Encapsulates the Tkinter application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
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
        self._day_cell_pointer_hover: DateKey | None = None
        self._active_day_header: DateKey | None = None
        self._drag_data: dict[str, object] = {}
        self._tree_selection_anchor: str | None = None
        self._day_selection_anchor: Dict[DateKey, int] = {}
        self._state_path: Path = STATE_PATH
        self._state_path.parent.mkdir(parents=True, exist_ok=True)
        self._state_save_after_id: str | None = None
        self._undo_stack: list[dict[str, Any]] = []
        self._undo_stack_limit = 100
        self._redo_stack: list[dict[str, Any]] = []
        self._redo_stack_limit = self._undo_stack_limit
        self._cached_monitor_bounds: tuple[int, int, int, int] | None = None
        self._cached_monitor_geometry: tuple[int, int, int, int] | None = None

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
        self.root.bind("<Configure>", self._invalidate_monitor_cache, add="+")
        self.root.bind_all("<Control-z>", self._undo_last_action)
        self.root.bind_all("<Command-z>", self._undo_last_action)
        self.root.bind_all("<Control-Shift-Z>", self._redo_last_action)
        self.root.bind_all("<Command-Shift-Z>", self._redo_last_action)
        self.root.bind_all("<Control-y>", self._redo_last_action)
        self.root.bind_all("<Command-y>", self._redo_last_action)
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

    @staticmethod
    def _normalize_assignment(values: Iterable[object]) -> Tuple[str, str]:
        sequence = tuple(str(value) for value in values)
        first = sequence[0] if len(sequence) > 0 else ""
        second = sequence[1] if len(sequence) > 1 else ""
        return (first, second)

    @staticmethod
    def _event_state_has_flag(event: tk.Event | None, mask: int) -> bool:
        try:
            state = int(getattr(event, "state", 0))
        except (TypeError, ValueError):
            return False
        return bool(state & mask)

    @classmethod
    def _is_shift_pressed(cls, event: tk.Event | None) -> bool:
        return cls._event_state_has_flag(event, 0x0001)

    @classmethod
    def _is_control_pressed(cls, event: tk.Event | None) -> bool:
        control_masks: tuple[int, ...] = (0x0004,)
        if sys.platform == "darwin":
            control_masks += (0x0010,)
        return any(cls._event_state_has_flag(event, mask) for mask in control_masks)

    def _invalidate_monitor_cache(self, *_: object) -> None:
        """Clear cached monitor information."""

        self._cached_monitor_bounds = None
        self._cached_monitor_geometry = None

    def _get_monitor_bounds(
        self, reference: tuple[int, int] | None = None
    ) -> tuple[int, int, int, int] | None:
        """Return the bounding rectangle for the monitor containing ``reference``."""

        try:
            self.root.update_idletasks()
        except tk.TclError:
            pass

        default_reference = reference is None
        geometry: tuple[int, int, int, int] | None = None
        if reference is None:
            try:
                root_x = int(self.root.winfo_rootx())
                root_y = int(self.root.winfo_rooty())
                root_width = int(self.root.winfo_width())
                root_height = int(self.root.winfo_height())
            except tk.TclError:
                reference = (0, 0)
            else:
                if root_width <= 1:
                    root_width = int(self.root.winfo_reqwidth())
                if root_height <= 1:
                    root_height = int(self.root.winfo_reqheight())
                geometry = (root_x, root_y, root_width, root_height)
                reference = (
                    root_x + root_width // 2,
                    root_y + root_height // 2,
                )

        assert reference is not None
        ref_x = int(reference[0])
        ref_y = int(reference[1])

        if (
            default_reference
            and geometry is not None
            and self._cached_monitor_bounds is not None
            and self._cached_monitor_geometry == geometry
        ):
            return self._cached_monitor_bounds

        bounds: tuple[int, int, int, int] | None = None

        if bounds is None and sys.platform.startswith("win"):
            try:
                import ctypes
                from ctypes import wintypes
            except Exception:
                pass
            else:
                MonitorFromPoint = getattr(
                    ctypes.windll.user32, "MonitorFromPoint", None
                )
                GetMonitorInfo = getattr(ctypes.windll.user32, "GetMonitorInfoW", None)
                if MonitorFromPoint and GetMonitorInfo:
                    MONITOR_DEFAULTTONEAREST = 2

                    class POINT(ctypes.Structure):
                        _fields_ = [
                            ("x", wintypes.LONG),
                            ("y", wintypes.LONG),
                        ]

                    class RECT(ctypes.Structure):
                        _fields_ = [
                            ("left", wintypes.LONG),
                            ("top", wintypes.LONG),
                            ("right", wintypes.LONG),
                            ("bottom", wintypes.LONG),
                        ]

                    class MONITORINFO(ctypes.Structure):
                        _fields_ = [
                            ("cbSize", wintypes.DWORD),
                            ("rcMonitor", RECT),
                            ("rcWork", RECT),
                            ("dwFlags", wintypes.DWORD),
                        ]

                    point = POINT(ref_x, ref_y)
                    monitor = MonitorFromPoint(point, MONITOR_DEFAULTTONEAREST)
                    if monitor:
                        info = MONITORINFO()
                        info.cbSize = ctypes.sizeof(info)
                        if GetMonitorInfo(monitor, ctypes.byref(info)):
                            rect = info.rcMonitor
                            left, top, right, bottom = (
                                int(rect.left),
                                int(rect.top),
                                int(rect.right),
                                int(rect.bottom),
                            )
                            if left < right and top < bottom:
                                bounds = (left, top, right, bottom)

        if bounds is None:
            try:
                from screeninfo import get_monitors  # type: ignore
            except Exception:
                pass
            else:
                monitors: list[tuple[int, int, int, int]] = []
                for monitor in get_monitors():
                    try:
                        left = int(getattr(monitor, "x"))
                        top = int(getattr(monitor, "y"))
                        width = int(getattr(monitor, "width"))
                        height = int(getattr(monitor, "height"))
                    except (AttributeError, TypeError, ValueError):
                        continue
                    if width <= 0 or height <= 0:
                        continue
                    monitors.append((left, top, left + width, top + height))
                for left, top, right, bottom in monitors:
                    if left <= ref_x < right and top <= ref_y < bottom:
                        bounds = (left, top, right, bottom)
                        break
                if bounds is None and monitors:
                    bounds = monitors[0]

        if (
            bounds is None
            and sys.platform.startswith(("linux", "freebsd"))
            and os.environ.get("DISPLAY")
        ):
            try:
                result = subprocess.run(
                    ["xrandr", "--current"],
                    capture_output=True,
                    text=True,
                    check=False,
                    timeout=0.5,
                )
            except (OSError, subprocess.SubprocessError):
                result = None
            if result and result.returncode == 0 and result.stdout:
                monitors: list[tuple[int, int, int, int]] = []
                for line in result.stdout.splitlines():
                    match = XRANDR_MONITOR_PATTERN.match(line)
                    if not match:
                        continue
                    try:
                        width = int(match.group("w"))
                        height = int(match.group("h"))
                        left = int(match.group("x"))
                        top = int(match.group("y"))
                    except (TypeError, ValueError):
                        continue
                    if width <= 0 or height <= 0:
                        continue
                    monitors.append((left, top, left + width, top + height))
                for left, top, right, bottom in monitors:
                    if left <= ref_x < right and top <= ref_y < bottom:
                        bounds = (left, top, right, bottom)
                        break
                if bounds is None and monitors:
                    bounds = monitors[0]

        if bounds is None:
            vroot_bounds: tuple[int, int, int, int] | None = None
            try:
                vroot_x = int(self.root.winfo_vrootx())
                vroot_y = int(self.root.winfo_vrooty())
                vroot_width = int(self.root.winfo_vrootwidth())
                vroot_height = int(self.root.winfo_vrootheight())
            except (tk.TclError, ValueError):
                vroot_width = vroot_height = 0
            else:
                if vroot_width > 0 and vroot_height > 0:
                    vroot_bounds = (
                        vroot_x,
                        vroot_y,
                        vroot_x + vroot_width,
                        vroot_y + vroot_height,
                    )

            try:
                screen_width = int(self.root.winfo_screenwidth())
                screen_height = int(self.root.winfo_screenheight())
            except (tk.TclError, ValueError):
                screen_width = screen_height = 0

            if screen_width > 0 and screen_height > 0:
                try:
                    root_left = int(self.root.winfo_rootx())
                    root_top = int(self.root.winfo_rooty())
                except tk.TclError:
                    root_left = root_top = 0

                left = root_left
                top = root_top
                if vroot_bounds is not None:
                    vleft, vtop, vright, vbottom = vroot_bounds
                    if vright - vleft >= screen_width:
                        max_left = vright - screen_width
                        left = max(min(left, max_left), vleft)
                    else:
                        left = vleft
                    if vbottom - vtop >= screen_height:
                        max_top = vbottom - screen_height
                        top = max(min(top, max_top), vtop)
                    else:
                        top = vtop

                bounds = (
                    left,
                    top,
                    left + screen_width,
                    top + screen_height,
                )
            elif vroot_bounds is not None:
                bounds = vroot_bounds

        if bounds is not None and default_reference and geometry is not None:
            self._cached_monitor_bounds = bounds
            self._cached_monitor_geometry = geometry

        return bounds

    def _constrain_to_monitor(
        self,
        x: int,
        y: int,
        width: int,
        height: int,
        reference: tuple[int, int] | None = None,
    ) -> tuple[int, int]:
        bounds = self._get_monitor_bounds(reference=reference)

        try:
            x_value = int(x)
        except (TypeError, ValueError):
            x_value = 0
        try:
            y_value = int(y)
        except (TypeError, ValueError):
            y_value = 0
        try:
            width_value = max(int(width), 1)
        except (TypeError, ValueError):
            width_value = 1
        try:
            height_value = max(int(height), 1)
        except (TypeError, ValueError):
            height_value = 1

        if not bounds:
            return (x_value, y_value)

        left, top, right, bottom = bounds
        if right <= left or bottom <= top:
            return (x_value, y_value)

        max_x = max(right - width_value, left)
        max_y = max(bottom - height_value, top)

        constrained_x = max(left, min(x_value, max_x))
        constrained_y = max(top, min(y_value, max_y))

        return (constrained_x, constrained_y)

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

    def _normalize_date_key(self, date_key: object) -> DateKey | None:
        if isinstance(date_key, (tuple, list)) and len(date_key) == 3:
            try:
                return (int(date_key[0]), int(date_key[1]), int(date_key[2]))
            except (TypeError, ValueError):
                return None
        return None

    def _capture_assignments_state(self, date_key: DateKey) -> dict[str, Any]:
        had_key = date_key in self._calendar_assignments
        if not had_key:
            return {"had_key": False, "previous": None}

        assignments = self._calendar_assignments.get(date_key, [])
        previous: list[Tuple[str, str]] = []
        for entry in assignments:
            if isinstance(entry, (list, tuple)):
                first = str(entry[0]) if len(entry) > 0 else ""
                second = str(entry[1]) if len(entry) > 1 else ""
                previous.append((first, second))
        return {"had_key": True, "previous": previous}

    def _capture_notes_state(self, date_key: DateKey) -> dict[str, Any]:
        had_key = date_key in self._calendar_notes
        previous = self._calendar_notes.get(date_key) if had_key else None
        if isinstance(previous, str):
            return {"had_key": had_key, "previous": previous}
        return {"had_key": had_key, "previous": None}

    def _normalize_history_action(
        self, action: dict[str, Any]
    ) -> dict[str, Any] | None:
        if not isinstance(action, dict):
            return None

        kind = action.get("kind")

        if kind == "assignments":
            raw_dates = action.get("dates")
            if not isinstance(raw_dates, dict) or not raw_dates:
                return None

            normalized_dates: dict[DateKey, dict[str, Any]] = {}
            for raw_key, info in raw_dates.items():
                normalized_key = self._normalize_date_key(raw_key)
                if normalized_key is None:
                    continue

                info_dict = info if isinstance(info, dict) else {}
                had_key = bool(info_dict.get("had_key"))
                previous_raw = info_dict.get("previous")
                previous_list: list[Tuple[str, str]] | None = None
                if isinstance(previous_raw, list):
                    previous_list = []
                    for entry in previous_raw:
                        if isinstance(entry, (list, tuple)):
                            first = str(entry[0]) if len(entry) > 0 else ""
                            second = str(entry[1]) if len(entry) > 1 else ""
                            previous_list.append((first, second))

                normalized_dates[normalized_key] = {
                    "had_key": had_key,
                    "previous": previous_list,
                }

            if not normalized_dates:
                return None

            return {"kind": "assignments", "dates": normalized_dates}

        if kind == "notes":
            normalized_key = self._normalize_date_key(action.get("date_key"))
            if normalized_key is None:
                return None

            had_key = bool(action.get("had_key"))
            previous_raw = action.get("previous")
            if isinstance(previous_raw, str):
                previous_value: str | None = previous_raw
            elif previous_raw is None:
                previous_value = None
            else:
                previous_value = str(previous_raw)

            return {
                "kind": "notes",
                "date_key": normalized_key,
                "had_key": had_key,
                "previous": previous_value,
            }

        return None

    def _push_undo_action(
        self, action: dict[str, Any], *, clear_redo: bool = True
    ) -> None:
        normalized_action = self._normalize_history_action(action)
        if normalized_action is None:
            return

        self._undo_stack.append(normalized_action)
        limit = getattr(self, "_undo_stack_limit", 0)
        if isinstance(limit, int) and limit > 0 and len(self._undo_stack) > limit:
            del self._undo_stack[: len(self._undo_stack) - limit]

        if clear_redo:
            self._redo_stack.clear()

    def _push_redo_action(self, action: dict[str, Any]) -> None:
        normalized_action = self._normalize_history_action(action)
        if normalized_action is None:
            return

        self._redo_stack.append(normalized_action)
        limit = getattr(self, "_redo_stack_limit", 0)
        if isinstance(limit, int) and limit > 0 and len(self._redo_stack) > limit:
            del self._redo_stack[: len(self._redo_stack) - limit]

    def _undo_last_action(self, event: tk.Event | None = None) -> str | None:
        widget = getattr(event, "widget", None)
        if isinstance(widget, tk.Text):
            try:
                undo_enabled = bool(widget.cget("undo"))
            except tk.TclError:
                undo_enabled = False
            if undo_enabled:
                try:
                    widget.edit_undo()
                except tk.TclError:
                    pass
                else:
                    return "break"

        if not self._undo_stack:
            self._set_status(FAIL_COLOR, "Nothing to undo.")
            return "break" if event is not None else None

        action = self._undo_stack.pop()
        kind = action.get("kind")
        restored = False
        status_message = ""

        if kind == "assignments":
            entries = action.get("dates")
            if isinstance(entries, dict):
                restored_labels: list[str] = []
                redo_dates: dict[DateKey, dict[str, Any]] = {}
                for raw_key, info in entries.items():
                    normalized_key = self._normalize_date_key(raw_key)
                    if normalized_key is None:
                        continue

                    redo_dates[normalized_key] = self._capture_assignments_state(
                        normalized_key
                    )

                    info_dict = info if isinstance(info, dict) else {}
                    previous_raw = info_dict.get("previous")
                    restored_assignments: list[Tuple[str, str]] = []
                    if isinstance(previous_raw, list):
                        for entry in previous_raw:
                            if isinstance(entry, (list, tuple)):
                                first = str(entry[0]) if len(entry) > 0 else ""
                                second = str(entry[1]) if len(entry) > 1 else ""
                                restored_assignments.append((first, second))

                    if restored_assignments:
                        self._calendar_assignments[normalized_key] = restored_assignments
                    else:
                        self._calendar_assignments.pop(normalized_key, None)

                    self._update_day_cell_display(normalized_key)
                    restored_labels.append(self._format_date_label(normalized_key))
                    restored = True

                if redo_dates:
                    self._push_redo_action({"kind": "assignments", "dates": redo_dates})

                if restored_labels:
                    if len(restored_labels) == 1:
                        status_message = (
                            f"Undo: restored assignments for {restored_labels[0]}."
                        )
                    else:
                        status_message = (
                            "Undo: restored assignments for "
                            + ", ".join(restored_labels)
                            + "."
                        )
        elif kind == "notes":
            normalized_key = self._normalize_date_key(action.get("date_key"))
            if normalized_key is not None:
                redo_snapshot = self._capture_notes_state(normalized_key)
                self._push_redo_action(
                    {
                        "kind": "notes",
                        "date_key": normalized_key,
                        "had_key": bool(redo_snapshot.get("had_key")),
                        "previous": redo_snapshot.get("previous"),
                    }
                )

                previous_raw = action.get("previous")
                had_key = bool(action.get("had_key"))

                if isinstance(previous_raw, str):
                    restored_text = previous_raw
                    self._calendar_notes[normalized_key] = restored_text
                elif previous_raw is None and not had_key:
                    restored_text = ""
                    self._calendar_notes.pop(normalized_key, None)
                else:
                    restored_text = str(previous_raw) if previous_raw is not None else ""
                    if restored_text:
                        self._calendar_notes[normalized_key] = restored_text
                    else:
                        self._calendar_notes.pop(normalized_key, None)

                day_cell = self._day_cells.get(normalized_key)
                if day_cell:
                    notes_widget = day_cell.notes_text
                    try:
                        focused_widget = self.root.focus_get()
                    except tk.TclError:
                        focused_widget = None

                    has_focus = focused_widget is notes_widget
                    if widget is notes_widget:
                        has_focus = True

                    notes_widget.delete("1.0", tk.END)
                    if restored_text:
                        notes_widget.insert("1.0", restored_text)
                    if has_focus:
                        try:
                            notes_widget.focus_set()
                        except tk.TclError:
                            pass

                status_message = (
                    f"Undo: restored notes for {self._format_date_label(normalized_key)}."
                )
                restored = True

        if restored:
            self._schedule_state_save()
            if status_message:
                self._set_status(SUCCESS_COLOR, status_message)
        else:
            self._set_status(FAIL_COLOR, "Nothing to undo.")

        return "break" if event is not None else None

    def _redo_last_action(self, event: tk.Event | None = None) -> str | None:
        widget = getattr(event, "widget", None)
        if isinstance(widget, tk.Text):
            try:
                redo_enabled = bool(widget.cget("undo"))
            except tk.TclError:
                redo_enabled = False
            if redo_enabled:
                try:
                    widget.edit_redo()
                except tk.TclError:
                    pass
                else:
                    return "break"

        if not self._redo_stack:
            self._set_status(FAIL_COLOR, "Nothing to redo.")
            return "break" if event is not None else None

        action = self._redo_stack.pop()
        kind = action.get("kind")
        applied = False
        status_message = ""

        if kind == "assignments":
            entries = action.get("dates")
            if isinstance(entries, dict):
                undo_entries: dict[DateKey, dict[str, Any]] = {}
                restored_labels: list[str] = []
                for raw_key, info in entries.items():
                    normalized_key = self._normalize_date_key(raw_key)
                    if normalized_key is None:
                        continue

                    undo_entries[normalized_key] = self._capture_assignments_state(
                        normalized_key
                    )

                    info_dict = info if isinstance(info, dict) else {}
                    previous_raw = info_dict.get("previous")
                    restored_assignments: list[Tuple[str, str]] = []
                    if isinstance(previous_raw, list):
                        for entry in previous_raw:
                            if isinstance(entry, (list, tuple)):
                                first = str(entry[0]) if len(entry) > 0 else ""
                                second = str(entry[1]) if len(entry) > 1 else ""
                                restored_assignments.append((first, second))

                    if restored_assignments:
                        self._calendar_assignments[normalized_key] = restored_assignments
                    else:
                        self._calendar_assignments.pop(normalized_key, None)

                    self._update_day_cell_display(normalized_key)
                    restored_labels.append(self._format_date_label(normalized_key))
                    applied = True

                if undo_entries:
                    self._push_undo_action(
                        {"kind": "assignments", "dates": undo_entries},
                        clear_redo=False,
                    )

                if restored_labels:
                    if len(restored_labels) == 1:
                        status_message = (
                            f"Redo: restored assignments for {restored_labels[0]}."
                        )
                    else:
                        status_message = (
                            "Redo: restored assignments for "
                            + ", ".join(restored_labels)
                            + "."
                        )

        elif kind == "notes":
            normalized_key = self._normalize_date_key(action.get("date_key"))
            if normalized_key is not None:
                undo_snapshot = self._capture_notes_state(normalized_key)
                self._push_undo_action(
                    {
                        "kind": "notes",
                        "date_key": normalized_key,
                        "had_key": bool(undo_snapshot.get("had_key")),
                        "previous": undo_snapshot.get("previous"),
                    },
                    clear_redo=False,
                )

                previous_raw = action.get("previous")
                had_key = bool(action.get("had_key"))

                if isinstance(previous_raw, str):
                    restored_text = previous_raw
                    self._calendar_notes[normalized_key] = restored_text
                elif previous_raw is None and not had_key:
                    restored_text = ""
                    self._calendar_notes.pop(normalized_key, None)
                else:
                    restored_text = str(previous_raw) if previous_raw is not None else ""
                    if restored_text:
                        self._calendar_notes[normalized_key] = restored_text
                    else:
                        self._calendar_notes.pop(normalized_key, None)

                day_cell = self._day_cells.get(normalized_key)
                if day_cell:
                    notes_widget = day_cell.notes_text
                    try:
                        focused_widget = self.root.focus_get()
                    except tk.TclError:
                        focused_widget = None

                    has_focus = focused_widget is notes_widget
                    if widget is notes_widget:
                        has_focus = True

                    notes_widget.delete("1.0", tk.END)
                    if restored_text:
                        notes_widget.insert("1.0", restored_text)
                    if has_focus:
                        try:
                            notes_widget.focus_set()
                        except tk.TclError:
                            pass

                status_message = (
                    f"Redo: restored notes for {self._format_date_label(normalized_key)}."
                )
                applied = True

        if applied:
            self._schedule_state_save()
            if status_message:
                self._set_status(SUCCESS_COLOR, status_message)
        else:
            self._set_status(FAIL_COLOR, "Nothing to redo.")

        return "break" if event is not None else None

    def _invoke_text_widget_undo(self, event: tk.Event) -> None:
        widget = getattr(event, "widget", None)
        if isinstance(widget, tk.Text):
            try:
                widget.edit_undo()
            except tk.TclError:
                pass

    def _invoke_text_widget_redo(self, event: tk.Event) -> None:
        widget = getattr(event, "widget", None)
        if isinstance(widget, tk.Text):
            try:
                widget.edit_redo()
            except tk.TclError:
                pass

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
        style.configure("Dark.TNotebook", background=BACKGROUND_COLOR, borderwidth=0)
        style.configure(
            "Dark.TNotebook.Tab",
            background=ACCENT_COLOR,
            foreground=TEXT_COLOR,
            padding=(12, 6),
        )
        style.map(
            "Dark.TNotebook.Tab",
            background=[("selected", ACTIVE_DAY_HEADER_BACKGROUND)],
            foreground=[("selected", TEXT_COLOR)],
        )
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
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Exit", command=self._on_close)
        menu_bar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(menu_bar, tearoff=False)
        edit_menu.add_command(label="Undo", command=lambda: self._undo_last_action(None))
        edit_menu.add_command(label="Redo", command=lambda: self._redo_last_action(None))
        menu_bar.add_cascade(label="Edit", menu=edit_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=False)
        settings_menu.add_command(label="Show Settings", command=self._focus_settings_tab)
        settings_menu.add_command(
            label="Show Orders & Calendar",
            command=self._focus_main_tab,
        )
        menu_bar.add_cascade(label="Settings", menu=settings_menu)

        help_menu = tk.Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="About", command=self._show_about_message)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        container = ttk.Frame(self.root, style="Dark.TFrame")
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.notebook = ttk.Notebook(container, style="Dark.TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.main_tab = ttk.Frame(self.notebook, style="Dark.TFrame")
        self.settings_tab = ttk.Frame(self.notebook, style="Dark.TFrame", padding=20)

        self.notebook.add(self.main_tab, text="Orders & Calendar")
        self.notebook.add(self.settings_tab, text="Settings")

        main_tab = self.main_tab
        settings_tab = self.settings_tab
        settings_tab.columnconfigure(1, weight=1)

        username_label = ttk.Label(settings_tab, text="Username", style="Dark.TLabel")
        username_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.username_entry = ttk.Entry(settings_tab, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=0, column=1, sticky=tk.W)
        self.username_entry.bind("<Return>", self._on_enter_pressed)

        password_label = ttk.Label(settings_tab, text="Password", style="Dark.TLabel")
        password_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))

        self.password_entry = ttk.Entry(
            settings_tab,
            textvariable=self.password_var,
            show="*",
            width=30,
        )
        self.password_entry.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        self.password_entry.bind("<Return>", self._on_enter_pressed)

        button_frame = ttk.Frame(settings_tab, style="Dark.TFrame")
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

        self.status_message = ttk.Label(settings_tab, text="", style="Dark.TLabel")
        self.status_message.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))

        self.last_refresh_label = ttk.Label(
            settings_tab,
            textvariable=self.last_refresh_var,
            style="Dark.TLabel",
            font=("TkDefaultFont", 8),
        )
        self.last_refresh_label.grid(
            row=3, column=0, columnspan=3, sticky=tk.W, pady=(2, 0)
        )

        content_paned = ttk.Panedwindow(main_tab, orient=tk.HORIZONTAL, style="Dark.TPanedwindow")
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
        self.tree.bind("<KeyPress-Up>", self._on_tree_key_navigate)
        self.tree.bind("<KeyPress-Down>", self._on_tree_key_navigate)
        self.tree.bind("<KeyPress-Left>", self._on_tree_key_navigate)
        self.tree.bind("<KeyPress-Right>", self._on_tree_key_navigate)

        self._orders_tooltip = HoverTooltip(
            self.tree,
            (
                "Shift-click to select a range. "
                "Ctrl/Cmd-click to toggle items. "
                "Drag the selection onto a day to schedule."
            ),
        )

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(2, weight=1)

        # Allocate roughly one fifth of the horizontal space to the orders table
        # so it is about twenty-five percent narrower than before, leaving more
        # room for the calendar view.
        content_paned.add(table_frame, weight=1)

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

        content_paned.add(calendar_frame, weight=4)

        self._set_panedwindow_ratio(content_paned, 0, 0.2)

    def _focus_main_tab(self) -> None:
        notebook = getattr(self, "notebook", None)
        main_tab = getattr(self, "main_tab", None)
        if notebook is None or main_tab is None:
            return

        try:
            notebook.select(main_tab)
        except tk.TclError:
            return

        tree_widget = getattr(self, "tree", None)
        if isinstance(tree_widget, ttk.Treeview):
            try:
                tree_widget.focus_set()
            except tk.TclError:
                pass

    def _set_panedwindow_ratio(
        self, paned: ttk.Panedwindow, sash_index: int, fraction: float
    ) -> None:
        """Set the horizontal position of a paned window sash."""

        clamped_fraction = max(0.0, min(1.0, fraction))

        def adjust() -> None:
            try:
                paned.update_idletasks()
                total_width = paned.winfo_width()
            except tk.TclError:
                return

            if total_width <= 0:
                self.root.after(50, adjust)
                return

            target_x = int(total_width * clamped_fraction)

            try:
                paned.sash_place(sash_index, target_x, 0)
            except tk.TclError:
                return

        self.root.after_idle(adjust)

    def _focus_settings_tab(self) -> None:
        notebook = getattr(self, "notebook", None)
        settings_tab = getattr(self, "settings_tab", None)
        if notebook is None or settings_tab is None:
            return

        try:
            notebook.select(settings_tab)
        except tk.TclError:
            return

        username_entry = getattr(self, "username_entry", None)
        if isinstance(username_entry, ttk.Entry):
            try:
                username_entry.focus_set()
            except tk.TclError:
                pass

    def _show_about_message(self) -> None:
        status_message = f"{APP_TITLE} - {APP_DESCRIPTION}"
        try:
            messagebox.showinfo(
                title=ABOUT_TITLE,
                message=ABOUT_MESSAGE,
                parent=self.root,
            )
        except tk.TclError:
            pass

        self._set_status(SUCCESS_COLOR, status_message)

    def _render_calendar(self) -> None:
        year = self._current_year
        month = self._current_month
        today = dt.date.today()
        today_key = (today.year, today.month, today.day)
        first_of_month = dt.date(year, month, 1)
        self.month_label_var.set(first_of_month.strftime("%B %Y"))

        self._remove_calendar_hover()
        self._day_cell_pointer_hover = None
        previous_active = self._active_day_header

        for date_key, day_cell in list(self._day_cells.items()):
            self._save_day_notes(date_key)
            day_cell.frame.destroy()

        self._day_cells.clear()

        for child in self.calendar_grid.winfo_children():
            child.destroy()

        month_structure = calendar.Calendar().monthdatescalendar(year, month)

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
            for column_index, day_date in enumerate(week):
                is_current_month = day_date.month == month
                date_key = (day_date.year, day_date.month, day_date.day)
                day_number = day_date.day
                is_today = date_key == today_key
                border_color = TODAY_BORDER_COLOR if is_today else ACCENT_COLOR
                border_thickness = 2 if is_today else 1
                cell_background = (
                    DAY_CELL_BACKGROUND
                    if is_current_month
                    else ADJACENT_MONTH_DAY_CELL_BACKGROUND
                )
                header_fg = TEXT_COLOR if is_current_month else ADJACENT_MONTH_TEXT_COLOR
                notes_bg = (
                    NOTES_TEXT_BACKGROUND
                    if is_current_month
                    else ADJACENT_MONTH_NOTES_BACKGROUND
                )
                notes_fg = TEXT_COLOR if is_current_month else ADJACENT_MONTH_TEXT_COLOR
                orders_bg = (
                    ORDERS_LIST_BACKGROUND
                    if is_current_month
                    else ADJACENT_MONTH_ORDERS_BACKGROUND
                )
                orders_fg = TEXT_COLOR if is_current_month else ADJACENT_MONTH_TEXT_COLOR
                cell_frame = tk.Frame(
                    self.calendar_grid,
                    bg=cell_background,
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
                    text=str(day_number),
                    anchor="nw",
                    bg=cell_background,
                    fg=header_fg,
                    font=("TkDefaultFont", 10, "bold"),
                    padx=4,
                    pady=2,
                )
                header_label.grid(row=0, column=0, sticky="ew")
                try:
                    header_label.configure(takefocus=True)
                except tk.TclError:
                    pass

                header_label.bind(
                    "<Button-1>",
                    lambda event, key=date_key: self._on_day_header_click(event, key),
                )
                header_label.bind(
                    "<FocusIn>",
                    lambda event, key=date_key: self._on_day_header_focus(key),
                )
                header_label.bind(
                    "<Delete>",
                    lambda event, key=date_key: self._on_day_clear_request(event, key),
                )
                header_label.bind(
                    "<Destroy>",
                    lambda event, key=date_key: self._on_day_header_destroy(event, key),
                )
                cell_frame.bind(
                    "<Enter>",
                    lambda event, key=date_key: self._on_day_cell_pointer_enter(event, key),
                )
                cell_frame.bind(
                    "<Leave>",
                    lambda event, key=date_key: self._on_day_cell_pointer_leave(event, key),
                )
                header_label.bind(
                    "<Enter>",
                    lambda event, key=date_key: self._on_day_cell_pointer_enter(event, key),
                )
                header_label.bind(
                    "<Leave>",
                    lambda event, key=date_key: self._on_day_cell_pointer_leave(event, key),
                )

                notes_text = tk.Text(
                    cell_frame,
                    height=3,
                    wrap=tk.WORD,
                    bg=notes_bg,
                    fg=notes_fg,
                    insertbackground=notes_fg,
                    relief="flat",
                    bd=0,
                    undo=True,
                    autoseparators=True,
                    maxundo=-1,
                )
                notes_text.grid(row=1, column=0, sticky="nsew", padx=4, pady=(2, 2))
                notes_text.bind(
                    "<Control-z>",
                    lambda event: (self._invoke_text_widget_undo(event), "break")[1],
                )
                notes_text.bind(
                    "<Command-z>",
                    lambda event: (self._invoke_text_widget_undo(event), "break")[1],
                )
                notes_text.bind(
                    "<Control-Shift-Z>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )
                notes_text.bind(
                    "<Command-Shift-Z>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )
                notes_text.bind(
                    "<Control-y>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )
                notes_text.bind(
                    "<Command-y>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )
                notes_text.bind(
                    "<Control-Y>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )
                notes_text.bind(
                    "<Command-Y>",
                    lambda event: (self._invoke_text_widget_redo(event), "break")[1],
                )

                orders_list = tk.Listbox(
                    cell_frame,
                    height=3,
                    activestyle="none",
                    exportselection=False,
                    selectmode=tk.EXTENDED,
                )
                orders_list.configure(
                    bg=orders_bg,
                    fg=orders_fg,
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
                    default_bg=cell_background,
                    header_fg=header_fg,
                    notes_bg=notes_bg,
                    notes_fg=notes_fg,
                    orders_bg=orders_bg,
                    orders_fg=orders_fg,
                )
                day_cell.border_color = border_color
                day_cell.border_thickness = border_thickness
                day_cell.is_today = is_today
                day_cell.in_current_month = is_current_month
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
                orders_list.bind(
                    "<KeyPress-Up>",
                    lambda event, key=date_key: self._on_day_order_key_navigate(
                        event, key, -1
                    ),
                )
                orders_list.bind(
                    "<KeyPress-Left>",
                    lambda event, key=date_key: self._on_day_order_key_navigate(
                        event, key, -1
                    ),
                )
                orders_list.bind(
                    "<KeyPress-Down>",
                    lambda event, key=date_key: self._on_day_order_key_navigate(
                        event, key, 1
                    ),
                )
                orders_list.bind(
                    "<KeyPress-Right>",
                    lambda event, key=date_key: self._on_day_order_key_navigate(
                        event, key, 1
                    ),
                )

                self._update_day_cell_display(date_key)

        self._calendar_hover = None
        self._set_active_day_header(previous_active)

    def _refresh_day_header_selection(
        self, previous_active: DateKey | None = None
    ) -> None:
        current_active = self._active_day_header

        if previous_active and previous_active != current_active:
            self._apply_day_cell_base_style(previous_active)

        if current_active is None:
            return

        if current_active not in self._day_cells:
            self._active_day_header = None
            return

        self._apply_day_cell_base_style(current_active)

    def _set_active_day_header(self, date_key: DateKey | None) -> None:
        previous_active = self._active_day_header
        if previous_active == date_key:
            self._refresh_day_header_selection()
            return

        self._active_day_header = date_key
        self._refresh_day_header_selection(previous_active)

    def _on_day_header_click(self, event: tk.Event, date_key: DateKey) -> None:
        self._set_active_day_header(date_key)
        widget = getattr(event, "widget", None)
        if widget is not None:
            try:
                widget.focus_set()
            except tk.TclError:
                pass

    def _on_day_header_focus(self, date_key: DateKey) -> None:
        self._set_active_day_header(date_key)

    def _on_day_header_destroy(
        self, event: tk.Event | None, date_key: DateKey
    ) -> None:
        widget = getattr(event, "widget", None)
        if widget is not None:
            try:
                widget.unbind("<Delete>")
            except tk.TclError:
                pass
        if self._active_day_header == date_key:
            self._set_active_day_header(None)

    def _on_day_clear_request(
        self, event: tk.Event | None, date_key: DateKey | None = None
    ) -> None:
        if date_key is None:
            date_key = self._active_day_header
        if date_key is None:
            return

        assignments = self._calendar_assignments.get(date_key)
        if not assignments:
            self._set_status(FAIL_COLOR, "No orders scheduled for this day.")
            return

        snapshot = self._capture_assignments_state(date_key)
        self._push_undo_action(
            {"kind": "assignments", "dates": {date_key: snapshot}}
        )
        removed_count = len(assignments)
        self._calendar_assignments.pop(date_key, None)
        self._update_day_cell_display(date_key)
        self._schedule_state_save()

        date_label_text = self._format_date_label(date_key)
        plural = "s" if removed_count != 1 else ""
        message = f"Cleared {removed_count} order{plural} from {date_label_text}."
        self._set_status(SUCCESS_COLOR, message)

    def _apply_day_cell_base_style(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        try:
            if not day_cell.frame.winfo_exists():
                return
        except tk.TclError:
            return

        border_color = getattr(day_cell, "border_color", ACCENT_COLOR)
        border_thickness = getattr(day_cell, "border_thickness", 1)
        is_active = self._active_day_header == date_key

        if is_active:
            border_color = ACTIVE_DAY_BORDER_COLOR
            border_thickness = max(border_thickness, 3)

        try:
            day_cell.frame.configure(
                bg=day_cell.default_bg,
                highlightbackground=border_color,
                highlightcolor=border_color,
                highlightthickness=border_thickness,
            )
        except tk.TclError:
            return

        assignments = self._calendar_assignments.get(date_key, [])
        has_assignments = bool(assignments)

        base_header_fg = getattr(day_cell, "header_fg", TEXT_COLOR)
        base_orders_fg = getattr(day_cell, "orders_fg", TEXT_COLOR)
        base_notes_bg = getattr(day_cell, "notes_bg", NOTES_TEXT_BACKGROUND)
        base_notes_fg = getattr(day_cell, "notes_fg", TEXT_COLOR)
        base_orders_bg = getattr(day_cell, "orders_bg", ORDERS_LIST_BACKGROUND)

        header_bg: str
        header_fg: str
        orders_bg: str
        orders_fg: str

        if has_assignments and getattr(day_cell, "in_current_month", True):
            header_bg = ASSIGNMENT_HEADER_BACKGROUND
            header_fg = TEXT_COLOR
            orders_bg = ASSIGNMENT_LIST_BACKGROUND
            orders_fg = TEXT_COLOR
        else:
            if getattr(day_cell, "is_today", False):
                header_bg = TODAY_HEADER_BACKGROUND
                header_fg = TEXT_COLOR
            else:
                header_bg = day_cell.default_bg
                header_fg = base_header_fg
            orders_bg = base_orders_bg
            orders_fg = base_orders_fg

        if is_active:
            header_bg = ACTIVE_DAY_HEADER_BACKGROUND
            header_fg = TEXT_COLOR

        try:
            day_cell.header_label.configure(bg=header_bg, fg=header_fg)
        except tk.TclError:
            return

        try:
            day_cell.orders_list.configure(bg=orders_bg, fg=orders_fg)
        except tk.TclError:
            return

        try:
            day_cell.notes_text.configure(
                bg=base_notes_bg,
                fg=base_notes_fg,
                insertbackground=base_notes_fg,
            )
        except tk.TclError:
            return

        if (
            self._day_cell_pointer_hover == date_key
            and not self._drag_data.get("active")
            and self._calendar_hover != date_key
            and self._is_pointer_over_day_cell(day_cell)
        ):
            self._apply_day_cell_pointer_hover(date_key)

    def _on_day_cell_pointer_enter(
        self, event: tk.Event | None, date_key: DateKey
    ) -> None:
        if self._drag_data.get("active"):
            return
        self._apply_day_cell_pointer_hover(date_key, check_pointer=False)

    def _on_day_cell_pointer_leave(
        self, event: tk.Event | None, date_key: DateKey
    ) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        destination_widget = None
        if event is not None:
            try:
                x_root = int(getattr(event, "x_root"))
                y_root = int(getattr(event, "y_root"))
            except (TypeError, ValueError):
                x_root = y_root = None
            if x_root is not None and y_root is not None:
                try:
                    destination_widget = self.root.winfo_containing(x_root, y_root)
                except tk.TclError:
                    destination_widget = None

        if self._widget_belongs_to_day_cell(destination_widget, day_cell):
            return

        if self._day_cell_pointer_hover == date_key:
            self._day_cell_pointer_hover = None

        if self._calendar_hover == date_key and self._drag_data.get("active"):
            return

        self._apply_day_cell_base_style(date_key)

    def _widget_belongs_to_day_cell(
        self, widget: tk.Misc | None, day_cell: DayCell | None
    ) -> bool:
        if widget is None or day_cell is None:
            return False

        frame = day_cell.frame
        current = widget
        while current is not None:
            if current == frame:
                return True
            current = getattr(current, "master", None)
        return False

    def _is_pointer_over_day_cell(self, day_cell: DayCell) -> bool:
        try:
            pointer_x, pointer_y = self.root.winfo_pointerxy()
        except tk.TclError:
            return False

        try:
            widget = self.root.winfo_containing(pointer_x, pointer_y)
        except tk.TclError:
            return False

        return self._widget_belongs_to_day_cell(widget, day_cell)

    def _apply_day_cell_pointer_hover(
        self, date_key: DateKey, *, check_pointer: bool = True
    ) -> None:
        if self._drag_data.get("active"):
            return

        if self._calendar_hover == date_key:
            return

        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        try:
            if not day_cell.frame.winfo_exists():
                return
        except tk.TclError:
            return

        if check_pointer and not self._is_pointer_over_day_cell(day_cell):
            return

        border_color = getattr(day_cell, "border_color", ACCENT_COLOR)
        border_thickness = getattr(day_cell, "border_thickness", 1)
        if self._active_day_header == date_key:
            border_color = ACTIVE_DAY_BORDER_COLOR
            border_thickness = max(border_thickness, 3)

        try:
            day_cell.frame.configure(
                bg=DAY_CELL_HOVER_VALID,
                highlightbackground=border_color,
                highlightcolor=border_color,
                highlightthickness=border_thickness,
            )
            header_hover_fg = (
                TEXT_COLOR
                if getattr(day_cell, "in_current_month", True)
                else getattr(day_cell, "header_fg", ADJACENT_MONTH_TEXT_COLOR)
            )
            day_cell.header_label.configure(
                bg=DAY_CELL_HOVER_VALID, fg=header_hover_fg
            )
        except tk.TclError:
            return

        self._day_cell_pointer_hover = date_key

    def _save_day_notes(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return

        notes_widget = day_cell.notes_text
        new_text = notes_widget.get("1.0", "end-1c")
        has_new_text = bool(new_text.strip())
        existing_text = self._calendar_notes.get(date_key)

        if has_new_text:
            if isinstance(existing_text, str) and existing_text == new_text:
                return
        elif existing_text is None:
            return

        snapshot = self._capture_notes_state(date_key)
        self._push_undo_action(
            {
                "kind": "notes",
                "date_key": date_key,
                "previous": snapshot.get("previous"),
                "had_key": snapshot.get("had_key"),
            }
        )

        if has_new_text:
            self._calendar_notes[date_key] = new_text
        else:
            self._calendar_notes.pop(date_key, None)

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
            self._set_status(
                FAIL_COLOR, "Please select at least one order to remove."
            )
            return

        selected_indices: set[int] = set()
        for item in selection:
            try:
                index = int(item)
            except (TypeError, ValueError):
                continue
            selected_indices.add(index)

        valid_indices = sorted(
            index for index in selected_indices if 0 <= index < len(assignments)
        )

        if not valid_indices:
            self._set_status(
                FAIL_COLOR, "Unable to determine which orders to remove."
            )
            return

        snapshot = self._capture_assignments_state(date_key)
        self._push_undo_action(
            {"kind": "assignments", "dates": {date_key: snapshot}}
        )

        removed_assignments: list[Tuple[str, str]] = []
        for index in reversed(valid_indices):
            removed_assignments.append(assignments.pop(index))
        removed_assignments.reverse()

        if assignments:
            self._calendar_assignments[date_key] = assignments
        else:
            self._calendar_assignments.pop(date_key, None)

        orders_list.selection_clear(0, tk.END)
        self._update_day_cell_display(date_key)
        self._schedule_state_save()

        message = self._format_bulk_removal_message(date_key, removed_assignments)
        self._set_status(SUCCESS_COLOR, message)

    def _open_day_details(self, date_key: DateKey) -> None:
        day_cell = self._day_cells.get(date_key)
        cell_bounds: tuple[int, int, int, int] | None = None
        if day_cell is not None:
            cell_widget = getattr(day_cell, "frame", None)
            if cell_widget is not None:
                try:
                    if cell_widget.winfo_exists():
                        if cell_widget.winfo_ismapped():
                            cell_widget.update_idletasks()
                        else:
                            self.root.update_idletasks()
                        cell_x = cell_widget.winfo_rootx()
                        cell_y = cell_widget.winfo_rooty()
                        cell_width = cell_widget.winfo_width()
                        cell_height = cell_widget.winfo_height()
                        if cell_width <= 1:
                            cell_width = cell_widget.winfo_reqwidth()
                        if cell_height <= 1:
                            cell_height = cell_widget.winfo_reqheight()
                        cell_bounds = (
                            cell_x,
                            cell_y,
                            max(int(cell_width), 1),
                            max(int(cell_height), 1),
                        )
                except tk.TclError:
                    cell_bounds = None

        window = tk.Toplevel(self.root)
        window.withdraw()
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
            selectmode=tk.EXTENDED,
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

            snapshot = self._capture_assignments_state(date_key)
            self._push_undo_action(
                {"kind": "assignments", "dates": {date_key: snapshot}}
            )
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

            snapshot = self._capture_assignments_state(date_key)
            self._push_undo_action(
                {"kind": "assignments", "dates": {date_key: snapshot}}
            )
            removed_count = len(assignments)
            self._calendar_assignments.pop(date_key, None)
            self._update_day_cell_display(date_key)
            self._schedule_state_save()
            refresh_list()

            date_label_text = self._format_date_label(date_key)
            plural = "s" if removed_count != 1 else ""
            message = f"Cleared {removed_count} order{plural} from {date_label_text}."
            self._set_status(SUCCESS_COLOR, message)
            close_dialog()

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
        window.update_idletasks()
        try:
            self.root.update_idletasks()
        except tk.TclError:
            pass
        try:
            root_x = int(self.root.winfo_rootx())
            root_y = int(self.root.winfo_rooty())
            root_width = int(self.root.winfo_width())
            root_height = int(self.root.winfo_height())
        except tk.TclError:
            root_x = root_y = 0
            root_width = root_height = 0
        else:
            if root_width <= 1:
                root_width = int(self.root.winfo_reqwidth())
            if root_height <= 1:
                root_height = int(self.root.winfo_reqheight())

        window_width = int(window.winfo_width())
        window_height = int(window.winfo_height())
        if window_width <= 1:
            window_width = int(window.winfo_reqwidth())
        if window_height <= 1:
            window_height = int(window.winfo_reqheight())

        if cell_bounds is not None:
            cell_x, cell_y, cell_width, cell_height = cell_bounds
            target_x = int(cell_x + (cell_width - window_width) / 2)
            target_y = int(cell_y + (cell_height - window_height) / 2)
        elif root_width > 0 and root_height > 0:
            target_x = int(root_x + (root_width - window_width) / 2)
            target_y = int(root_y + (root_height - window_height) / 2)
        else:
            target_x = int(root_x)
            target_y = int(root_y)

        target_x, target_y = self._constrain_to_monitor(
            target_x,
            target_y,
            window_width,
            window_height,
        )
        window.geometry(f"+{target_x}+{target_y}")
        window.deiconify()
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
            "items": (),
            "values": (),
            "start_x": 0,
            "start_y": 0,
            "widget": None,
            "active": False,
            "source": None,
            "source_date_key": None,
            "source_indices": (),
            "source_assignments": (),
            "selection_snapshot": (),
            "selection_anchor": None,
            "focus_item": None,
            "active_index": None,
            "pending_tree_toggle": False,
        }

    def _normalize_tree_anchor(self, children: list[str]) -> str | None:
        anchor = self._tree_selection_anchor
        if anchor and anchor in children:
            return anchor
        self._tree_selection_anchor = None
        return None

    def _restore_drag_selection(self) -> None:
        snapshot = self._drag_data.get("selection_snapshot")
        if not isinstance(snapshot, (tuple, list)):
            return

        source = self._drag_data.get("source")
        if source == "tree":
            children = list(self.tree.get_children(""))
            preserved = [item for item in snapshot if item in children]
            try:
                if preserved:
                    self.tree.selection_set(preserved)
                else:
                    self.tree.selection_remove(self.tree.selection())
            except tk.TclError:
                return

            focus_item = self._drag_data.get("focus_item")
            if isinstance(focus_item, str) and focus_item in children:
                try:
                    self.tree.focus(focus_item)
                except tk.TclError:
                    pass

            anchor = self._drag_data.get("selection_anchor")
            if isinstance(anchor, str) and anchor in children:
                self._tree_selection_anchor = anchor
            elif not preserved:
                self._tree_selection_anchor = None
        elif source == "calendar":
            date_key = self._drag_data.get("source_date_key")
            day_cell = self._day_cells.get(date_key) if date_key is not None else None
            orders_list = getattr(day_cell, "orders_list", None)
            if orders_list is None:
                return

            try:
                orders_list.selection_clear(0, tk.END)
            except tk.TclError:
                return

            for raw_index in snapshot:
                try:
                    index = int(raw_index)
                except (TypeError, ValueError):
                    continue
                try:
                    orders_list.selection_set(index)
                except tk.TclError:
                    continue

            anchor_index = self._drag_data.get("selection_anchor")
            try:
                anchor_value = int(anchor_index)
            except (TypeError, ValueError):
                anchor_value = None

            if anchor_value is not None:
                try:
                    orders_list.selection_anchor(anchor_value)
                except tk.TclError:
                    pass

            active_index = self._drag_data.get("active_index")
            try:
                active_value = int(active_index)
            except (TypeError, ValueError):
                active_value = None

            if active_value is not None:
                try:
                    orders_list.activate(active_value)
                except tk.TclError:
                    pass

    def _on_tree_key_navigate(self, event: tk.Event) -> str | None:
        keysym = str(getattr(event, "keysym", ""))
        if keysym not in {"Up", "Down", "Left", "Right"}:
            return None

        direction = -1 if keysym in {"Up", "Left"} else 1
        return self._navigate_tree_with_keyboard(direction, event)

    def _navigate_tree_with_keyboard(
        self, direction: int, event: tk.Event | None
    ) -> str | None:
        children = list(self.tree.get_children(""))
        if not children:
            return "break"

        extend_selection = self._is_shift_pressed(event)

        focus_item = self.tree.focus()
        if focus_item not in children:
            focus_index = 0 if direction >= 0 else len(children) - 1
            focus_item = children[focus_index]
        else:
            focus_index = children.index(focus_item)

        new_index = focus_index + direction
        new_index = max(0, min(new_index, len(children) - 1))
        target = children[new_index]

        if extend_selection:
            anchor = self._normalize_tree_anchor(children)
            if anchor is None or anchor not in children:
                anchor = focus_item if focus_item in children else target
                if anchor not in children:
                    anchor = target
                self._tree_selection_anchor = anchor
            start_index = children.index(anchor)
            end_index = new_index
            if start_index > end_index:
                start_index, end_index = end_index, start_index
            selection = children[start_index : end_index + 1]
            try:
                self.tree.selection_set(selection)
            except tk.TclError:
                pass
        else:
            try:
                self.tree.selection_set((target,))
            except tk.TclError:
                pass
            self._tree_selection_anchor = target

        try:
            self.tree.focus(target)
        except tk.TclError:
            pass

        try:
            self.tree.see(target)
        except tk.TclError:
            pass

        return "break"

    def _on_order_press(self, event: tk.Event) -> str | None:
        self._end_drag()
        self._clear_other_day_selections(None)
        item_id = self.tree.identify_row(event.y)
        ctrl_pressed = self._is_control_pressed(event)
        shift_pressed = self._is_shift_pressed(event)

        self._drag_data["pending_tree_toggle"] = False

        if not item_id:
            if not ctrl_pressed and not shift_pressed:
                self.tree.selection_remove(self.tree.selection())
                self._tree_selection_anchor = None
            self._drag_data.update(
                {
                    "items": (),
                    "values": (),
                    "start_x": event.x_root,
                    "start_y": event.y_root,
                    "widget": None,
                    "active": False,
                    "source": "tree",
                    "source_date_key": None,
                    "source_indices": (),
                    "source_assignments": (),
                    "selection_snapshot": (),
                    "selection_anchor": self._tree_selection_anchor,
                    "focus_item": None,
                    "active_index": None,
                }
            )
            return "break"

        children = list(self.tree.get_children(""))
        anchor = self._normalize_tree_anchor(children)
        current_selection = set(self.tree.selection())

        if shift_pressed and children:
            if anchor is None:
                anchor = item_id
            if anchor not in children:
                anchor = item_id
            try:
                start_index = children.index(anchor)
                end_index = children.index(item_id)
            except ValueError:
                selection = (item_id,)
            else:
                if start_index > end_index:
                    start_index, end_index = end_index, start_index
                selection = tuple(children[start_index : end_index + 1])
            try:
                self.tree.selection_set(selection)
            except tk.TclError:
                pass
            self._tree_selection_anchor = anchor
        elif ctrl_pressed:
            if item_id in current_selection:
                self._drag_data["pending_tree_toggle"] = True
            else:
                try:
                    self.tree.selection_add(item_id)
                except tk.TclError:
                    pass
                self._drag_data["pending_tree_toggle"] = False
                if anchor is None:
                    self._tree_selection_anchor = item_id
        else:
            if item_id in current_selection and current_selection:
                self._tree_selection_anchor = item_id
            else:
                try:
                    self.tree.selection_set((item_id,))
                except tk.TclError:
                    pass
                self._tree_selection_anchor = item_id

        try:
            self.tree.focus(item_id)
        except tk.TclError:
            pass

        selected_set = set(self.tree.selection())
        ordered_selection = tuple(
            child for child in children if child in selected_set
        )

        normalized_values: list[Tuple[str, str]] = []
        for selected_id in ordered_selection:
            values = self.tree.item(selected_id, "values") or ()
            normalized_values.append(self._normalize_assignment(values))

        assignments_tuple = tuple(normalized_values)

        self._drag_data.update(
            {
                "items": ordered_selection,
                "values": assignments_tuple,
                "start_x": event.x_root,
                "start_y": event.y_root,
                "widget": None,
                "active": False,
                "source": "tree",
                "source_date_key": None,
                "source_indices": (),
                "source_assignments": assignments_tuple,
                "selection_snapshot": ordered_selection,
                "selection_anchor": self._tree_selection_anchor,
                "focus_item": item_id,
                "active_index": None,
            }
        )

        return "break"

    def _on_order_drag(self, event: tk.Event) -> str | None:
        items = self._drag_data.get("items")
        if not items:
            return None

        if self._drag_data.get("pending_tree_toggle"):
            pressed_item = self._drag_data.get("focus_item")
            tree_widget = getattr(self, "tree", None)
            if tree_widget is not None and isinstance(pressed_item, str):
                try:
                    hovered_item = tree_widget.identify_row(event.y)
                except tk.TclError:
                    hovered_item = ""
                if hovered_item != pressed_item:
                    self._drag_data["pending_tree_toggle"] = False

        self._restore_drag_selection()

        x_root = int(getattr(event, "x_root", 0))
        y_root = int(getattr(event, "y_root", 0))

        drag_active = bool(self._drag_data.get("active"))
        if not drag_active:
            start_x = int(self._drag_data.get("start_x", x_root))
            start_y = int(self._drag_data.get("start_y", y_root))
            if (
                abs(x_root - start_x) >= DRAG_THRESHOLD
                or abs(y_root - start_y) >= DRAG_THRESHOLD
            ):
                self._begin_drag()
                drag_active = bool(self._drag_data.get("active"))
                if drag_active:
                    self._restore_drag_selection()

        if drag_active:
            self._position_drag_window(x_root, y_root)

        target_info = self._detect_calendar_target(x_root, y_root)
        self._update_calendar_hover(target_info)
        return "break"

    def _on_order_release(self, event: tk.Event) -> str | None:
        items = self._drag_data.get("items")
        drag_was_active = bool(self._drag_data.get("active"))

        if items:
            self._restore_drag_selection()

        if not items:
            if drag_was_active:
                self._end_drag()
                return "break"
            self._end_drag()
            return "break"

        if not drag_was_active:
            if self._drag_data.get("pending_tree_toggle"):
                pressed_item = self._drag_data.get("focus_item")
                tree_widget = getattr(self, "tree", None)
                same_row = False
                if tree_widget is not None and isinstance(pressed_item, str):
                    try:
                        hovered_item = tree_widget.identify_row(event.y)
                    except tk.TclError:
                        hovered_item = ""
                    same_row = hovered_item == pressed_item
                if same_row and tree_widget is not None:
                    try:
                        tree_widget.selection_remove(pressed_item)
                    except tk.TclError:
                        pass
                    try:
                        remaining = tree_widget.selection()
                    except tk.TclError:
                        remaining = ()
                    if not remaining:
                        self._tree_selection_anchor = None
                self._drag_data["pending_tree_toggle"] = False
            self._end_drag()
            return "break"

        target_info = self._detect_calendar_target(event.x_root, event.y_root)
        normalized_key: DateKey | None = None
        if target_info:
            raw_key = target_info.get("date_key")
            if isinstance(raw_key, tuple) and len(raw_key) == 3:
                try:
                    normalized_key = (
                        int(raw_key[0]),
                        int(raw_key[1]),
                        int(raw_key[2]),
                    )
                except (TypeError, ValueError):
                    normalized_key = None

        if normalized_key is not None:
            raw_orders = self._drag_data.get("values", ())
            if not isinstance(raw_orders, (tuple, list)) or not raw_orders:
                self._queue.put(
                    (
                        "calendar_drop",
                        False,
                        "Unable to determine which order was dragged.",
                        None,
                    )
                )
            else:
                normalized_orders = tuple(
                    self._normalize_assignment(order)
                    for order in raw_orders
                )
                target_label = self._format_date_label(normalized_key)
                message = self._format_assignment_move_message(
                    normalized_orders,
                    target_label,
                )
                payload: dict[str, object] = {
                    "date_key": normalized_key,
                    "orders": normalized_orders,
                    "source_kind": "tree",
                }
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
        return "break"

    def _on_day_order_press(self, event: tk.Event, date_key: DateKey) -> str | None:
        self._end_drag()
        self._clear_other_day_selections(date_key)

        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return "break"

        orders_list = day_cell.orders_list
        assignments = self._calendar_assignments.get(date_key, [])
        try:
            normalized_date_key: DateKey = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            normalized_date_key = date_key
        if not assignments:
            orders_list.selection_clear(0, tk.END)
            self._day_selection_anchor.pop(normalized_date_key, None)
            self._drag_data.update(
                {
                    "items": (),
                    "values": (),
                    "start_x": event.x_root,
                    "start_y": event.y_root,
                    "widget": None,
                    "active": False,
                    "source": "calendar",
                    "source_date_key": date_key,
                    "source_indices": (),
                    "source_assignments": (),
                    "selection_snapshot": (),
                    "selection_anchor": None,
                    "focus_item": None,
                    "active_index": None,
                }
            )
            return "break"

        try:
            index = int(orders_list.nearest(event.y))
        except (tk.TclError, ValueError):
            return "break"

        ctrl_pressed = self._is_control_pressed(event)
        shift_pressed = self._is_shift_pressed(event)

        if index < 0 or index >= len(assignments):
            if not ctrl_pressed and not shift_pressed:
                orders_list.selection_clear(0, tk.END)
            return "break"

        try:
            bbox = orders_list.bbox(index)
        except tk.TclError:
            bbox = None

        if not bbox or not (bbox[1] <= event.y <= bbox[1] + bbox[3]):
            if not ctrl_pressed and not shift_pressed:
                orders_list.selection_clear(0, tk.END)
            return "break"

        if shift_pressed:
            anchor = self._day_selection_anchor.get(normalized_date_key)
            if not isinstance(anchor, int) or not (0 <= anchor < len(assignments)):
                anchor = index
                self._day_selection_anchor[normalized_date_key] = anchor
            start = min(anchor, index)
            end = max(anchor, index)
            orders_list.selection_clear(0, tk.END)
            for idx in range(start, end + 1):
                orders_list.selection_set(idx)
            orders_list.selection_anchor(anchor)
        elif ctrl_pressed:
            if orders_list.selection_includes(index):
                orders_list.selection_clear(index)
            else:
                orders_list.selection_set(index)
                orders_list.selection_anchor(index)
        else:
            if not orders_list.selection_includes(index):
                orders_list.selection_clear(0, tk.END)
                orders_list.selection_set(index)
            orders_list.selection_anchor(index)
            self._day_selection_anchor[normalized_date_key] = index

        orders_list.activate(index)

        try:
            anchor_value = int(orders_list.index(tk.ANCHOR))
        except (tk.TclError, ValueError):
            anchor_value = None

        current_selection = orders_list.curselection()
        selected_indices: tuple[int, ...] = tuple(int(i) for i in current_selection)

        normalized_assignments: list[Tuple[str, str]] = []
        for idx in selected_indices:
            if 0 <= idx < len(assignments):
                normalized_assignments.append(
                    self._normalize_assignment(assignments[idx])
                )

        assignments_tuple = tuple(normalized_assignments)

        self._drag_data.update(
            {
                "items": selected_indices,
                "values": assignments_tuple,
                "start_x": event.x_root,
                "start_y": event.y_root,
                "widget": None,
                "active": False,
                "source": "calendar",
                "source_date_key": normalized_date_key,
                "source_indices": selected_indices,
                "source_assignments": assignments_tuple,
                "selection_snapshot": selected_indices,
                "selection_anchor": anchor_value,
                "focus_item": None,
                "active_index": index,
            }
        )

        return "break"

    def _on_day_order_key_navigate(
        self, event: tk.Event, date_key: DateKey, direction: int
    ) -> str | None:
        day_cell = self._day_cells.get(date_key)
        if not day_cell:
            return "break"

        orders_list = day_cell.orders_list
        try:
            size = int(orders_list.size())
        except (tk.TclError, ValueError):
            return "break"

        try:
            normalized_date_key: DateKey = (
                int(date_key[0]),
                int(date_key[1]),
                int(date_key[2]),
            )
        except (TypeError, ValueError):
            normalized_date_key = date_key

        if size <= 0:
            self._day_selection_anchor.pop(normalized_date_key, None)
            return "break"

        try:
            active_index = int(orders_list.index(tk.ACTIVE))
        except (tk.TclError, ValueError):
            active_index = None

        if active_index is None or not (0 <= active_index < size):
            selection = orders_list.curselection()
            if selection:
                try:
                    active_index = int(selection[-1])
                except (TypeError, ValueError):
                    active_index = None

        if active_index is None or not (0 <= active_index < size):
            anchor_hint = self._day_selection_anchor.get(normalized_date_key)
            if isinstance(anchor_hint, int) and 0 <= anchor_hint < size:
                active_index = anchor_hint

        if active_index is None:
            active_index = 0 if direction >= 0 else size - 1

        target_index = active_index + direction
        if target_index < 0:
            target_index = 0
        elif target_index >= size:
            target_index = size - 1

        extend_selection = self._is_shift_pressed(event)
        ctrl_pressed = self._is_control_pressed(event)

        if extend_selection:
            anchor = self._day_selection_anchor.get(normalized_date_key)
            if not isinstance(anchor, int) or not (0 <= anchor < size):
                anchor = active_index if 0 <= active_index < size else target_index
                self._day_selection_anchor[normalized_date_key] = anchor
            start = min(anchor, target_index)
            end = max(anchor, target_index)
            orders_list.selection_clear(0, tk.END)
            for idx in range(start, end + 1):
                orders_list.selection_set(idx)
            try:
                orders_list.selection_anchor(anchor)
            except tk.TclError:
                pass
        elif ctrl_pressed:
            try:
                orders_list.activate(target_index)
                orders_list.see(target_index)
            except tk.TclError:
                pass
            return "break"
        else:
            orders_list.selection_clear(0, tk.END)
            orders_list.selection_set(target_index)
            try:
                orders_list.selection_anchor(target_index)
            except tk.TclError:
                pass
            self._day_selection_anchor[normalized_date_key] = target_index

        try:
            orders_list.activate(target_index)
        except tk.TclError:
            pass

        try:
            orders_list.see(target_index)
        except tk.TclError:
            pass

        return "break"

    def _on_day_order_drag(self, event: tk.Event, date_key: DateKey) -> str | None:
        items = self._drag_data.get("items")
        if not items:
            return None

        self._restore_drag_selection()

        x_root = int(getattr(event, "x_root", 0))
        y_root = int(getattr(event, "y_root", 0))

        drag_active = bool(self._drag_data.get("active"))
        if not drag_active:
            start_x = int(self._drag_data.get("start_x", x_root))
            start_y = int(self._drag_data.get("start_y", y_root))
            if (
                abs(x_root - start_x) >= DRAG_THRESHOLD
                or abs(y_root - start_y) >= DRAG_THRESHOLD
            ):
                self._begin_drag()
                drag_active = bool(self._drag_data.get("active"))
                if drag_active:
                    self._restore_drag_selection()

        if drag_active:
            self._position_drag_window(x_root, y_root)

        target_info = self._detect_calendar_target(x_root, y_root)
        self._update_calendar_hover(target_info)
        return "break"

    def _on_day_order_release(self, event: tk.Event, date_key: DateKey) -> str | None:
        drag_was_active = bool(self._drag_data.get("active"))

        if self._drag_data.get("source") != "calendar":
            self._end_drag()
            return "break" if drag_was_active else None

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
            return "break" if drag_was_active else None

        items = self._drag_data.get("items")
        if items:
            self._restore_drag_selection()
        if not items:
            self._end_drag()
            return "break" if drag_was_active else None

        if not drag_was_active:
            self._end_drag()
            return None

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
            raw_orders = self._drag_data.get("values", ())
            if not isinstance(raw_orders, (tuple, list)) or not raw_orders:
                self._queue.put(
                    (
                        "calendar_drop",
                        False,
                        "Unable to determine which order was dragged.",
                        None,
                    )
                )
            else:
                normalized_orders = tuple(
                    self._normalize_assignment(order)
                    for order in raw_orders
                )

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
                    else None
                )

                message = self._format_assignment_move_message(
                    normalized_orders,
                    target_label,
                    source_label=source_label,
                    same_day=normalized_source == normalized_key,
                )

                payload: dict[str, object] = {
                    "date_key": normalized_key,
                    "orders": normalized_orders,
                    "source_kind": "calendar",
                }

                if normalized_source is not None:
                    payload["source_date_key"] = normalized_source

                    source_indices = self._drag_data.get("source_indices", ())
                    if isinstance(source_indices, (tuple, list)):
                        payload["source_indices"] = tuple(
                            int(index) for index in source_indices
                        )

                    source_assignments = self._drag_data.get(
                        "source_assignments", ()
                    )
                    if isinstance(source_assignments, (tuple, list)):
                        payload["source_orders"] = tuple(
                            self._normalize_assignment(order)
                            for order in source_assignments
                        )

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
        return "break"

    def _begin_drag(self) -> None:
        items = self._drag_data.get("items")
        values = self._drag_data.get("values", ())
        if not items or not isinstance(values, (tuple, list)):
            return

        self._drag_data["pending_tree_toggle"] = False

        normalized_orders = [
            self._normalize_assignment(value) for value in values
        ]
        labels = [self._format_assignment_label(order) for order in normalized_orders]
        if not labels:
            label_text = ""
        elif len(labels) == 1:
            label_text = labels[0]
        else:
            preview = ", ".join(labels[:3])
            if len(labels) > 3:
                preview += ", ..."
            label_text = f"{len(labels)} orders: {preview}"

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
        self._drag_data["values"] = tuple(normalized_orders)
        self._drag_data["active"] = True

        start_x = int(self._drag_data.get("start_x", 0))
        start_y = int(self._drag_data.get("start_y", 0))
        self._position_drag_window(start_x, start_y)

    def _position_drag_window(self, x_root: int, y_root: int) -> None:
        widget = self._drag_data.get("widget")
        if widget is None:
            return
        try:
            widget.update_idletasks()
        except tk.TclError:
            pass

        try:
            window_width = int(widget.winfo_width())
            window_height = int(widget.winfo_height())
        except tk.TclError:
            window_width = window_height = 0

        if window_width <= 1:
            try:
                window_width = int(widget.winfo_reqwidth())
            except (tk.TclError, ValueError):
                window_width = 1
        if window_height <= 1:
            try:
                window_height = int(widget.winfo_reqheight())
            except (tk.TclError, ValueError):
                window_height = 1

        base_x = int(x_root) + 16
        base_y = int(y_root) + 16
        target_x, target_y = self._constrain_to_monitor(
            base_x,
            base_y,
            window_width,
            window_height,
        )
        widget.geometry(f"+{target_x}+{target_y}")

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
        header_fg = (
            TEXT_COLOR
            if getattr(day_cell, "in_current_month", True)
            else getattr(day_cell, "header_fg", ADJACENT_MONTH_TEXT_COLOR)
        )
        day_cell.header_label.configure(bg=hover_color, fg=header_fg)

        self._calendar_hover = date_key

    def _remove_calendar_hover(self) -> None:
        if self._calendar_hover is None:
            return

        date_key = self._calendar_hover
        self._apply_day_cell_base_style(date_key)

        self._calendar_hover = None

    def _clear_other_day_selections(self, active_key: DateKey | None) -> None:
        """Clear selections on all day listboxes except the active one."""

        normalized_active: DateKey | None = None
        if isinstance(active_key, (tuple, list)) and len(active_key) == 3:
            try:
                normalized_active = (
                    int(active_key[0]),
                    int(active_key[1]),
                    int(active_key[2]),
                )
            except (TypeError, ValueError):
                try:
                    normalized_active = tuple(active_key)  # type: ignore[arg-type]
                except TypeError:
                    normalized_active = None

        for key, day_cell in self._day_cells.items():
            if active_key is not None and key == active_key:
                continue
            if normalized_active is not None and key == normalized_active:
                continue

            orders_list = getattr(day_cell, "orders_list", None)
            if orders_list is None:
                continue

            try:
                orders_list.selection_clear(0, tk.END)
            except tk.TclError:
                continue

    def _clear_tree_selection(self) -> None:
        """Clear the selection state for the orders tree."""

        tree = getattr(self, "tree", None)
        if tree is None:
            return

        try:
            selected_items = tree.selection()
        except tk.TclError:
            selected_items = ()

        if selected_items:
            try:
                tree.selection_remove(selected_items)
            except tk.TclError:
                pass

        try:
            tree.selection_anchor("")
        except tk.TclError:
            pass

        try:
            tree.focus("")
        except tk.TclError:
            pass

        self._tree_selection_anchor = None

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
        orders = payload.get("orders")
        if (
            not isinstance(date_key, (tuple, list))
            or len(date_key) != 3
            or not isinstance(orders, (tuple, list))
            or not orders
        ):
            return

        try:
            normalized_key = (int(date_key[0]), int(date_key[1]), int(date_key[2]))
        except (TypeError, ValueError):
            return

        normalized_orders = [
            self._normalize_assignment(order) for order in orders
        ]
        if not normalized_orders:
            return

        source_kind = payload.get("source_kind")
        clear_tree_after_drop = (
            source_kind == "tree"
            or (source_kind is None and "source_date_key" not in payload)
        )

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

        target_snapshot = self._capture_assignments_state(normalized_key)
        source_snapshot: dict[str, Any] | None = (
            self._capture_assignments_state(normalized_source)
            if normalized_source is not None
            else None
        )
        deselect_after_cross_day_move = (
            source_kind == "calendar"
            and normalized_source is not None
            and normalized_source != normalized_key
        )

        def current_selection(listbox: tk.Listbox | None) -> list[int]:
            if listbox is None:
                return []
            try:
                return [int(i) for i in listbox.curselection()]
            except (tk.TclError, TypeError, ValueError):
                return []

        target_day_cell = self._day_cells.get(normalized_key)
        target_listbox = (
            target_day_cell.orders_list if target_day_cell else None
        )
        target_selection_before = current_selection(target_listbox)

        source_listbox: tk.Listbox | None = None
        source_selection_before: list[int] = []
        if normalized_source is not None:
            source_day_cell = self._day_cells.get(normalized_source)
            source_listbox = (
                source_day_cell.orders_list if source_day_cell else None
            )
            source_selection_before = current_selection(source_listbox)

        target_assignments = self._calendar_assignments.setdefault(
            normalized_key, []
        )

        if normalized_source is not None and normalized_source == normalized_key:
            source_assignments_ref = target_assignments
        elif normalized_source is not None:
            source_assignments_ref = self._calendar_assignments.get(
                normalized_source
            )
        else:
            source_assignments_ref = None

        raw_source_orders = payload.get("source_orders")
        normalized_source_orders: list[Tuple[str, str]] = []
        if isinstance(raw_source_orders, (tuple, list)):
            normalized_source_orders = [
                self._normalize_assignment(order)
                for order in raw_source_orders
            ]
        if not normalized_source_orders:
            normalized_source_orders = list(normalized_orders)

        source_indices_raw = payload.get("source_indices")
        index_hints: list[int | None] = []
        if isinstance(source_indices_raw, (tuple, list)):
            for raw_index in source_indices_raw:
                try:
                    index_hints.append(int(raw_index))
                except (TypeError, ValueError):
                    index_hints.append(None)

        removed_indices: list[int] = []
        removed_from_source = False
        if normalized_source is not None and source_assignments_ref:
            assignments_list = source_assignments_ref
            used_indices: set[int] = set()
            for position, order in enumerate(normalized_source_orders):
                index_hint = index_hints[position] if position < len(index_hints) else None
                removal_index: int | None = None
                if (
                    isinstance(index_hint, int)
                    and 0 <= index_hint < len(assignments_list)
                    and index_hint not in used_indices
                ):
                    if self._normalize_assignment(assignments_list[index_hint]) == order:
                        removal_index = index_hint
                if removal_index is None:
                    for idx, assignment in enumerate(assignments_list):
                        if idx in used_indices:
                            continue
                        if self._normalize_assignment(assignment) == order:
                            removal_index = idx
                            break
                if removal_index is not None:
                    used_indices.add(removal_index)
                    removed_indices.append(removal_index)

            if removed_indices:
                for idx in sorted(removed_indices, reverse=True):
                    if 0 <= idx < len(assignments_list):
                        assignments_list.pop(idx)
                removed_from_source = True
                if normalized_source != normalized_key:
                    if assignments_list:
                        self._calendar_assignments[normalized_source] = assignments_list
                    else:
                        self._calendar_assignments.pop(normalized_source, None)

        added_to_target = False
        target_indices: list[int] = []
        for order in normalized_orders:
            if order in target_assignments:
                index = target_assignments.index(order)
            else:
                target_assignments.append(order)
                added_to_target = True
                index = len(target_assignments) - 1
            target_indices.append(index)

        removed_sorted = sorted(removed_indices)

        def adjust_selection(selection: list[int], removed: list[int]) -> list[int]:
            if not selection or not removed:
                return list(selection)
            adjusted: list[int] = []
            removed_sorted_local = sorted(removed)
            removed_set = set(removed_sorted_local)
            for index in selection:
                if index in removed_set:
                    continue
                shift = sum(1 for value in removed_sorted_local if value < index)
                new_index = index - shift
                if new_index >= 0:
                    adjusted.append(new_index)
            return adjusted

        if normalized_source == normalized_key:
            base_target_selection = adjust_selection(
                target_selection_before, removed_sorted
            )
        else:
            base_target_selection = list(target_selection_before)

        combined_target_selection: list[int] = []
        for idx in base_target_selection:
            if idx not in combined_target_selection:
                combined_target_selection.append(idx)
        for idx in target_indices:
            if idx not in combined_target_selection:
                combined_target_selection.append(idx)
        if deselect_after_cross_day_move:
            combined_target_selection = []

        source_selection_after: list[int] = []
        if (
            normalized_source is not None
            and normalized_source != normalized_key
            and source_selection_before
        ):
            source_selection_after = adjust_selection(
                source_selection_before, removed_sorted
            )
        if deselect_after_cross_day_move:
            source_selection_after = []
            self._day_selection_anchor.pop(normalized_key, None)
            self._day_selection_anchor.pop(normalized_source, None)

        self._update_day_cell_display(normalized_key)
        if normalized_source is not None and normalized_source != normalized_key:
            self._update_day_cell_display(normalized_source)

        if clear_tree_after_drop and (added_to_target or removed_from_source):
            self._clear_tree_selection()

        def apply_selection(listbox: tk.Listbox | None, indices: list[int]) -> None:
            if listbox is None:
                return
            try:
                size = listbox.size()
            except tk.TclError:
                return
            listbox.selection_clear(0, tk.END)
            valid_indices = [idx for idx in indices if 0 <= idx < size]
            for idx in valid_indices:
                listbox.selection_set(idx)
            if valid_indices:
                anchor_index = valid_indices[-1]
                try:
                    listbox.selection_anchor(anchor_index)
                    listbox.activate(anchor_index)
                except tk.TclError:
                    pass

        apply_selection(target_listbox, combined_target_selection)
        if normalized_source is not None and normalized_source != normalized_key:
            apply_selection(source_listbox, source_selection_after)

        self._clear_other_day_selections(normalized_key)

        undo_entries: dict[DateKey, dict[str, Any]] = {}
        if added_to_target or (
            removed_from_source and normalized_source == normalized_key
        ):
            undo_entries[normalized_key] = dict(target_snapshot)
        if removed_from_source and (
            normalized_source is not None and normalized_source != normalized_key
        ):
            entry_snapshot = source_snapshot or {"had_key": False, "previous": None}
            undo_entries[normalized_source] = dict(entry_snapshot)

        if undo_entries:
            self._push_undo_action({"kind": "assignments", "dates": undo_entries})

        if added_to_target or removed_from_source:
            self._schedule_state_save()

    def _assign_order_to_day(
        self,
        date_key: DateKey,
        order_values: Tuple[str, ...],
        *,
        push_undo: bool = True,
    ) -> bool:
        normalized = self._normalize_assignment(order_values)

        snapshot = self._capture_assignments_state(date_key)
        assignments = self._calendar_assignments.setdefault(date_key, [])

        if normalized in assignments:
            self._update_day_cell_display(date_key)
            return False

        if push_undo:
            self._push_undo_action(
                {"kind": "assignments", "dates": {date_key: snapshot}}
            )

        assignments.append(normalized)
        self._update_day_cell_display(date_key)
        self._schedule_state_save()
        return True

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

    def _format_assignment_move_message(
        self,
        assignments: Iterable[Tuple[str, str]],
        target_label: str,
        *,
        source_label: str | None = None,
        same_day: bool = False,
    ) -> str:
        assignment_list = list(assignments)
        if not assignment_list:
            return ""

        count = len(assignment_list)
        if count == 1:
            order_number = assignment_list[0][0].strip()
            company = assignment_list[0][1].strip()
            order_label = "order"
            if order_number:
                order_label += f" {order_number}"
            if company:
                order_label += f" ({company})"

            if source_label and not same_day:
                return f"Moved {order_label} from {source_label} to {target_label}."
            if source_label:
                capitalized = order_label[0].upper() + order_label[1:]
                return f"{capitalized} remains scheduled for {target_label}."
            return f"Assigned {order_label} to {target_label}."

        order_phrase = f"{count} orders"
        if source_label and not same_day:
            return f"Moved {order_phrase} from {source_label} to {target_label}."
        if source_label:
            capitalized = order_phrase[0].upper() + order_phrase[1:]
            return f"{capitalized} remain scheduled for {target_label}."
        return f"Assigned {order_phrase} to {target_label}."

    def _format_bulk_removal_message(
        self, date_key: DateKey, assignments: Iterable[Tuple[str, str]]
    ) -> str:
        assignment_list = list(assignments)
        if not assignment_list:
            return ""

        if len(assignment_list) == 1:
            return self._format_removal_message(date_key, assignment_list[0])

        count = len(assignment_list)
        date_label = self._format_date_label(date_key)
        return f"Removed {count} orders from {date_label}."

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

        self._tree_selection_anchor = None

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
