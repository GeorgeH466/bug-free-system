"""Custom grid-based monthly calendar and reusable Tkinter date picker."""
from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from typing import Any, Iterable, Mapping

import database
from utils import format_date, normalize_date, parse_date

try:  # Keep pure helpers importable on headless CI systems without Tkinter.
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    messagebox = None
    ttk = None

_FrameBase = ttk.Frame if ttk is not None else object
_DialogBase = tk.Toplevel if tk is not None else object
WEEKDAYS = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
SCHEDULED_STATUSES = {"scheduled", "pending", "overdue"}

COLORS = {
    "empty": "#FFFFFF",
    "outside": "#F1F3F4",
    "today": "#FFF2A8",
    "event": "#DDEBFF",
    "overdue": "#FCE8E6",
    "text": "#1F2937",
    "muted": "#6B7280",
    "today_border": "#EAB308",
    "event_border": "#5B8DEF",
    "overdue_border": "#DC2626",
    "default_border": "#D1D5DB",
}


def month_grid(year: int, month: int) -> list[list[date]]:
    """Return six Monday-first weeks covering the requested month."""
    first = date(year, month, 1)
    start = first - timedelta(days=first.weekday())
    return [[start + timedelta(days=row * 7 + column) for column in range(7)] for row in range(6)]


def is_overdue_event(event: Mapping[str, Any], today: date | None = None) -> bool:
    reference = today or date.today()
    try:
        event_day = parse_date(event.get("event_date"))
    except (TypeError, ValueError):
        return False
    status = str(event.get("status") or "Scheduled").lower()
    return bool(event_day and event_day < reference and status in SCHEDULED_STATUSES)


def day_visual_state(
    day: date,
    events: Iterable[Mapping[str, Any]],
    *,
    today: date | None = None,
    display_month: int | None = None,
) -> dict[str, Any]:
    """Pure color/border decision helper, also used by tests."""
    reference = today or date.today()
    month = display_month or day.month
    event_list = list(events)
    overdue = any(is_overdue_event(event, reference) for event in event_list)
    is_today = day == reference
    if overdue:
        background, border, width = COLORS["overdue"], COLORS["today_border"] if is_today else COLORS["overdue_border"], 3 if is_today else 2
    elif event_list:
        background, border, width = COLORS["event"], COLORS["today_border"] if is_today else COLORS["event_border"], 3 if is_today else 2
    elif is_today:
        background, border, width = COLORS["today"], COLORS["today_border"], 3
    elif day.month != month:
        background, border, width = COLORS["outside"], COLORS["default_border"], 1
    else:
        background, border, width = COLORS["empty"], COLORS["default_border"], 1
    return {
        "background": background,
        "foreground": COLORS["muted"] if day.month != month else COLORS["text"],
        "border": border,
        "border_width": width,
        "event_count": len(event_list),
        "overdue": overdue,
        "is_today": is_today,
    }


def group_events_by_date(events: Iterable[Mapping[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        key = normalize_date(event.get("event_date"))
        if key:
            grouped[key].append(dict(event))
    return dict(grouped)


class DayEventsDialog(_DialogBase):
    """Popup showing events and service due dates for one calendar day."""

    def __init__(self, parent: Any, day: date, events: list[dict[str, Any]], due_services: list[dict[str, Any]]):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the calendar.")
        super().__init__(parent)
        self.day = day
        self.title(f"Events — {format_date(day)}")
        self.geometry("760x420")
        self.transient(parent)
        self.resizable(True, True)
        container = ttk.Frame(self, padding=14)
        container.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        ttk.Label(container, text=f"Schedule for {format_date(day)}", style="Heading.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )
        columns = ("source", "customer", "type", "time", "status", "details")
        self.tree = ttk.Treeview(container, columns=columns, show="headings")
        headings = (
            ("source", "Source", 95), ("customer", "Customer", 180), ("type", "Type", 110),
            ("time", "Time", 80), ("status", "Status", 100), ("details", "Details", 220),
        )
        for column, heading, width in headings:
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        for event in events:
            status = str(event.get("status") or "Scheduled")
            if is_overdue_event(event, date.today()):
                status = "Overdue"
            self.tree.insert("", "end", values=(
                "Event",
                event.get("customer_name") or "Unknown customer",
                event.get("event_type") or "Service",
                event.get("event_time") or "",
                status,
                event.get("event_notes") or "",
            ))
        for customer in due_services:
            self.tree.insert("", "end", values=(
                "Customer schedule",
                customer.get("customer_name") or "",
                customer.get("service_type") or "Service due",
                "",
                "Due",
                customer.get("customer_location") or "",
            ))
        if not self.tree.get_children():
            self.tree.insert("", "end", values=("", "No events scheduled", "", "", "", ""))
        buttons = ttk.Frame(container)
        buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Add Event…", command=self.add_event).pack(side="right", padx=6)
        self.grab_set()
        self.wait_window(self)

    def add_event(self) -> None:  # pragma: no cover - requires desktop GUI
        try:
            from customer_form import EventDialog
        except ImportError:
            messagebox.showerror("Calendar", "The event editor is unavailable.", parent=self)
            return
        EventDialog(
            self,
            customer_id=None,
            initial_date=self.day,
            on_saved=lambda *_: self.destroy(),
        )


class DatePickerDialog(_DialogBase):
    """A standalone month grid for selecting a date without third-party widgets."""

    def __init__(self, parent: Any, initial: Any = None):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the date picker.")
        super().__init__(parent)
        self.selected: date | None = None
        self.cleared = False
        try:
            self.current = parse_date(initial) or date.today()
        except (TypeError, ValueError):
            self.current = date.today()
        self.title("Select date")
        self.transient(parent)
        self.resizable(False, False)
        frame = ttk.Frame(self, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")
        self.month_var = tk.StringVar()
        ttk.Button(frame, text="‹", width=3, command=lambda: self._move(-1)).grid(row=0, column=0)
        ttk.Label(frame, textvariable=self.month_var, width=20, anchor="center", font=("Segoe UI", 11, "bold")).grid(row=0, column=1, columnspan=5)
        ttk.Button(frame, text="›", width=3, command=lambda: self._move(1)).grid(row=0, column=6)
        for index, name in enumerate(WEEKDAYS):
            ttk.Label(frame, text=name, width=4, anchor="center").grid(row=1, column=index, pady=4)
        self.buttons: list[list[Any]] = []
        for row in range(6):
            button_row: list[Any] = []
            for column in range(7):
                button = tk.Button(
                    frame, text="", width=4, height=2, relief="flat",
                    bg=COLORS["empty"], activebackground=COLORS["event"],
                    command=lambda r=row, c=column: self._select(r, c),
                )
                button.grid(row=row + 2, column=column, padx=1, pady=1)
                button_row.append(button)
            self.buttons.append(button_row)
        ttk.Button(frame, text="Today", command=lambda: self._pick(date.today())).grid(row=8, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(frame, text="Clear", command=self._clear).grid(row=8, column=2, columnspan=2, pady=(8, 0))
        ttk.Button(frame, text="Cancel", command=self.destroy).grid(row=8, column=4, columnspan=3, sticky="e", pady=(8, 0))
        self._draw()
        self.grab_set()
        self.wait_window(self)

    def _move(self, delta: int) -> None:
        month_index = self.current.year * 12 + self.current.month - 1 + delta
        self.current = date(month_index // 12, month_index % 12 + 1, 1)
        self._draw()

    def _draw(self) -> None:
        self.month_var.set(self.current.strftime("%B %Y"))
        grid = month_grid(self.current.year, self.current.month)
        for row in range(6):
            for column in range(7):
                day = grid[row][column]
                button = self.buttons[row][column]
                button.configure(text=str(day.day))
                state = day_visual_state(day, [], today=date.today(), display_month=self.current.month)
                button.configure(
                    bg=state["background"],
                    fg=state["foreground"],
                    highlightbackground=state["border"],
                    highlightthickness=state["border_width"],
                )
                button._picked_date = day  # type: ignore[attr-defined]

    def _select(self, row: int, column: int) -> None:
        self._pick(self.buttons[row][column]._picked_date)

    def _pick(self, value: date | None) -> None:
        self.selected = value
        self.destroy()

    def _clear(self) -> None:
        self.selected = None
        self.cleared = True
        self.destroy()


def show_date_picker(parent: Any, initial: Any = None) -> date | None:
    dialog = DatePickerDialog(parent, initial)
    return dialog.selected


class DateField(ttk.Frame if ttk is not None else object):  # type: ignore[misc]
    """ISO date Entry plus a custom calendar popup button."""

    def __init__(self, parent: Any, **kwargs: Any):
        if ttk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for DateField.")
        super().__init__(parent, **kwargs)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.var, width=14)
        self.entry.grid(row=0, column=0, sticky="ew")
        self.button = ttk.Button(self, text="▦", width=3, command=self._open_picker)
        self.button.grid(row=0, column=1, padx=(3, 0))
        self.columnconfigure(0, weight=1)
        try:
            style = ttk.Style(self)
            style.configure("Invalid.TEntry", fieldbackground="#FDE8E6", foreground="#991B1B")
            style.configure("Invalid.TCombobox", fieldbackground="#FDE8E6", foreground="#991B1B")
        except Exception:
            pass

    def _open_picker(self) -> None:
        picker = DatePickerDialog(self, self.get() or date.today())
        if picker.selected:
            self.set(picker.selected)
        elif picker.cleared:
            self.set("")

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, value: Any) -> None:
        try:
            self.var.set(normalize_date(value) or "")
        except (TypeError, ValueError):
            self.var.set("" if value is None else str(value))

    def set_error_style(self, invalid: bool) -> None:
        self.entry.configure(style="Invalid.TEntry" if invalid else "TEntry")


class CalendarFrame(_FrameBase):
    """Monthly events calendar with colored day cells."""

    def __init__(self, parent: Any, app: Any = None, **kwargs: Any):
        if ttk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the calendar screen.")
        super().__init__(parent, **kwargs)
        self.app = app
        today = date.today()
        self.year = today.year
        self.month = today.month
        self.day_buttons: dict[date, Any] = {}
        self.grouped_events: dict[str, list[dict[str, Any]]] = {}
        self.status_var = tk.StringVar(value="Ready")
        self._build()
        self.refresh()

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        header = ttk.Frame(self, padding=(18, 14, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Calendar", style="Heading.TLabel").pack(side="left")
        ttk.Button(header, text="‹ Previous", command=lambda: self._move(-1)).pack(side="left", padx=(24, 4))
        ttk.Button(header, text="Today", command=self._today).pack(side="left", padx=4)
        ttk.Button(header, text="Next ›", command=lambda: self._move(1)).pack(side="left", padx=4)
        self.month_label = ttk.Label(header, text="", style="Month.TLabel")
        self.month_label.pack(side="left", padx=20)
        ttk.Button(header, text="Refresh", command=self.refresh).pack(side="right")
        ttk.Label(header, textvariable=self.status_var).pack(side="right", padx=12)

        days = ttk.Frame(self, padding=(18, 4))
        days.grid(row=1, column=0, sticky="ew")
        for column in range(7):
            days.columnconfigure(column, weight=1)
            ttk.Label(days, text=WEEKDAYS[column], anchor="center", font=("Segoe UI", 10, "bold")).grid(
                row=0, column=column, sticky="ew", padx=2, pady=(0, 4)
            )
        self.grid_frame = ttk.Frame(self, padding=(18, 4, 18, 8))
        self.grid_frame.grid(row=2, column=0, sticky="nsew")
        for row in range(6):
            self.grid_frame.rowconfigure(row, weight=1)
            for column in range(7):
                self.grid_frame.columnconfigure(column, weight=1)
                button = tk.Button(
                    self.grid_frame, text="", anchor="nw", justify="left",
                    relief="flat", font=("Segoe UI", 10), cursor="hand2",
                )
                button.grid(row=row, column=column, sticky="nsew", padx=2, pady=2)
                self.day_buttons[date(1900, 1, 1) + timedelta(days=row * 7 + column)] = button
        # Keys are replaced on each refresh; this dict is rebuilt from the grid.
        self.day_buttons = {}
        self._create_buttons()

        legend = ttk.Frame(self, padding=(18, 8, 18, 18))
        legend.grid(row=3, column=0, sticky="ew")
        for label, color in (
            ("Today", COLORS["today"]), ("Events", COLORS["event"]),
            ("Overdue", COLORS["overdue"]), ("No events", COLORS["empty"]),
        ):
            swatch = tk.Label(legend, text="   ", bg=color, relief="groove", width=3)
            swatch.pack(side="left", padx=(0, 4))
            ttk.Label(legend, text=label).pack(side="left", padx=(0, 18))

    def _create_buttons(self) -> None:
        for child in self.grid_frame.winfo_children():
            child.destroy()
        grid = month_grid(self.year, self.month)
        self.day_buttons = {}
        for row, week in enumerate(grid):
            for column, day in enumerate(week):
                button = tk.Button(
                    self.grid_frame, text="", anchor="nw", justify="left",
                    relief="flat", font=("Segoe UI", 10), cursor="hand2",
                    command=lambda selected=day: self.open_day(selected),
                )
                button.grid(row=row, column=column, sticky="nsew", padx=2, pady=2)
                self.day_buttons[day] = button
        for row in range(6):
            self.grid_frame.rowconfigure(row, weight=1)
        for column in range(7):
            self.grid_frame.columnconfigure(column, weight=1)

    def _move(self, delta: int) -> None:
        month_index = self.year * 12 + self.month - 1 + delta
        self.year, self.month = month_index // 12, month_index % 12 + 1
        self._create_buttons()
        self.refresh(recreate=False)

    def _today(self) -> None:
        today = date.today()
        self.year, self.month = today.year, today.month
        self._create_buttons()
        self.refresh(recreate=False)

    def refresh(self, *, recreate: bool = True) -> None:
        try:
            grid = month_grid(self.year, self.month)
            if recreate:
                self._create_buttons()
            start, end = grid[0][0], grid[-1][-1]
            events = database.get_calendar_events(start, end)
            self.grouped_events = group_events_by_date(events)
            today = date.today()
            for day, button in self.day_buttons.items():
                day_events = self.grouped_events.get(day.isoformat(), [])
                state = day_visual_state(day, day_events, today=today, display_month=self.month)
                label = str(day.day)
                if day_events:
                    label += f"\n{len(day_events)} event{'s' if len(day_events) != 1 else ''}"
                if state["overdue"]:
                    label += "\nOverdue"
                button.configure(
                    text=label,
                    bg=state["background"],
                    fg=state["foreground"],
                    activebackground=state["background"],
                    highlightbackground=state["border"],
                    highlightcolor=state["border"],
                    highlightthickness=state["border_width"],
                    bd=0,
                )
            self.month_label.configure(text=date(self.year, self.month, 1).strftime("%B %Y"))
            self.status_var.set(f"{len(events)} events in view")
        except Exception as exc:
            self.status_var.set("Could not load calendar")
            if messagebox:
                messagebox.showerror("Calendar error", f"Could not load events:\n{exc}", parent=self)

    def open_day(self, day: date) -> None:
        try:
            events = database.get_events_by_date(day)
            due_services = database.get_services_due_between(day, day)
            DayEventsDialog(self, day, events, due_services)
            self.refresh(recreate=False)
        except Exception as exc:
            messagebox.showerror("Calendar error", f"Could not open this day:\n{exc}", parent=self)
