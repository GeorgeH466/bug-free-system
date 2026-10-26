"""Customer add/edit form, detail view, and event/payment editors."""
from __future__ import annotations

from datetime import date
from typing import Any, Callable, Mapping

import analytics
import database
import validation
from utils import format_currency, format_date

try:  # Keep the validation-facing code importable on headless CI systems.
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    messagebox = None
    ttk = None

_DialogBase = tk.Toplevel if tk is not None else object

SERVICE_TYPES = ("Maintenance", "Repair", "Install", "Inspection", "Follow-up", "Emergency")
SERVICE_FREQUENCIES = ("Monthly", "Quarterly", "Biannually", "Annually", "Ad hoc")
FAULT_STATUSES = ("None", "Active", "Resolved")
EVENT_TYPES = ("Service", "Installation", "Survey", "Follow-up", "Inspection", "Repair")

FIELD_SECTIONS = (
    ("Customer", (
        ("customer_name", "Customer name *", "entry"),
        ("customer_location", "Customer location / address", "text"),
        ("customer_email", "Customer email", "entry"),
        ("customer_phone", "Customer phone", "entry"),
        ("id_code", "Customer ID code", "entry"),
    )),
    ("Important dates (YYYY-MM-DD)", (
        ("survey_date", "Survey date", "date"),
        ("installation_date", "Installation date", "date"),
        ("last_service_date", "Last service date", "date"),
        ("next_service_date", "Next service date", "date"),
    )),
    ("Service and system", (
        ("service_type", "Service type", "service_type"),
        ("service_frequency", "Service frequency", "frequency"),
        ("fault_status", "Fault status", "fault_status"),
        ("installed_system", "Installed system", "text"),
        ("individual_hardware", "Individual hardware", "text"),
        ("hardware_locations", "Hardware locations", "text"),
    )),
    ("Codes and batteries", (
        ("master_code", "Master code", "entry"),
        ("battery_alarm", "Alarm battery", "entry"),
        ("battery_capacity", "Battery capacity", "entry"),
    )),
    ("Accounts", (
        ("outstanding_invoices", "Outstanding invoices", "entry"),
        ("paid_invoices", "Paid invoices", "entry"),
    )),
)


class FieldTooltip:
    """Lightweight hover tooltip used to explain field validation errors."""

    def __init__(self, widget: Any):
        self.widget = widget
        self.text = ""
        self.tip: Any = None
        widget.bind("<Enter>", lambda _: self.show(), add="+")
        widget.bind("<Leave>", lambda _: self.hide(), add="+")
        widget.bind("<ButtonPress>", lambda _: self.hide(), add="+")

    def set_text(self, text: str) -> None:
        self.text = text
        if self.tip is not None:
            self.hide()

    def show(self) -> None:
        if not self.text or self.tip is not None or tk is None:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip, text=self.text, justify="left", bg="#7F1D1D", fg="white",
            relief="solid", borderwidth=1, font=("Segoe UI", 9), padx=8, pady=5,
            wraplength=300,
        )
        label.pack()

    def hide(self) -> None:
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


def _read_widget(widget: Any) -> str:
    kind = getattr(widget, "_field_kind", "entry")
    if kind == "text":
        return widget.get("1.0", "end").strip()
    return str(widget.get()).strip()


def _write_widget(widget: Any, value: Any) -> None:
    kind = getattr(widget, "_field_kind", "entry")
    if value is None:
        value = ""
    if kind == "text":
        widget.delete("1.0", "end")
        widget.insert("1.0", str(value))
    elif kind == "date":
        widget.set(value)
    elif kind in {"entry", "combo"}:
        widget.delete(0, "end")
        widget.insert(0, str(value))
    else:
        widget.set(value)


class CustomerForm(_DialogBase):
    """Validated add/edit form containing every writable customer column."""

    def __init__(
        self,
        parent: Any,
        *,
        customer_id: int | None = None,
        on_saved: Callable[[], None] | None = None,
        app: Any = None,
    ):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the customer form.")
        super().__init__(parent)
        self.customer_id = customer_id
        self.on_saved = on_saved
        self.app = app
        self.widgets: dict[str, Any] = {}
        self.error_labels: dict[str, Any] = {}
        self.tooltips: dict[str, FieldTooltip] = {}
        self.status_var = tk.StringVar(value="Complete the fields and click Save.")
        self.title("Edit Customer" if customer_id else "Add New Customer")
        self.geometry("820x720")
        self.transient(parent)
        self.resizable(True, True)
        self._build()
        if customer_id:
            self._load(customer_id)
        if not self.winfo_exists():
            return
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _build(self) -> None:
        from calendar_widget import DateField

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        outer = ttk.Frame(self, padding=(16, 14))
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        canvas = tk.Canvas(outer, highlightthickness=0, bg="#f8fafc")
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)
        self.form = ttk.Frame(canvas, padding=(8, 8, 14, 8))
        self.form.columnconfigure(1, weight=1)
        canvas_window = canvas.create_window((0, 0), window=self.form, anchor="nw")
        self.form.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfig(canvas_window, width=event.width))
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"))

        row = 0
        system = ttk.LabelFrame(self.form, text="Record", padding=10)
        system.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        system.columnconfigure(1, weight=1)
        system.columnconfigure(3, weight=1)
        self.record_id_var = tk.StringVar(value="New record" if not self.customer_id else str(self.customer_id))
        self.created_var = tk.StringVar(value="Set on save")
        self.updated_var = tk.StringVar(value="Set on save")
        for index, (label, variable) in enumerate((
            ("Internal ID", self.record_id_var), ("Created", self.created_var), ("Updated", self.updated_var),
        )):
            ttk.Label(system, text=label).grid(row=0, column=index * 2, sticky="w", padx=(0, 6))
            ttk.Label(system, textvariable=variable).grid(row=0, column=index * 2 + 1, sticky="w", padx=(0, 18))
        row += 1

        for section_title, fields in FIELD_SECTIONS:
            frame = ttk.LabelFrame(self.form, text=section_title, padding=10)
            frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 10))
            frame.columnconfigure(1, weight=1)
            for field_row, (name, label, kind) in enumerate(fields):
                ttk.Label(frame, text=label).grid(row=field_row * 2, column=0, sticky="w", padx=(0, 10), pady=(4, 0))
                widget = self._make_widget(frame, kind, DateField)
                widget._field_kind = "text" if kind == "text" else ("date" if kind == "date" else ("combo" if kind not in {"entry"} else "entry"))
                widget.grid(row=field_row * 2, column=1, sticky="ew", pady=(4, 0))
                error = ttk.Label(frame, text="", foreground="#B91C1C", wraplength=480)
                error.grid(row=field_row * 2 + 1, column=1, sticky="w", pady=(0, 2))
                self.widgets[name] = widget
                self.error_labels[name] = error
                self.tooltips[name] = FieldTooltip(widget)
            row += 1

        footer = ttk.Frame(self, padding=(16, 8, 16, 14))
        footer.grid(row=1, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, textvariable=self.status_var, foreground="#475569").grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Cancel", command=self.destroy).grid(row=0, column=1, padx=4)
        ttk.Button(footer, text="Save Customer", style="Accent.TButton", command=self.save).grid(row=0, column=2)

    def _make_widget(self, parent: Any, kind: str, DateField: Any) -> Any:
        if kind == "text":
            return tk.Text(parent, height=3, width=48, wrap="word", relief="solid", borderwidth=1)
        if kind == "date":
            return DateField(parent)
        if kind == "service_type":
            combo = ttk.Combobox(parent, values=SERVICE_TYPES, width=46)
            return combo
        if kind == "frequency":
            return ttk.Combobox(parent, values=SERVICE_FREQUENCIES, width=46)
        if kind == "fault_status":
            combo = ttk.Combobox(parent, values=FAULT_STATUSES, width=46)
            combo.set("None")
            return combo
        return ttk.Entry(parent, width=50)

    def _load(self, customer_id: int) -> None:
        try:
            customer = database.get_customer(customer_id)
            if not customer:
                raise ValueError("This customer no longer exists.")
            self.created_var.set(str(customer.get("created_at") or "—"))
            self.updated_var.set(str(customer.get("updated_at") or "—"))
            for name, widget in self.widgets.items():
                _write_widget(widget, customer.get(name))
            if not self.widgets["outstanding_invoices"].get():
                _write_widget(self.widgets["outstanding_invoices"], "0")
            if not self.widgets["paid_invoices"].get():
                _write_widget(self.widgets["paid_invoices"], "0")
        except Exception as exc:
            messagebox.showerror("Customer form", f"Could not load this customer:\n{exc}", parent=self)
            self.destroy()

    def collect_data(self) -> dict[str, Any]:
        return {name: _read_widget(widget) for name, widget in self.widgets.items()}

    def clear_errors(self) -> None:
        for name, widget in self.widgets.items():
            self.error_labels[name].configure(text="")
            self.tooltips[name].set_text("")
            kind = getattr(widget, "_field_kind", "entry")
            if kind == "text":
                widget.configure(background="#FFFFFF")
            elif kind == "date":
                widget.set_error_style(False)
            elif kind == "combo":
                widget.configure(style="TCombobox")
            else:
                widget.configure(style="TEntry")

    def show_errors(self, errors: Mapping[str, str]) -> None:
        self.clear_errors()
        for name, message in errors.items():
            widget = self.widgets.get(name)
            if widget is None:
                continue
            self.error_labels[name].configure(text=message)
            self.tooltips[name].set_text(message)
            kind = getattr(widget, "_field_kind", "entry")
            if kind == "text":
                widget.configure(background="#FDE8E6")
            elif kind == "date":
                widget.set_error_style(True)
            elif kind == "combo":
                widget.configure(style="Invalid.TCombobox")
            else:
                widget.configure(style="Invalid.TEntry")

    def save(self) -> None:
        result = validation.validate_customer(self.collect_data())
        if not result.is_valid:
            self.show_errors(result.errors)
            self.status_var.set(f"{len(result.errors)} field(s) need attention.")
            messagebox.showerror(
                "Validation failed",
                "Please correct the highlighted fields:\n\n" + result.error_summary,
                parent=self,
            )
            return
        self.clear_errors()
        try:
            if self.customer_id:
                database.update_customer(self.customer_id, result.data)
            else:
                database.add_customer(result.data)
        except Exception as exc:
            messagebox.showerror("Save failed", f"Could not save the customer:\n{exc}", parent=self)
            return
        if self.on_saved:
            self.on_saved()
        messagebox.showinfo("Customer saved", "The customer record was saved successfully.", parent=self)
        self.destroy()


class EventDialog(_DialogBase):
    """Add/edit one event; past scheduled events warn but are not blocked."""

    def __init__(
        self,
        parent: Any,
        *,
        customer_id: int | None = None,
        event_id: int | None = None,
        initial_date: Any = None,
        on_saved: Callable[[], None] | None = None,
    ):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the event editor.")
        super().__init__(parent)
        from calendar_widget import DateField

        self.event_id = event_id
        self.fixed_customer_id = customer_id
        self.on_saved = on_saved
        self.customer_map: dict[str, int] = {}
        self.customer_var = tk.StringVar()
        self.time_var = tk.StringVar()
        self.type_var = tk.StringVar(value="Service")
        self.status_var = tk.StringVar(value="Scheduled")
        self.title("Edit Event" if event_id else "Add Event")
        self.transient(parent)
        self.resizable(False, False)
        frame = ttk.Frame(self, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        row = 0
        ttk.Label(frame, text="Customer").grid(row=row, column=0, sticky="w", pady=5)
        self.customer_combo = ttk.Combobox(frame, textvariable=self.customer_var, state="readonly", width=42)
        self.customer_combo.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1
        ttk.Label(frame, text="Date").grid(row=row, column=0, sticky="w", pady=5)
        self.date_field = DateField(frame)
        self.date_field.grid(row=row, column=1, sticky="w", pady=5)
        row += 1
        ttk.Label(frame, text="Time").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.time_var, width=16).grid(row=row, column=1, sticky="w", pady=5)
        row += 1
        ttk.Label(frame, text="Type").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Combobox(frame, textvariable=self.type_var, values=EVENT_TYPES, width=39).grid(row=row, column=1, sticky="w", pady=5)
        row += 1
        ttk.Label(frame, text="Status").grid(row=row, column=0, sticky="w", pady=5)
        ttk.Combobox(frame, textvariable=self.status_var, values=validation.EVENT_STATUSES, state="readonly", width=39).grid(row=row, column=1, sticky="w", pady=5)
        row += 1
        ttk.Label(frame, text="Notes").grid(row=row, column=0, sticky="nw", pady=5)
        self.notes = tk.Text(frame, width=42, height=5, wrap="word")
        self.notes.grid(row=row, column=1, sticky="ew", pady=5)
        row += 1
        self.error_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.error_var, foreground="#B91C1C", wraplength=360).grid(row=row, column=0, columnspan=2, sticky="w", pady=(8, 0))
        row += 1
        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="Save", style="Accent.TButton", command=self.save).pack(side="right")

        choices = database.get_customer_choices()
        labels = [f"{c['customer_name']} (#{c['id']})" for c in choices]
        self.customer_map = dict(zip(labels, [c["id"] for c in choices]))
        self.customer_combo.configure(values=labels)
        if customer_id:
            match = next((label for label, cid in self.customer_map.items() if cid == customer_id), None)
            if match:
                self.customer_combo.set(match)
            self.customer_combo.configure(state="disabled")
        elif labels:
            self.customer_combo.current(0)
        existing = database.get_event(event_id) if event_id else None
        self.date_field.set(initial_date or (existing or {}).get("event_date") or date.today())
        if event_id and existing:
            self._load_event(existing)
        elif event_id:
            messagebox.showerror("Event", "This event no longer exists.", parent=self)
            self.destroy()
            return
        self.grab_set()
        self.wait_window(self)

    def _load_event(self, event: Mapping[str, Any]) -> None:
        match = next((label for label, cid in self.customer_map.items() if cid == event["customer_id"]), None)
        if match:
            self.customer_combo.set(match)
        self.date_field.set(event.get("event_date"))
        self.time_var.set(event.get("event_time") or "")
        self.type_var.set(event.get("event_type") or "Service")
        self.status_var.set(event.get("status") or "Scheduled")
        self.notes.insert("1.0", event.get("event_notes") or "")

    def save(self) -> None:
        customer_id = self.fixed_customer_id or self.customer_map.get(self.customer_var.get())
        payload = {
            "customer_id": customer_id,
            "event_date": self.date_field.get(),
            "event_time": self.time_var.get(),
            "event_type": self.type_var.get(),
            "event_notes": self.notes.get("1.0", "end").strip(),
            "status": self.status_var.get(),
        }
        result = validation.validate_event(payload, is_new=not self.event_id)
        if not result.is_valid:
            self.error_var.configure(text=result.error_summary)
            messagebox.showerror("Event validation", result.error_summary, parent=self)
            return
        if result.warnings:
            messagebox.showwarning(
                "Scheduled date warning",
                "\n".join(result.warnings.values()) + "\n\nThe event will still be saved.",
                parent=self,
            )
        try:
            if self.event_id:
                database.update_event(self.event_id, result.data)
            else:
                database.add_event(result.data)
        except Exception as exc:
            messagebox.showerror("Save failed", f"Could not save the event:\n{exc}", parent=self)
            return
        if self.on_saved:
            self.on_saved()
        self.destroy()


class PaymentDialog(_DialogBase):
    """Add/edit one invoice/payment record."""

    def __init__(
        self,
        parent: Any,
        *,
        customer_id: int,
        payment_id: int | None = None,
        on_saved: Callable[[], None] | None = None,
    ):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the payment editor.")
        super().__init__(parent)
        from calendar_widget import DateField

        self.customer_id = customer_id
        self.payment_id = payment_id
        self.on_saved = on_saved
        self.amount_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Outstanding")
        self.title("Edit Invoice" if payment_id else "Add Invoice")
        self.transient(parent)
        self.resizable(False, False)
        frame = ttk.Frame(self, padding=16)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Invoice date").grid(row=0, column=0, sticky="w", pady=5)
        self.invoice_date = DateField(frame)
        self.invoice_date.grid(row=0, column=1, sticky="w", pady=5)
        ttk.Label(frame, text="Payment date").grid(row=1, column=0, sticky="w", pady=5)
        self.payment_date = DateField(frame)
        self.payment_date.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(frame, text="Amount").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.amount_var, width=18).grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(frame, text="Status").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Combobox(frame, textvariable=self.status_var, values=validation.PAYMENT_STATUSES, state="readonly", width=17).grid(row=3, column=1, sticky="w", pady=5)
        self.error_var = tk.StringVar()
        ttk.Label(frame, textvariable=self.error_var, foreground="#B91C1C", wraplength=360).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="Save", style="Accent.TButton", command=self.save).pack(side="right")
        if payment_id:
            payment = database.get_payment(payment_id) or {}
            self.invoice_date.set(payment.get("invoice_date"))
            self.payment_date.set(payment.get("payment_date"))
            self.amount_var.set(str(payment.get("amount") or 0))
            self.status_var.set(payment.get("status") or "Outstanding")
        else:
            self.invoice_date.set(date.today())
            self.amount_var.set("0")
        self.grab_set()
        self.wait_window(self)

    def save(self) -> None:
        payload = {
            "customer_id": self.customer_id,
            "invoice_date": self.invoice_date.get(),
            "payment_date": self.payment_date.get() or None,
            "amount": self.amount_var.get(),
            "status": self.status_var.get(),
        }
        result = validation.validate_payment(payload)
        if not result.is_valid:
            self.error_var.configure(text=result.error_summary)
            messagebox.showerror("Payment validation", result.error_summary, parent=self)
            return
        try:
            if self.payment_id:
                database.update_payment(self.payment_id, result.data)
            else:
                database.add_payment(result.data)
        except Exception as exc:
            messagebox.showerror("Save failed", f"Could not save the invoice:\n{exc}", parent=self)
            return
        if self.on_saved:
            self.on_saved()
        self.destroy()


class CustomerDetailWindow(_DialogBase):
    """All customer fields plus event history, payment history, and analytics."""

    def __init__(self, parent: Any, customer_id: int, *, app: Any = None):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the customer detail view.")
        super().__init__(parent)
        self.customer_id = customer_id
        self.app = app
        self.field_labels: dict[str, Any] = {}
        self.title("Customer details")
        self.geometry("900x680")
        self.transient(parent)
        self._build()
        self.refresh()
        if self.winfo_exists():
            self.wait_window(self)

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        outer = ttk.Frame(self, padding=(14, 12))
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.columnconfigure(0, weight=1)
        self.title_var = tk.StringVar(value="Customer")
        ttk.Label(header, textvariable=self.title_var, style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="Refresh", command=self.refresh).grid(row=0, column=1, padx=4)
        ttk.Button(header, text="Edit Customer", style="Accent.TButton", command=self.edit_customer).grid(row=0, column=2, padx=4)
        ttk.Button(header, text="Close", command=self.destroy).grid(row=0, column=3)

        notebook = ttk.Notebook(outer)
        notebook.grid(row=1, column=0, sticky="nsew")
        details = ttk.Frame(notebook, padding=10)
        events = ttk.Frame(notebook, padding=10)
        payments = ttk.Frame(notebook, padding=10)
        notebook.add(details, text="All fields")
        notebook.add(events, text="Event history")
        notebook.add(payments, text="Payment history")
        self._build_details(details)
        self._build_events(events)
        self._build_payments(payments)

    def _build_details(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        metrics = ttk.LabelFrame(parent, text="Payment analytics", padding=10)
        metrics.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.metric_vars = {
            "on_time_rate": tk.StringVar(value="N/A"),
            "average_days": tk.StringVar(value="N/A"),
            "invoice_count": tk.StringVar(value="0"),
        }
        for index, (label, key) in enumerate((
            ("On-time payment rate", "on_time_rate"),
            ("Average days to pay", "average_days"),
            ("Invoices analysed", "invoice_count"),
        )):
            card = ttk.Frame(metrics, padding=10)
            card.grid(row=0, column=index, sticky="ew", padx=6)
            metrics.columnconfigure(index, weight=1)
            ttk.Label(card, text=label).pack()
            ttk.Label(card, textvariable=self.metric_vars[key], style="Metric.TLabel").pack()

        canvas = tk.Canvas(parent, highlightthickness=0, bg="#f8fafc")
        canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)
        form = ttk.Frame(canvas, padding=8)
        canvas.create_window((0, 0), window=form, anchor="nw")
        form.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
        form.columnconfigure(1, weight=1)
        for index, column in enumerate(database.CUSTOMER_COLUMNS):
            ttk.Label(form, text=column.replace("_", " ").title(), font=("Segoe UI", 9, "bold")).grid(
                row=index, column=0, sticky="nw", padx=(0, 14), pady=3
            )
            label = ttk.Label(form, text="—", wraplength=650, justify="left")
            label.grid(row=index, column=1, sticky="w", pady=3)
            self.field_labels[column] = label

    def _build_events(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        columns = ("date", "time", "type", "status", "notes")
        self.event_tree = ttk.Treeview(parent, columns=columns, show="headings")
        for column, heading, width in (
            ("date", "Date", 110), ("time", "Time", 90), ("type", "Type", 120),
            ("status", "Status", 110), ("notes", "Notes", 440),
        ):
            self.event_tree.heading(column, text=heading)
            self.event_tree.column(column, width=width, anchor="w")
        self.event_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.event_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.event_tree.configure(yscrollcommand=scrollbar.set)
        self.event_tree.bind("<Double-1>", lambda _: self.edit_event())
        actions = ttk.Frame(parent)
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Add Event", command=self.add_event).pack(side="left")
        ttk.Button(actions, text="Edit", command=self.edit_event).pack(side="left", padx=4)
        ttk.Button(actions, text="Delete", command=self.delete_event).pack(side="left")

    def _build_payments(self, parent: Any) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        columns = ("invoice_date", "payment_date", "amount", "status")
        self.payment_tree = ttk.Treeview(parent, columns=columns, show="headings")
        for column, heading, width in (
            ("invoice_date", "Invoice date", 150), ("payment_date", "Payment date", 150),
            ("amount", "Amount", 140), ("status", "Status", 140),
        ):
            self.payment_tree.heading(column, text=heading)
            self.payment_tree.column(column, width=width, anchor="w")
        self.payment_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.payment_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.payment_tree.configure(yscrollcommand=scrollbar.set)
        self.payment_tree.bind("<Double-1>", lambda _: self.edit_payment())
        actions = ttk.Frame(parent)
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Add Invoice", command=self.add_payment).pack(side="left")
        ttk.Button(actions, text="Edit", command=self.edit_payment).pack(side="left", padx=4)
        ttk.Button(actions, text="Delete", command=self.delete_payment).pack(side="left")

    def refresh(self) -> None:
        try:
            customer = database.get_customer(self.customer_id)
            if not customer:
                messagebox.showerror("Customer", "This customer no longer exists.", parent=self)
                self.destroy()
                return
            self.title_var.set(f"{customer['customer_name']}  •  #{customer['id']}")
            for column, label in self.field_labels.items():
                value = customer.get(column)
                if column in database.CUSTOMER_DATE_COLUMNS:
                    text = format_date(value, fallback="—")
                elif column in database.CUSTOMER_NUMBER_COLUMNS:
                    text = format_currency(value)
                else:
                    text = str(value) if value not in (None, "") else "—"
                label.configure(text=text)
            stats = analytics.get_payment_stats(self.customer_id)
            self.metric_vars["on_time_rate"].set(stats["formatted_on_time_rate"])
            self.metric_vars["average_days"].set(stats["formatted_average_days"])
            self.metric_vars["invoice_count"].set(str(stats["invoice_count"]))

            for child in self.event_tree.get_children():
                self.event_tree.delete(child)
            for event in database.get_events_for_customer(self.customer_id):
                status = event.get("status") or "Scheduled"
                if event.get("event_date") and event["event_date"] < date.today().isoformat() and status.lower() in validation.SCHEDULED_EVENT_STATUSES:
                    status = "Overdue"
                self.event_tree.insert("", "end", iid=str(event["id"]), values=(
                    format_date(event.get("event_date")), event.get("event_time") or "—",
                    event.get("event_type") or "—", status, event.get("event_notes") or "—",
                ))
            for child in self.payment_tree.get_children():
                self.payment_tree.delete(child)
            for payment in database.get_payments_for_customer(self.customer_id):
                self.payment_tree.insert("", "end", iid=str(payment["id"]), values=(
                    format_date(payment.get("invoice_date")),
                    format_date(payment.get("payment_date")),
                    format_currency(payment.get("amount")),
                    payment.get("status") or "Outstanding",
                ))
        except Exception as exc:
            messagebox.showerror("Customer details", f"Could not load details:\n{exc}", parent=self)

    def edit_customer(self) -> None:
        CustomerForm(self, customer_id=self.customer_id, on_saved=self.refresh, app=self.app)

    def add_event(self) -> None:
        EventDialog(self, customer_id=self.customer_id, on_saved=self.refresh)

    def edit_event(self) -> None:
        selection = self.event_tree.selection()
        if not selection:
            messagebox.showinfo("Select an event", "Select an event to edit.", parent=self)
            return
        EventDialog(self, customer_id=self.customer_id, event_id=int(selection[0]), on_saved=self.refresh)

    def delete_event(self) -> None:
        selection = self.event_tree.selection()
        if not selection:
            return
        if not messagebox.askyesno("Delete event", "Delete the selected event?", parent=self):
            return
        try:
            database.delete_event(int(selection[0]))
            self.refresh()
        except Exception as exc:
            messagebox.showerror("Delete failed", str(exc), parent=self)

    def add_payment(self) -> None:
        PaymentDialog(self, customer_id=self.customer_id, on_saved=self.refresh)

    def edit_payment(self) -> None:
        selection = self.payment_tree.selection()
        if not selection:
            messagebox.showinfo("Select an invoice", "Select an invoice to edit.", parent=self)
            return
        PaymentDialog(self, customer_id=self.customer_id, payment_id=int(selection[0]), on_saved=self.refresh)

    def delete_payment(self) -> None:
        selection = self.payment_tree.selection()
        if not selection:
            return
        if not messagebox.askyesno("Delete invoice", "Delete the selected invoice record?", parent=self):
            return
        try:
            database.delete_payment(int(selection[0]))
            self.refresh()
        except Exception as exc:
            messagebox.showerror("Delete failed", str(exc), parent=self)
