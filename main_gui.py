"""Main Tkinter window, sidebar navigation, dashboard, and customer table."""
from __future__ import annotations

import queue
import shutil
import sys
import threading
import traceback
from datetime import date, datetime
from typing import Any, Callable

import analytics
import database
import email_notifier
import invoice_generator
import validation
from calendar_widget import CalendarFrame
from customer_form import CustomerDetailWindow, CustomerForm
from spreadsheet_handler import ImportExportFrame
from utils import format_currency, format_date

try:  # Keep module importable on headless CI systems without Tkinter.
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    messagebox = None
    ttk = None

_FrameBase = ttk.Frame if ttk is not None else object

NAVIGATION = (
    ("dashboard", "Dashboard"),
    ("customers", "Customers"),
    ("calendar", "Calendar"),
    ("invoices", "Invoices"),
    ("import_export", "Import/Export"),
    ("settings", "Settings"),
)


class DashboardFrame(_FrameBase):
    """Home screen with metrics, urgent customers, and payment-risk flags."""

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        super().__init__(parent, **kwargs)
        self.app = app
        self.card_vars: dict[str, Any] = {}
        self._build()
        self.refresh()

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)
        header = ttk.Frame(self, padding=(20, 16, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Dashboard", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Today's priorities, service schedule, and payment health", foreground="#64748B").grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Button(header, text="Refresh", command=self.refresh).grid(row=0, column=1, rowspan=2, sticky="e")
        self.week_label = ttk.Label(header, text="", foreground="#475569")
        self.week_label.grid(row=0, column=2, rowspan=2, sticky="e", padx=(14, 0))

        cards = ttk.Frame(self, padding=(20, 8))
        cards.grid(row=1, column=0, sticky="ew")
        for column in range(3):
            cards.columnconfigure(column, weight=1)
        card_specs = (
            ("total_customers", "Total customers"),
            ("due_this_week", "Services due this week"),
            ("overdue_services", "Overdue services"),
            ("active_faults", "Active faults"),
            ("total_outstanding", "Outstanding invoices"),
            ("on_time_rate", "Global on-time payment rate"),
        )
        for index, (key, label) in enumerate(card_specs):
            card = ttk.Frame(cards, style="Card.TFrame", padding=14)
            card.grid(row=index // 3, column=index % 3, sticky="ew", padx=6, pady=6)
            variable = tk.StringVar(value="—")
            self.card_vars[key] = variable
            ttk.Label(card, text=label, foreground="#64748B").pack(anchor="w")
            ttk.Label(card, textvariable=variable, style="Metric.TLabel").pack(anchor="w", pady=(5, 0))

        urgent = ttk.LabelFrame(self, text="Urgent & Pending — active faults first, then soonest service date", padding=10)
        urgent.grid(row=2, column=0, sticky="nsew", padx=20, pady=(8, 6))
        urgent.columnconfigure(0, weight=1)
        urgent.rowconfigure(0, weight=1)
        columns = ("name", "location", "service_type", "next_service", "fault", "outstanding")
        self.urgent_tree = ttk.Treeview(urgent, columns=columns, show="headings", height=9)
        for column, heading, width in (
            ("name", "Customer", 210), ("location", "Location", 260),
            ("service_type", "Service", 130), ("next_service", "Next service", 130),
            ("fault", "Fault", 100), ("outstanding", "Outstanding", 130),
        ):
            self.urgent_tree.heading(column, text=heading)
            self.urgent_tree.column(column, width=width, anchor="w")
        self.urgent_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(urgent, orient="vertical", command=self.urgent_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.urgent_tree.configure(yscrollcommand=scrollbar.set)
        self.urgent_tree.bind("<Double-1>", self._open_urgent_customer)
        self.urgent_tree.tag_configure("fault", background="#FEE2E2")
        self.urgent_tree.tag_configure("overdue", background="#FEF3C7")

        flagged = ttk.LabelFrame(self, text="Payment watchlist — on-time payment rate below 50%", padding=10)
        flagged.grid(row=3, column=0, sticky="nsew", padx=20, pady=(6, 16))
        flagged.columnconfigure(0, weight=1)
        flagged.rowconfigure(0, weight=1)
        self.flagged_tree = ttk.Treeview(flagged, columns=("name", "invoices", "rate", "outstanding"), show="headings", height=6)
        for column, heading, width in (
            ("name", "Customer", 300), ("invoices", "Invoices", 130),
            ("rate", "On-time rate", 160), ("outstanding", "Outstanding", 180),
        ):
            self.flagged_tree.heading(column, text=heading)
            self.flagged_tree.column(column, width=width, anchor="w")
        self.flagged_tree.grid(row=0, column=0, sticky="nsew")
        flagged_scroll = ttk.Scrollbar(flagged, orient="vertical", command=self.flagged_tree.yview)
        flagged_scroll.grid(row=0, column=1, sticky="ns")
        self.flagged_tree.configure(yscrollcommand=flagged_scroll.set)
        self.flagged_tree.bind("<Double-1>", self._open_flagged_customer)

    def refresh(self) -> None:
        try:
            metrics = analytics.get_dashboard_analytics()
            self.card_vars["total_customers"].set(str(metrics["total_customers"]))
            self.card_vars["due_this_week"].set(str(metrics["due_this_week"]))
            self.card_vars["overdue_services"].set(str(metrics["overdue_services"]))
            self.card_vars["active_faults"].set(str(metrics["active_faults"]))
            self.card_vars["total_outstanding"].set(metrics["formatted_outstanding"])
            self.card_vars["on_time_rate"].set(metrics["global_payment_stats"]["formatted_on_time_rate"])
            self.week_label.configure(text=f"Week: {metrics['week_label']}")

            for child in self.urgent_tree.get_children():
                self.urgent_tree.delete(child)
            today = date.today().isoformat()
            for customer in database.get_urgent_customers(75):
                active = str(customer.get("fault_status") or "").lower() == "active"
                overdue = bool(customer.get("next_service_date") and customer["next_service_date"] < today)
                self.urgent_tree.insert("", "end", iid=str(customer["id"]), tags=(
                    "fault" if active else "overdue" if overdue else "",
                ), values=(
                    customer.get("customer_name") or "",
                    customer.get("customer_location") or "—",
                    customer.get("service_type") or "—",
                    format_date(customer.get("next_service_date")),
                    customer.get("fault_status") or "None",
                    format_currency(customer.get("outstanding_invoices")),
                ))
            for child in self.flagged_tree.get_children():
                self.flagged_tree.delete(child)
            for customer in metrics["flagged_customers"]:
                self.flagged_tree.insert("", "end", iid=str(customer["id"]), values=(
                    customer.get("customer_name") or "",
                    customer.get("invoice_count") or 0,
                    f"{customer.get('on_time_rate') or 0:.1f}%",
                    format_currency(customer.get("outstanding_invoices")),
                ))
            if hasattr(self.app, "set_status"):
                self.app.set_status("Dashboard refreshed")
        except Exception as exc:
            messagebox.showerror("Dashboard", f"Could not refresh the dashboard:\n{exc}", parent=self)

    def _open_urgent_customer(self, _: Any) -> None:
        selection = self.urgent_tree.selection()
        if selection:
            CustomerDetailWindow(self, int(selection[0]), app=self.app)
            self.refresh()

    def _open_flagged_customer(self, _: Any) -> None:
        selection = self.flagged_tree.selection()
        if selection:
            CustomerDetailWindow(self, int(selection[0]), app=self.app)
            self.refresh()


class CustomersFrame(_FrameBase):
    """Searchable, paged customer table suitable for thousands of records."""

    COLUMNS = (
        ("id", "ID", 60, "id"),
        ("customer_name", "Customer", 230, "customer_name"),
        ("customer_location", "Location", 230, "customer_location"),
        ("customer_phone", "Phone", 130, "id"),
        ("service_type", "Service", 120, "service_type"),
        ("next_service_date", "Next service", 130, "next_service_date"),
        ("fault_status", "Fault", 100, "fault_status"),
        ("outstanding_invoices", "Outstanding", 130, "outstanding_invoices"),
        ("actions", "Actions", 130, "id"),
    )

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        super().__init__(parent, **kwargs)
        self.app = app
        self.page = 0
        self.total = 0
        self.order_by = "customer_name"
        self.descending = False
        self.search_var = tk.StringVar()
        self.outstanding_var = tk.BooleanVar(value=False)
        self.faults_var = tk.BooleanVar(value=False)
        self.page_var = tk.StringVar(value="Page 1 of 1")
        self.status_var = tk.StringVar(value="Ready")
        self._search_job: str | None = None
        self._build()
        self.refresh()

    @property
    def page_size(self) -> int:
        try:
            return max(50, min(1000, int(database.get_setting("customer_page_size", "200") or 200)))
        except (TypeError, ValueError):
            return 200

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        header = ttk.Frame(self, padding=(20, 16, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Customers", style="Heading.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var, foreground="#64748B").pack(side="right")

        controls = ttk.Frame(self, padding=(20, 8))
        controls.grid(row=1, column=0, sticky="ew")
        controls.columnconfigure(1, weight=1)
        ttk.Label(controls, text="Search").grid(row=0, column=0, sticky="w")
        self.search_entry = ttk.Entry(controls, textvariable=self.search_var, width=42)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 12))
        self.search_var.trace_add("write", lambda *_: self._schedule_search())
        ttk.Checkbutton(controls, text="Outstanding only", variable=self.outstanding_var, command=self._filters_changed).grid(row=0, column=2, padx=4)
        ttk.Checkbutton(controls, text="Active faults only", variable=self.faults_var, command=self._filters_changed).grid(row=0, column=3, padx=4)
        ttk.Button(controls, text="Clear", command=self.clear_filters).grid(row=0, column=4, padx=4)

        table_frame = ttk.Frame(self, padding=(20, 8, 20, 4))
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        columns = tuple(column[0] for column in self.COLUMNS)
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for column, heading, width, sort_key in self.COLUMNS:
            self.tree.heading(column, text=heading, command=lambda key=sort_key: self.sort_by(key))
            self.tree.column(column, width=width, anchor="w", stretch=column in {"customer_name", "customer_location"})
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.bind("<Double-1>", lambda _: self.open_details())
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.tag_configure("fault", background="#FEE2E2")
        self.tree.tag_configure("overdue", background="#FEF3C7")

        footer = ttk.Frame(self, padding=(20, 8, 20, 16))
        footer.grid(row=3, column=0, sticky="ew")
        ttk.Button(footer, text="Add New Customer", style="Accent.TButton", command=self.add_customer).pack(side="left")
        ttk.Button(footer, text="View Details", command=self.open_details).pack(side="left", padx=4)
        ttk.Button(footer, text="Edit", command=self.edit_selected).pack(side="left")
        ttk.Button(footer, text="Delete", command=self.delete_selected).pack(side="left", padx=4)
        ttk.Button(footer, text="Refresh", command=lambda: self.refresh(reset_page=False)).pack(side="left", padx=4)
        ttk.Button(footer, text="‹ Previous", command=self.previous_page).pack(side="right")
        ttk.Label(footer, textvariable=self.page_var).pack(side="right", padx=10)
        ttk.Button(footer, text="Next ›", command=self.next_page).pack(side="right")

    def _filters(self) -> dict[str, Any]:
        return {
            "outstanding_only": bool(self.outstanding_var.get()),
            "active_faults_only": bool(self.faults_var.get()),
        }

    def _schedule_search(self) -> None:
        if self._search_job:
            try:
                self.after_cancel(self._search_job)
            except Exception:
                pass
        self._search_job = self.after(250, self._filters_changed)

    def _filters_changed(self) -> None:
        self._search_job = None
        self.refresh(reset_page=True)

    def clear_filters(self) -> None:
        self.search_var.set("")
        self.outstanding_var.set(False)
        self.faults_var.set(False)
        self.refresh(reset_page=True)

    def sort_by(self, key: str) -> None:
        if key == "id":
            return
        if self.order_by == key:
            self.descending = not self.descending
        else:
            self.order_by = key
            self.descending = False
        self.refresh(reset_page=True)

    def refresh(self, *, reset_page: bool = False) -> None:
        if reset_page:
            self.page = 0
        try:
            query = self.search_var.get().strip() or None
            rows, self.total = database.get_customer_page(
                query,
                self._filters(),
                self.page_size,
                self.page * self.page_size,
                self.order_by,
                self.descending,
            )
            for child in self.tree.get_children():
                self.tree.delete(child)
            today = date.today().isoformat()
            for row in rows:
                active = str(row.get("fault_status") or "").lower() == "active"
                overdue = bool(row.get("next_service_date") and row["next_service_date"] < today)
                self.tree.insert("", "end", iid=str(row["id"]), tags=(
                    "fault" if active else "overdue" if overdue else "",
                ), values=(
                    row["id"],
                    row.get("customer_name") or "",
                    row.get("customer_location") or "—",
                    row.get("customer_phone") or "—",
                    row.get("service_type") or "—",
                    format_date(row.get("next_service_date")),
                    row.get("fault_status") or "None",
                    format_currency(row.get("outstanding_invoices")),
                    "✎ Edit   🗑 Delete",
                ))
            pages = max(1, (self.total + self.page_size - 1) // self.page_size)
            self.page = min(self.page, pages - 1)
            self.page_var.set(f"Page {self.page + 1} of {pages}")
            self.status_var.set(f"Showing {len(rows)} of {self.total} matching customers")
        except Exception as exc:
            self.status_var.set("Could not load customers")
            messagebox.showerror("Customers", f"Could not load customers:\n{exc}", parent=self)

    def _on_click(self, event: Any) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        action_column = f"#{len(self.COLUMNS)}"
        if self.tree.identify_column(event.x) != action_column:
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        self.tree.selection_set(row)
        bbox = self.tree.bbox(row, action_column)
        if bbox and event.x - bbox[0] < bbox[2] / 2:
            self.edit_selected()
        else:
            self.delete_selected()

    def selected_id(self) -> int | None:
        selection = self.tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def add_customer(self) -> None:
        CustomerForm(self, on_saved=lambda: self.refresh(reset_page=True), app=self.app)
        self.refresh(reset_page=True)

    def edit_selected(self) -> None:
        customer_id = self.selected_id()
        if not customer_id:
            messagebox.showinfo("Select a customer", "Select a customer row first.", parent=self)
            return
        CustomerForm(
            self,
            customer_id=customer_id,
            on_saved=lambda: self.refresh(reset_page=False),
            app=self.app,
        )
        self.refresh(reset_page=False)

    def delete_selected(self) -> None:
        customer_id = self.selected_id()
        if not customer_id:
            messagebox.showinfo("Select a customer", "Select a customer row first.", parent=self)
            return
        values = self.tree.item(str(customer_id), "values")
        name = values[1] if len(values) > 1 else "this customer"
        if not messagebox.askyesno(
            "Delete customer",
            f"Delete {name}?\n\nThis also deletes linked events and payment records.",
            parent=self,
        ):
            return
        try:
            database.delete_customer(customer_id)
            self.refresh(reset_page=False)
        except Exception as exc:
            messagebox.showerror("Delete failed", f"Could not delete the customer:\n{exc}", parent=self)

    def open_details(self) -> None:
        customer_id = self.selected_id()
        if not customer_id:
            messagebox.showinfo("Select a customer", "Select a customer row first.", parent=self)
            return
        CustomerDetailWindow(self, customer_id, app=self.app)
        self.refresh(reset_page=False)

    def previous_page(self) -> None:
        if self.page > 0:
            self.page -= 1
            self.refresh()

    def next_page(self) -> None:
        if (self.page + 1) * self.page_size < self.total:
            self.page += 1
            self.refresh()


class SettingsFrame(_FrameBase):
    """Local settings, company identity, maintenance actions, and email settings."""

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        super().__init__(parent, **kwargs)
        self.app = app
        self.due_days_var = tk.StringVar()
        self.page_size_var = tk.StringVar()
        self.company_name_var = tk.StringVar()
        self.company_address_var = tk.StringVar()
        self.company_email_var = tk.StringVar()
        self.company_phone_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self._build()
        self.refresh()

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        header = ttk.Frame(self, padding=(20, 16, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Settings", style="Heading.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var, foreground="#475569").pack(side="right")

        general = ttk.LabelFrame(self, text="Application settings", padding=12)
        general.grid(row=1, column=0, sticky="ew", padx=20, pady=6)
        general.columnconfigure(1, weight=1)
        fields = (
            ("Invoice due days", self.due_days_var, 0),
            ("Customer page size", self.page_size_var, 1),
            ("Company name", self.company_name_var, 2),
            ("Company address", self.company_address_var, 3),
            ("Company email", self.company_email_var, 4),
            ("Company phone", self.company_phone_var, 5),
        )
        for label, variable, row in fields:
            ttk.Label(general, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(general, textvariable=variable, width=48).grid(row=row, column=1, sticky="ew", pady=4)
        buttons = ttk.Frame(general)
        buttons.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(buttons, text="Save Application Settings", style="Accent.TButton", command=self.save).pack(side="left")
        ttk.Button(buttons, text="Recreate Default Invoice Template", command=self.recreate_template).pack(side="left", padx=6)
        ttk.Button(buttons, text="Back Up Database", command=self.backup_database).pack(side="left")

        info = ttk.LabelFrame(self, text="Local data", padding=12)
        info.grid(row=2, column=0, sticky="ew", padx=20, pady=6)
        ttk.Label(info, text=f"Database: {database.get_db_path()}").pack(anchor="w")
        ttk.Label(info, text=f"Exports: {database.get_exports_dir()}").pack(anchor="w")
        ttk.Label(info, text=f"Invoice template: {invoice_generator.get_configured_template_path()}").pack(anchor="w")
        outlook_status = "available" if email_notifier.is_outlook_available() else "unavailable (Windows + Outlook + pywin32 required)"
        ttk.Label(info, text=f"Outlook automation: {outlook_status}").pack(anchor="w")

        self.email_frame = email_notifier.EmailNotificationFrame(self, self.app)
        self.email_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(6, 16))
        self.rowconfigure(3, weight=1)

    def refresh(self) -> None:
        settings = database.get_all_settings()
        self.due_days_var.set(settings.get("invoice_due_days", "30"))
        self.page_size_var.set(settings.get("customer_page_size", "200"))
        self.company_name_var.set(settings.get("company_name", ""))
        self.company_address_var.set(settings.get("company_address", ""))
        self.company_email_var.set(settings.get("company_email", ""))
        self.company_phone_var.set(settings.get("company_phone", ""))
        self.status_var.set("Settings loaded")

    def save(self) -> None:
        try:
            due_days = int(self.due_days_var.get().strip())
            page_size = int(self.page_size_var.get().strip())
            if due_days < 0 or not (50 <= page_size <= 1000):
                raise ValueError("Due days must be 0 or more and page size must be between 50 and 1000.")
            if self.company_email_var.get().strip():
                validation.validate_email_address(self.company_email_var.get(), required=True)
        except (TypeError, ValueError) as exc:
            messagebox.showerror("Settings", str(exc), parent=self)
            return
        for key, value in (
            ("invoice_due_days", due_days),
            ("customer_page_size", page_size),
            ("company_name", self.company_name_var.get()),
            ("company_address", self.company_address_var.get()),
            ("company_email", self.company_email_var.get()),
            ("company_phone", self.company_phone_var.get()),
        ):
            database.set_setting(key, value)
        self.status_var.set("Settings saved")
        messagebox.showinfo("Settings", "Application settings were saved.", parent=self)

    def recreate_template(self) -> None:
        if not messagebox.askyesno(
            "Replace template",
            "Replace the default Word template with a fresh copy? Custom templates will remain on disk.",
            parent=self,
        ):
            return
        try:
            path = invoice_generator.ensure_default_template(force=True)
            database.set_setting("invoice_template_path", "")
            messagebox.showinfo("Template created", f"Created:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Template error", str(exc), parent=self)

    def backup_database(self) -> None:
        source = database.get_db_path()
        if not source.is_file():
            messagebox.showerror("Backup", "The database file does not exist yet.", parent=self)
            return
        backup_dir = database.get_exports_dir() / "Backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        destination = backup_dir / f"service_tracker_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        try:
            shutil.copy2(source, destination)
            messagebox.showinfo("Backup created", f"Database copied to:\n{destination}", parent=self)
        except Exception as exc:
            messagebox.showerror("Backup failed", str(exc), parent=self)


class ServiceTrackerApp:
    """Application controller: navigation, status, background work, and errors."""

    def __init__(self, root: Any):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required to launch Service Tracker.")
        self.root = root
        self.root.title("Service Tracker")
        self.root.geometry("1280x800")
        self.root.minsize(1024, 680)
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.report_callback_exception = self._handle_tk_exception
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._configure_styles()
        self._queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._async_callbacks: tuple[Callable[[Any], None] | None, Callable[[Exception], None] | None] = (None, None)
        self._busy = False
        self._build_menu()
        self._build_layout()
        self._build_screens()
        self.show_screen("dashboard")
        self.root.bind("<F5>", lambda _: self.refresh_current_screen())
        self.set_status(f"Database: {database.get_db_path()}")

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        available = set(style.theme_names())
        if "clam" in available:
            style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background="#F8FAFC")
        style.configure("TLabel", background="#F8FAFC", foreground="#1E293B")
        style.configure("Heading.TLabel", font=("Segoe UI", 17, "bold"), foreground="#0F172A")
        style.configure("Month.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1E40AF")
        style.configure("Metric.TLabel", font=("Segoe UI", 18, "bold"), foreground="#1E40AF")
        style.configure("Card.TFrame", background="#FFFFFF", relief="groove", borderwidth=1)
        style.configure("TButton", padding=(9, 6))
        style.configure("Accent.TButton", background="#2563EB", foreground="#FFFFFF", borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#1D4ED8"), ("pressed", "#1E40AF")])
        style.configure("Sidebar.TButton", background="#1E293B", foreground="#E2E8F0", borderwidth=0, anchor="w", padding=(16, 11), font=("Segoe UI", 10))
        style.map("Sidebar.TButton", background=[("active", "#334155")])
        style.configure("SidebarActive.TButton", background="#2563EB", foreground="#FFFFFF", borderwidth=0, anchor="w", padding=(16, 11), font=("Segoe UI", 10, "bold"))
        style.map("SidebarActive.TButton", background=[("active", "#1D4ED8")])
        style.configure("Treeview", rowheight=27, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E2E8F0", foreground="#1E293B")
        style.map("Treeview", background=[("selected", "#BFDBFE")], foreground=[("selected", "#0F172A")])
        style.configure("Invalid.TEntry", fieldbackground="#FDE8E6", foreground="#991B1B")
        style.configure("Invalid.TCombobox", fieldbackground="#FDE8E6", foreground="#991B1B")
        style.configure("TProgressbar", troughcolor="#E2E8F0", background="#2563EB")

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        app_menu = tk.Menu(menubar, tearoff=0)
        app_menu.add_command(label="Refresh  F5", command=self.refresh_current_screen)
        app_menu.add_separator()
        app_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=app_menu)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About Service Tracker", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.configure(menu=menubar)

    def _build_layout(self) -> None:
        self.sidebar = tk.Frame(self.root, bg="#0F172A", width=220)
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)
        brand = tk.Label(
            self.sidebar, text="Service Tracker", bg="#0F172A", fg="#FFFFFF",
            font=("Segoe UI", 15, "bold"), anchor="w", padx=18, pady=20,
        )
        brand.pack(fill="x")
        tk.Label(
            self.sidebar, text="Customer & Service Manager", bg="#0F172A",
            fg="#94A3B8", font=("Segoe UI", 9), anchor="w", padx=18,
        ).pack(fill="x", pady=(0, 18))
        self.nav_buttons: dict[str, Any] = {}
        for key, label in NAVIGATION:
            button = ttk.Button(
                self.sidebar, text=label, style="Sidebar.TButton",
                command=lambda selected=key: self.show_screen(selected),
            )
            button.pack(fill="x", padx=10, pady=3)
            self.nav_buttons[key] = button
        tk.Label(
            self.sidebar, text="Local-first • SQLite", bg="#0F172A", fg="#64748B",
            font=("Segoe UI", 9), anchor="w", padx=18,
        ).pack(side="bottom", fill="x", pady=16)

        self.content = ttk.Frame(self.root)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        status = ttk.Frame(self.root, padding=(10, 6), relief="groove")
        status.grid(row=1, column=0, columnspan=2, sticky="ew")
        status.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=170)
        self.progress.grid(row=0, column=1, sticky="e", padx=(10, 0))

    def _build_screens(self) -> None:
        self.screens = {
            "dashboard": DashboardFrame(self.content, self),
            "customers": CustomersFrame(self.content, self),
            "calendar": CalendarFrame(self.content, self),
            "invoices": invoice_generator.InvoiceFrame(self.content, self),
            "import_export": ImportExportFrame(self.content, self),
            "settings": SettingsFrame(self.content, self),
        }
        for screen in self.screens.values():
            screen.grid(row=0, column=0, sticky="nsew")

    @property
    def current_screen(self) -> Any:
        return self.screens.get(self.active_screen)

    def show_screen(self, key: str) -> None:
        if key not in self.screens:
            return
        self.active_screen = key
        self.screens[key].tkraise()
        for other, button in self.nav_buttons.items():
            button.configure(style="SidebarActive.TButton" if other == key else "Sidebar.TButton")
        if hasattr(self.screens[key], "refresh"):
            try:
                self.screens[key].refresh()
            except Exception as exc:
                self._show_error("Could not refresh this screen", exc)
        self.set_status(f"{dict(NAVIGATION)[key]} screen")

    def refresh_current_screen(self) -> None:
        screen = self.current_screen
        if screen is not None and hasattr(screen, "refresh"):
            try:
                screen.refresh()
            except Exception as exc:
                self._show_error("Could not refresh this screen", exc)

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def run_async(
        self,
        work: Callable[[], Any],
        on_success: Callable[[Any], None] | None = None,
        *,
        on_error: Callable[[Exception], None] | None = None,
        message: str = "Working…",
    ) -> None:
        """Run I/O or Outlook work off the Tk thread with a loading indicator."""
        if self._busy:
            messagebox.showinfo("Please wait", "Another operation is already running.", parent=self.root)
            return
        self._busy = True
        self._async_callbacks = (on_success, on_error)
        self.set_status(message)
        self.progress.start(12)

        def worker() -> None:
            try:
                result = work()
                self._queue.put(("success", result))
            except Exception as exc:  # noqa: BLE001 - surfaced in the GUI thread
                traceback.print_exc()
                self._queue.put(("error", exc))

        threading.Thread(target=worker, daemon=True, name="ServiceTrackerWorker").start()
        self.root.after(50, self._poll_async)

    def _poll_async(self) -> None:
        try:
            kind, payload = self._queue.get_nowait()
        except queue.Empty:
            if self._busy:
                self.root.after(50, self._poll_async)
            return
        self._busy = False
        self.progress.stop()
        on_success, on_error = self._async_callbacks
        self._async_callbacks = (None, None)
        if kind == "success":
            self.set_status("Ready")
            if on_success:
                try:
                    on_success(payload)
                except Exception as exc:
                    self._show_error("Operation completed but the screen could not update", exc)
        else:
            error = payload if isinstance(payload, Exception) else RuntimeError(str(payload))
            self.set_status(f"Error: {error}")
            if on_error:
                try:
                    on_error(error)
                except Exception as exc:
                    self._show_error("Error handler failed", exc)
            else:
                self._show_error("Operation failed", error)

    def _show_error(self, title: str, error: Exception) -> None:
        messagebox.showerror(title, f"{error}\n\nIf this persists, check the database and generated files.", parent=self.root)

    def _handle_tk_exception(self, exc_type: Any, exc_value: Any, exc_tb: Any) -> None:
        traceback.print_exception(exc_type, exc_value, exc_tb)
        messagebox.showerror("Unexpected error", f"{exc_type.__name__}: {exc_value}", parent=self.root)

    def show_about(self) -> None:
        messagebox.showinfo(
            "About Service Tracker",
            "Service Tracker\nLocal customer, service, invoice, and payment management.\n\n"
            f"Python {sys.version.split()[0]} | SQLite | Tkinter",
            parent=self.root,
        )

    def show_welcome(self) -> None:
        if database.get_setting("welcome_seen", "0") == "1":
            return
        messagebox.showinfo(
            "Welcome to Service Tracker",
            "Your local database, exports folder, and Word invoice template are ready.\n\n"
            "Start by adding a customer, scheduling an event, or importing an existing workbook. "
            "All data stays on this computer.",
            parent=self.root,
        )
        database.set_setting("welcome_seen", "1")

    def on_close(self) -> None:
        if self._busy and not messagebox.askyesno(
            "Close Service Tracker",
            "An operation is still running. Close anyway?",
            parent=self.root,
        ):
            return
        self.root.destroy()
