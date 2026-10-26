"""Weekly Outlook summary emails and their Excel action-list attachments."""
from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Mapping

import database
import spreadsheet_handler
import validation
from utils import format_currency, format_date, parse_date, week_bounds

try:  # Keep summary/attachment logic importable on headless CI systems.
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    messagebox = None
    ttk = None

_FrameBase = ttk.Frame if ttk is not None else object
DAYS_OF_WEEK = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")


class EmailNotAvailableError(RuntimeError):
    """Raised when Outlook automation is unavailable on the current platform."""


@dataclass
class WeeklySummary:
    subject: str
    body: str
    start_date: date
    end_date: date
    attachment_path: Path
    due_services: list[dict[str, Any]]
    overdue_services: list[dict[str, Any]]
    active_faults: list[dict[str, Any]]
    total_outstanding: float
    generated_at: datetime


def get_email_settings() -> dict[str, str]:
    return database.get_all_settings()


def save_email_settings(values: Mapping[str, Any]) -> None:
    allowed = {
        "email_recipient", "email_schedule_enabled", "email_schedule_day",
        "email_schedule_time", "email_display_before_send", "email_last_sent",
    }
    for key, value in values.items():
        if key in allowed:
            if isinstance(value, bool):
                value = "1" if value else "0"
            database.set_setting(key, value)


def is_outlook_available() -> bool:
    if sys.platform != "win32":
        return False
    try:  # pragma: no cover - Windows only
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def _service_line(customer: Mapping[str, Any]) -> str:
    parts = [
        customer.get("customer_name") or "Unnamed customer",
        customer.get("customer_location") or "No location",
        customer.get("service_type") or "Service",
        format_date(customer.get("next_service_date")),
    ]
    return " — ".join(str(part) for part in parts)


def build_summary_body(
    *,
    start_date: date,
    end_date: date,
    due_services: list[dict[str, Any]],
    overdue_services: list[dict[str, Any]],
    active_faults: list[dict[str, Any]],
    total_outstanding: float,
) -> str:
    lines = [
        "WEEKLY SERVICE TRACKER SUMMARY",
        f"Week: {start_date.strftime('%d %B %Y')} to {end_date.strftime('%d %B %Y')}",
        "",
        f"1. Services due this week ({len(due_services)})",
    ]
    lines.extend(f"   - {_service_line(row)}" for row in due_services) or lines.append("   - None")
    lines.extend(["", f"2. Overdue services ({len(overdue_services)})"])
    lines.extend(f"   - {_service_line(row)}" for row in overdue_services) or lines.append("   - None")
    lines.extend(["", f"3. Active faults ({len(active_faults)})"])
    for row in active_faults:
        lines.append(
            f"   - {row.get('customer_name') or 'Unnamed customer'} — "
            f"{row.get('customer_location') or 'No location'} — "
            f"{row.get('fault_status') or 'Active'} — Next service: {format_date(row.get('next_service_date'))}"
        )
    if not active_faults:
        lines.append("   - None")
    lines.extend([
        "",
        "4. Accounts",
        f"   - Total outstanding invoices: {format_currency(total_outstanding)}",
        "",
        "The attached workbook contains the detailed weekly action list.",
        f"Generated {datetime.now().strftime('%d %B %Y at %H:%M')} by Service Tracker.",
    ])
    return "\n".join(lines)


def generate_weekly_action_list_xlsx(summary: WeeklySummary, path: str | Path | None = None) -> Path:
    """Create the multi-sheet .xlsx attachment for the weekly Outlook email."""
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    overview = workbook.active
    overview.title = "Summary"
    overview.append(["Weekly Service Tracker Action List"])
    overview.append(["Week", f"{summary.start_date.isoformat()} to {summary.end_date.isoformat()}"])
    overview.append(["Services due this week", len(summary.due_services)])
    overview.append(["Overdue services", len(summary.overdue_services)])
    overview.append(["Active faults", len(summary.active_faults)])
    overview.append(["Total outstanding invoices", float(summary.total_outstanding)])
    overview.append(["Generated", summary.generated_at.strftime("%Y-%m-%d %H:%M:%S")])
    overview["A1"].font = spreadsheet_handler._openpyxl_modules()[1][0]["Font"](bold=True, size=14, color="1F4E79")
    overview["B6"].number_format = '"£"#,##0.00'
    overview.column_dimensions["A"].width = 30
    overview.column_dimensions["B"].width = 34

    def add_sheet(title: str, headers: list[str], rows: list[list[Any]], widths: list[int]) -> None:
        sheet = workbook.create_sheet(title)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        spreadsheet_handler._style_header(sheet, len(headers))
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, date):
                    cell.number_format = "yyyy-mm-dd"
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width

    def service_rows(records: list[dict[str, Any]]) -> list[list[Any]]:
        return [[
            row.get("customer_name") or "",
            row.get("customer_location") or "",
            row.get("service_type") or "",
            parse_date(row.get("next_service_date")),
            row.get("customer_phone") or "",
            row.get("customer_email") or "",
        ] for row in records]

    service_headers = ["Customer", "Location", "Service Type", "Service Date", "Phone", "Email"]
    service_widths = [28, 34, 20, 15, 18, 32]
    add_sheet("Services Due", service_headers, service_rows(summary.due_services), service_widths)
    add_sheet("Overdue Services", service_headers, service_rows(summary.overdue_services), service_widths)
    add_sheet(
        "Active Faults",
        ["Customer", "Location", "Fault Status", "Next Service", "Installed System", "Outstanding"],
        [[
            row.get("customer_name") or "", row.get("customer_location") or "",
            row.get("fault_status") or "Active", parse_date(row.get("next_service_date")),
            row.get("installed_system") or "", float(row.get("outstanding_invoices") or 0),
        ] for row in summary.active_faults],
        [28, 34, 16, 15, 34, 16],
    )
    outstanding_customers = database.search_customers(filters={"outstanding_only": True})
    add_sheet(
        "Outstanding Invoices",
        ["Customer", "Location", "Outstanding", "Paid", "Email", "Phone"],
        [[
            row.get("customer_name") or "", row.get("customer_location") or "",
            float(row.get("outstanding_invoices") or 0), float(row.get("paid_invoices") or 0),
            row.get("customer_email") or "", row.get("customer_phone") or "",
        ] for row in outstanding_customers],
        [28, 34, 16, 16, 32, 18],
    )

    destination = Path(path) if path else (
        database.get_exports_dir() / "Reports" /
        f"weekly_action_list_{summary.start_date.isoformat()}.xlsx"
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(str(destination))
    finally:
        workbook.close()
    return destination


def build_weekly_summary(
    reference: Any = None,
    *,
    attachment_path: str | Path | None = None,
) -> WeeklySummary:
    reference_date = parse_date(reference) or date.today()
    start, end = week_bounds(reference_date)
    due_services = database.get_services_due_between(start, end)
    overdue_services = database.get_overdue_services(reference_date)
    active_faults = database.get_active_faults()
    total_outstanding = database.get_total_outstanding()
    generated_at = datetime.now()
    subject = (
        f"Weekly Service Tracker Summary — "
        f"{start.strftime('%d %b')} to {end.strftime('%d %b %Y')}"
    )
    body = build_summary_body(
        start_date=start,
        end_date=end,
        due_services=due_services,
        overdue_services=overdue_services,
        active_faults=active_faults,
        total_outstanding=total_outstanding,
    )
    proposal_path = attachment_path or (
        database.get_exports_dir() / "Reports" /
        f"weekly_action_list_{start.isoformat()}.xlsx"
    )
    summary = WeeklySummary(
        subject=subject,
        body=body,
        start_date=start,
        end_date=end,
        attachment_path=Path(proposal_path),
        due_services=due_services,
        overdue_services=overdue_services,
        active_faults=active_faults,
        total_outstanding=total_outstanding,
        generated_at=generated_at,
    )
    summary.attachment_path = generate_weekly_action_list_xlsx(summary, proposal_path)
    return summary


def send_weekly_summary(
    summary: WeeklySummary,
    recipient: str,
    *,
    display_before_send: bool = False,
) -> bool:
    """Send a prepared summary through the local Windows Outlook application."""
    recipient = validation.validate_email_address(recipient, required=True) or ""
    if not summary.attachment_path.is_file():
        raise FileNotFoundError(f"Summary attachment is missing: {summary.attachment_path}")
    if not is_outlook_available():
        raise EmailNotAvailableError(
            "Outlook automation requires Windows, a desktop installation of Microsoft Outlook, "
            "and pywin32. The summary can still be previewed and its Excel attachment generated."
        )
    try:  # pragma: no cover - Windows/Outlook only
        import win32com.client
    except ImportError as exc:
        raise EmailNotAvailableError("pywin32 could not be imported. Install requirements.txt on Windows.") from exc

    pythoncom = None
    com_started = False
    try:  # pragma: no cover - Windows/Outlook only
        try:
            import pythoncom
            pythoncom.CoInitialize()
            com_started = True
        except Exception:
            pythoncom = None
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = summary.subject
        mail.Body = summary.body
        mail.Attachments.Add(str(summary.attachment_path.resolve()))
        if display_before_send:
            mail.Display(False)
        mail.Send()
        return True
    except Exception as exc:
        raise RuntimeError(f"Outlook could not send the weekly summary: {exc}") from exc
    finally:
        if com_started and pythoncom is not None:  # pragma: no cover
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def send_weekly_summary_now(
    recipient: str | None = None,
    *,
    reference: Any = None,
    display_before_send: bool | None = None,
) -> WeeklySummary:
    settings = get_email_settings()
    recipient = recipient or settings.get("email_recipient") or ""
    if display_before_send is None:
        display_before_send = settings.get("email_display_before_send") == "1"
    summary = build_weekly_summary(reference)
    send_weekly_summary(summary, recipient, display_before_send=bool(display_before_send))
    return summary


def should_send_scheduled(settings: Mapping[str, str], now: datetime | None = None) -> bool:
    """Pure scheduler predicate: weekly, after configured time, once per day."""
    if str(settings.get("email_schedule_enabled", "0")) not in {"1", "true", "yes"}:
        return False
    moment = now or datetime.now()
    configured_day = str(settings.get("email_schedule_day", "monday")).lower()
    if moment.strftime("%A").lower() != configured_day:
        return False
    try:
        target = datetime.strptime(str(settings.get("email_schedule_time", "09:00")), "%H:%M").time()
    except ValueError:
        return False
    if moment.time() < target:
        return False
    return str(settings.get("email_last_sent", "")) != moment.date().isoformat()


def mark_email_sent(moment: datetime | None = None) -> None:
    database.set_setting("email_last_sent", (moment or datetime.now()).date().isoformat())


class EmailNotificationFrame(_FrameBase):
    """Settings UI for recipient, manual send, preview, and weekly scheduling."""

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        if ttk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for email settings.")
        super().__init__(parent, **kwargs)
        self.app = app
        self.recipient_var = tk.StringVar()
        self.schedule_var = tk.BooleanVar(value=False)
        self.day_var = tk.StringVar(value="Monday")
        self.time_var = tk.StringVar(value="09:00")
        self.display_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="")
        self._sending = False
        self._build()
        self.load_settings()
        self.after(5000, self._check_schedule)

    def _build(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        form = ttk.LabelFrame(self, text="Email notifications", padding=12)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="Recipient email").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.recipient_var, width=42).grid(row=0, column=1, columnspan=3, sticky="ew", pady=5)
        ttk.Checkbutton(form, text="Enable automatic weekly send", variable=self.schedule_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Label(form, text="Day").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=self.day_var, values=[day.title() for day in DAYS_OF_WEEK], state="readonly", width=16).grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(form, text="Time (24h HH:MM)").grid(row=2, column=2, sticky="w", padx=(12, 6), pady=5)
        ttk.Entry(form, textvariable=self.time_var, width=10).grid(row=2, column=3, sticky="w", pady=5)
        ttk.Checkbutton(form, text="Show Outlook draft before sending", variable=self.display_var).grid(row=3, column=0, columnspan=3, sticky="w", pady=5)
        ttk.Label(form, textvariable=self.status_var, foreground="#475569").grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))

        actions = ttk.Frame(self)
        actions.grid(row=1, column=0, sticky="ew", pady=8)
        ttk.Button(actions, text="Save Email Settings", command=self.save_settings).pack(side="left")
        ttk.Button(actions, text="Preview Summary", command=self.preview).pack(side="left", padx=6)
        ttk.Button(actions, text="Send Now", style="Accent.TButton", command=self.send_now).pack(side="left")

        preview = ttk.LabelFrame(self, text="Email preview", padding=8)
        preview.grid(row=2, column=0, sticky="nsew")
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.preview_text = tk.Text(preview, height=12, wrap="word", bg="#f8fafc")
        self.preview_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(preview, orient="vertical", command=self.preview_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.preview_text.configure(yscrollcommand=scrollbar.set)

    def load_settings(self) -> None:
        settings = get_email_settings()
        self.recipient_var.set(settings.get("email_recipient", ""))
        self.schedule_var.set(str(settings.get("email_schedule_enabled", "0")) in {"1", "true", "yes"})
        self.day_var.set(str(settings.get("email_schedule_day", "monday")).title())
        self.time_var.set(settings.get("email_schedule_time", "09:00"))
        self.display_var.set(settings.get("email_display_before_send", "0") == "1")
        state = "available" if is_outlook_available() else "unavailable on this platform"
        self.status_var.set(f"Outlook automation is {state}. Previews and Excel attachments work everywhere.")

    def save_settings(self, *, notify: bool = True) -> bool:
        try:
            validation.validate_email_address(self.recipient_var.get(), required=True)
            datetime.strptime(self.time_var.get().strip(), "%H:%M")
            day = self.day_var.get().lower()
            if day not in DAYS_OF_WEEK:
                raise ValueError("Select a valid day of the week.")
        except (TypeError, ValueError) as exc:
            messagebox.showerror("Email settings", str(exc), parent=self)
            return False
        save_email_settings({
            "email_recipient": self.recipient_var.get().strip(),
            "email_schedule_enabled": self.schedule_var.get(),
            "email_schedule_day": day,
            "email_schedule_time": self.time_var.get().strip(),
            "email_display_before_send": self.display_var.get(),
        })
        if notify:
            messagebox.showinfo("Email settings", "Email settings were saved.", parent=self)
        return True

    def preview(self) -> None:
        self.status_var.set("Building weekly summary…")

        def done(summary: WeeklySummary) -> None:
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", f"Subject: {summary.subject}\nTo: {self.recipient_var.get()}\n\n{summary.body}")
            self.status_var.set(f"Attachment: {summary.attachment_path}")

        self._run(build_weekly_summary, done, message="Building weekly summary…")

    def send_now(self) -> None:
        try:
            recipient = validation.validate_email_address(self.recipient_var.get(), required=True)
        except ValueError as exc:
            messagebox.showerror("Email settings", str(exc), parent=self)
            return
        if not self.save_settings(notify=False):
            return
        self.status_var.set("Sending weekly summary…")
        display = bool(self.display_var.get())

        def work() -> WeeklySummary:
            return send_weekly_summary_now(recipient, display_before_send=display)

        def done(summary: WeeklySummary) -> None:
            self.status_var.set(f"Sent: {summary.subject}")
            messagebox.showinfo("Email sent", f"Weekly summary sent to {recipient}.", parent=self)

        self._run(work, done, message="Sending weekly summary…")

    def _check_schedule(self) -> None:
        try:
            if not self.winfo_exists():
                return
            settings = get_email_settings()
            if should_send_scheduled(settings) and not self._sending:
                recipient = settings.get("email_recipient", "")
                if recipient and validation.is_valid_email(recipient):
                    self._sending = True

                    def work() -> WeeklySummary:
                        return send_weekly_summary_now(recipient)

                    def done(summary: WeeklySummary) -> None:
                        mark_email_sent()
                        self._sending = False
                        self.status_var.set(f"Scheduled email sent: {summary.subject}")

                    def failed(error: Exception) -> None:
                        self._sending = False
                        self.status_var.set(f"Scheduled email failed: {error}")

                    if hasattr(self.app, "run_async"):
                        self.app.run_async(work, done, on_error=failed, message="Sending scheduled summary…")
                    else:  # pragma: no cover
                        try:
                            done(work())
                        except Exception as exc:
                            failed(exc)
        finally:
            try:
                if self.winfo_exists():
                    self.after(60000, self._check_schedule)
            except Exception:
                # The application may be closing while a scheduled check is active.
                pass

    def _run(self, work: Callable[[], Any], done: Callable[[Any], None], *, message: str) -> None:
        def failed(error: Exception) -> None:
            self.status_var.set(f"Operation failed: {error}")
            messagebox.showerror("Email", str(error), parent=self)

        if hasattr(self.app, "run_async"):
            self.app.run_async(work, done, on_error=failed, message=message)
        else:  # pragma: no cover
            try:
                done(work())
            except Exception as exc:
                self.status_var.set(f"Operation failed: {exc}")
                messagebox.showerror("Email", str(exc), parent=self)
