"""Excel import/export for customer records using openpyxl."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping

import database
import validation
from utils import normalize_date, parse_date

try:  # Keep backend importable on headless CI systems without Tkinter.
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    filedialog = None
    messagebox = None
    ttk = None

_FrameBase = ttk.Frame if ttk is not None else object

CUSTOMER_EXPORT_COLUMNS = database.CUSTOMER_COLUMNS
COLUMN_HEADERS = {
    "id": "ID",
    "customer_name": "Customer Name",
    "customer_location": "Customer Location",
    "customer_email": "Customer Email",
    "customer_phone": "Customer Phone",
    "survey_date": "Survey Date",
    "installation_date": "Installation Date",
    "last_service_date": "Last Service Date",
    "next_service_date": "Next Service Date",
    "service_type": "Service Type",
    "service_frequency": "Service Frequency",
    "installed_system": "Installed System",
    "individual_hardware": "Individual Hardware",
    "hardware_locations": "Hardware Locations",
    "fault_status": "Fault Status",
    "outstanding_invoices": "Outstanding Invoices",
    "paid_invoices": "Paid Invoices",
    "id_code": "ID Code",
    "master_code": "Master Code",
    "battery_alarm": "Battery Alarm",
    "battery_capacity": "Battery Capacity",
    "created_at": "Created At",
    "updated_at": "Updated At",
}


def _normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


HEADER_ALIASES = {
    _normalize_header(column): column for column in database.CUSTOMER_WRITABLE_COLUMNS
}
HEADER_ALIASES.update({
    "name": "customer_name",
    "client": "customer_name",
    "clientname": "customer_name",
    "businessname": "customer_name",
    "address": "customer_location",
    "site": "customer_location",
    "siteaddress": "customer_location",
    "email": "customer_email",
    "emailaddress": "customer_email",
    "phone": "customer_phone",
    "telephone": "customer_phone",
    "mobile": "customer_phone",
    "mobilenumber": "customer_phone",
    "survey": "survey_date",
    "installed": "installation_date",
    "installation": "installation_date",
    "lastservice": "last_service_date",
    "nextservice": "next_service_date",
    "servicedue": "next_service_date",
    "frequency": "service_frequency",
    "system": "installed_system",
    "hardware": "individual_hardware",
    "fault": "fault_status",
    "outstanding": "outstanding_invoices",
    "outstandingamount": "outstanding_invoices",
    "unpaid": "outstanding_invoices",
    "paid": "paid_invoices",
    "paidamount": "paid_invoices",
    "customeridcode": "id_code",
    "alarmbattery": "battery_alarm",
})


@dataclass
class ExportFilters:
    query: str = ""
    outstanding_only: bool = False
    active_faults_only: bool = False
    due_before: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "query": self.query.strip(),
            "outstanding_only": self.outstanding_only,
            "active_faults_only": self.active_faults_only,
            "due_before": normalize_date(self.due_before) if self.due_before else None,
        }


@dataclass
class ImportSummary:
    path: Path
    total_rows: int = 0
    imported: int = 0
    skipped: int = 0
    errors: list[str] = field(default_factory=list)
    dry_run: bool = False

    @property
    def message(self) -> str:
        verb = "would be imported" if self.dry_run else "rows imported"
        return f"{self.imported} {verb}, {self.skipped} rows skipped due to errors."

    @property
    def details(self) -> str:
        lines = [self.message, f"Source: {self.path}", f"Data rows: {self.total_rows}"]
        if self.errors:
            lines.extend(["", "Skipped rows:"])
            lines.extend(f"• {error}" for error in self.errors)
        return "\n".join(lines)


def _openpyxl_modules() -> tuple[Any, Any]:
    try:
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError as exc:  # pragma: no cover - dependency is in requirements
        raise RuntimeError("openpyxl is required for spreadsheet operations.") from exc
    styles = {"Alignment": Alignment, "Border": Border, "Font": Font, "PatternFill": PatternFill, "Side": Side}
    return (Workbook, load_workbook), (styles, get_column_letter)


def _style_header(worksheet: Any, column_count: int, *, freeze: bool = True) -> None:
    _, (styles, _) = _openpyxl_modules()
    fill = styles["PatternFill"](fill_type="solid", fgColor="1F4E79")
    font = styles["Font"](color="FFFFFF", bold=True, size=11)
    thin = styles["Side"](style="thin", color="D9E2F3")
    for cell in worksheet[1][:column_count]:
        cell.fill = fill
        cell.font = font
        cell.alignment = styles["Alignment"](horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles["Border"](bottom=thin)
    worksheet.row_dimensions[1].height = 28
    if freeze:
        worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions


def _autosize_columns(worksheet: Any, headers: list[str], rows: Iterable[Mapping[str, Any]]) -> None:
    _, (_, get_column_letter) = _openpyxl_modules()
    materialized = list(rows)
    for index, header in enumerate(headers, start=1):
        values = [str(row.get(header, "") or "") for row in materialized[:500]]
        width = max([len(str(header))] + [len(value) for value in values]) + 3
        worksheet.column_dimensions[get_column_letter(index)].width = min(max(width, 11), 38)


def default_export_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return database.get_exports_dir() / f"customers_{stamp}.xlsx"


def export_customers(
    path: str | Path | None = None,
    *,
    filters: ExportFilters | Mapping[str, Any] | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> dict[str, Any]:
    """Export all customer fields to a styled, filterable .xlsx workbook."""
    if filters is None:
        export_filters = ExportFilters()
    elif isinstance(filters, ExportFilters):
        export_filters = filters
    else:
        export_filters = ExportFilters(**dict(filters))
    options = export_filters.as_dict()
    rows = database.search_customers(options["query"] or None, filters=options)
    (Workbook, _), _ = _openpyxl_modules()
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Customers"
    headers = [COLUMN_HEADERS.get(column, column.replace("_", " ").title()) for column in CUSTOMER_EXPORT_COLUMNS]
    worksheet.append(headers)
    for index, row in enumerate(rows, start=1):
        values: list[Any] = []
        for column in CUSTOMER_EXPORT_COLUMNS:
            value = row.get(column)
            if column in database.CUSTOMER_DATE_COLUMNS:
                try:
                    value = parse_date(value)
                except (TypeError, ValueError):
                    value = None
            values.append(value)
        worksheet.append(values)
        if progress_callback and (index % 100 == 0 or index == len(rows)):
            progress_callback(index, len(rows))
    for row in worksheet.iter_rows(min_row=2):
        for column, cell in enumerate(row, start=1):
            field_name = CUSTOMER_EXPORT_COLUMNS[column - 1]
            if field_name in database.CUSTOMER_DATE_COLUMNS:
                cell.number_format = "yyyy-mm-dd"
            elif field_name in database.CUSTOMER_NUMBER_COLUMNS:
                cell.number_format = '"£"#,##0.00'
    _style_header(worksheet, len(headers))
    _autosize_columns(worksheet, headers, rows)
    info = workbook.create_sheet("Export Info")
    info.append(["Service Tracker customer export"])
    info.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    info.append(["Rows", len(rows)])
    info.append(["Search", options["query"] or "All"])
    info.append(["Outstanding only", "Yes" if options["outstanding_only"] else "No"])
    info.append(["Active faults only", "Yes" if options["active_faults_only"] else "No"])
    info.append(["Due before", options["due_before"] or "Any"])
    info.column_dimensions["A"].width = 24
    info.column_dimensions["B"].width = 40
    info["A1"].font = _openpyxl_modules()[1][0]["Font"](bold=True, size=14, color="1F4E79")

    destination = Path(path) if path else default_export_path()
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(str(destination))
    finally:
        workbook.close()
    return {"path": destination, "row_count": len(rows), "filters": options}


def create_import_template(path: str | Path | None = None) -> Path:
    (Workbook, _), (styles, get_column_letter) = _openpyxl_modules()
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Customers"
    headers = [COLUMN_HEADERS[c] for c in database.CUSTOMER_WRITABLE_COLUMNS]
    worksheet.append(headers)
    worksheet.append([
        "Example Customer Ltd", "1 High Street, St Albans", "accounts@example.com",
        "01727 123456", "2025-01-10", "2025-02-01", "2026-01-15", "2027-01-15",
        "Maintenance", "Annually", "Intruder alarm", "Panel and detectors",
        "Panel: reception; Detectors: hall", "None", 0, 250, "CUST-001", "", "", "",
    ][:len(headers)])
    _style_header(worksheet, len(headers))
    for index, header in enumerate(headers, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = min(max(len(header) + 3, 14), 34)
    worksheet["A1"].comment = None
    notes = workbook.create_sheet("Instructions")
    for row in (
        ["Service Tracker import instructions"],
        ["1.", "Dates may be ISO YYYY-MM-DD, DD/MM/YYYY, or real Excel dates."],
        ["2.", "Customer Name is required; email and phone are validated."],
        ["3.", "Date order: survey <= installation <= last service < next service."],
        ["4.", "Financial values must be non-negative numbers."],
        ["5.", "Invalid rows are skipped and listed in the import report."],
    ):
        notes.append(row)
    notes.column_dimensions["A"].width = 6
    notes.column_dimensions["B"].width = 85
    notes["A1"].font = styles["Font"](bold=True, size=14, color="1F4E79")
    destination = Path(path) if path else database.get_exports_dir() / "import_template.xlsx"
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(str(destination))
    workbook.close()
    return destination


def _map_headers(raw_headers: Iterable[Any]) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for index, raw in enumerate(raw_headers):
        name = HEADER_ALIASES.get(_normalize_header(raw))
        if name and name in database.CUSTOMER_WRITABLE_COLUMNS:
            mapping[index] = name
    if "customer_name" not in mapping.values():
        raise ValueError("The spreadsheet needs a Customer Name column (or a recognized alias).")
    return mapping


def _row_is_blank(values: Iterable[Any]) -> bool:
    return all(value is None or (isinstance(value, str) and not value.strip()) for value in values)


def import_customers(
    path: str | Path,
    *,
    dry_run: bool = False,
    progress_callback: Callable[[int, int], None] | None = None,
) -> ImportSummary:
    """Validate and import customer rows, reporting row-level errors."""
    source = Path(path)
    if not source.is_file():
        raise FileNotFoundError(f"Workbook not found: {source}")
    (_, load_workbook), _ = _openpyxl_modules()
    workbook = load_workbook(str(source), read_only=True, data_only=True)
    summary = ImportSummary(path=source, dry_run=dry_run)
    try:
        worksheet = workbook.worksheets[0]
        iterator = worksheet.iter_rows(values_only=True)
        header_row = None
        for _ in range(20):
            candidate = next(iterator, None)
            if candidate and not _row_is_blank(candidate):
                header_row = candidate
                break
        if header_row is None:
            raise ValueError("The workbook has no header row.")
        mapping = _map_headers(header_row)

        valid_rows: list[tuple[dict[str, Any], int]] = []
        estimated_total = max(0, (worksheet.max_row or 1) - 1)
        for excel_row, values in enumerate(iterator, start=2):
            if values is None or _row_is_blank(values):
                continue
            summary.total_rows += 1
            raw: dict[str, Any] = {}
            for index, field_name in mapping.items():
                raw[field_name] = values[index] if index < len(values) else None
            checked = validation.validate_customer(raw)
            if not checked.is_valid:
                summary.skipped += 1
                detail = "; ".join(f"{name}: {message}" for name, message in checked.errors.items())
                summary.errors.append(f"Row {excel_row}: {detail}")
            else:
                valid_rows.append((checked.data, excel_row))
            if progress_callback:
                progress_callback(summary.total_rows, estimated_total or summary.total_rows)

        if not dry_run:
            with database.transaction() as conn:
                for data, excel_row in valid_rows:
                    try:
                        database._insert_customer(conn, data)
                        summary.imported += 1
                    except Exception as exc:
                        summary.skipped += 1
                        summary.errors.append(f"Row {excel_row}: database error: {exc}")
        else:
            summary.imported = len(valid_rows)
        return summary
    finally:
        workbook.close()


class ImportExportFrame(_FrameBase):
    """Import/Export screen with filter controls and an operation report."""

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        if ttk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the import/export screen.")
        super().__init__(parent, **kwargs)
        self.app = app
        self.query_var = tk.StringVar()
        self.outstanding_var = tk.BooleanVar(value=False)
        self.faults_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready")
        self._build()

    def _build(self) -> None:
        from calendar_widget import DateField

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        header = ttk.Frame(self, padding=(18, 14, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Import / Export", style="Heading.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var).pack(side="right")

        filters = ttk.LabelFrame(self, text="Export filters", padding=12)
        filters.grid(row=1, column=0, sticky="ew", padx=18, pady=8)
        filters.columnconfigure(1, weight=1)
        ttk.Label(filters, text="Search").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(filters, textvariable=self.query_var).grid(row=0, column=1, columnspan=3, sticky="ew", pady=4)
        ttk.Checkbutton(filters, text="Outstanding invoices only", variable=self.outstanding_var).grid(row=1, column=0, sticky="w", pady=4)
        ttk.Checkbutton(filters, text="Active faults only", variable=self.faults_var).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(filters, text="Next service due before").grid(row=2, column=0, sticky="w", pady=4)
        self.due_before = DateField(filters)
        self.due_before.grid(row=2, column=1, sticky="w", pady=4)

        actions = ttk.LabelFrame(self, text="Actions", padding=12)
        actions.grid(row=2, column=0, sticky="ew", padx=18, pady=8)
        ttk.Button(actions, text="Export Filters to Excel", style="Accent.TButton", command=self.export_filtered).pack(side="left")
        ttk.Button(actions, text="Export All Customers", command=lambda: self.export_filtered(clear_filters=True)).pack(side="left", padx=6)
        ttk.Button(actions, text="Import from Excel…", command=self.import_file).pack(side="left", padx=6)
        ttk.Button(actions, text="Download Import Template", command=self.create_template).pack(side="left", padx=6)
        ttk.Checkbutton(actions, text="Validate import only (no changes)", variable=self.dry_run_var).pack(side="right")

        report = ttk.LabelFrame(self, text="Operation report", padding=10)
        report.grid(row=3, column=0, sticky="nsew", padx=18, pady=(8, 18))
        report.columnconfigure(0, weight=1)
        report.rowconfigure(0, weight=1)
        self.report_text = tk.Text(report, wrap="word", height=12, bg="#f8fafc")
        self.report_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(report, orient="vertical", command=self.report_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.report_text.configure(yscrollcommand=scrollbar.set)
        self._append("Ready. Exports include all customer fields, styled headers, filters, and frozen headings.")

    def _append(self, text: str) -> None:
        self.report_text.insert("end", text.rstrip() + "\n")
        self.report_text.see("end")

    def current_filters(self) -> ExportFilters:
        return ExportFilters(
            query=self.query_var.get(),
            outstanding_only=bool(self.outstanding_var.get()),
            active_faults_only=bool(self.faults_var.get()),
            due_before=self.due_before.get() or None,
        )

    def export_filtered(self, *, clear_filters: bool = False) -> None:
        filters = ExportFilters() if clear_filters else self.current_filters()
        destination = filedialog.asksaveasfilename(
            title="Save customer export",
            defaultextension=".xlsx",
            initialdir=str(database.get_exports_dir()),
            initialfile=default_export_path().name,
            filetypes=[("Excel workbook", "*.xlsx")],
            parent=self,
        )
        if not destination:
            return
        self.status_var.set("Exporting…")

        def work() -> dict[str, Any]:
            return export_customers(destination, filters=filters)

        def done(result: dict[str, Any]) -> None:
            self.status_var.set(f"Exported {result['row_count']} customers")
            self._append(f"Exported {result['row_count']} rows to {result['path']}")

        self._run(work, done)

    def create_template(self) -> None:
        destination = filedialog.asksaveasfilename(
            title="Save import template",
            defaultextension=".xlsx",
            initialdir=str(database.get_exports_dir()),
            initialfile="import_template.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            parent=self,
        )
        if not destination:
            return
        self._run(
            lambda: create_import_template(destination),
            lambda path: self._append(f"Created import template: {path}"),
        )

    def import_file(self) -> None:
        source = filedialog.askopenfilename(
            title="Select customer workbook",
            initialdir=str(database.get_exports_dir()),
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
            parent=self,
        )
        if not source:
            return
        dry_run = bool(self.dry_run_var.get())
        self.status_var.set("Validating workbook…")

        def work() -> ImportSummary:
            return import_customers(source, dry_run=dry_run)

        def done(summary: ImportSummary) -> None:
            self.status_var.set(summary.message)
            self._append(summary.details)
            if not dry_run and hasattr(self.app, "refresh_current_screen"):
                self.app.refresh_current_screen()
            messagebox.showinfo("Import complete", summary.message, parent=self)

        self._run(work, done)

    def _run(self, work: Callable[[], Any], done: Callable[[Any], None]) -> None:
        if hasattr(self.app, "run_async"):
            self.app.run_async(work, done, message="Working with spreadsheet…")
        else:  # pragma: no cover - standalone use
            try:
                done(work())
            except Exception as exc:
                self.status_var.set("Operation failed")
                messagebox.showerror("Spreadsheet error", str(exc), parent=self)
