"""Invoice document generation using docxtpl and a bundled Word template."""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import Any, Mapping

import database
import validation
from utils import (
    format_currency,
    format_date,
    parse_number,
    safe_filename,
)

try:  # Keep backend importable on headless CI systems without Tkinter.
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # pragma: no cover - exercised only on Windows desktops
    tk = None
    filedialog = None
    messagebox = None
    ttk = None

_FrameBase = ttk.Frame if ttk is not None else object
_DialogBase = tk.Toplevel if tk is not None else object

INVOICE_STATUSES = ("Outstanding", "Paid")


def get_company_details() -> dict[str, str]:
    settings = database.get_all_settings()
    return {
        "company_name": settings.get("company_name") or "Your Business Name",
        "company_address": settings.get("company_address") or "Business address",
        "company_email": settings.get("company_email") or "",
        "company_phone": settings.get("company_phone") or "",
    }


def get_configured_template_path() -> Path:
    configured = database.get_setting("invoice_template_path", "")
    return Path(configured) if configured else database.get_template_path()


def ensure_default_template(path: str | Path | None = None, *, force: bool = False) -> Path:
    """Create a clean, editable Word invoice template with Jinja2 placeholders."""
    target = Path(path) if path else database.get_template_path()
    if target.exists() and not force:
        return target
    try:
        from docx import Document
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt, RGBColor
    except ImportError as exc:  # pragma: no cover - dependency is in requirements
        raise RuntimeError("python-docx is required to create the invoice template.") from exc

    target.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = title.add_run("INVOICE")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(31, 78, 121)

    company = doc.add_paragraph()
    company.add_run("{{company_name}}\n").bold = True
    company.add_run("{{company_address}}\n")
    company.add_run("{{company_email}}  |  {{company_phone}}")

    details = doc.add_table(rows=0, cols=2)
    details.alignment = WD_TABLE_ALIGNMENT.LEFT
    left, right = details.add_row().cells
    left.text = (
        "Bill To\n"
        "{{customer_name}}\n"
        "{{customer_location}}\n"
        "Customer ID: {{id_code}}"
    )
    right.text = (
        "Invoice Number: {{invoice_number}}\n"
        "Invoice Date: {{invoice_date}}\n"
        "Due Date: {{due_date}}\n"
        "Service Type: {{service_type}}"
    )
    for cell in (left, right):
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_after = Pt(2)

    doc.add_paragraph("")
    items = doc.add_table(rows=1, cols=5)
    items.style = "Table Grid"
    items.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ("Description", "Qty", "Unit Price", "Tax %", "Line Total")
    widths = (Inches(3.3), Inches(0.6), Inches(1.0), Inches(0.7), Inches(1.1))
    for cell, label, width in zip(items.rows[0].cells, headers, widths):
        cell.text = label
        cell.width = width
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    # docxtpl requires the row-level Jinja tags to occupy their own table rows.
    # Those two rows disappear during rendering, leaving one row per line item.
    start_row = items.add_row().cells
    start_row[0].text = "{%tr for item in line_items %}"
    row = items.add_row().cells
    row[0].text = "{{item.description}}"
    row[1].text = "{{item.quantity}}"
    row[2].text = "{{item.unit_price}}"
    row[3].text = "{{item.tax_rate}}"
    row[4].text = "{{item.line_total}}"
    end_row = items.add_row().cells
    end_row[0].text = "{%tr endfor %}"

    doc.add_paragraph("")
    summary = doc.add_table(rows=0, cols=2)
    summary.alignment = WD_TABLE_ALIGNMENT.RIGHT
    labels = (
        ("Subtotal:", "{{subtotal}}"),
        ("Adjustment:", "{{adjustment}}"),
        ("Tax ({{tax_rate}}%):", "{{tax_amount}}"),
        ("Amount Due:", "{{amount}}"),
    )
    for label, value in labels:
        cells = summary.add_row().cells
        cells[0].text = label
        cells[1].text = value
        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if label == "Amount Due:":
            for cell in cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

    doc.add_paragraph("")
    notes = doc.add_paragraph()
    notes.add_run("Notes\n").bold = True
    notes.add_run("{{notes}}")
    footer = doc.add_paragraph("Thank you for your business. Please quote {{invoice_number}} with payment.")
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.runs[0].italic = True
    doc.save(str(target))
    return target


def build_invoice_context(data: Mapping[str, Any]) -> dict[str, Any]:
    """Normalize validated invoice data into strings suitable for Word."""
    tax_rate = float(data.get("tax_rate") or 0)
    line_items = []
    for item in data.get("line_items", []):
        line_items.append({
            "description": str(item.get("description", "")),
            "quantity": f"{float(item.get('quantity') or 0):g}",
            "unit_price": format_currency(item.get("unit_price") or 0),
            "tax_rate": f"{tax_rate:g}",
            "line_total": format_currency(item.get("line_total") or 0),
        })
    context = dict(data)
    context.update(get_company_details())
    context.update({
        "line_items": line_items,
        "customer_name": data.get("customer_name") or "",
        "customer_location": data.get("customer_location") or "",
        "id_code": data.get("id_code") or "",
        "service_type": data.get("service_type") or "Service",
        "invoice_number": data.get("invoice_number") or "",
        "invoice_date": format_date(data.get("invoice_date")),
        "due_date": format_date(data.get("due_date"), fallback="On receipt"),
        "notes": data.get("notes") or "",
        "subtotal": format_currency(data.get("subtotal") or 0),
        "adjustment": format_currency(data.get("adjustment") or 0),
        "tax_rate": f"{tax_rate:g}",
        "tax_amount": format_currency(data.get("tax_amount") or 0),
        "amount": format_currency(data.get("amount") or 0),
        "total_amount": format_currency(data.get("amount") or 0),
        "generated_date": format_date(date.today()),
    })
    return context


def default_invoice_data(customer: Mapping[str, Any], reference: date | None = None) -> dict[str, Any]:
    today = reference or date.today()
    try:
        due_days = int(database.get_setting("invoice_due_days", "30") or 30)
    except (TypeError, ValueError):
        due_days = 30
    amount = float(customer.get("outstanding_invoices") or 0)
    return {
        "customer_id": customer.get("id"),
        "customer_name": customer.get("customer_name") or "",
        "customer_location": customer.get("customer_location") or "",
        "id_code": customer.get("id_code") or "",
        "service_type": customer.get("service_type") or "Service",
        "invoice_number": f"INV-{today.strftime('%Y%m%d')}-{customer.get('id') or 'NEW'}",
        "invoice_date": today.isoformat(),
        "due_date": (today + timedelta(days=due_days)).isoformat(),
        "tax_rate": 0.0,
        "adjustment": 0.0,
        "notes": "Payment is due by the date shown above.",
        "line_items": [{
            "description": customer.get("service_type") or "Service visit",
            "quantity": 1.0,
            "unit_price": amount,
        }],
        "status": "Outstanding",
    }


def generate_invoice(
    data: Mapping[str, Any],
    output_path: str | Path | None = None,
    *,
    template_path: str | Path | None = None,
    update_payment: bool = False,
) -> dict[str, Any]:
    """Render a validated invoice to an independently editable .docx file."""
    checked = validation.validate_invoice(data)
    if not checked.is_valid:
        raise ValueError(checked.error_summary or "Invoice data is invalid.")
    payload = checked.data
    template = Path(template_path) if template_path else get_configured_template_path()
    if not template.exists():
        if template == database.get_template_path() or template_path is None:
            template = ensure_default_template(template)
        else:
            raise FileNotFoundError(f"Invoice template not found: {template}")

    if output_path:
        destination = Path(output_path)
    else:
        name = safe_filename(str(payload.get("invoice_number") or "invoice"))
        destination = database.get_exports_dir() / "Invoices" / f"{name}.docx"
    destination.parent.mkdir(parents=True, exist_ok=True)
    context = build_invoice_context(payload)

    try:
        from docxtpl import DocxTemplate
    except ImportError as exc:  # pragma: no cover - dependency is in requirements
        raise RuntimeError("docxtpl is required to generate invoices.") from exc
    document = DocxTemplate(str(template))
    document.render(context)
    document.save(str(destination))

    payment_id = None
    if update_payment:
        customer_id = payload.get("customer_id")
        if not customer_id:
            raise ValueError("Select a saved customer before recording an invoice payment.")
        payment_id = database.add_payment({
            "customer_id": int(customer_id),
            "invoice_date": payload.get("invoice_date"),
            "payment_date": None,
            "amount": payload.get("amount"),
            "status": "Outstanding",
        })
    return {
        "path": destination,
        "payment_id": payment_id,
        "amount": payload.get("amount"),
        "context": context,
    }


class LineItemDialog(_DialogBase):
    """Small modal editor for one invoice line item."""

    def __init__(self, parent: Any, item: Mapping[str, Any] | None = None):
        if tk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the invoice editor.")
        super().__init__(parent)
        self.result: dict[str, Any] | None = None
        self.title("Line item")
        self.transient(parent)
        self.resizable(False, False)
        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        self.description_var = tk.StringVar(value=str((item or {}).get("description", "")))
        self.quantity_var = tk.StringVar(value=str((item or {}).get("quantity", 1)))
        self.unit_price_var = tk.StringVar(value=str((item or {}).get("unit_price", 0)))
        ttk.Label(frame, text="Description").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.description_var, width=42).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Quantity").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.quantity_var, width=14).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(frame, text="Unit price").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frame, textvariable=self.unit_price_var, width=14).grid(row=2, column=1, sticky="w", pady=4)
        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="Save", command=self._save).pack(side="right")
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _save(self) -> None:
        try:
            quantity = parse_number(self.quantity_var.get(), default=None)
            unit_price = parse_number(self.unit_price_var.get(), default=None)
            description = self.description_var.get().strip()
            if not description:
                raise ValueError("Description is required.")
            if quantity is None or quantity <= 0:
                raise ValueError("Quantity must be greater than zero.")
            if unit_price is None or unit_price < 0:
                raise ValueError("Unit price cannot be negative.")
        except (TypeError, ValueError) as exc:
            messagebox.showerror("Invalid line item", str(exc), parent=self)
            return
        self.result = {
            "description": description,
            "quantity": float(quantity),
            "unit_price": float(unit_price),
        }
        self.destroy()


class InvoiceFrame(_FrameBase):
    """Invoices screen: customer selection, editable lines, and Word generation."""

    def __init__(self, parent: Any, app: Any, **kwargs: Any):
        if ttk is None:  # pragma: no cover
            raise RuntimeError("Tkinter is required for the invoices screen.")
        super().__init__(parent, **kwargs)
        self.app = app
        self.customers: list[dict[str, Any]] = []
        self.customer_map: dict[str, int] = {}
        self.line_items: list[dict[str, Any]] = []
        self.customer_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.id_code_var = tk.StringVar()
        self.service_type_var = tk.StringVar()
        self.invoice_number_var = tk.StringVar()
        self.tax_rate_var = tk.StringVar(value="0")
        self.adjustment_var = tk.StringVar(value="0")
        self.update_payment_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")
        self._build()
        self.refresh()

    def _build(self) -> None:
        from calendar_widget import DateField

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        header = ttk.Frame(self, padding=(18, 14, 18, 8))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Invoice Generator", style="Heading.TLabel").pack(side="left")
        ttk.Button(header, text="Choose Template…", command=self.choose_template).pack(side="right", padx=4)
        ttk.Button(header, text="Refresh Customers", command=self.refresh).pack(side="right")

        top = ttk.Frame(self, padding=(18, 8))
        top.grid(row=1, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        ttk.Label(top, text="Customer").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.customer_combo = ttk.Combobox(
            top, textvariable=self.customer_var, state="readonly", width=40
        )
        self.customer_combo.grid(row=0, column=1, columnspan=3, sticky="ew", pady=4)
        self.customer_combo.bind("<<ComboboxSelected>>", lambda _: self.on_customer_changed())

        labels = [
            ("Customer name", self.name_var, 1, 0),
            ("Location", self.location_var, 2, 0),
            ("ID code", self.id_code_var, 3, 0),
            ("Service type", self.service_type_var, 4, 0),
            ("Invoice number", self.invoice_number_var, 1, 2),
            ("Tax rate (%)", self.tax_rate_var, 3, 2),
            ("Adjustment", self.adjustment_var, 4, 2),
        ]
        for text, variable, row, column in labels:
            ttk.Label(top, text=text).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=4)
            ttk.Entry(top, textvariable=variable, width=34).grid(row=row, column=column + 1, sticky="ew", pady=4)
        self.invoice_date = DateField(top)
        self.due_date = DateField(top)
        ttk.Label(top, text="Invoice date").grid(row=2, column=2, sticky="w", padx=(0, 8), pady=4)
        self.invoice_date.grid(row=2, column=3, sticky="w", pady=4)
        ttk.Label(top, text="Due date").grid(row=5, column=2, sticky="w", padx=(0, 8), pady=4)
        self.due_date.grid(row=5, column=3, sticky="w", pady=4)
        ttk.Label(top, text="Notes").grid(row=5, column=0, sticky="nw", padx=(0, 8), pady=4)
        self.notes_text = tk.Text(top, height=3, width=36, wrap="word")
        self.notes_text.grid(row=5, column=1, rowspan=2, sticky="ew", pady=4)

        middle = ttk.LabelFrame(self, text="Line items", padding=10)
        middle.grid(row=2, column=0, sticky="nsew", padx=18, pady=8)
        middle.columnconfigure(0, weight=1)
        middle.rowconfigure(0, weight=1)
        columns = ("description", "quantity", "unit_price", "line_total")
        self.tree = ttk.Treeview(middle, columns=columns, show="headings", height=8)
        for column, heading, width in (
            ("description", "Description", 420),
            ("quantity", "Qty", 80),
            ("unit_price", "Unit price", 120),
            ("line_total", "Line total", 120),
        ):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, anchor="center" if column != "description" else "w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(middle, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<Double-1>", lambda _: self.edit_line_item())
        actions = ttk.Frame(middle)
        actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Add", command=self.add_line_item).pack(side="left")
        ttk.Button(actions, text="Edit", command=self.edit_line_item).pack(side="left", padx=4)
        ttk.Button(actions, text="Remove", command=self.remove_line_item).pack(side="left")
        self.total_label = ttk.Label(actions, text="Total: £0.00", style="Metric.TLabel")
        self.total_label.pack(side="right")

        bottom = ttk.Frame(self, padding=(18, 8, 18, 16))
        bottom.grid(row=3, column=0, sticky="ew")
        ttk.Checkbutton(
            bottom,
            text="Record as an outstanding payment",
            variable=self.update_payment_var,
        ).pack(side="left")
        ttk.Button(bottom, text="Generate Invoice", style="Accent.TButton", command=self.generate).pack(side="right")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="right", padx=12)

    def refresh(self) -> None:
        try:
            self.customers = database.get_customer_choices()
            labels = [
                f"{row['customer_name']} — {row.get('customer_location') or 'No location'} (#{row['id']})"
                for row in self.customers
            ]
            self.customer_map = dict(zip(labels, [row["id"] for row in self.customers]))
            self.customer_combo.configure(values=labels)
            if labels and not self.customer_var.get():
                self.customer_combo.current(0)
                self.on_customer_changed()
            self.status_var.set(f"{len(labels)} customers available")
        except Exception as exc:
            messagebox.showerror("Invoice error", f"Could not load customers:\n{exc}", parent=self)
            self.status_var.set("Could not load customers")

    def on_customer_changed(self) -> None:
        customer_id = self.customer_map.get(self.customer_var.get())
        customer = next((row for row in self.customers if row["id"] == customer_id), None)
        if not customer:
            return
        full = database.get_customer(int(customer_id)) or customer
        data = default_invoice_data(full)
        self.name_var.set(data["customer_name"])
        self.location_var.set(data["customer_location"])
        self.id_code_var.set(data["id_code"])
        self.service_type_var.set(data["service_type"])
        self.invoice_number_var.set(data["invoice_number"])
        self.invoice_date.set(data["invoice_date"])
        self.due_date.set(data["due_date"])
        self.tax_rate_var.set(str(data["tax_rate"]))
        self.adjustment_var.set(str(data["adjustment"]))
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", data["notes"])
        self.line_items = [dict(item) for item in data["line_items"]]
        self._refresh_line_items()

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Word invoice template",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
            parent=self,
        )
        if not path:
            return
        database.set_setting("invoice_template_path", path)
        messagebox.showinfo("Template saved", f"Using template:\n{path}", parent=self)

    def _refresh_line_items(self) -> None:
        for child in self.tree.get_children():
            self.tree.delete(child)
        for index, item in enumerate(self.line_items):
            quantity = float(item.get("quantity") or 0)
            unit_price = float(item.get("unit_price") or 0)
            total = quantity * unit_price
            self.tree.insert(
                "", "end", iid=str(index),
                values=(
                    item.get("description", ""),
                    f"{quantity:g}",
                    format_currency(unit_price),
                    format_currency(total),
                ),
            )
        self._update_total_label()

    def _update_total_label(self) -> None:
        data = self.collect_data()
        checked = validation.validate_invoice(data)
        total = checked.data.get("amount", 0) if checked.is_valid else 0
        self.total_label.configure(text=f"Total: {format_currency(total)}")

    def add_line_item(self) -> None:
        dialog = LineItemDialog(self)
        if dialog.result:
            self.line_items.append(dialog.result)
            self._refresh_line_items()

    def edit_line_item(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Select a line", "Select a line item to edit.", parent=self)
            return
        index = int(selection[0])
        dialog = LineItemDialog(self, self.line_items[index])
        if dialog.result:
            self.line_items[index] = dialog.result
            self._refresh_line_items()

    def remove_line_item(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        index = int(selection[0])
        if messagebox.askyesno("Remove line", "Remove the selected line item?", parent=self):
            self.line_items.pop(index)
            self._refresh_line_items()

    def collect_data(self) -> dict[str, Any]:
        return {
            "customer_id": self.customer_map.get(self.customer_var.get()),
            "customer_name": self.name_var.get().strip(),
            "customer_location": self.location_var.get().strip(),
            "id_code": self.id_code_var.get().strip(),
            "service_type": self.service_type_var.get().strip(),
            "invoice_number": self.invoice_number_var.get().strip(),
            "invoice_date": self.invoice_date.get(),
            "due_date": self.due_date.get(),
            "tax_rate": self.tax_rate_var.get(),
            "adjustment": self.adjustment_var.get(),
            "notes": self.notes_text.get("1.0", "end").strip(),
            "line_items": [dict(item) for item in self.line_items],
            "status": "Outstanding",
        }

    def generate(self) -> None:
        checked = validation.validate_invoice(self.collect_data())
        if not checked.is_valid:
            messagebox.showerror(
                "Invoice validation",
                "Please correct these fields before generating:\n\n" + checked.error_summary,
                parent=self,
            )
            return
        default_name = safe_filename(str(checked.data.get("invoice_number") or "invoice")) + ".docx"
        destination = filedialog.asksaveasfilename(
            title="Save invoice",
            defaultextension=".docx",
            initialdir=str(database.get_exports_dir() / "Invoices"),
            initialfile=default_name,
            filetypes=[("Word document", "*.docx")],
            parent=self,
        )
        if not destination:
            return
        update_payment = bool(self.update_payment_var.get())
        self.status_var.set("Generating invoice…")

        def work() -> dict[str, Any]:
            return generate_invoice(
                checked.data,
                destination,
                update_payment=update_payment,
            )

        def done(result: dict[str, Any]) -> None:
            self.status_var.set(f"Generated {Path(result['path']).name}")
            messagebox.showinfo(
                "Invoice generated",
                f"Saved to:\n{result['path']}\n\nTotal: {format_currency(result.get('amount'))}",
                parent=self,
            )
            if hasattr(self.app, "refresh_current_screen"):
                self.app.refresh_current_screen()

        if hasattr(self.app, "run_async"):
            self.app.run_async(work, done, message="Generating invoice…")
        else:  # pragma: no cover - standalone use
            try:
                done(work())
            except Exception as exc:
                messagebox.showerror("Invoice error", str(exc), parent=self)
