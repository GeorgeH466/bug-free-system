"""Validation engine used before any customer, event, payment, or invoice save."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Mapping

from utils import normalize_date, parse_date, parse_number, parse_time

EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
PHONE_PATTERN = re.compile(r"^[0-9 +()\-]*$")
CUSTOMER_DATE_FIELDS = (
    "survey_date", "installation_date", "last_service_date", "next_service_date",
)
CUSTOMER_NUMBER_FIELDS = ("outstanding_invoices", "paid_invoices")
SCHEDULED_EVENT_STATUSES = {"scheduled", "pending", "overdue"}
EVENT_STATUSES = ("Scheduled", "Completed", "Pending", "Overdue", "Cancelled")
PAYMENT_STATUSES = ("Paid", "Outstanding")


@dataclass
class ValidationResult:
    data: dict[str, Any]
    errors: dict[str, str] = field(default_factory=dict)
    warnings: dict[str, str] = field(default_factory=dict)

    @property
    def is_valid(self) -> bool:
        return not self.errors

    def add_error(self, field_name: str, message: str) -> None:
        self.errors.setdefault(field_name, message)

    def add_warning(self, field_name: str, message: str) -> None:
        self.warnings.setdefault(field_name, message)

    def first_error(self) -> str:
        return next(iter(self.errors.values()), "")

    @property
    def error_summary(self) -> str:
        return "\n".join(f"• {name}: {message}" for name, message in self.errors.items())


def is_valid_email(value: Any) -> bool:
    if value is None or not str(value).strip():
        return True
    return bool(EMAIL_PATTERN.fullmatch(str(value).strip()))


def validate_email_address(value: Any, *, required: bool = False) -> str | None:
    text = "" if value is None else str(value).strip()
    if not text:
        if required:
            raise ValueError("An email address is required.")
        return None
    if not is_valid_email(text):
        raise ValueError("Enter a valid email address.")
    return text


def _clean_optional_text(value: Any) -> str | None:
    if value is None:
        return None
    return str(value).strip() or None


def _add_number(result: ValidationResult, field_name: str, *, default: float = 0.0) -> None:
    try:
        number = parse_number(result.data.get(field_name), default=default)
        result.data[field_name] = default if number is None else float(number)
        if result.data[field_name] < 0:
            result.add_error(field_name, "This amount cannot be negative.")
    except (TypeError, ValueError) as exc:
        result.add_error(field_name, str(exc))


def validate_customer(data: Mapping[str, Any]) -> ValidationResult:
    """Validate and normalize a customer record.

    Returns cleaned ISO date strings, normalized numbers, and trimmed text.
    Date sequence rules are only checked when both dates are present.
    """
    result = ValidationResult(data=dict(data))
    name = _clean_optional_text(result.data.get("customer_name"))
    result.data["customer_name"] = name
    if not name:
        result.add_error("customer_name", "Customer name is required.")

    # Text fields: normalize empty strings to NULL.
    text_fields = (
        "customer_location", "customer_email", "customer_phone", "service_type",
        "service_frequency", "installed_system", "individual_hardware",
        "hardware_locations", "fault_status", "id_code", "master_code",
        "battery_alarm", "battery_capacity",
    )
    for field_name in text_fields:
        result.data[field_name] = _clean_optional_text(result.data.get(field_name))

    if not is_valid_email(result.data.get("customer_email")):
        result.add_error("customer_email", "Enter a valid email address (name@example.com).")

    phone = result.data.get("customer_phone")
    if phone:
        if not PHONE_PATTERN.fullmatch(str(phone)):
            result.add_error("customer_phone", "Only digits, spaces, +, -, ( and ) are allowed.")
        elif not any(ch.isdigit() for ch in str(phone)):
            result.add_error("customer_phone", "Enter at least one digit.")

    for field_name in CUSTOMER_DATE_FIELDS:
        try:
            result.data[field_name] = normalize_date(result.data.get(field_name))
        except (TypeError, ValueError) as exc:
            result.add_error(field_name, str(exc))
            result.data[field_name] = None

    for field_name in CUSTOMER_NUMBER_FIELDS:
        _add_number(result, field_name)

    parsed_dates: dict[str, date | None] = {}
    for field_name in CUSTOMER_DATE_FIELDS:
        try:
            parsed_dates[field_name] = parse_date(result.data.get(field_name))
        except (TypeError, ValueError):
            parsed_dates[field_name] = None

    survey, installation = parsed_dates["survey_date"], parsed_dates["installation_date"]
    last_service, next_service = parsed_dates["last_service_date"], parsed_dates["next_service_date"]
    if survey and installation and survey > installation:
        result.add_error("installation_date", "Installation date must be on or after the survey date.")
    if installation and last_service and installation > last_service:
        result.add_error("last_service_date", "Last service date must be on or after installation.")
    if last_service and next_service and last_service >= next_service:
        result.add_error("next_service_date", "Next service date must be after the last service date.")

    return result


def validate_event(data: Mapping[str, Any], *, is_new: bool = True, today: date | None = None) -> ValidationResult:
    """Validate an event. A past scheduled event is a warning, not a blocking error."""
    result = ValidationResult(data=dict(data))
    try:
        result.data["customer_id"] = int(result.data.get("customer_id"))
    except (TypeError, ValueError):
        result.add_error("customer_id", "Select a customer.")

    try:
        event_date = normalize_date(result.data.get("event_date"))
        if not event_date:
            raise ValueError("Event date is required.")
        result.data["event_date"] = event_date
    except (TypeError, ValueError) as exc:
        result.add_error("event_date", str(exc))
        event_date = None

    try:
        result.data["event_time"] = parse_time(result.data.get("event_time"))
    except (TypeError, ValueError) as exc:
        result.add_error("event_time", str(exc))

    result.data["event_type"] = _clean_optional_text(result.data.get("event_type")) or "Service"
    result.data["event_notes"] = _clean_optional_text(result.data.get("event_notes"))
    status = _clean_optional_text(result.data.get("status")) or "Scheduled"
    if status not in EVENT_STATUSES:
        result.add_error("status", f"Status must be one of: {', '.join(EVENT_STATUSES)}.")
    result.data["status"] = status

    if is_new and event_date:
        parsed = parse_date(event_date)
        reference = today or date.today()
        if parsed and parsed < reference and status.lower() in SCHEDULED_EVENT_STATUSES:
            result.add_warning(
                "event_date",
                "This scheduled event is in the past. It is allowed, but will appear as overdue.",
            )
    return result


def validate_payment(data: Mapping[str, Any]) -> ValidationResult:
    result = ValidationResult(data=dict(data))
    try:
        result.data["customer_id"] = int(result.data.get("customer_id"))
    except (TypeError, ValueError):
        result.add_error("customer_id", "Select a customer.")

    for field_name in ("invoice_date", "payment_date"):
        try:
            result.data[field_name] = normalize_date(result.data.get(field_name))
        except (TypeError, ValueError) as exc:
            result.add_error(field_name, str(exc))
            result.data[field_name] = None

    _add_number(result, "amount")
    if result.data.get("amount", 0) <= 0:
        result.add_error("amount", "Invoice amount must be greater than zero.")

    status = _clean_optional_text(result.data.get("status")) or "Outstanding"
    if status not in PAYMENT_STATUSES:
        result.add_error("status", f"Status must be one of: {', '.join(PAYMENT_STATUSES)}.")
    result.data["status"] = status

    invoice_date = parse_date(result.data.get("invoice_date"))
    payment_date = parse_date(result.data.get("payment_date"))
    if status == "Paid" and not payment_date:
        result.add_error("payment_date", "A paid invoice needs a payment date.")
    if status == "Outstanding" and payment_date:
        result.add_error("status", "Remove the payment date or mark the invoice as Paid.")
    if invoice_date and payment_date and payment_date < invoice_date:
        result.add_error("payment_date", "Payment date cannot be before the invoice date.")
    return result


def validate_invoice(data: Mapping[str, Any]) -> ValidationResult:
    """Validate an invoice and normalize its editable line items."""
    result = ValidationResult(data=dict(data))
    result.data["customer_name"] = _clean_optional_text(result.data.get("customer_name"))
    if not result.data["customer_name"]:
        result.add_error("customer_name", "Customer name is required.")
    for field_name in ("customer_location", "id_code", "service_type", "invoice_number", "notes"):
        result.data[field_name] = _clean_optional_text(result.data.get(field_name))

    for field_name in ("invoice_date", "due_date"):
        try:
            value = normalize_date(result.data.get(field_name))
            if not value and field_name == "invoice_date":
                raise ValueError("Invoice date is required.")
            result.data[field_name] = value
        except (TypeError, ValueError) as exc:
            result.add_error(field_name, str(exc))
            result.data[field_name] = None

    _add_number(result, "tax_rate")
    _add_number(result, "adjustment")
    if result.data.get("tax_rate", 0) < 0:
        result.add_error("tax_rate", "Tax rate cannot be negative.")

    raw_items = result.data.get("line_items") or []
    if not isinstance(raw_items, (list, tuple)) or not raw_items:
        result.add_error("line_items", "Add at least one line item.")
        raw_items = []
    cleaned_items: list[dict[str, Any]] = []
    for index, raw in enumerate(raw_items, start=1):
        item = dict(raw) if isinstance(raw, Mapping) else {}
        description = _clean_optional_text(item.get("description"))
        if not description:
            result.add_error("line_items", f"Line {index} needs a description.")
        errors_before = len(result.errors)
        try:
            quantity = parse_number(item.get("quantity"), default=None)
            unit_price = parse_number(item.get("unit_price"), default=None)
            if quantity is None or unit_price is None:
                raise ValueError("Line items need a quantity and unit price.")
            if quantity <= 0:
                raise ValueError("Quantity must be greater than zero.")
            if unit_price < 0:
                raise ValueError("Unit price cannot be negative.")
            cleaned_items.append({
                "description": description or "",
                "quantity": float(quantity),
                "unit_price": float(unit_price),
                "line_total": round(float(quantity) * float(unit_price), 2),
            })
        except (TypeError, ValueError) as exc:
            if len(result.errors) == errors_before:
                result.add_error("line_items", f"Line {index}: {exc}")
    result.data["line_items"] = cleaned_items

    invoice_date = parse_date(result.data.get("invoice_date"))
    due_date = parse_date(result.data.get("due_date"))
    if due_date and invoice_date and due_date < invoice_date:
        result.add_error("due_date", "Due date cannot be before the invoice date.")

    if not result.errors:
        subtotal = round(sum(float(item["line_total"]) for item in cleaned_items), 2)
        adjustment = float(result.data.get("adjustment") or 0)
        tax = round((subtotal + adjustment) * float(result.data.get("tax_rate") or 0) / 100.0, 2)
        result.data["subtotal"] = subtotal
        result.data["tax_amount"] = tax
        result.data["amount"] = round(subtotal + adjustment + tax, 2)
    return result
