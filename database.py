"""SQLite persistence layer for Service Tracker.

All SQL uses bound parameters; column names passed to dynamic UPDATE/INSERT
statements are selected from module-level whitelists.
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import Any, Iterator, Mapping

from utils import normalize_date, parse_number, week_bounds

APP_DIR = Path(__file__).resolve().parent
_db_path = APP_DIR / "service_tracker.db"
_exports_dir = APP_DIR / "exports"
_template_path = APP_DIR / "templates" / "invoice_template.docx"

CUSTOMER_COLUMNS = (
    "id", "customer_name", "customer_location", "customer_email", "customer_phone",
    "survey_date", "installation_date", "last_service_date", "next_service_date",
    "service_type", "service_frequency", "installed_system", "individual_hardware",
    "hardware_locations", "fault_status", "outstanding_invoices", "paid_invoices",
    "id_code", "master_code", "battery_alarm", "battery_capacity", "created_at", "updated_at",
)
CUSTOMER_WRITABLE_COLUMNS = tuple(
    c for c in CUSTOMER_COLUMNS if c not in {"id", "created_at", "updated_at"}
)
CUSTOMER_DATE_COLUMNS = (
    "survey_date", "installation_date", "last_service_date", "next_service_date",
)
CUSTOMER_NUMBER_COLUMNS = ("outstanding_invoices", "paid_invoices")
EVENT_COLUMNS = ("id", "customer_id", "event_date", "event_time", "event_type", "event_notes", "status")
EVENT_WRITABLE_COLUMNS = tuple(c for c in EVENT_COLUMNS if c != "id")
PAYMENT_COLUMNS = ("id", "customer_id", "invoice_date", "payment_date", "amount", "status")
PAYMENT_WRITABLE_COLUMNS = tuple(c for c in PAYMENT_COLUMNS if c != "id")
PAYMENT_STATUSES = {"Paid", "Outstanding"}

SCHEMA = """
CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_name TEXT NOT NULL,
    customer_location TEXT,
    customer_email TEXT,
    customer_phone TEXT,
    survey_date DATE,
    installation_date DATE,
    last_service_date DATE,
    next_service_date DATE,
    service_type TEXT,
    service_frequency TEXT,
    installed_system TEXT,
    individual_hardware TEXT,
    hardware_locations TEXT,
    fault_status TEXT,
    outstanding_invoices REAL NOT NULL DEFAULT 0,
    paid_invoices REAL NOT NULL DEFAULT 0,
    id_code TEXT,
    master_code TEXT,
    battery_alarm TEXT,
    battery_capacity TEXT,
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    event_date DATE NOT NULL,
    event_time TIME,
    event_type TEXT,
    event_notes TEXT,
    status TEXT NOT NULL DEFAULT 'Scheduled',
    FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    invoice_date DATE,
    payment_date DATE,
    amount REAL,
    status TEXT NOT NULL DEFAULT 'Outstanding' CHECK (status IN ('Paid', 'Outstanding')),
    FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE INDEX IF NOT EXISTS idx_customers_name ON customers(customer_name COLLATE NOCASE);
CREATE INDEX IF NOT EXISTS idx_customers_next_service ON customers(next_service_date);
CREATE INDEX IF NOT EXISTS idx_customers_fault_status ON customers(fault_status);
CREATE INDEX IF NOT EXISTS idx_customers_outstanding ON customers(outstanding_invoices);
CREATE INDEX IF NOT EXISTS idx_events_date ON events(event_date);
CREATE INDEX IF NOT EXISTS idx_events_customer_date ON events(customer_id, event_date);
CREATE INDEX IF NOT EXISTS idx_payments_customer ON payments(customer_id);
CREATE INDEX IF NOT EXISTS idx_payments_invoice_date ON payments(invoice_date);
CREATE INDEX IF NOT EXISTS idx_payments_status ON payments(status);

CREATE TRIGGER IF NOT EXISTS trg_customers_updated_at
AFTER UPDATE ON customers
FOR EACH ROW
WHEN NEW.updated_at = OLD.updated_at
BEGIN
    UPDATE customers SET updated_at = CURRENT_TIMESTAMP WHERE id = OLD.id;
END;
"""


def configure(
    *,
    db_path: str | Path | None = None,
    exports_dir: str | Path | None = None,
    template_path: str | Path | None = None,
) -> None:
    """Override application paths (used by tests and portable installations)."""
    global _db_path, _exports_dir, _template_path
    if db_path is not None:
        _db_path = Path(db_path)
    if exports_dir is not None:
        _exports_dir = Path(exports_dir)
    if template_path is not None:
        _template_path = Path(template_path)


def set_database_path(path: str | Path) -> None:
    configure(db_path=path)


def reset_paths() -> None:
    configure(
        db_path=APP_DIR / "service_tracker.db",
        exports_dir=APP_DIR / "exports",
        template_path=APP_DIR / "templates" / "invoice_template.docx",
    )


def get_db_path() -> Path:
    return _db_path


def get_exports_dir() -> Path:
    return _exports_dir


def get_template_path() -> Path:
    return _template_path


def ensure_directories() -> None:
    get_db_path().parent.mkdir(parents=True, exist_ok=True)
    get_exports_dir().mkdir(parents=True, exist_ok=True)
    (get_exports_dir() / "Invoices").mkdir(parents=True, exist_ok=True)
    (get_exports_dir() / "Reports").mkdir(parents=True, exist_ok=True)
    get_template_path().parent.mkdir(parents=True, exist_ok=True)


@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    """Yield a configured connection, committing or rolling back automatically."""
    ensure_directories()
    conn = sqlite3.connect(str(get_db_path()), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 10000")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# Alias used by bulk importers so all rows can share one transaction.
transaction = get_connection


def initialize_database() -> None:
    """Create/upgrade the local database and indexes."""
    ensure_directories()
    with get_connection() as conn:
        conn.executescript(SCHEMA)
        # Lightweight forward migration for databases created before event status.
        event_columns = {row["name"] for row in conn.execute("PRAGMA table_info(events)")}
        if "status" not in event_columns:
            conn.execute("ALTER TABLE events ADD COLUMN status TEXT NOT NULL DEFAULT 'Scheduled'")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_events_status ON events(status)")


def _as_text(value: Any, *, required: bool = False) -> str | None:
    if value is None:
        text = ""
    else:
        text = str(value).strip()
    if required and not text:
        raise ValueError("A value is required.")
    return text or None


def _clean_customer_payload(data: Mapping[str, Any], *, for_update: bool = False) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for column in CUSTOMER_WRITABLE_COLUMNS:
        if column not in data:
            continue
        value = data[column]
        if column == "customer_name":
            name = _as_text(value, required=True)
            assert name is not None
            cleaned[column] = name
        elif column in CUSTOMER_DATE_COLUMNS:
            cleaned[column] = normalize_date(value)
        elif column in CUSTOMER_NUMBER_COLUMNS:
            parsed = parse_number(value, default=0.0)
            cleaned[column] = 0.0 if parsed is None else parsed
        else:
            cleaned[column] = _as_text(value)
    if not for_update and "customer_name" not in cleaned:
        raise ValueError("customer_name is required.")
    return cleaned


def _clean_event_payload(data: Mapping[str, Any], *, for_update: bool = False) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for column in EVENT_WRITABLE_COLUMNS:
        if column not in data:
            continue
        value = data[column]
        if column == "customer_id":
            cleaned[column] = int(value)
        elif column == "event_date":
            parsed = normalize_date(value)
            if not parsed:
                raise ValueError("event_date is required.")
            cleaned[column] = parsed
        elif column == "event_time":
            cleaned[column] = _as_text(value)
        else:
            cleaned[column] = _as_text(value)
    if not for_update and ("customer_id" not in cleaned or "event_date" not in cleaned):
        raise ValueError("customer_id and event_date are required.")
    if not for_update and not cleaned.get("status"):
        cleaned["status"] = "Scheduled"
    return cleaned


def _clean_payment_payload(data: Mapping[str, Any], *, for_update: bool = False) -> dict[str, Any]:
    cleaned: dict[str, Any] = {}
    for column in PAYMENT_WRITABLE_COLUMNS:
        if column not in data:
            continue
        value = data[column]
        if column == "customer_id":
            cleaned[column] = int(value)
        elif column in {"invoice_date", "payment_date"}:
            cleaned[column] = normalize_date(value)
        elif column == "amount":
            parsed = parse_number(value, default=0.0)
            cleaned[column] = 0.0 if parsed is None else parsed
        elif column == "status":
            status = _as_text(value) or "Outstanding"
            if status not in PAYMENT_STATUSES:
                status = "Paid" if status.lower() in {"paid", "yes", "true"} else "Outstanding"
            cleaned[column] = status
    if not for_update and "customer_id" not in cleaned:
        raise ValueError("customer_id is required.")
    return cleaned


def _row_to_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    return dict(row) if row is not None else None


def _insert_customer(conn: sqlite3.Connection, data: Mapping[str, Any]) -> int:
    payload = _clean_customer_payload(data)
    columns = ", ".join(payload)
    placeholders = ", ".join("?" for _ in payload)
    cursor = conn.execute(
        f"INSERT INTO customers ({columns}) VALUES ({placeholders})",
        tuple(payload.values()),
    )
    if cursor.lastrowid is None:
        raise sqlite3.Error("Could not determine the new customer ID.")
    return int(cursor.lastrowid)


# ---------------------------------------------------------------------------
# Customer CRUD and queries
# ---------------------------------------------------------------------------
def add_customer(data: Mapping[str, Any]) -> int:
    with get_connection() as conn:
        return _insert_customer(conn, data)


def update_customer(customer_id: int, data: Mapping[str, Any]) -> bool:
    payload = _clean_customer_payload(data, for_update=True)
    if not payload:
        return False
    assignments = ", ".join(f"{column} = ?" for column in payload)
    params = (*payload.values(), int(customer_id))
    with get_connection() as conn:
        cursor = conn.execute(
            f"UPDATE customers SET {assignments}, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            params,
        )
        return cursor.rowcount > 0


def delete_customer(customer_id: int) -> bool:
    with get_connection() as conn:
        cursor = conn.execute("DELETE FROM customers WHERE id = ?", (int(customer_id),))
        return cursor.rowcount > 0


def get_customer(customer_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM customers WHERE id = ?", (int(customer_id),)).fetchone()
        return _row_to_dict(row)


def _escape_like(value: str) -> str:
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _build_customer_where(
    query: str | None = None,
    filters: Mapping[str, Any] | None = None,
) -> tuple[str, list[Any]]:
    filters = filters or {}
    clauses: list[str] = []
    params: list[Any] = []
    if query and query.strip():
        like = f"%{_escape_like(query.strip())}%"
        searchable = (
            "customer_name", "customer_location", "customer_email", "customer_phone",
            "id_code", "installed_system", "individual_hardware",
        )
        clauses.append("(" + " OR ".join(
            f"COALESCE({column}, '') LIKE ? ESCAPE '\\'" for column in searchable
        ) + ")")
        params.extend([like] * len(searchable))
    if filters.get("outstanding_only"):
        clauses.append("COALESCE(outstanding_invoices, 0) > 0")
    if filters.get("active_faults_only"):
        clauses.append("LOWER(COALESCE(fault_status, '')) = 'active'")
    if filters.get("due_before"):
        clauses.append("next_service_date IS NOT NULL AND date(next_service_date) <= date(?)")
        params.append(normalize_date(filters["due_before"]))
    if filters.get("due_from"):
        clauses.append("next_service_date IS NOT NULL AND date(next_service_date) >= date(?)")
        params.append(normalize_date(filters["due_from"]))
    return (" WHERE " + " AND ".join(clauses) if clauses else ""), params


SORT_COLUMNS = {
    "id": "id",
    "customer_name": "customer_name COLLATE NOCASE",
    "customer_location": "customer_location COLLATE NOCASE",
    "customer_email": "customer_email COLLATE NOCASE",
    "next_service_date": "next_service_date",
    "last_service_date": "last_service_date",
    "service_type": "service_type COLLATE NOCASE",
    "fault_status": "fault_status COLLATE NOCASE",
    "outstanding_invoices": "outstanding_invoices",
    "paid_invoices": "paid_invoices",
    "id_code": "id_code COLLATE NOCASE",
}


def count_customers(query: str | None = None, filters: Mapping[str, Any] | None = None) -> int:
    where, params = _build_customer_where(query, filters)
    with get_connection() as conn:
        row = conn.execute(f"SELECT COUNT(*) AS total FROM customers{where}", tuple(params)).fetchone()
        return int(row["total"])


def search_customers(
    query: str | None = None,
    *,
    filters: Mapping[str, Any] | None = None,
    limit: int | None = None,
    offset: int = 0,
    order_by: str = "customer_name",
    descending: bool = False,
) -> list[dict[str, Any]]:
    where, params = _build_customer_where(query, filters)
    order = SORT_COLUMNS.get(order_by, "customer_name COLLATE NOCASE")
    direction = "DESC" if descending else "ASC"
    sql = f"SELECT * FROM customers{where} ORDER BY {order} {direction}, id ASC"
    if limit is not None:
        sql += " LIMIT ? OFFSET ?"
        params.extend([int(limit), int(offset)])
    elif offset:
        sql += " LIMIT -1 OFFSET ?"
        params.append(int(offset))
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(sql, tuple(params))]


def get_all_customers(limit: int | None = None, offset: int = 0) -> list[dict[str, Any]]:
    return search_customers(limit=limit, offset=offset)


def get_customer_page(
    query: str | None,
    filters: Mapping[str, Any] | None,
    limit: int,
    offset: int,
    order_by: str = "customer_name",
    descending: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    """Return one page of customers plus the total matching count."""
    rows = search_customers(
        query,
        filters=filters,
        limit=limit,
        offset=offset,
        order_by=order_by,
        descending=descending,
    )
    return rows, count_customers(query, filters)


def get_customer_choices(limit: int = 10000) -> list[dict[str, Any]]:
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            "SELECT id, customer_name, customer_location, id_code, service_type "
            "FROM customers ORDER BY customer_name COLLATE NOCASE LIMIT ?",
            (int(limit),),
        )]


# ---------------------------------------------------------------------------
# Event CRUD and queries
# ---------------------------------------------------------------------------
def add_event(data: Mapping[str, Any]) -> int:
    payload = _clean_event_payload(data)
    columns = ", ".join(payload)
    placeholders = ", ".join("?" for _ in payload)
    with get_connection() as conn:
        cursor = conn.execute(
            f"INSERT INTO events ({columns}) VALUES ({placeholders})",
            tuple(payload.values()),
        )
        return int(cursor.lastrowid)


def update_event(event_id: int, data: Mapping[str, Any]) -> bool:
    payload = _clean_event_payload(data, for_update=True)
    if not payload:
        return False
    assignments = ", ".join(f"{column} = ?" for column in payload)
    with get_connection() as conn:
        cursor = conn.execute(
            f"UPDATE events SET {assignments} WHERE id = ?",
            (*payload.values(), int(event_id)),
        )
        return cursor.rowcount > 0


def delete_event(event_id: int) -> bool:
    with get_connection() as conn:
        cursor = conn.execute("DELETE FROM events WHERE id = ?", (int(event_id),))
        return cursor.rowcount > 0


def get_event(event_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM events WHERE id = ?", (int(event_id),)).fetchone()
        return _row_to_dict(row)


def get_events_for_customer(customer_id: int, limit: int | None = None) -> list[dict[str, Any]]:
    sql = "SELECT * FROM events WHERE customer_id = ? ORDER BY date(event_date) DESC, event_time DESC, id DESC"
    params: list[Any] = [int(customer_id)]
    if limit is not None:
        sql += " LIMIT ?"
        params.append(int(limit))
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(sql, tuple(params))]


def get_calendar_events(start_date: Any, end_date: Any) -> list[dict[str, Any]]:
    start = normalize_date(start_date)
    end = normalize_date(end_date)
    if not start or not end:
        raise ValueError("A valid start and end date are required.")
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            """
            SELECT e.*, c.customer_name, c.customer_location
            FROM events e
            LEFT JOIN customers c ON c.id = e.customer_id
            WHERE date(e.event_date) BETWEEN date(?) AND date(?)
            ORDER BY date(e.event_date), COALESCE(e.event_time, '99:99'), e.id
            """,
            (start, end),
        )]


def get_events_by_date(event_date: Any) -> list[dict[str, Any]]:
    day = normalize_date(event_date)
    return get_calendar_events(day, day) if day else []


# ---------------------------------------------------------------------------
# Payment CRUD and queries
# ---------------------------------------------------------------------------
def add_payment(data: Mapping[str, Any]) -> int:
    payload = _clean_payment_payload(data)
    columns = ", ".join(payload)
    placeholders = ", ".join("?" for _ in payload)
    with get_connection() as conn:
        cursor = conn.execute(
            f"INSERT INTO payments ({columns}) VALUES ({placeholders})",
            tuple(payload.values()),
        )
        return int(cursor.lastrowid)


def update_payment(payment_id: int, data: Mapping[str, Any]) -> bool:
    payload = _clean_payment_payload(data, for_update=True)
    if not payload:
        return False
    assignments = ", ".join(f"{column} = ?" for column in payload)
    with get_connection() as conn:
        cursor = conn.execute(
            f"UPDATE payments SET {assignments} WHERE id = ?",
            (*payload.values(), int(payment_id)),
        )
        return cursor.rowcount > 0


def delete_payment(payment_id: int) -> bool:
    with get_connection() as conn:
        cursor = conn.execute("DELETE FROM payments WHERE id = ?", (int(payment_id),))
        return cursor.rowcount > 0


def get_payment(payment_id: int) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM payments WHERE id = ?", (int(payment_id),)).fetchone()
        return _row_to_dict(row)


def get_payments_for_customer(customer_id: int, limit: int | None = None) -> list[dict[str, Any]]:
    sql = "SELECT * FROM payments WHERE customer_id = ? ORDER BY date(invoice_date) DESC, id DESC"
    params: list[Any] = [int(customer_id)]
    if limit is not None:
        sql += " LIMIT ?"
        params.append(int(limit))
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(sql, tuple(params))]


# ---------------------------------------------------------------------------
# Dashboard/email helper queries
# ---------------------------------------------------------------------------
def get_services_due_between(start_date: Any, end_date: Any) -> list[dict[str, Any]]:
    start = normalize_date(start_date)
    end = normalize_date(end_date)
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            """
            SELECT * FROM customers
            WHERE next_service_date IS NOT NULL
              AND date(next_service_date) BETWEEN date(?) AND date(?)
            ORDER BY date(next_service_date), customer_name COLLATE NOCASE
            """,
            (start, end),
        )]


def get_overdue_services(reference: Any = None) -> list[dict[str, Any]]:
    today = normalize_date(reference) or date.today().isoformat()
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            """
            SELECT * FROM customers
            WHERE next_service_date IS NOT NULL
              AND date(next_service_date) < date(?)
            ORDER BY date(next_service_date), customer_name COLLATE NOCASE
            """,
            (today,),
        )]


def get_active_faults() -> list[dict[str, Any]]:
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            "SELECT * FROM customers WHERE LOWER(COALESCE(fault_status, '')) = 'active' "
            "ORDER BY next_service_date IS NULL, next_service_date, customer_name COLLATE NOCASE"
        )]


def get_total_outstanding() -> float:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT COALESCE(SUM(outstanding_invoices), 0) AS total FROM customers"
        ).fetchone()
        return float(row["total"] or 0)


def get_urgent_customers(limit: int = 100) -> list[dict[str, Any]]:
    with get_connection() as conn:
        return [dict(row) for row in conn.execute(
            """
            SELECT * FROM customers
            ORDER BY
              CASE WHEN LOWER(COALESCE(fault_status, '')) = 'active' THEN 0 ELSE 1 END,
              next_service_date IS NULL,
              date(next_service_date),
              customer_name COLLATE NOCASE
            LIMIT ?
            """,
            (int(limit),),
        )]


def get_dashboard_metrics(reference: Any = None) -> dict[str, Any]:
    today = normalize_date(reference) or date.today().isoformat()
    start, end = week_bounds(date.fromisoformat(today))
    with get_connection() as conn:
        total = int(conn.execute("SELECT COUNT(*) AS n FROM customers").fetchone()["n"])
        due = int(conn.execute(
            "SELECT COUNT(*) AS n FROM customers WHERE date(next_service_date) BETWEEN date(?) AND date(?)",
            (start.isoformat(), end.isoformat()),
        ).fetchone()["n"])
        overdue = int(conn.execute(
            "SELECT COUNT(*) AS n FROM customers WHERE next_service_date IS NOT NULL AND date(next_service_date) < date(?)",
            (today,),
        ).fetchone()["n"])
        outstanding = float(conn.execute(
            "SELECT COALESCE(SUM(outstanding_invoices), 0) AS n FROM customers"
        ).fetchone()["n"] or 0)
        active_faults = int(conn.execute(
            "SELECT COUNT(*) AS n FROM customers WHERE LOWER(COALESCE(fault_status, '')) = 'active'"
        ).fetchone()["n"])
    return {
        "total_customers": total,
        "due_this_week": due,
        "overdue_services": overdue,
        "total_outstanding": outstanding,
        "active_faults": active_faults,
        "week_start": start.isoformat(),
        "week_end": end.isoformat(),
        "today": today,
    }


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
DEFAULT_SETTINGS = {
    "welcome_seen": "0",
    "email_recipient": "",
    "email_schedule_enabled": "0",
    "email_schedule_day": "monday",
    "email_schedule_time": "09:00",
    "email_last_sent": "",
    "email_display_before_send": "0",
    "invoice_due_days": "30",
    "invoice_template_path": "",
    "customer_page_size": "200",
}


def get_setting(key: str, default: str | None = None) -> str | None:
    fallback = DEFAULT_SETTINGS.get(key, default)
    with get_connection() as conn:
        row = conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row and row["value"] is not None else fallback


def set_setting(key: str, value: Any) -> None:
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO app_settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, "" if value is None else str(value)),
        )


def get_all_settings() -> dict[str, str]:
    settings = dict(DEFAULT_SETTINGS)
    with get_connection() as conn:
        for row in conn.execute("SELECT key, value FROM app_settings"):
            settings[row["key"]] = row["value"]
    return settings
