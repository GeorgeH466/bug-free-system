"""Payment analytics and dashboard support calculations."""
from __future__ import annotations

from typing import Any

import database
from utils import format_currency, week_bounds

DEFAULT_DUE_DAYS = 30


def get_due_days() -> int:
    try:
        value = int(database.get_setting("invoice_due_days", str(DEFAULT_DUE_DAYS)) or DEFAULT_DUE_DAYS)
        return max(0, value)
    except (TypeError, ValueError):
        return DEFAULT_DUE_DAYS


def get_payment_stats(customer_id: int, *, due_days: int | None = None) -> dict[str, Any]:
    """Return on-time payment rate and average days to pay for one customer."""
    days = get_due_days() if due_days is None else max(0, int(due_days))
    with database.get_connection() as conn:
        rows = conn.execute(
            """
            SELECT
                p.*,
                CASE
                    WHEN p.payment_date IS NOT NULL
                     AND date(p.payment_date) <= date(p.invoice_date, '+' || ? || ' days')
                    THEN 1 ELSE 0
                END AS on_time,
                CASE WHEN p.payment_date IS NOT NULL AND p.invoice_date IS NOT NULL
                     THEN julianday(p.payment_date) - julianday(p.invoice_date)
                END AS days_to_pay
            FROM payments p
            WHERE p.customer_id = ? AND p.invoice_date IS NOT NULL AND date(p.invoice_date) IS NOT NULL
            """,
            (days, int(customer_id)),
        ).fetchall()
    total = len(rows)
    paid = sum(1 for row in rows if row["payment_date"])
    on_time = sum(int(row["on_time"]) for row in rows)
    delays = [float(row["days_to_pay"]) for row in rows if row["days_to_pay"] is not None]
    return {
        "customer_id": int(customer_id),
        "due_days": days,
        "invoice_count": total,
        "paid_count": paid,
        "outstanding_count": total - paid,
        "on_time_count": on_time,
        "on_time_rate": round(100.0 * on_time / total, 1) if total else None,
        "average_days_to_pay": round(sum(delays) / len(delays), 1) if delays else None,
        "formatted_on_time_rate": "N/A" if not total else f"{100.0 * on_time / total:.1f}%",
        "formatted_average_days": "N/A" if not delays else f"{sum(delays) / len(delays):.1f} days",
    }


def get_global_payment_stats(*, due_days: int | None = None) -> dict[str, Any]:
    days = get_due_days() if due_days is None else max(0, int(due_days))
    with database.get_connection() as conn:
        row = conn.execute(
            """
            SELECT
                COUNT(*) AS invoice_count,
                SUM(CASE WHEN payment_date IS NOT NULL THEN 1 ELSE 0 END) AS paid_count,
                SUM(CASE
                    WHEN payment_date IS NOT NULL
                     AND date(payment_date) <= date(invoice_date, '+' || ? || ' days')
                    THEN 1 ELSE 0 END) AS on_time_count,
                AVG(CASE WHEN payment_date IS NOT NULL AND invoice_date IS NOT NULL
                    THEN julianday(payment_date) - julianday(invoice_date) END) AS avg_days
            FROM payments
            WHERE invoice_date IS NOT NULL AND date(invoice_date) IS NOT NULL
            """,
            (days,),
        ).fetchone()
    total = int(row["invoice_count"] or 0)
    on_time = int(row["on_time_count"] or 0)
    avg_days = float(row["avg_days"]) if row["avg_days"] is not None else None
    return {
        "due_days": days,
        "invoice_count": total,
        "paid_count": int(row["paid_count"] or 0),
        "on_time_count": on_time,
        "on_time_rate": round(100.0 * on_time / total, 1) if total else None,
        "average_days_to_pay": round(avg_days, 1) if avg_days is not None else None,
        "formatted_on_time_rate": "N/A" if not total else f"{100.0 * on_time / total:.1f}%",
        "formatted_average_days": "N/A" if avg_days is None else f"{avg_days:.1f} days",
    }


def get_low_on_time_customers(threshold: float = 50.0, *, due_days: int | None = None, limit: int = 100) -> list[dict[str, Any]]:
    days = get_due_days() if due_days is None else max(0, int(due_days))
    with database.get_connection() as conn:
        rows = conn.execute(
            """
            SELECT c.id, c.customer_name,
                   COUNT(p.id) AS invoice_count,
                   ROUND(100.0 * SUM(CASE
                       WHEN p.payment_date IS NOT NULL
                        AND date(p.payment_date) <= date(p.invoice_date, '+' || ? || ' days')
                       THEN 1 ELSE 0 END) / COUNT(p.id), 1) AS on_time_rate,
                   COALESCE(c.outstanding_invoices, 0) AS outstanding_invoices
            FROM customers c
            JOIN payments p ON p.customer_id = c.id
            WHERE p.invoice_date IS NOT NULL AND date(p.invoice_date) IS NOT NULL
            GROUP BY c.id, c.customer_name
            HAVING on_time_rate < ?
            ORDER BY on_time_rate ASC, c.customer_name COLLATE NOCASE
            LIMIT ?
            """,
            (days, float(threshold), int(limit)),
        ).fetchall()
    return [dict(row) for row in rows]


def get_dashboard_analytics() -> dict[str, Any]:
    metrics = database.get_dashboard_metrics()
    start, end = week_bounds()
    metrics["global_payment_stats"] = get_global_payment_stats()
    metrics["flagged_customers"] = get_low_on_time_customers(50.0)
    metrics["week_label"] = f"{start.strftime('%d %b')} – {end.strftime('%d %b %Y')}"
    metrics["formatted_outstanding"] = format_currency(metrics["total_outstanding"])
    return metrics
