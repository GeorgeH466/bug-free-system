from __future__ import annotations

from datetime import date, timedelta

import database
from utils import week_bounds

try:
    from .base import ServiceTrackerTestCase
except ImportError:  # `unittest discover -s tests` puts tests/ directly on sys.path
    from base import ServiceTrackerTestCase


class DatabaseTests(ServiceTrackerTestCase):
    def test_customer_event_payment_crud_and_cascade(self) -> None:
        customer_id = database.add_customer({
            "customer_name": "Northgate Alarms",
            "customer_location": "14 Bridge Road",
            "customer_email": "office@northgate.example",
            "customer_phone": "+44 1727 555010",
            "survey_date": "2025-01-01",
            "installation_date": "2025-01-15",
            "last_service_date": "2026-01-15",
            "next_service_date": "2027-01-15",
            "service_type": "Maintenance",
            "outstanding_invoices": "£125.50",
            "paid_invoices": 0,
            "fault_status": "Active",
        })
        customer = database.get_customer(customer_id)
        self.assertEqual(customer["customer_name"], "Northgate Alarms")
        self.assertEqual(customer["survey_date"], "2025-01-01")
        self.assertEqual(customer["outstanding_invoices"], 125.5)
        self.assertIsNotNone(customer["created_at"])
        self.assertIsNotNone(customer["updated_at"])

        self.assertTrue(database.update_customer(customer_id, {"next_service_date": "2027-02-15"}))
        self.assertEqual(database.get_customer(customer_id)["next_service_date"], "2027-02-15")

        event_id = database.add_event({
            "customer_id": customer_id,
            "event_date": "2027-02-15",
            "event_time": "09:30",
            "event_type": "Service",
            "event_notes": "Annual visit",
            "status": "Scheduled",
        })
        self.assertEqual(database.get_event(event_id)["event_type"], "Service")
        self.assertTrue(database.update_event(event_id, {"status": "Completed"}))
        self.assertEqual(database.get_event(event_id)["status"], "Completed")

        payment_id = database.add_payment({
            "customer_id": customer_id,
            "invoice_date": "2026-01-20",
            "payment_date": "2026-02-01",
            "amount": 250,
            "status": "Paid",
        })
        self.assertEqual(database.get_payment(payment_id)["status"], "Paid")
        self.assertEqual(len(database.get_payments_for_customer(customer_id)), 1)

        self.assertTrue(database.delete_customer(customer_id))
        self.assertIsNone(database.get_customer(customer_id))
        self.assertIsNone(database.get_event(event_id))
        self.assertIsNone(database.get_payment(payment_id))

    def test_search_is_parameterized_and_index_friendly(self) -> None:
        for index in range(4):
            database.add_customer({
                "customer_name": f"Customer {index}",
                "customer_location": "St Albans",
                "outstanding_invoices": index * 10,
                "fault_status": "Active" if index == 1 else "None",
            })
        self.assertEqual(database.count_customers(), 4)
        self.assertEqual(len(database.search_customers("Customer")), 4)
        self.assertEqual(database.count_customers(filters={"outstanding_only": True}), 3)
        self.assertEqual(database.count_customers(filters={"active_faults_only": True}), 1)
        page, total = database.get_customer_page(None, {}, 2, 0)
        self.assertEqual(total, 4)
        self.assertEqual(len(page), 2)

        malicious = "'; DROP TABLE customers; --"
        self.assertEqual(database.count_customers(malicious), 0)
        self.assertEqual(database.count_customers(), 4)
        self.assertEqual(database.count_customers("100%_"), 0)

    def test_dashboard_and_settings_queries(self) -> None:
        today = date.today()
        _, end = week_bounds(today)
        due_date = min(end, today + timedelta(days=1))
        database.add_customer({
            "customer_name": "Due This Week",
            "next_service_date": due_date.isoformat(),
            "fault_status": "Active",
            "outstanding_invoices": 100,
        })
        database.add_customer({
            "customer_name": "Overdue",
            "next_service_date": (today - timedelta(days=10)).isoformat(),
            "outstanding_invoices": 50,
        })
        metrics = database.get_dashboard_metrics(today)
        self.assertEqual(metrics["total_customers"], 2)
        self.assertEqual(metrics["due_this_week"], 1)
        self.assertEqual(metrics["overdue_services"], 1)
        self.assertEqual(metrics["total_outstanding"], 150.0)
        self.assertEqual(len(database.get_urgent_customers()), 2)
        self.assertEqual(database.get_urgent_customers(1)[0]["customer_name"], "Due This Week")
        self.assertEqual(database.get_overdue_services(today)[0]["customer_name"], "Overdue")

        self.assertIsNone(database.get_setting("missing", None))
        database.set_setting("email_recipient", "manager@example.com")
        self.assertEqual(database.get_setting("email_recipient"), "manager@example.com")
        self.assertIn("invoice_due_days", database.get_all_settings())
