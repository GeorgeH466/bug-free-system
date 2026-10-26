import analytics
import database

try:
    from .base import ServiceTrackerTestCase
except ImportError:  # `unittest discover -s tests` puts tests/ directly on sys.path
    from base import ServiceTrackerTestCase


class AnalyticsTests(ServiceTrackerTestCase):
    def setUp(self) -> None:
        super().setUp()
        database.set_setting("invoice_due_days", "30")
        self.customer_id = database.add_customer({"customer_name": "Payment Analytics Ltd"})
        database.add_payment({
            "customer_id": self.customer_id,
            "invoice_date": "2026-01-01",
            "payment_date": "2026-01-10",
            "amount": 100,
            "status": "Paid",
        })
        database.add_payment({
            "customer_id": self.customer_id,
            "invoice_date": "2026-02-01",
            "payment_date": "2026-03-20",
            "amount": 200,
            "status": "Paid",
        })
        database.add_payment({
            "customer_id": self.customer_id,
            "invoice_date": "2026-03-01",
            "payment_date": None,
            "amount": 300,
            "status": "Outstanding",
        })

    def test_customer_on_time_rate_and_average_days(self) -> None:
        stats = analytics.get_payment_stats(self.customer_id)
        self.assertEqual(stats["invoice_count"], 3)
        self.assertEqual(stats["paid_count"], 2)
        self.assertEqual(stats["on_time_count"], 1)
        self.assertEqual(stats["on_time_rate"], 33.3)
        self.assertEqual(stats["average_days_to_pay"], 28.0)
        self.assertEqual(stats["formatted_on_time_rate"], "33.3%")
        self.assertEqual(stats["formatted_average_days"], "28.0 days")

    def test_global_stats_and_low_rate_flags(self) -> None:
        stats = analytics.get_global_payment_stats()
        self.assertEqual(stats["invoice_count"], 3)
        self.assertEqual(stats["on_time_rate"], 33.3)
        flagged = analytics.get_low_on_time_customers(50.0)
        self.assertEqual(len(flagged), 1)
        self.assertEqual(flagged[0]["customer_name"], "Payment Analytics Ltd")
        self.assertEqual(flagged[0]["on_time_rate"], 33.3)

    def test_customer_without_invoices_has_no_rate(self) -> None:
        other_id = database.add_customer({"customer_name": "No Invoices"})
        stats = analytics.get_payment_stats(other_id)
        self.assertIsNone(stats["on_time_rate"])
        self.assertIsNone(stats["average_days_to_pay"])
        self.assertEqual(stats["formatted_on_time_rate"], "N/A")

    def test_due_days_setting_changes_late_payment(self) -> None:
        database.set_setting("invoice_due_days", "60")
        stats = analytics.get_payment_stats(self.customer_id)
        self.assertEqual(stats["on_time_count"], 2)
        self.assertEqual(stats["on_time_rate"], 66.7)
