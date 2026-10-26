import unittest
from datetime import date, timedelta

import validation


class ValidationTests(unittest.TestCase):
    def test_valid_customer_normalizes_dates_and_currency(self) -> None:
        result = validation.validate_customer({
            "customer_name": "  Example Client  ",
            "customer_email": "hello@example.com",
            "customer_phone": "+44 (0) 1727-555 010",
            "survey_date": "01/01/2025",
            "installation_date": "2025-01-10",
            "last_service_date": "10/01/2026",
            "next_service_date": "2027-01-10",
            "outstanding_invoices": "£1,234.50",
            "paid_invoices": "",
        })
        self.assertTrue(result.is_valid, result.errors)
        self.assertEqual(result.data["customer_name"], "Example Client")
        self.assertEqual(result.data["survey_date"], "2025-01-01")
        self.assertEqual(result.data["last_service_date"], "2026-01-10")
        self.assertEqual(result.data["outstanding_invoices"], 1234.5)
        self.assertEqual(result.data["paid_invoices"], 0.0)

    def test_customer_validation_reports_each_rule(self) -> None:
        result = validation.validate_customer({
            "customer_name": " ",
            "customer_email": "not-an-email",
            "customer_phone": "call me",
            "survey_date": "2025-05-01",
            "installation_date": "2025-04-01",
            "last_service_date": "2025-03-01",
            "next_service_date": "2025-03-01",
            "outstanding_invoices": -5,
            "paid_invoices": "abc",
            "installation_date_extra": "2025-02-30",
        })
        self.assertFalse(result.is_valid)
        for field in (
            "customer_name", "customer_email", "customer_phone", "installation_date",
            "next_service_date", "outstanding_invoices", "paid_invoices",
        ):
            self.assertIn(field, result.errors)
        self.assertIn("•", result.error_summary)

    def test_invalid_calendar_dates_are_rejected(self) -> None:
        result = validation.validate_customer({
            "customer_name": "Client", "next_service_date": "2026-02-30",
        })
        self.assertFalse(result.is_valid)
        self.assertIn("next_service_date", result.errors)

    def test_past_scheduled_event_warns_but_remains_valid(self) -> None:
        yesterday = (date.today() - timedelta(days=1)).isoformat()
        result = validation.validate_event({
            "customer_id": 1,
            "event_date": yesterday,
            "event_time": "9:30 AM",
            "event_type": "Service",
            "status": "Scheduled",
        })
        self.assertTrue(result.is_valid)
        self.assertIn("event_date", result.warnings)
        self.assertEqual(result.data["event_time"], "09:30:00")

    def test_payment_and_invoice_validation(self) -> None:
        payment = validation.validate_payment({
            "customer_id": 1,
            "invoice_date": "2026-01-01",
            "payment_date": "2026-02-01",
            "amount": "£250.00",
            "status": "Paid",
        })
        self.assertTrue(payment.is_valid, payment.errors)
        self.assertEqual(payment.data["amount"], 250.0)

        bad_payment = validation.validate_payment({
            "customer_id": 1,
            "invoice_date": "2026-02-01",
            "payment_date": "2026-01-01",
            "amount": 0,
            "status": "Paid",
        })
        self.assertIn("amount", bad_payment.errors)
        self.assertIn("payment_date", bad_payment.errors)

        invoice = validation.validate_invoice({
            "customer_name": "Client",
            "invoice_date": "2026-01-01",
            "due_date": "2026-01-31",
            "tax_rate": 20,
            "line_items": [
                {"description": "Service", "quantity": 2, "unit_price": 100},
                {"description": "Battery", "quantity": 1, "unit_price": 50},
            ],
        })
        self.assertTrue(invoice.is_valid, invoice.errors)
        self.assertEqual(invoice.data["subtotal"], 250.0)
        self.assertEqual(invoice.data["tax_amount"], 50.0)
        self.assertEqual(invoice.data["amount"], 300.0)

        bad_invoice = validation.validate_invoice({
            "customer_name": "Client", "invoice_date": "2026-02-01",
            "due_date": "2026-01-01", "line_items": [],
        })
        self.assertIn("line_items", bad_invoice.errors)
        self.assertIn("due_date", bad_invoice.errors)


if __name__ == "__main__":
    unittest.main()
