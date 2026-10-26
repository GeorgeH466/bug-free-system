import importlib
import py_compile
import sqlite3
import unittest
from pathlib import Path

import database

try:
    from .base import ServiceTrackerTestCase
except ImportError:  # `unittest discover -s tests` puts tests/ directly on sys.path
    from base import ServiceTrackerTestCase

ROOT = Path(__file__).resolve().parents[1]
REQUIRED_FILES = (
    "main.py", "database.py", "validation.py", "main_gui.py", "customer_form.py",
    "calendar_widget.py", "invoice_generator.py", "spreadsheet_handler.py",
    "email_notifier.py", "analytics.py", "templates/invoice_template.docx",
    "exports/.gitkeep", "requirements.txt", "README.md",
)
REQUIRED_MODULES = (
    "utils", "database", "validation", "analytics", "calendar_widget",
    "customer_form", "invoice_generator", "spreadsheet_handler",
    "email_notifier", "main_gui", "main",
)


class ProjectStructureTests(ServiceTrackerTestCase):
    def test_required_files_exist(self) -> None:
        for relative in REQUIRED_FILES:
            with self.subTest(path=relative):
                self.assertTrue((ROOT / relative).is_file(), f"Missing {relative}")
        template = ROOT / "templates" / "invoice_template.docx"
        self.assertGreater(template.stat().st_size, 10_000)

    def test_all_python_files_compile_and_import_without_tkinter(self) -> None:
        for path in ROOT.glob("*.py"):
            with self.subTest(file=path.name):
                py_compile.compile(str(path), doraise=True)
        for module_name in REQUIRED_MODULES:
            with self.subTest(module=module_name):
                self.assertIsNotNone(importlib.import_module(module_name))

    def test_requirements_match_stack(self) -> None:
        requirements = (ROOT / "requirements.txt").read_text().lower()
        for package in ("openpyxl", "docxtpl", "python-docx", "pywin32"):
            self.assertIn(package, requirements)

    def test_required_indexes_created(self) -> None:
        connection = sqlite3.connect(str(database.get_db_path()))
        try:
            indexes = {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='index'")}
        finally:
            connection.close()
        self.assertIn("idx_customers_name", indexes)
        self.assertIn("idx_customers_next_service", indexes)
        self.assertIn("idx_customers_fault_status", indexes)


if __name__ == "__main__":
    unittest.main()
