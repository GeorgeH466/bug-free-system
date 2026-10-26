"""Shared temporary-database fixture for backend tests."""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import database


class ServiceTrackerTestCase(unittest.TestCase):
    def setUp(self) -> None:
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.workspace = Path(temporary.name)
        database.configure(
            db_path=self.workspace / "service_tracker.db",
            exports_dir=self.workspace / "exports",
            template_path=self.workspace / "templates" / "invoice_template.docx",
        )
        self.addCleanup(database.reset_paths)
        database.initialize_database()
