"""Service Tracker desktop application entry point."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

import database
import invoice_generator


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Launch the Service Tracker desktop app.")
    parser.add_argument("--db-path", type=Path, help="Use an alternate SQLite database file.")
    parser.add_argument("--exports-dir", type=Path, help="Use an alternate generated-files folder.")
    parser.add_argument("--template-path", type=Path, help="Use an alternate default Word template.")
    parser.add_argument("--reset-template", action="store_true", help="Recreate the bundled Word template before launch.")
    return parser


def _show_startup_error(message: str) -> None:
    print(f"Service Tracker startup error: {message}", file=sys.stderr)
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Service Tracker", message)
        root.destroy()
    except Exception:
        # Console fallback is already printed above.
        pass


def main(argv: Sequence[str] | None = None) -> int:
    if sys.version_info < (3, 10):
        print("Service Tracker requires Python 3.10 or newer.", file=sys.stderr)
        return 1
    args = build_parser().parse_args(argv)
    database.configure(
        db_path=args.db_path,
        exports_dir=args.exports_dir,
        template_path=args.template_path,
    )
    try:
        database.initialize_database()
        invoice_generator.ensure_default_template(force=bool(args.reset_template))
    except Exception as exc:
        _show_startup_error(f"Could not initialize its local files:\n{exc}")
        return 1

    try:
        import tkinter as tk
    except ImportError:
        print(
            "Tkinter is not installed. On Windows, reinstall Python with the "
            "'tcl/tk and IDLE' option enabled.",
            file=sys.stderr,
        )
        return 1

    try:
        from main_gui import ServiceTrackerApp
    except Exception as exc:
        _show_startup_error(f"Could not load the application:\n{exc}")
        return 1

    root = tk.Tk()
    app = ServiceTrackerApp(root)
    app.show_welcome()
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
