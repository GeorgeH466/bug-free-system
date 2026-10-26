# Service Tracker

**Service Tracker** is a local-first Windows desktop application for small trades businesses that need to replace spreadsheets with a searchable customer/service database, validated records, a calendar, Word invoices, Excel import/export, payment analytics, and Outlook email summaries.

All application data is stored in one local SQLite database. There is no cloud dependency.

---

## 1. Requirements

- Windows 10/11
- Python **3.10 or newer**
- Microsoft Outlook desktop installed (required only for sending email)
- `pip` and Tkinter (include the **tcl/tk and IDLE** option when installing Python on Windows)

Python packages are listed in `requirements.txt`:

```text
openpyxl
docxtpl
python-docx
pywin32   # installed only on Windows by the environment marker
```

---

## 2. Installation and launch

From the project directory:

```powershell
python -m pip install -r requirements.txt
python main.py
```

On first launch the application will:

1. create `service_tracker.db`;
2. create `exports/`, `exports/Invoices/`, and `exports/Reports/`;
3. create `templates/invoice_template.docx` if it is missing;
4. create the required tables, indexes, and settings;
5. display a one-time welcome message.

Optional launch options:

```powershell
python main.py --db-path "D:\Data\service_tracker.db"
python main.py --reset-template
python main.py --help
```

---

## 3. Features

### Dashboard

- Total customers, services due this week, overdue services, active faults, outstanding invoices, and global on-time payment rate.
- **Urgent & Pending** list sorted with active faults first and then by the soonest `next_service_date`.
- Payment watchlist flagging customers with an on-time payment rate below 50%.

### Customers

- Searchable, sortable, paged `ttk.Treeview` table designed for thousands of records.
- Add, edit, delete, and detail views; deletes require confirmation.
- A validated form with a field for every writable customer column, custom date pickers, dropdowns, inline error messages, and red highlighting with tooltips.
- Customer detail view shows all fields, event history, invoice/payment history, and per-customer payment analytics.
- Event and invoice records can be added, edited, and deleted from the detail view.

### Calendar

- Custom six-week, Monday-first Tkinter grid; no third-party calendar package.
- Previous/next month, today, and refresh controls.
- Yellow today, light-blue event days, light-red overdue event days, and default white/gray empty days. Today also keeps a gold border when it contains events.
- Click a day to see events, customer names, event types, times, status, notes, and service due dates from customer records.

### Invoices

- Bundled, editable `templates/invoice_template.docx` using Jinja2-style placeholders such as `{{customer_name}}`, `{{customer_location}}`, `{{id_code}}`, `{{amount}}`, `{{invoice_date}}`, `{{service_type}}`, and `{{due_date}}`.
- Customer selection auto-populates customer details.
- Editable invoice number, dates, notes, tax rate, adjustment, and line items.
- `docxtpl` renders a standalone `.docx` invoice that remains fully editable in Microsoft Word.
- Optionally records the generated total in the `payments` table as an outstanding invoice.

### Spreadsheets

- Styled `.xlsx` export with all customer columns, frozen headings, autofilters, date/currency formats, column widths, and an export-information sheet.
- Export filters for search, outstanding invoices, active faults, and next service due before a date.
- Import validates every row using the same rules as the customer form. Invalid rows are skipped and reported with row numbers and field-level reasons.
- Dry-run import validation and a downloadable import template are included.

### Email notifications

- A manually triggered or scheduled weekly Outlook email with:
  - services due this week;
  - overdue services;
  - active faults;
  - total outstanding invoice amount;
  - a multi-sheet `.xlsx` weekly action list attachment.
- Recipient, weekly day/time, automatic-send toggle, and optional Outlook draft preview are configurable in **Settings → Email notifications**.
- `win32com.client.Dispatch("Outlook.Application")` is used on Windows. Outlook automation is never imported at startup, so previews and attachment generation can still be tested on non-Windows systems.

### Payment analytics

The requested `payments` schema has no separate due-date column, so the due date is derived as:

```text
invoice_date + Settings → Invoice due days (default: 30)
```

- **On-time payment rate** = invoices with `payment_date <= derived due date` divided by all dated invoices. Unpaid invoices are included in the denominator.
- **Average days to pay** = mean `payment_date - invoice_date` for paid invoices.
- Per-customer metrics appear in the detail view; the global rate and customers below 50% appear on the dashboard.

---

## 4. Validation rules

Validation runs before customer, event, payment, and invoice data reaches the database.

- `survey_date <= installation_date`
- `installation_date <= last_service_date`
- `last_service_date < next_service_date`
- Email addresses must match a basic email pattern.
- Phone numbers may contain only digits, spaces, `+`, `-`, `(`, and `)`.
- Invoice amounts must be non-negative numbers.
- All date fields must be valid calendar dates. ISO `YYYY-MM-DD` is recommended; `DD/MM/YYYY`, Excel date cells, and several common formats are also accepted.
- A **new scheduled event in the past** produces a warning but is still saved and shown as overdue.

The `events` table includes one small extension column, `status` (`Scheduled`, `Pending`, `Overdue`, `Completed`, or `Cancelled`), so historical completed events are not mistaken for overdue work. All fields requested in the original schema are present.

---

## 5. Project structure

```text
service_tracker/
├── main.py                       # Entry point and first-run initialization
├── database.py                   # SQLite schema, indexes, settings, and CRUD
├── validation.py                 # Pure validation/normalization engine
├── utils.py                      # Date, number, week, and filename helpers
├── main_gui.py                   # Main window, navigation, dashboard, customers
├── customer_form.py              # Add/edit form, detail view, event/payment editors
├── calendar_widget.py            # Monthly calendar, day popup, and DateField
├── invoice_generator.py          # Word invoice template and generation screen
├── spreadsheet_handler.py        # Excel import/export and screen
├── email_notifier.py             # Outlook summary and scheduler
├── analytics.py                  # Payment statistics
├── templates/
│   └── invoice_template.docx     # Editable Word invoice template
├── exports/                      # Runtime exports, invoices, reports, backups
├── service_tracker.db            # Created on first run; ignored by Git
├── requirements.txt
├── README.md
└── tests/                        # Headless backend and document tests
```

Generated database files and most files under `exports/` are ignored by Git. The bundled Word template is tracked.

---

## 6. Importing an Excel workbook

The importer uses the first worksheet and recognizes the export headers or common aliases such as `Name`, `Address`, `Email`, `Phone`, `Survey`, `Installation`, `Next Service`, and `Outstanding`.

Minimum recommended columns:

| Customer Name | Customer Email | Survey Date | Installation Date | Next Service Date | Outstanding Invoices |
|---|---|---|---|---|---|

Use **Import/Export → Download Import Template** to generate a complete workbook with instructions. Select **Validate import only** to check a workbook without writing to the database.

---

## 7. Performance and reliability

- Indexes exist on `customer_name`, `next_service_date`, `fault_status`, event dates, payment dates, and payment status.
- The customer table is paged (200 records by default; configurable from 50–1000 in Settings), so loading the GUI does not deserialize 10,000 records at once.
- Searching and counting are performed in SQL with escaped `LIKE` wildcards and bound parameters.
- Dynamic SQL column names come from fixed whitelists; all values are parameterized.
- Database, file, document, and email operations are wrapped in error handling and show user-friendly dialogs.
- Exports, imports, invoice generation, and Outlook work run off the Tkinter UI thread with a loading indicator.

---

## 8. Tests

The test suite covers the database, validation, analytics, calendar color logic, Excel import/export, invoice rendering, weekly email attachment, scheduler, project structure, and headless module imports.

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

The GUI classes are designed to import cleanly on headless CI systems. Rendering Tkinter windows requires a local desktop session.

---

## 9. Outlook notes

- `pywin32` talks to the locally installed **desktop** version of Outlook; it does not use SMTP credentials or a cloud API.
- Outlook security prompts may appear the first time an external application sends mail.
- The scheduler checks once per minute while the application is open and sends at most once on the configured day after the configured time.
- If Outlook is unavailable, **Preview Summary** still builds the email body and Excel attachment; only the final send is disabled.

---

## 10. Troubleshooting

| Problem | Resolution |
|---|---|
| `No module named tkinter` | Reinstall Python on Windows with “tcl/tk and IDLE” enabled. |
| `No module named docxtpl/openpyxl` | Run `python -m pip install -r requirements.txt`. |
| Outlook email error | Confirm Windows, desktop Outlook, and `pywin32`; try sending a test message manually in Outlook first. |
| Invoice placeholders remain visible | Choose a valid `.docx` template in **Invoices → Choose Template**, or recreate the default template in Settings. |
| Import skipped rows | Open the operation report; each skipped row includes its field-level validation error. |
| Need a clean database | Use Settings → Back Up Database, then remove `service_tracker.db` while the app is closed. |
