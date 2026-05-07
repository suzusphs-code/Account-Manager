# 🏠 Society Maintenance Manager

A full-featured desktop application for managing apartment society or housing complex maintenance — built entirely with Python and Tkinter, with no web server or database server required.

Handles everything from recording monthly payments and generating receipts, to tracking arrears, managing expenditures, sending WhatsApp reminders, and exporting ledger statements as PDF or Excel — all stored locally in a SQLite database.



## Features

### 🏢 Multi-Society Support
- Manage multiple societies / apartments from a single installation
- Each society gets its own isolated SQLite `.db` file
- Switch between societies at any time from the launcher screen
- Rename or remove societies from the registry without deleting their data

### 💳 Payment Recording
- Record monthly maintenance payments with auto-generated receipt numbers
- Set payment period (month from → month to) with automatic month count and amount validation
- Late fee field with an auto-suggest button based on outstanding arrears
- Detects and flags payments for past financial years (arrears)
- Cross-financial-year payments are automatically split into two linked records

### 🧾 Receipts
- Live receipt preview before saving
- Export receipts as formatted PDF (A4, printable)
- Send receipts directly via WhatsApp (opens WhatsApp Web or the desktop app with the message pre-filled)
- Receipts include society name, address, flat details, payment period breakdown, and late fee line

### 📅 Arrears Tracking
- Automatically calculates unpaid months for every flat from the society's start financial year to today
- Banner on the payment form alerts when a selected flat has outstanding dues
- Pre-fill button populates the form with arrears data in one click
- Handles non-contiguous unpaid months safely — warns instead of overwriting already-paid months
- Choice to book arrears under the original past FY (marks months paid in the matrix) or the current FY (lump-sum receipt)
- Bulk WhatsApp arrears reminders to all defaulters in one action

### 🔔 Unpaid Tracker
- Month-wise view: pick any month + year and see exactly which flats haven't paid
- One-click WhatsApp reminder to individual unpaid residents

### 📊 Yearly Payment Matrix
- Visual grid: all flats as rows, all 12 FY months as columns
- Green = paid (shows amount), Red ✗ = unpaid
- Per-row totals and unpaid count
- Switch between financial years

### 📒 Accounts Ledger
- Full double-entry style ledger per flat account
- Expenditure accounts (one per category)
- Admission fees account
- Master ledger view: all accounts combined with opening balance, Dr, Cr, closing balance
- All-flats summary view with status (Due / Clear / Advance)
- Manual journal entries (Dr/Cr) for adjustments, corrections, and Tally-style entries
- Export any account to Excel (.xlsx) or PDF

### 📊 Dashboard
- Financial year selector
- Summary stat cards: Expected, Collected, Outstanding Dues, Expenditure, Net Position (Surplus/Deficit)
- Month-by-month collection progress bars
- Top defaulters list sorted by amount owed
- Expenditure breakdown by account

### 💰 Admission Fees
- Separate tracking for one-time joining/admission fees
- Receipt generation and PDF export
- WhatsApp message support

### 📄 Reports & Exports
- Individual flat account statement (PDF)
- Expenditure account statement (PDF)
- Batch export: all flat PDFs in one click to a chosen folder
- Master ledger PDF
- All-flats Excel workbook (one sheet per flat + summary sheet)
- Payment records CSV export with filters

### ⚙️ Settings
- Society name and address (shown on all receipts and PDFs)
- Monthly fee amount
- Society start financial year (controls how far back arrears are calculated)
- Unit label (`Flat` / `Shop` / `Office` / `Member` — changes all labels app-wide)
- Expenditure categories (add/remove accounts)

### 🗃️ Data & Storage
- All data stored in a local SQLite database — no internet required
- Automatic daily backup on launch (keeps last 30 days) in a `backups/` subfolder
- `societies.json` registry maps society names to their database files

---

## Requirements

**Python 3.8 or later**

### Required (built-in — no install needed)
- `tkinter` — UI
- `sqlite3` — database
- `json`, `csv`, `datetime`, `webbrowser`, `urllib` — standard library

### Optional
| Package | Purpose | Install |
|---|---|---|
| `reportlab` | PDF receipt and statement generation | `pip install reportlab` |
| `openpyxl` | Excel ledger export | `pip install openpyxl` |

The app runs without either optional package — PDF and Excel buttons simply show an install prompt if the library is missing.

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/society-maintenance-manager.git
cd society-maintenance-manager

# 2. (Optional) Install extras for PDF and Excel support
pip install reportlab openpyxl

# 3. Run
python main.py
```

No virtual environment is strictly required, but recommended:

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install reportlab openpyxl
python main.py
```

### Windows note
The app handles UTF-8 console encoding automatically on Windows. Just run it directly — no extra setup needed.

---

## First Launch

1. The **Society Launcher** opens. Click **+ New Society**.
2. Enter a society name, choose where to save the `.db` file, and set a unit label (`Flat` by default).
3. The main app opens. Go to **⚙ Settings** to set the monthly fee, address, and start financial year.
4. Go to **🏠 Manage Flats** to add your residents (flat number, owner name, mobile).
5. Start recording payments from the main **New Payment** form.

---

## Project Structure

```
society-maintenance-manager/
│
├── main.py   # Main application (single-file)
├── societies.json             # Auto-created: registry of society databases
├── README.md
│
└── (per society, wherever you choose to save them)
    ├── myapartment.db         # SQLite database
    └── backups/
        ├── myapartment_20250501.db
        └── myapartment_20250502.db
```

### Database Tables

| Table | Purpose |
|---|---|
| `flats` | Registered units — flat number, owner name, mobile, display order |
| `payments` | Maintenance payment records with receipt number, period, fee |
| `expenditures` | Expense entries linked to accounts |
| `expenditure_accounts` | Expense categories/accounts |
| `admission_fees` | One-time joining fee records |
| `manual_journals` | Manual Dr/Cr adjustment entries |
| `settings` | Society configuration (name, fee, FY start, etc.) |

---

## Financial Year Convention

The app follows the **Indian financial year**: April to March.

- FY 2024 = April 2024 → March 2025, labelled `FY 2024-25`
- Arrears are calculated from `SOCIETY_START_FY` (set in Settings) up to the current month
- Cross-FY payments (e.g. February → May spanning two FYs) are automatically split

---

## How Arrears Pre-fill Works

1. Select a flat on the payment form — if unpaid months exist, a yellow banner appears.
2. Click **Pre-fill**. If arrears span multiple financial years, a dialog lets you pick which FY to settle.
3. Choose whether to book under the **original past FY** (months get marked paid in the matrix) or the **current FY** (recorded as a lump-sum, months remain unpaid in the historical matrix).
4. The form is populated with flat details, FY, payment period, and calculated amount.
5. If unpaid months are non-contiguous (e.g. Feb paid, Jan + Mar unpaid), the period fields are left blank and a warning is shown — to prevent accidentally marking the paid month as paid again.

---

## WhatsApp Integration

The app opens WhatsApp with a pre-written message — it does **not** send messages automatically. It uses the `whatsapp://` URI scheme (desktop app) with a fallback to `wa.me/` (WhatsApp Web).

Mobile numbers are normalised automatically: `9876543210` → `919876543210`.

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first.

Some areas that could be extended:
- Multi-language / locale support
- Email receipt delivery
- Cloud sync / shared database support
- PyInstaller packaging for a standalone `.exe`

---

## License

MIT License — free to use, modify, and distribute.

---

## Author

Built for real-world apartment society management. All data stays on your machine.
