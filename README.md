Loan Schedule
=============

A Tkinter desktop application for managing, tracking, reporting, and charting loan schedules.

The app now uses a notebook/tab layout.
Current tab order: Loan Management | Track Loan | Purchases | Reports | Machine Learning | Charts.

Features
--------

### 1) Loan Management Tab

- Entry fields:
  - Bank Name
  - Loan Number
  - Loan Amount
  - Annual Interest Rate (%)
  - Term (years)
  - Start Date (`YYYY-MM-DD`)
- Bank Name and Loan Number support type-ahead auto-suggestions.
- Buttons:
  - Add Loan
  - Update Loan
  - Remove Loan
  - Search Loan
  - Clear Results
  - Export (`.csv` or `.xlsx`)
- Update Loan and Remove Loan prompt for confirmation before applying changes.
- Treeview columns:
  - Bank Name
  - Loan Number
  - Loan Amount
  - Annual Interest Rate (%)
  - Term (years)
  - Start Date
  - End Date
  - Duration Remaining (years or days)

### 2) Track Loan Tab

- Entry fields:
  - Bank Name
  - Loan Number
  - Status (`Paid` / `Upcoming`)
- Bank Name and Loan Number support type-ahead auto-suggestions.
- Buttons:
  - Mark Selected Paid
  - Refresh
  - Search
  - Clear
  - Edit
  - Import (`.csv` or `.xlsx`)
  - Export (`.csv` or `.xlsx`)
  - Template
- Treeview:
  - Grouped by Bank Name + Loan Number
  - Columns: Payment #, Date, Payment, Principal Paid, Interest Paid, Remaining Balance
- Group headers include per-loan summaries:
  - Average Payment
  - Average Principal Paid
  - Average Interest Paid
  - Remaining Balance (current total outstanding for the loan)
- Displays only future/upcoming payments in the tracking tree.

### 3) Purchases Tab

- Entry fields:
  - Number
  - Description
  - Type
  - Purchase Date (`YYYY-MM-DD`)
  - Purchase Price
  - VIN
  - QTY
- Type is an editable dropdown; users can add/edit purchase type options.
- Buttons:
  - Add
  - Update
  - Delete
  - Search
  - Clear
  - Export (`.csv` or `.xlsx`)
- Treeview columns:
  - Number
  - Description
  - Type
  - Purchase Date
  - Purchase Price
  - VIN
  - QTY

### 4) Reports Tab

- Entry fields:
  - Bank Name
  - Loan Number
- Bank Name and Loan Number support type-ahead auto-suggestions.
- Buttons:
  - Run Filtered Report
  - All Active Loan
  - Loan History
  - Export (`.csv` or `.xlsx`)
  - Clear
- Treeview columns:
  - Bank Name, Loan Number, Start Date, End Date, Duration, Status, Loan Amount, Remaining Balance
- Summary view includes:
  - Data Mode (PRIVATE/PUBLIC)
  - Total Open Loans
  - Total Loan Amount
  - Remaining Loan Balance
  - Payment Status Counts (Paid, Upcoming)
  - Styled with Times New Roman, size 12

### 5) Charts Tab

- Entry fields:
  - Bank Name
  - Loan Number
  - Chart Option
  - Start Date
  - End Date
  - Status
- Bank Name and Loan Number support type-ahead auto-suggestions.
- Buttons:
  - Search
  - Display Chart
  - Clear
- Search auto-populates chart Start Date and End Date based on selected Bank Name/Loan Number matches.
- Search updates Chart View and Chart Details to only show rows matching the selected Bank Name and/or Loan Number.
- Display Chart opens Matplotlib in a separate figure window.
- Chart output includes distinct executive dashboard layouts by **Chart Option**:
  - **Overview**: top remaining balances (bar), payment status distribution (pie), cumulative payment/interest trend (line), and loan amount vs remaining (bar)
  - **By Bank**: exposure by bank (grouped bar), payments vs interest by bank (line), remaining-balance share (pie), and payment status counts by bank (stacked bar)
  - **By Loan**: remaining-balance trajectory by payment number (line), total payments by loan (bar), total interest paid by loan (bar), and remaining-balance share (pie)
- In-app "Chart View" shows KPI totals/averages in Times New Roman, size 12, including:
  - Total Payments
  - Total Interest Paid
- Chart details tree columns:
  - Bank Name, Loan Number, Loan Amount, Remaining Balance, Remaining Duration

Data Storage Mode
-----------------

Loan data is stored as JSON:

- If `json_files/` exists in the project root, the app uses that folder for **PRIVATE** data.
- Otherwise, the app uses `data/` for **PUBLIC** sample/fictitious data.

Files are written as:

- `loan_schedule_private.json` or
- `loan_schedule_public.json`

Project Structure
-----------------

- `loan_schedule.py`: Main application code (GUI + logic + persistence).
- `tests/test_loan_schedule.py`: Pytest unit tests for schedule/validation helpers.
- `requirements.txt`: Python dependencies.
- `LICENSE.txt`: MIT license.

Requirements
------------

- Python 3.9+
- Pip

Installation
------------

1. Create and activate a virtual environment (recommended).

Windows (cmd):

```bat
python -m venv .venv
.venv\Scripts\activate
```

2. Install dependencies:

```bat
pip install -r requirements.txt
```

Run the App
-----------

From the project root:

```bat
python loan_schedule.py
```

Run Tests
---------

From the project root:

```bat
pytest -q
```

Validation and Security Notes
-----------------------------

- Input validation enforces:
  - Loan amount > 0
  - Interest rate >= 0
  - Loan term (years) > 0
  - Date fields must use `YYYY-MM-DD`
- User-facing messages are sanitized and avoid exposing raw exceptions.
- Centralized error handling is used for GUI actions.
- Export/import support `.csv` and `.xlsx`.
- Export sanitizes string cells that begin with spreadsheet formula prefixes (`=`, `+`, `-`, `@`) to reduce CSV/XLSX formula-injection risk.

Formula Used
------------

Monthly payment:

`payment = P * (r * (1 + r)^n) / ((1 + r)^n - 1)`

Where:

- `P` = principal
- `r` = monthly interest rate (`annual_rate / 12 / 100`)
- `n` = total number of payments (`years * 12`)

Notes
-----

- Payment dates advance by 30 days per cycle (`timedelta(days=30)`), which is an approximation rather than true calendar months.
- `0%` annual interest is supported (equal principal-only payments).

License
-------

This project is licensed under the MIT License. See `LICENSE.txt` for details.
