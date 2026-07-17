Loan Schedule
=============

A simple desktop app for generating a loan amortization schedule using Python and Tkinter.

The application lets you:
- Enter loan amount, annual interest rate, loan term, and start date.
- Calculate a full payment-by-payment amortization table.
- Export the generated schedule to an Excel file.

Features
--------

- GUI built with Tkinter.
- Fixed monthly payment calculation using the standard amortization formula.
- Full schedule output including:
	- Payment number
	- Payment date
	- Total payment
	- Principal paid
	- Interest paid
	- Remaining balance
- One-click export to `.xlsx`.

Project Structure
-----------------

- `loan_schedule.py`: Main application code (GUI + schedule logic).
- `requirements.txt`: Python dependencies.
- `LICENSE.txt`: MIT license.

Requirements
------------

- Python 3.9+
- Pip

Dependencies (from `requirements.txt`):
- numpy==2.5.1
- pandas==3.0.3
- python-dateutil==2.9.0.post0
- six==1.17.0
- tzdata==2026.3

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

How to Use
----------

1. Fill in the input fields:
	 - Loan Amount
	 - Annual Interest Rate (%)
	 - Term (years)
	 - Start Date (`YYYY-MM-DD`)
2. Click **Calculate Schedule**.
3. Review the amortization table in the text area.
4. Click **Save to Excel** to export the table.

Notes and Limitations
---------------------

- Payment dates are advanced by 30 days each cycle (`timedelta(days=30)`), 
  which is an approximation and not true calendar-month handling.
- Invalid input values (for example malformed dates or non-numeric amounts) 
  are shown via an error dialog.
- The schedule must be calculated before export is available.

Formula Used
------------

Monthly payment is calculated as:

`payment = P * (r * (1 + r)^n) / ((1 + r)^n - 1)`

Where:
- `P` = principal
- `r` = monthly interest rate (`annual_rate / 12 / 100`)
- `n` = total number of payments (`years * 12`)

License
-------

This project is licensed under the MIT License. See `LICENSE.txt` for details.
 
