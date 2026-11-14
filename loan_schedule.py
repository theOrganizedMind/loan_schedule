import calendar
import csv
import json
import logging
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd

try:
    import matplotlib.pyplot as plt
except ImportError:  # pragma: no cover - optional runtime dependency
    plt = None

try:
    from openpyxl import Workbook, load_workbook
except ImportError:  # pragma: no cover - optional runtime dependency
    Workbook = None
    load_workbook = None

LOGGER = logging.getLogger(__name__)
if not LOGGER.handlers:
    logging.basicConfig(level=logging.INFO)

DATE_FORMAT = "%Y-%m-%d"
TRACK_STATUSES = ("Paid", "Upcoming")
LOAN_HEADERS = [
    "Bank Name",
    "Loan Number",
    "Loan Amount",
    "Annual Interest Rate (%)",
    "Term (years)",
    "Start Date (YYYY-MM-DD)",
    "End Date (YYYY-MM-DD)",
    "Duration Remaining",
]
PAYMENT_HEADERS = [
    "Bank Name",
    "Loan Number",
    "Payment #",
    "Date",
    "Payment",
    "Principal Paid",
    "Interest Paid",
    "Remaining Balance",
    "Status",
]
PURCHASE_HEADERS = [
    "Number",
    "Description",
    "Type",
    "Purchase Date",
    "Purchase Price",
    "VIN",
    "QTY",
]
REPORT_HEADERS = [
    "Bank Name",
    "Loan Number",
    "Start Date",
    "End Date",
    "Duration",
    "Status",
    "Loan Amount",
    "Remaining Balance",
]
CHART_OPTIONS = ("Overview", "By Bank", "By Loan")
ML_OPTION_HEADERS = ("Extra/Month", "Months Saved", "Interest Saved", "New Payoff Date")
DEFAULT_PURCHASE_TYPES = ["Truck", "Dumpsters", "Equipment"]


def parse_date(value: str, label: str) -> datetime:
    """Parse a YYYY-MM-DD string into a datetime, raising a labeled validation error."""
    try:
        return datetime.strptime(value, DATE_FORMAT)
    except ValueError as exc:
        raise ValueError(f"{label} must use YYYY-MM-DD format.") from exc


def resolve_data_directory(base_dir: str | None = None) -> tuple[str, str]:
    """Return the preferred data directory path and visibility mode (PRIVATE or PUBLIC)."""
    root = base_dir or os.path.dirname(os.path.abspath(__file__))
    private_dir = os.path.join(root, "json_files")
    public_dir = os.path.join(root, "data")
    if os.path.isdir(private_dir):
        return private_dir, "PRIVATE"
    os.makedirs(public_dir, exist_ok=True)
    return public_dir, "PUBLIC"


def _validate_loan_inputs(principal: float, annual_rate: float, years: int, start_date: datetime) -> None:
    """Validate core loan calculation inputs before building an amortization schedule."""
    if principal <= 0:
        raise ValueError("Loan amount must be greater than 0.")
    if annual_rate < 0:
        raise ValueError("Annual interest rate cannot be negative.")
    if years <= 0:
        raise ValueError("Loan term (years) must be greater than 0.")
    if not isinstance(start_date, datetime):
        raise TypeError("Start date must be a datetime object.")


def next_monthly_payment_date(payment_date: datetime, anchor_day: int) -> datetime:
    """Return the next monthly payment date while preserving the original day-of-month."""
    if payment_date.month == 12:
        next_year = payment_date.year + 1
        next_month = 1
    else:
        next_year = payment_date.year
        next_month = payment_date.month + 1
    last_day = calendar.monthrange(next_year, next_month)[1]
    next_day = min(anchor_day, last_day)
    return payment_date.replace(year=next_year, month=next_month, day=next_day)


def calculate_schedule(principal: float, annual_rate: float, years: int, start_date: datetime) -> pd.DataFrame:
    """Calculate and return a payment-by-payment amortization schedule as a DataFrame."""
    _validate_loan_inputs(principal, annual_rate, years, start_date)
    monthly_rate = annual_rate / 12 / 100
    n_payments = years * 12
    if monthly_rate == 0:
        payment = principal / n_payments
    else:
        payment = principal * (monthly_rate * (1 + monthly_rate) ** n_payments) / ((1 + monthly_rate) ** n_payments - 1)

    schedule = []
    balance = principal
    payment_date = start_date
    anchor_day = start_date.day
    for i in range(1, n_payments + 1):
        interest = balance * monthly_rate
        principal_paid = payment - interest
        balance -= principal_paid
        if balance < 1e-2:
            principal_paid += balance
            balance = 0
        schedule.append(
            {
                "Payment #": i,
                "Date": payment_date.strftime(DATE_FORMAT),
                "Payment": round(payment, 2),
                "Principal Paid": round(principal_paid, 2),
                "Interest Paid": round(interest, 2),
                "Remaining Balance": round(balance, 2),
            }
        )
        payment_date = next_monthly_payment_date(payment_date, anchor_day)
        if balance <= 0:
            break
    return pd.DataFrame(schedule)


def _monthly_payment_amount(principal: float, annual_rate: float, years: int) -> float:
    """Return the scheduled monthly payment amount for an amortizing loan."""
    monthly_rate = annual_rate / 12 / 100
    n_payments = years * 12
    if monthly_rate == 0:
        return principal / n_payments
    return principal * (monthly_rate * (1 + monthly_rate) ** n_payments) / ((1 + monthly_rate) ** n_payments - 1)


def simulate_extra_payment_metrics(
    principal: float,
    annual_rate: float,
    years: int,
    start_date: datetime,
    extra_monthly_payment: float,
) -> dict[str, object]:
    """Simulate payoff metrics when adding a fixed extra monthly payment."""
    _validate_loan_inputs(principal, annual_rate, years, start_date)
    if extra_monthly_payment < 0:
        raise ValueError("Extra monthly payment cannot be negative.")

    monthly_rate = annual_rate / 12 / 100
    scheduled_payment = _monthly_payment_amount(principal, annual_rate, years)
    payment_amount = scheduled_payment + extra_monthly_payment

    balance = principal
    payment_date = start_date
    anchor_day = start_date.day
    months = 0
    total_interest = 0.0
    payoff_date = start_date
    max_iterations = years * 12 * 2

    while balance > 1e-8 and months < max_iterations:
        interest = balance * monthly_rate
        principal_paid = payment_amount - interest
        if principal_paid <= 0:
            raise ValueError("Monthly payment is too low to reduce principal.")
        principal_paid = min(principal_paid, balance)
        balance -= principal_paid
        total_interest += interest
        months += 1
        payoff_date = payment_date
        payment_date = next_monthly_payment_date(payment_date, anchor_day)

    if months >= max_iterations and balance > 1e-8:
        raise ValueError("Unable to compute payoff metrics for the provided loan values.")

    return {
        "months": months,
        "total_interest": round(total_interest, 2),
        "payoff_date": payoff_date.strftime(DATE_FORMAT),
        "scheduled_payment": round(scheduled_payment, 2),
        "total_payment": round(payment_amount, 2),
    }


def _fit_linear_model(features: list[float], targets: list[float]) -> tuple[float, float]:
    """Fit a simple y = intercept + slope*x linear model."""
    if not features or len(features) != len(targets):
        raise ValueError("Feature and target lists must be non-empty and aligned.")
    n = len(features)
    mean_x = sum(features) / n
    mean_y = sum(targets) / n
    numerator = sum((x - mean_x) * (y - mean_y) for x, y in zip(features, targets))
    denominator = sum((x - mean_x) ** 2 for x in features)
    slope = 0.0 if denominator == 0 else numerator / denominator
    intercept = mean_y - slope * mean_x
    return intercept, slope


def format_duration_remaining(end_date_value: str, today: datetime | None = None) -> str:
    """Format remaining duration to an end date using days for short spans and years for long spans."""
    today_value = (today or datetime.today()).date()
    end_date = parse_date(end_date_value, "End Date").date()
    delta_days = (end_date - today_value).days
    if delta_days <= 0:
        return "0 days"
    if delta_days > 365:
        years = round(delta_days / 365, 1)
        return f"{years} years"
    return f"{delta_days} days"


def build_payment_rows(bank_name: str, loan_number: str, schedule: pd.DataFrame, today: datetime | None = None) -> list[dict[str, object]]:
    """Convert a schedule DataFrame into persisted payment row dictionaries with computed status."""
    today_date = (today or datetime.today()).date()
    rows: list[dict[str, object]] = []
    for _, row in schedule.iterrows():
        payment_date = parse_date(str(row["Date"]), "Payment Date").date()
        status = "Upcoming" if payment_date >= today_date else "Paid"
        rows.append(
            {
                "bank_name": bank_name,
                "loan_number": loan_number,
                "payment_number": int(row["Payment #"]),
                "date": str(row["Date"]),
                "payment": float(row["Payment"]),
                "principal_paid": float(row["Principal Paid"]),
                "interest_paid": float(row["Interest Paid"]),
                "remaining_balance": float(row["Remaining Balance"]),
                "status": status,
            }
        )
    return rows


class LoanScheduleApp:
    def __init__(self, root: tk.Tk):
        """Initialize application state, load persisted data, and build the tabbed interface."""
        self.root = root
        self.root.title("Loan Management and Tracking")
        self.root.geometry("1400x900")

        data_directory, visibility = resolve_data_directory()
        self.data_directory = data_directory
        self.data_visibility = visibility
        self.data_file = os.path.join(data_directory, f"loan_schedule_{visibility.lower()}.json")
        self.data = {"loans": [], "payments": [], "purchases": [], "purchase_types": list(DEFAULT_PURCHASE_TYPES)}
        self.current_report_rows: list[dict[str, object]] = []

        self.load_data()
        self.build_ui()
        self.refresh_all()

    def _safe_execute(self, action, *, title: str = "Operation Failed") -> None:
        """Execute a callback with centralized user-safe exception handling."""
        try:
            action()
        except ValueError as exc:
            LOGGER.warning("Validation error during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showwarning("Validation Error", str(exc))
        except (OSError, json.JSONDecodeError) as exc:
            LOGGER.exception("Data I/O error during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showerror(title, "A data access error occurred. Please try again.")
        except ImportError as exc:
            LOGGER.warning("Missing dependency during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showerror(title, str(exc))
        except Exception as exc:  # pragma: no cover - Tk runtime protection
            LOGGER.exception("Unexpected error during %s: %s", getattr(action, "__name__", "action"), exc)
            messagebox.showerror(title, "An unexpected application error occurred.")

    def load_data(self) -> None:
        """Load loans and payment records from the configured JSON file."""
        if not os.path.exists(self.data_file):
            self.save_data()
            return
        with open(self.data_file, "r", encoding="utf-8") as handle:
            loaded = json.load(handle)
        loans = loaded.get("loans", [])
        payments = loaded.get("payments", [])
        purchases = loaded.get("purchases", [])
        purchase_types = loaded.get("purchase_types", list(DEFAULT_PURCHASE_TYPES))
        if (
            not isinstance(loans, list)
            or not isinstance(payments, list)
            or not isinstance(purchases, list)
            or not isinstance(purchase_types, list)
        ):
            raise ValueError("Data file has invalid structure.")
        cleaned_types = [str(item).strip() for item in purchase_types if str(item).strip()]
        self.data = {
            "loans": loans,
            "payments": payments,
            "purchases": purchases,
            "purchase_types": cleaned_types or list(DEFAULT_PURCHASE_TYPES),
        }

    def save_data(self) -> None:
        """Persist current in-memory loan and payment data to disk."""
        os.makedirs(self.data_directory, exist_ok=True)
        with open(self.data_file, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, indent=2)

    def build_ui(self) -> None:
        """Create the notebook container and build all tab UIs."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_loans = ttk.Frame(self.notebook)
        self.tab_track = ttk.Frame(self.notebook)
        self.tab_purchases = ttk.Frame(self.notebook)
        self.tab_reports = ttk.Frame(self.notebook)
        self.tab_ml = ttk.Frame(self.notebook)
        self.tab_charts = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_loans, text="📕Loan Management")
        self.notebook.add(self.tab_track, text="📍Track Loan")
        self.notebook.add(self.tab_purchases, text="🛒Purchases")
        self.notebook.add(self.tab_reports, text="📄Reports")
        self.notebook.add(self.tab_ml, text="🤖Machine Learning")
        self.notebook.add(self.tab_charts, text="📈Charts")

        self.build_loan_tab()
        self.build_track_tab()
        self.build_purchases_tab()
        self.build_reports_tab()
        self.build_machine_learning_tab()
        self.build_charts_tab()

    def build_loan_tab(self) -> None:
        """Build the Loan Management tab fields, actions, and results table."""
        form = ttk.LabelFrame(
            self.tab_loans,
            text=f"Loan Management ({self.data_visibility} data file: {os.path.basename(self.data_file)})",
        )
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(form, text="Bank Name").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.loan_bank_entry = ttk.Combobox(form, state="normal", width=24)
        self.loan_bank_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Loan Number").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.loan_number_entry = ttk.Combobox(form, state="normal", width=24)
        self.loan_number_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Loan Amount").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self.loan_amount_entry = ttk.Entry(form, width=18)
        self.loan_amount_entry.grid(row=0, column=5, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Annual Interest Rate (%)").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.loan_rate_entry = ttk.Entry(form, width=24)
        self.loan_rate_entry.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Term (years)").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.loan_years_entry = ttk.Entry(form, width=24)
        self.loan_years_entry.grid(row=1, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Start Date (YYYY-MM-DD)").grid(row=1, column=4, sticky="e", padx=6, pady=6)
        self.loan_start_entry = ttk.Entry(form, width=18)
        self.loan_start_entry.grid(row=1, column=5, sticky="w", padx=6, pady=6)
        self.loan_start_entry.insert(0, datetime.today().strftime(DATE_FORMAT))

        button_row = ttk.Frame(form)
        button_row.grid(row=2, column=0, columnspan=6, sticky="ew", padx=6, pady=6)

        ttk.Button(button_row, text="Add Loan", command=lambda: self._safe_execute(self.add_loan)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="Update Loan", command=lambda: self._safe_execute(self.update_loan)).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(button_row, text="Remove Loan", command=lambda: self._safe_execute(self.remove_loan)).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(button_row, text="Search Loan", command=lambda: self._safe_execute(self.search_loans)).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(button_row, text="Clear Results", command=lambda: self._safe_execute(self.clear_loan_results)).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(button_row, text="Export", command=lambda: self._safe_execute(self.export_loans)).grid(row=0, column=5, padx=4, pady=4)
        button_row.columnconfigure(6, weight=1)

        table_frame = ttk.LabelFrame(self.tab_loans, text="Loan List")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = ("bank_name", "loan_number", "loan_amount", "annual_rate", "term_years", "start_date", "end_date", "duration_remaining")
        self.loan_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        widths = (180, 140, 130, 160, 120, 140, 140, 160)
        for col, heading, width in zip(columns, LOAN_HEADERS, widths):
            self.loan_tree.heading(col, text=heading)
            self.loan_tree.column(col, width=width, anchor="w")
        loan_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.loan_tree.yview)
        self.loan_tree.configure(yscrollcommand=loan_scroll.set)
        self.loan_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        loan_scroll.grid(row=0, column=1, sticky="ns", pady=6)
        self.loan_tree.bind("<<TreeviewSelect>>", self.on_loan_selected)

        self.tab_loans.rowconfigure(1, weight=1)
        self.tab_loans.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def build_track_tab(self) -> None:
        """Build the Track Loan tab fields, actions, and grouped payment tree."""
        form = ttk.LabelFrame(self.tab_track, text="Track Loan")
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(form, text="Bank Name").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.track_bank_entry = ttk.Combobox(form, state="normal", width=24)
        self.track_bank_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Loan Number").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.track_loan_entry = ttk.Combobox(form, state="normal", width=24)
        self.track_loan_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Status").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self.track_status_combo = ttk.Combobox(form, state="readonly", values=TRACK_STATUSES, width=18)
        self.track_status_combo.grid(row=0, column=5, sticky="w", padx=6, pady=6)
        self.track_status_combo.set("Upcoming")

        button_row = ttk.Frame(form)
        button_row.grid(row=1, column=0, columnspan=6, sticky="ew", padx=6, pady=6)
        ttk.Button(button_row, text="Mark Selected Paid", command=lambda: self._safe_execute(self.mark_selected_paid)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="Refresh", command=lambda: self._safe_execute(self.refresh_track_tree)).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(button_row, text="Search", command=lambda: self._safe_execute(self.search_track_rows)).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(button_row, text="Clear", command=lambda: self._safe_execute(self.clear_track_filters)).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(button_row, text="Edit", command=lambda: self._safe_execute(self.edit_selected_payment)).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(button_row, text="Import", command=lambda: self._safe_execute(self.import_payments)).grid(row=0, column=5, padx=4, pady=4)
        ttk.Button(button_row, text="Export", command=lambda: self._safe_execute(self.export_track_rows)).grid(row=0, column=6, padx=4, pady=4)
        ttk.Button(button_row, text="Template", command=lambda: self._safe_execute(self.export_payment_template)).grid(row=0, column=7, padx=4, pady=4)
        button_row.columnconfigure(8, weight=1)

        table_frame = ttk.LabelFrame(self.tab_track, text="Upcoming Payments Grouped by Bank and Loan")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = ("payment_number", "date", "payment", "principal_paid", "interest_paid", "remaining_balance")
        self.track_tree = ttk.Treeview(table_frame, columns=columns, show="tree headings", height=15)
        self.track_tree.heading("#0", text="Bank Name / Loan Number")
        self.track_tree.column("#0", width=280, anchor="w")
        headings = ("Payment #", "Date", "Payment", "Principal Paid", "Interest Paid", "Remaining Balance")
        widths = (90, 130, 120, 130, 120, 140)
        for col, heading, width in zip(columns, headings, widths):
            self.track_tree.heading(col, text=heading)
            self.track_tree.column(col, width=width, anchor="w")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.track_tree.yview)
        self.track_tree.configure(yscrollcommand=scroll.set)
        self.track_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scroll.grid(row=0, column=1, sticky="ns", pady=6)

        self.tab_track.rowconfigure(1, weight=1)
        self.tab_track.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def _purchase_types(self) -> list[str]:
        """Return normalized purchase type values for the Purchases tab dropdown."""
        values = [str(item).strip() for item in self.data.get("purchase_types", []) if str(item).strip()]
        return values or list(DEFAULT_PURCHASE_TYPES)

    def _set_purchase_types(self, raw_values: list[str]) -> None:
        """Store unique purchase types while preserving the first-seen order."""
        unique_values: list[str] = []
        seen: set[str] = set()
        for value in raw_values:
            cleaned = str(value).strip()
            if not cleaned:
                continue
            key = cleaned.lower()
            if key in seen:
                continue
            seen.add(key)
            unique_values.append(cleaned)
        self.data["purchase_types"] = unique_values or list(DEFAULT_PURCHASE_TYPES)

    def refresh_purchase_type_values(self) -> None:
        """Refresh purchase type combobox suggestions from persisted type values."""
        if hasattr(self, "purchase_type_entry"):
            self.purchase_type_entry.configure(values=self._purchase_types())

    def edit_purchase_types(self) -> None:
        """Edit purchase type options using a comma-separated dialog list."""
        current = ", ".join(self._purchase_types())
        entered = simpledialog.askstring(
            "Edit Purchase Types",
            "Enter purchase types separated by commas:",
            initialvalue=current,
        )
        if entered is None:
            return
        normalized = entered.replace("\n", ",").replace(";", ",")
        values = [item.strip() for item in normalized.split(",") if item.strip()]
        if not values:
            raise ValueError("At least one purchase type is required.")
        self._set_purchase_types(values)
        self.save_data()
        self.refresh_purchase_type_values()

    def build_purchases_tab(self) -> None:
        """Build the Purchases tab fields, actions, and purchase table."""
        form = ttk.LabelFrame(self.tab_purchases, text="Purchases")
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(form, text="Number").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.purchase_number_entry = ttk.Entry(form, width=24)
        self.purchase_number_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Description").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.purchase_description_entry = ttk.Entry(form, width=24)
        self.purchase_description_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Type").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self.purchase_type_entry = ttk.Combobox(form, state="normal", width=22)
        self.purchase_type_entry.grid(row=0, column=5, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Purchase Date").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.purchase_date_entry = ttk.Entry(form, width=24)
        self.purchase_date_entry.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        self.purchase_date_entry.insert(0, datetime.today().strftime(DATE_FORMAT))

        ttk.Label(form, text="Purchase Price").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.purchase_price_entry = ttk.Entry(form, width=24)
        self.purchase_price_entry.grid(row=1, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="VIN").grid(row=1, column=4, sticky="e", padx=6, pady=6)
        self.purchase_vin_entry = ttk.Entry(form, width=22)
        self.purchase_vin_entry.grid(row=1, column=5, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="QTY").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.purchase_qty_entry = ttk.Entry(form, width=24)
        self.purchase_qty_entry.grid(row=2, column=1, sticky="w", padx=6, pady=6)

        ttk.Button(form, text="Edit Types", command=lambda: self._safe_execute(self.edit_purchase_types)).grid(row=2, column=5, sticky="e", padx=6, pady=6)

        button_row = ttk.Frame(form)
        button_row.grid(row=3, column=0, columnspan=6, sticky="ew", padx=6, pady=6)
        ttk.Button(button_row, text="Add", command=lambda: self._safe_execute(self.add_purchase)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="Update", command=lambda: self._safe_execute(self.update_purchase)).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(button_row, text="Delete", command=lambda: self._safe_execute(self.delete_purchase)).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(button_row, text="Search", command=lambda: self._safe_execute(self.search_purchases)).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(button_row, text="Clear", command=lambda: self._safe_execute(self.clear_purchase_results)).grid(row=0, column=4, padx=4, pady=4)
        ttk.Button(button_row, text="Export", command=lambda: self._safe_execute(self.export_purchases)).grid(row=0, column=5, padx=4, pady=4)
        button_row.columnconfigure(6, weight=1)

        table_frame = ttk.LabelFrame(self.tab_purchases, text="Purchase List")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = ("number", "description", "type", "purchase_date", "purchase_price", "vin", "qty")
        self.purchase_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        widths = (140, 260, 140, 150, 140, 200, 100)
        for col, heading, width in zip(columns, PURCHASE_HEADERS, widths):
            self.purchase_tree.heading(col, text=heading)
            self.purchase_tree.column(col, width=width, anchor="w")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.purchase_tree.yview)
        self.purchase_tree.configure(yscrollcommand=scroll.set)
        self.purchase_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scroll.grid(row=0, column=1, sticky="ns", pady=6)
        self.purchase_tree.bind("<<TreeviewSelect>>", self.on_purchase_selected)

        self.tab_purchases.rowconfigure(1, weight=1)
        self.tab_purchases.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.refresh_purchase_type_values()

    def _filtered_purchases(
        self,
        number_filter: str = "",
        description_filter: str = "",
        type_filter: str = "",
        date_filter: str = "",
        price_filter: str = "",
        vin_filter: str = "",
        qty_filter: str = "",
    ) -> list[dict[str, object]]:
        """Return purchases matching the provided filter values."""
        number_filter = number_filter.strip().lower()
        description_filter = description_filter.strip().lower()
        type_filter = type_filter.strip().lower()
        date_filter = date_filter.strip().lower()
        price_filter = price_filter.strip().lower()
        vin_filter = vin_filter.strip().lower()
        qty_filter = qty_filter.strip().lower()
        rows: list[dict[str, object]] = []
        for purchase in self.data["purchases"]:
            number = str(purchase.get("number", ""))
            description = str(purchase.get("description", ""))
            purchase_type = str(purchase.get("type", ""))
            purchase_date = str(purchase.get("purchase_date", ""))
            purchase_price = f"{float(purchase.get('purchase_price', 0.0)):.2f}"
            vin = str(purchase.get("vin", ""))
            qty = str(int(purchase.get("qty", 0)))
            if number_filter and number_filter not in number.lower():
                continue
            if description_filter and description_filter not in description.lower():
                continue
            if type_filter and type_filter not in purchase_type.lower():
                continue
            if date_filter and date_filter not in purchase_date.lower():
                continue
            if price_filter and price_filter not in purchase_price.lower():
                continue
            if vin_filter and vin_filter not in vin.lower():
                continue
            if qty_filter and qty_filter not in qty.lower():
                continue
            rows.append(purchase)
        rows.sort(key=lambda item: str(item.get("number", "")).lower())
        return rows

    def refresh_purchases_tree(self, purchases: list[dict[str, object]]) -> None:
        """Render the provided purchases into the Purchases treeview."""
        self.purchase_tree.delete(*self.purchase_tree.get_children())
        for purchase in purchases:
            self.purchase_tree.insert(
                "",
                "end",
                values=(
                    purchase.get("number", ""),
                    purchase.get("description", ""),
                    purchase.get("type", ""),
                    purchase.get("purchase_date", ""),
                    f"{float(purchase.get('purchase_price', 0.0)):.2f}",
                    purchase.get("vin", ""),
                    int(purchase.get("qty", 0)),
                ),
            )

    def parse_purchase_form(self) -> dict[str, object]:
        """Parse and validate purchase form values into a normalized purchase record."""
        number = self.purchase_number_entry.get().strip()
        description = self.purchase_description_entry.get().strip()
        purchase_type = self.purchase_type_entry.get().strip()
        purchase_date = self.purchase_date_entry.get().strip()
        purchase_price = float(self.purchase_price_entry.get().strip())
        vin = self.purchase_vin_entry.get().strip()
        qty = int(self.purchase_qty_entry.get().strip())

        if not number:
            raise ValueError("Number is required.")
        if not description:
            raise ValueError("Description is required.")
        if not purchase_type:
            raise ValueError("Type is required.")
        parse_date(purchase_date, "Purchase Date")
        if purchase_price < 0:
            raise ValueError("Purchase Price cannot be negative.")
        if qty <= 0:
            raise ValueError("QTY must be greater than 0.")

        current_types = self._purchase_types()
        if purchase_type.lower() not in {item.lower() for item in current_types}:
            current_types.append(purchase_type)
            self._set_purchase_types(current_types)
            self.refresh_purchase_type_values()

        return {
            "number": number,
            "description": description,
            "type": purchase_type,
            "purchase_date": purchase_date,
            "purchase_price": round(purchase_price, 2),
            "vin": vin,
            "qty": qty,
        }

    def add_purchase(self) -> None:
        """Add a new purchase record."""
        parsed = self.parse_purchase_form()
        number_key = str(parsed["number"]).lower()
        for purchase in self.data["purchases"]:
            if str(purchase.get("number", "")).lower() == number_key:
                raise ValueError("Purchase Number already exists. Use Update instead.")
        self.data["purchases"].append(parsed)
        self.save_data()
        self.refresh_purchases_tree(self.data["purchases"])
        messagebox.showinfo("Saved", "Purchase added successfully.")

    def _selected_purchase_number(self) -> str:
        """Return the selected purchase number from the Purchases tree."""
        selected = self.purchase_tree.selection()
        if not selected:
            raise ValueError("Select a purchase row.")
        values = self.purchase_tree.item(selected[0], "values")
        if not values:
            raise ValueError("Select a valid purchase row.")
        return str(values[0]).strip()

    def update_purchase(self) -> None:
        """Update the selected purchase record."""
        selected_number = self._selected_purchase_number()
        parsed = self.parse_purchase_form()
        selected_key = selected_number.lower()
        updated_key = str(parsed["number"]).lower()

        target_index = next(
            (idx for idx, purchase in enumerate(self.data["purchases"]) if str(purchase.get("number", "")).lower() == selected_key),
            None,
        )
        if target_index is None:
            raise ValueError("Selected purchase no longer exists.")
        if updated_key != selected_key:
            for idx, purchase in enumerate(self.data["purchases"]):
                if idx == target_index:
                    continue
                if str(purchase.get("number", "")).lower() == updated_key:
                    raise ValueError("Purchase Number already exists. Use a unique Number.")

        self.data["purchases"][target_index] = parsed
        self.save_data()
        self.refresh_purchases_tree(self.data["purchases"])
        messagebox.showinfo("Updated", "Purchase updated successfully.")

    def delete_purchase(self) -> None:
        """Delete the selected purchase record after confirmation."""
        selected_number = self._selected_purchase_number()
        confirmed = messagebox.askokcancel("Confirm Delete", f"Delete purchase '{selected_number}'?", icon="warning")
        if not confirmed:
            return
        selected_key = selected_number.lower()
        self.data["purchases"] = [
            purchase for purchase in self.data["purchases"] if str(purchase.get("number", "")).lower() != selected_key
        ]
        self.save_data()
        self.refresh_purchases_tree(self.data["purchases"])
        self.clear_purchase_entries()
        messagebox.showinfo("Deleted", "Purchase deleted successfully.")

    def search_purchases(self) -> None:
        """Filter and display purchase rows matching current form filters."""
        rows = self._filtered_purchases(
            number_filter=self.purchase_number_entry.get(),
            description_filter=self.purchase_description_entry.get(),
            type_filter=self.purchase_type_entry.get(),
            date_filter=self.purchase_date_entry.get(),
            price_filter=self.purchase_price_entry.get(),
            vin_filter=self.purchase_vin_entry.get(),
            qty_filter=self.purchase_qty_entry.get(),
        )
        self.refresh_purchases_tree(rows)

    def clear_purchase_entries(self, *, include_default_date: bool = True) -> None:
        """Clear purchase form entries and optionally reset purchase date to today."""
        self.purchase_number_entry.delete(0, tk.END)
        self.purchase_description_entry.delete(0, tk.END)
        self.purchase_type_entry.delete(0, tk.END)
        self.purchase_date_entry.delete(0, tk.END)
        self.purchase_price_entry.delete(0, tk.END)
        self.purchase_vin_entry.delete(0, tk.END)
        self.purchase_qty_entry.delete(0, tk.END)
        if include_default_date:
            self.purchase_date_entry.insert(0, datetime.today().strftime(DATE_FORMAT))

    def clear_purchase_results(self) -> None:
        """Clear purchase filters and show all purchase rows."""
        self.clear_purchase_entries()
        self.refresh_purchases_tree(self.data["purchases"])

    def on_purchase_selected(self, _event=None) -> None:
        """Populate purchase form fields from the selected purchase row."""
        selected = self.purchase_tree.selection()
        if not selected:
            return
        values = self.purchase_tree.item(selected[0], "values")
        if not values:
            return
        self.clear_purchase_entries(include_default_date=False)
        self.purchase_number_entry.insert(0, str(values[0]))
        self.purchase_description_entry.insert(0, str(values[1]))
        self.purchase_type_entry.insert(0, str(values[2]))
        self.purchase_date_entry.insert(0, str(values[3]))
        self.purchase_price_entry.insert(0, str(values[4]))
        self.purchase_vin_entry.insert(0, str(values[5]))
        self.purchase_qty_entry.insert(0, str(values[6]))

    def export_purchases(self) -> None:
        """Export filtered purchase rows to CSV or XLSX."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        rows = self._filtered_purchases(
            number_filter=self.purchase_number_entry.get(),
            description_filter=self.purchase_description_entry.get(),
            type_filter=self.purchase_type_entry.get(),
            date_filter=self.purchase_date_entry.get(),
            price_filter=self.purchase_price_entry.get(),
            vin_filter=self.purchase_vin_entry.get(),
            qty_filter=self.purchase_qty_entry.get(),
        )
        serialized = [
            [
                row.get("number", ""),
                row.get("description", ""),
                row.get("type", ""),
                row.get("purchase_date", ""),
                row.get("purchase_price", 0.0),
                row.get("vin", ""),
                row.get("qty", 0),
            ]
            for row in rows
        ]
        self._write_table(file_path, PURCHASE_HEADERS, serialized)
        messagebox.showinfo("Export Complete", f"Purchase rows exported to {file_path}")

    def build_reports_tab(self) -> None:
        """Build the Reports tab filters, action buttons, result tree, and summary panel."""
        form = ttk.LabelFrame(self.tab_reports, text="Report Filters")
        form.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(form, text="Bank Name").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.report_bank_entry = ttk.Combobox(form, state="normal", width=30)
        self.report_bank_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(form, text="Loan Number").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.report_loan_entry = ttk.Combobox(form, state="normal", width=30)
        self.report_loan_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        button_row = ttk.Frame(form)
        button_row.grid(row=1, column=0, columnspan=8, sticky="ew", padx=6, pady=6)
        ttk.Button(button_row, text="Run Filtered Report", command=lambda: self._safe_execute(self.run_filtered_report)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="All Active Loan", command=lambda: self._safe_execute(self.run_active_loan_report)).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(button_row, text="Loan History", command=lambda: self._safe_execute(self.run_loan_history_report)).grid(row=0, column=2, padx=4, pady=4)
        ttk.Button(button_row, text="Export", command=lambda: self._safe_execute(self.export_reports)).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(button_row, text="Clear", command=lambda: self._safe_execute(self.clear_report_filters)).grid(row=0, column=4, padx=4, pady=4)
        button_row.columnconfigure(5, weight=1)

        table_frame = ttk.LabelFrame(self.tab_reports, text="Report Results")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        columns = ("bank_name", "loan_number", "start_date", "end_date", "duration", "status", "loan_amount", "remaining_balance")
        self.report_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        widths = (180, 130, 120, 120, 120, 110, 130, 150)
        for col, heading, width in zip(columns, REPORT_HEADERS, widths):
            self.report_tree.heading(col, text=heading)
            self.report_tree.column(col, width=width, anchor="w")
        report_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=report_scroll.set)
        self.report_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        report_scroll.grid(row=0, column=1, sticky="ns", pady=6)
        self.report_tree.bind("<Double-1>", self.on_report_double_click)

        summary_frame = ttk.LabelFrame(self.tab_reports, text="Summary View")
        summary_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        self.summary_text = tk.Text(summary_frame, height=8, wrap="word", font=("Times New Roman", 12))
        self.summary_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.summary_text.configure(state="disabled")

        self.tab_reports.rowconfigure(1, weight=1)
        self.tab_reports.rowconfigure(2, weight=1)
        self.tab_reports.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        summary_frame.columnconfigure(0, weight=1)

    def build_machine_learning_tab(self) -> None:
        """Build the Machine Learning tab for payment optimization analysis."""
        controls = ttk.LabelFrame(self.tab_ml, text="Payment Optimization")
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(controls, text="Bank Name").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.ml_bank_entry = ttk.Combobox(controls, state="normal", width=24)
        self.ml_bank_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(controls, text="Loan Number").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.ml_loan_entry = ttk.Combobox(controls, state="normal", width=24)
        self.ml_loan_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(controls, text="Extra / Month ($)").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.ml_extra_entry = ttk.Entry(controls, width=24)
        self.ml_extra_entry.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        self.ml_extra_entry.insert(0, "200")

        ttk.Label(controls, text="Max Extra to Evaluate ($)").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.ml_max_extra_entry = ttk.Entry(controls, width=24)
        self.ml_max_extra_entry.grid(row=1, column=3, sticky="w", padx=6, pady=6)
        self.ml_max_extra_entry.insert(0, "500")

        button_row = ttk.Frame(controls)
        button_row.grid(row=2, column=0, columnspan=4, sticky="ew", padx=6, pady=6)
        ttk.Button(button_row, text="Suggest Options", command=lambda: self._safe_execute(self.run_ml_payment_analysis)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="Clear", command=lambda: self._safe_execute(self.clear_ml_analysis)).grid(row=0, column=1, padx=4, pady=4)
        button_row.columnconfigure(2, weight=1)

        summary_frame = ttk.LabelFrame(self.tab_ml, text="Model Insights")
        summary_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.ml_summary_text = tk.Text(summary_frame, height=8, wrap="word", font=("Times New Roman", 12))
        self.ml_summary_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.ml_summary_text.configure(state="disabled")

        table_frame = ttk.LabelFrame(self.tab_ml, text="Best and Alternate Payment Options")
        table_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        columns = ("extra_payment", "months_saved", "interest_saved", "new_payoff_date")
        self.ml_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
        widths = (140, 140, 160, 160)
        for col, heading, width in zip(columns, ML_OPTION_HEADERS, widths):
            self.ml_tree.heading(col, text=heading)
            self.ml_tree.column(col, width=width, anchor="w")
        ml_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.ml_tree.yview)
        self.ml_tree.configure(yscrollcommand=ml_scroll.set)
        self.ml_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ml_scroll.grid(row=0, column=1, sticky="ns", pady=6)

        self.tab_ml.rowconfigure(1, weight=1)
        self.tab_ml.rowconfigure(2, weight=1)
        self.tab_ml.columnconfigure(0, weight=1)
        summary_frame.rowconfigure(0, weight=1)
        summary_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def build_charts_tab(self) -> None:
        """Build the Charts tab filters, KPI view, and chart details table."""
        controls = ttk.LabelFrame(self.tab_charts, text="Chart Controls")
        controls.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(controls, text="Bank Name").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.chart_bank_entry = ttk.Combobox(controls, state="normal", width=22)
        self.chart_bank_entry.grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(controls, text="Loan Number").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.chart_loan_entry = ttk.Combobox(controls, state="normal", width=22)
        self.chart_loan_entry.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(controls, text="Chart Option").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self.chart_option_combo = ttk.Combobox(controls, values=CHART_OPTIONS, state="readonly", width=18)
        self.chart_option_combo.grid(row=0, column=5, sticky="w", padx=6, pady=6)
        self.chart_option_combo.set("Overview")

        ttk.Label(controls, text="Start Date").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.chart_start_entry = ttk.Entry(controls, width=22)
        self.chart_start_entry.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        self.chart_start_entry.insert(0, datetime.today().strftime(DATE_FORMAT))

        ttk.Label(controls, text="End Date").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.chart_end_entry = ttk.Entry(controls, width=22)
        self.chart_end_entry.grid(row=1, column=3, sticky="w", padx=6, pady=6)
        self.chart_end_entry.insert(0, (datetime.today() + timedelta(days=365)).strftime(DATE_FORMAT))

        ttk.Label(controls, text="Status").grid(row=1, column=4, sticky="e", padx=6, pady=6)
        self.chart_status_combo = ttk.Combobox(controls, values=("All",) + TRACK_STATUSES, state="readonly", width=18)
        self.chart_status_combo.grid(row=1, column=5, sticky="w", padx=6, pady=6)
        self.chart_status_combo.set("All")

        button_row = ttk.Frame(controls)
        button_row.grid(row=2, column=0, columnspan=9, sticky="ew", padx=6, pady=6)
        ttk.Button(button_row, text="Search", command=lambda: self._safe_execute(self.search_chart_details)).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(button_row, text="Display Chart", command=lambda: self._safe_execute(self.display_chart)).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(button_row, text="Clear", command=lambda: self._safe_execute(self.clear_chart_filters)).grid(row=0, column=2, padx=4, pady=4)
        button_row.columnconfigure(3, weight=1)

        chart_frame = ttk.LabelFrame(self.tab_charts, text="Chart View")
        chart_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.chart_kpi_text = tk.Text(chart_frame, height=8, wrap="word", font=("Times New Roman", 12))
        self.chart_kpi_text.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.chart_kpi_text.configure(state="disabled")

        details_frame = ttk.LabelFrame(self.tab_charts, text="Chart Details")
        details_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        columns = ("bank_name", "loan_number", "loan_amount", "remaining_balance", "remaining_duration")
        self.chart_tree = ttk.Treeview(details_frame, columns=columns, show="headings", height=8)
        headings = ("Bank Name", "Loan Number", "Loan Amount", "Remaining Balance", "Remaining Duration")
        widths = (220, 150, 140, 160, 170)
        for col, heading, width in zip(columns, headings, widths):
            self.chart_tree.heading(col, text=heading)
            self.chart_tree.column(col, width=width, anchor="w")
        chart_scroll = ttk.Scrollbar(details_frame, orient="vertical", command=self.chart_tree.yview)
        self.chart_tree.configure(yscrollcommand=chart_scroll.set)
        self.chart_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        chart_scroll.grid(row=0, column=1, sticky="ns", pady=6)

        self.tab_charts.rowconfigure(1, weight=1)
        self.tab_charts.rowconfigure(2, weight=1)
        self.tab_charts.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(0, weight=1)
        chart_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(0, weight=1)
        details_frame.columnconfigure(0, weight=1)

        self._setup_autosuggestions()

    def refresh_all(self) -> None:
        """Refresh all tab views and suggestion sources from current data."""
        self.refresh_autosuggestion_values()
        self.refresh_purchase_type_values()
        self.refresh_loan_tree(self.data["loans"])
        self.refresh_track_tree()
        self.refresh_purchases_tree(self.data["purchases"])
        self.run_loan_history_report()
        self.refresh_ml_tab()
        self.refresh_chart_details()

    def refresh_ml_tab(self) -> None:
        """Clear ML suggestion outputs while preserving entered filter values."""
        self.ml_tree.delete(*self.ml_tree.get_children())
        self.ml_summary_text.configure(state="normal")
        self.ml_summary_text.delete("1.0", tk.END)
        self.ml_summary_text.insert(tk.END, "Select a loan and run Suggest Options to generate payment recommendations.")
        self.ml_summary_text.configure(state="disabled")

    def _setup_autosuggestions(self) -> None:
        """Register autosuggestion widgets and bind key handlers for filtering suggestions."""
        self.bank_suggestion_widgets = [
            self.loan_bank_entry,
            self.track_bank_entry,
            self.report_bank_entry,
            self.ml_bank_entry,
            self.chart_bank_entry,
        ]
        self.loan_suggestion_widgets = [
            self.loan_number_entry,
            self.track_loan_entry,
            self.report_loan_entry,
            self.ml_loan_entry,
            self.chart_loan_entry,
        ]
        for widget in self.bank_suggestion_widgets:
            widget.bind("<KeyRelease>", lambda _event, target=widget: self._filter_suggestions(target, "bank"))
        for widget in self.loan_suggestion_widgets:
            widget.bind("<KeyRelease>", lambda _event, target=widget: self._filter_suggestions(target, "loan"))
        self.refresh_autosuggestion_values()

    def _all_bank_names(self) -> list[str]:
        """Return a sorted unique list of bank names from loan and payment data."""
        values = {
            str(item.get("bank_name", "")).strip()
            for item in (self.data.get("loans", []) + self.data.get("payments", []))
            if str(item.get("bank_name", "")).strip()
        }
        return sorted(values, key=str.lower)

    def _all_loan_numbers(self) -> list[str]:
        """Return a sorted unique list of loan numbers from loan and payment data."""
        values = {
            str(item.get("loan_number", "")).strip()
            for item in (self.data.get("loans", []) + self.data.get("payments", []))
            if str(item.get("loan_number", "")).strip()
        }
        return sorted(values, key=str.lower)

    def refresh_autosuggestion_values(self) -> None:
        """Refresh all combobox suggestion lists with current bank and loan values."""
        bank_names = self._all_bank_names()
        loan_numbers = self._all_loan_numbers()
        if hasattr(self, "bank_suggestion_widgets"):
            for widget in self.bank_suggestion_widgets:
                widget.configure(values=bank_names)
        if hasattr(self, "loan_suggestion_widgets"):
            for widget in self.loan_suggestion_widgets:
                widget.configure(values=loan_numbers)

    def _filter_suggestions(self, widget: ttk.Combobox, kind: str) -> None:
        """Filter combobox suggestions by typed text for bank or loan values."""
        source = self._all_bank_names() if kind == "bank" else self._all_loan_numbers()
        typed = widget.get().strip().lower()
        if not typed:
            widget.configure(values=source)
            return
        filtered = [value for value in source if typed in value.lower()]
        widget.configure(values=filtered)

    def get_loan_key(self, loan: dict[str, object]) -> tuple[str, str]:
        """Return a normalized (bank_name, loan_number) key for a loan-like record."""
        return str(loan.get("bank_name", "")).strip().lower(), str(loan.get("loan_number", "")).strip().lower()

    def parse_loan_form(self) -> dict[str, object]:
        """Parse and validate loan form inputs, returning normalized fields and calculated schedule."""
        bank_name = self.loan_bank_entry.get().strip()
        loan_number = self.loan_number_entry.get().strip()
        if not bank_name:
            raise ValueError("Bank Name is required.")
        if not loan_number:
            raise ValueError("Loan Number is required.")
        principal = float(self.loan_amount_entry.get().strip())
        rate = float(self.loan_rate_entry.get().strip())
        years = int(self.loan_years_entry.get().strip())
        start_date = parse_date(self.loan_start_entry.get().strip(), "Start Date")
        schedule = calculate_schedule(principal, rate, years, start_date)
        end_date = str(schedule.iloc[-1]["Date"])
        return {
            "bank_name": bank_name,
            "loan_number": loan_number,
            "loan_amount": round(principal, 2),
            "annual_interest_rate": round(rate, 4),
            "term_years": years,
            "start_date": start_date.strftime(DATE_FORMAT),
            "end_date": end_date,
            "duration_remaining": format_duration_remaining(end_date),
            "schedule": schedule,
        }

    def add_loan(self) -> None:
        """Add a new loan and replace any existing payment rows for the same loan key."""
        parsed = self.parse_loan_form()
        new_key = (str(parsed["bank_name"]).lower(), str(parsed["loan_number"]).lower())
        for existing in self.data["loans"]:
            if self.get_loan_key(existing) == new_key:
                raise ValueError("Loan already exists. Use Update Loan instead.")
        self.data["loans"].append({k: v for k, v in parsed.items() if k != "schedule"})
        self.data["payments"] = [
            row
            for row in self.data["payments"]
            if not (
                str(row.get("bank_name", "")).lower() == new_key[0]
                and str(row.get("loan_number", "")).lower() == new_key[1]
            )
        ]
        self.data["payments"].extend(build_payment_rows(parsed["bank_name"], parsed["loan_number"], parsed["schedule"]))
        self.save_data()
        self.refresh_all()
        messagebox.showinfo("Saved", "Loan added successfully.")

    def update_loan(self) -> None:
        """Update an existing loan and regenerate its payment rows after user confirmation."""
        parsed = self.parse_loan_form()
        key = (str(parsed["bank_name"]).lower(), str(parsed["loan_number"]).lower())
        confirmed = messagebox.askokcancel(
            "Confirm Update",
            f"Update loan '{parsed['loan_number']}' for '{parsed['bank_name']}'?",
            icon="warning",
        )
        if not confirmed:
            return
        for idx, loan in enumerate(self.data["loans"]):
            if self.get_loan_key(loan) == key:
                self.data["loans"][idx] = {k: v for k, v in parsed.items() if k != "schedule"}
                break
        else:
            raise ValueError("Loan not found. Add the loan first.")
        self.data["payments"] = [
            row
            for row in self.data["payments"]
            if not (
                str(row.get("bank_name", "")).lower() == key[0]
                and str(row.get("loan_number", "")).lower() == key[1]
            )
        ]
        self.data["payments"].extend(build_payment_rows(parsed["bank_name"], parsed["loan_number"], parsed["schedule"]))
        self.save_data()
        self.refresh_all()
        messagebox.showinfo("Updated", "Loan updated successfully.")

    def remove_loan(self) -> None:
        """Remove the selected loan and associated payments after user confirmation."""
        selected = self.loan_tree.selection()
        if not selected:
            raise ValueError("Select a loan to remove.")
        values = self.loan_tree.item(selected[0], "values")
        if not values:
            raise ValueError("Select a valid loan row.")
        bank_name, loan_number = str(values[0]), str(values[1])
        confirmed = messagebox.askokcancel(
            "Confirm Remove",
            f"Remove loan '{loan_number}' for '{bank_name}'?",
            icon="warning",
        )
        if not confirmed:
            return
        key = (bank_name.lower(), loan_number.lower())
        self.data["loans"] = [loan for loan in self.data["loans"] if self.get_loan_key(loan) != key]
        self.data["payments"] = [
            row
            for row in self.data["payments"]
            if not (
                str(row.get("bank_name", "")).lower() == key[0]
                and str(row.get("loan_number", "")).lower() == key[1]
            )
        ]
        self.save_data()
        self.refresh_all()
        self.clear_loan_entries()
        messagebox.showinfo("Removed", "Loan removed successfully.")

    def search_loans(self) -> None:
        """Filter and display loan rows matching bank and loan form filters."""
        bank_filter = self.loan_bank_entry.get().strip().lower()
        loan_filter = self.loan_number_entry.get().strip().lower()
        results = []
        for loan in self.data["loans"]:
            bank_name = str(loan.get("bank_name", "")).lower()
            loan_number = str(loan.get("loan_number", "")).lower()
            if bank_filter and bank_filter not in bank_name:
                continue
            if loan_filter and loan_filter not in loan_number:
                continue
            results.append(loan)
        self.refresh_loan_tree(results)

    def clear_loan_entries(self, *, include_default_start: bool = True) -> None:
        """Clear loan form fields and optionally reset start date to today."""
        self.loan_bank_entry.delete(0, tk.END)
        self.loan_number_entry.delete(0, tk.END)
        self.loan_amount_entry.delete(0, tk.END)
        self.loan_rate_entry.delete(0, tk.END)
        self.loan_years_entry.delete(0, tk.END)
        self.loan_start_entry.delete(0, tk.END)
        if include_default_start:
            self.loan_start_entry.insert(0, datetime.today().strftime(DATE_FORMAT))

    def clear_loan_results(self) -> None:
        """Reset loan filters and restore the full loan table."""
        self.clear_loan_entries()
        self.refresh_loan_tree(self.data["loans"])

    def refresh_loan_tree(self, loans: list[dict[str, object]]) -> None:
        """Render the provided loans into the Loan Management treeview."""
        self.loan_tree.delete(*self.loan_tree.get_children())
        for loan in sorted(loans, key=lambda item: (str(item.get("bank_name", "")).lower(), str(item.get("loan_number", "")).lower())):
            end_date = str(loan.get("end_date", ""))
            self.loan_tree.insert(
                "",
                "end",
                values=(
                    loan.get("bank_name", ""),
                    loan.get("loan_number", ""),
                    f"{float(loan.get('loan_amount', 0.0)):.2f}",
                    f"{float(loan.get('annual_interest_rate', 0.0)):.4f}",
                    loan.get("term_years", ""),
                    loan.get("start_date", ""),
                    end_date,
                    format_duration_remaining(end_date) if end_date else "",
                ),
            )

    def on_loan_selected(self, _event=None) -> None:
        """Populate loan form fields from the currently selected loan row."""
        selected = self.loan_tree.selection()
        if not selected:
            return
        values = self.loan_tree.item(selected[0], "values")
        if not values:
            return
        self.clear_loan_entries(include_default_start=False)
        self.loan_bank_entry.insert(0, values[0])
        self.loan_number_entry.insert(0, values[1])
        self.loan_amount_entry.insert(0, values[2])
        self.loan_rate_entry.insert(0, values[3])
        self.loan_years_entry.insert(0, values[4])
        self.loan_start_entry.insert(0, values[5])

    def export_loans(self) -> None:
        """Export all loan rows to a user-selected CSV or XLSX file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        rows = []
        for loan in self.data["loans"]:
            end_date = str(loan.get("end_date", ""))
            rows.append(
                [
                    loan.get("bank_name", ""),
                    loan.get("loan_number", ""),
                    loan.get("loan_amount", 0.0),
                    loan.get("annual_interest_rate", 0.0),
                    loan.get("term_years", 0),
                    loan.get("start_date", ""),
                    end_date,
                    format_duration_remaining(end_date) if end_date else "",
                ]
            )
        self._write_table(file_path, LOAN_HEADERS, rows)
        messagebox.showinfo("Export Complete", f"Loan data exported to {file_path}")

    def _filtered_future_payments(self, bank_filter: str = "", loan_filter: str = "", status_filter: str = "") -> list[dict[str, object]]:
        """Return payments filtered by bank/loan/status with future-payment rules applied."""
        today_date = datetime.today().date()
        bank_filter = bank_filter.strip().lower()
        loan_filter = loan_filter.strip().lower()
        status_filter = status_filter.strip()
        include_paid = status_filter == "Paid"
        rows = []
        for payment in self.data["payments"]:
            payment_date = parse_date(str(payment.get("date", "")), "Payment Date").date()
            payment_status = str(payment.get("status", "Upcoming"))
            if include_paid:
                if payment_status != "Paid":
                    continue
            else:
                if payment_status == "Paid":
                    continue
                if payment_date < today_date:
                    continue
            bank_name = str(payment.get("bank_name", ""))
            loan_number = str(payment.get("loan_number", ""))
            if bank_filter and bank_filter not in bank_name.lower():
                continue
            if loan_filter and loan_filter not in loan_number.lower():
                continue
            if status_filter and payment_status != status_filter:
                continue
            rows.append(payment)
        rows.sort(key=lambda item: (str(item.get("bank_name", "")).lower(), str(item.get("loan_number", "")).lower(), int(item.get("payment_number", 0))))
        return rows

    def refresh_track_tree(self) -> None:
        """Refresh the Track Loan grouped tree using current status filter selection."""
        self.track_tree.delete(*self.track_tree.get_children())
        rows = self._filtered_future_payments(status_filter=self.track_status_combo.get())
        self._insert_track_rows(rows)

    def _insert_track_rows(self, rows: list[dict[str, object]]) -> None:
        """Insert grouped payment rows and parent summary values into the Track Loan tree."""
        grouped: dict[tuple[str, str], str] = {}
        grouped_rows: dict[tuple[str, str], list[dict[str, object]]] = {}
        for payment in rows:
            key = (str(payment.get("bank_name", "")), str(payment.get("loan_number", "")))
            if key not in grouped:
                parent = self.track_tree.insert("", "end", text=f"{key[0]} | {key[1]}", values=("", "", "", "", "", ""))
                grouped[key] = parent
                grouped_rows[key] = []
            grouped_rows[key].append(payment)
            item_id = f"{key[0]}::{key[1]}::{int(payment.get('payment_number', 0))}"
            self.track_tree.insert(
                grouped[key],
                "end",
                iid=item_id,
                text="",
                values=(
                    int(payment.get("payment_number", 0)),
                    payment.get("date", ""),
                    f"{float(payment.get('payment', 0.0)):.2f}",
                    f"{float(payment.get('principal_paid', 0.0)):.2f}",
                    f"{float(payment.get('interest_paid', 0.0)):.2f}",
                    f"{float(payment.get('remaining_balance', 0.0)):.2f}",
                ),
            )
        for key, parent in grouped.items():
            group_items = grouped_rows[key]
            if not group_items:
                continue
            avg_payment = sum(float(item.get("payment", 0.0)) for item in group_items) / len(group_items)
            avg_principal = sum(float(item.get("principal_paid", 0.0)) for item in group_items) / len(group_items)
            avg_interest = sum(float(item.get("interest_paid", 0.0)) for item in group_items) / len(group_items)
            loan_info = next(
                (
                    loan
                    for loan in self.data["loans"]
                    if str(loan.get("bank_name", "")).lower() == key[0].lower()
                    and str(loan.get("loan_number", "")).lower() == key[1].lower()
                ),
                None,
            )
            loan_amount = float(loan_info.get("loan_amount", 0.0)) if loan_info else 0.0
            _status, total_remaining = self._loan_status_and_remaining(key[0], key[1], loan_amount)
            self.track_tree.item(
                parent,
                values=(
                    "",
                    "",
                    f"{avg_payment:.2f}",
                    f"{avg_principal:.2f}",
                    f"{avg_interest:.2f}",
                    f"{total_remaining:.2f}",
                ),
            )
        for parent in grouped.values():
            self.track_tree.item(parent, open=True)

    def search_track_rows(self) -> None:
        """Apply Track Loan filters and render matching grouped payment rows."""
        rows = self._filtered_future_payments(
            bank_filter=self.track_bank_entry.get(),
            loan_filter=self.track_loan_entry.get(),
            status_filter=self.track_status_combo.get(),
        )
        self.track_tree.delete(*self.track_tree.get_children())
        self._insert_track_rows(rows)

    def clear_track_filters(self) -> None:
        """Clear Track Loan filters and restore default upcoming payment view."""
        self.track_bank_entry.delete(0, tk.END)
        self.track_loan_entry.delete(0, tk.END)
        self.track_status_combo.set("Upcoming")
        self.refresh_track_tree()

    def _selected_payment_record(self) -> dict[str, object]:
        """Return the payment record represented by the currently selected tree item."""
        selected = self.track_tree.selection()
        if not selected:
            raise ValueError("Select a payment row.")
        item_id = selected[0]
        if "::" not in item_id:
            raise ValueError("Select a payment entry, not a group header.")
        bank_name, loan_number, payment_number = item_id.split("::", maxsplit=2)
        payment_number_value = int(payment_number)
        for payment in self.data["payments"]:
            if (
                str(payment.get("bank_name", "")) == bank_name
                and str(payment.get("loan_number", "")) == loan_number
                and int(payment.get("payment_number", 0)) == payment_number_value
            ):
                return payment
        raise ValueError("Selected payment record no longer exists.")

    def mark_selected_paid(self) -> None:
        """Mark the selected payment as paid and refresh dependent views."""
        payment = self._selected_payment_record()
        payment["status"] = "Paid"
        self.save_data()
        self.refresh_autosuggestion_values()
        self.refresh_track_tree()
        self.run_filtered_report()
        self.refresh_chart_details()
        messagebox.showinfo("Updated", "Selected payment marked as Paid.")

    def edit_selected_payment(self) -> None:
        """Edit selected payment date and status using dialog prompts."""
        payment = self._selected_payment_record()
        new_date = simpledialog.askstring("Edit Payment", "Date (YYYY-MM-DD):", initialvalue=str(payment.get("date", "")))
        if new_date is None:
            return
        parse_date(new_date, "Payment Date")
        new_status = simpledialog.askstring("Edit Payment", "Status (Paid/Upcoming):", initialvalue=str(payment.get("status", "Upcoming")))
        if new_status is None:
            return
        if new_status not in TRACK_STATUSES:
            raise ValueError("Status must be 'Paid' or 'Upcoming'.")
        payment["date"] = new_date
        payment["status"] = new_status
        self.save_data()
        self.refresh_autosuggestion_values()
        self.refresh_track_tree()
        self.run_filtered_report()
        self.refresh_chart_details()

    def import_payments(self) -> None:
        """Import payment rows from CSV/XLSX and merge them into existing payment data."""
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if not file_path:
            return
        rows = self._read_table(file_path)
        if not rows:
            raise ValueError("Selected file has no rows.")
        required = set(PAYMENT_HEADERS)
        for row in rows:
            if not required.issubset(set(row.keys())):
                raise ValueError("Import file is missing required payment columns.")
        for row in rows:
            status = str(row.get("Status", "Upcoming")).strip() or "Upcoming"
            if status not in TRACK_STATUSES:
                raise ValueError("Payment status must be Paid or Upcoming.")
            record = {
                "bank_name": str(row.get("Bank Name", "")).strip(),
                "loan_number": str(row.get("Loan Number", "")).strip(),
                "payment_number": int(float(row.get("Payment #", 0))),
                "date": str(row.get("Date", "")).strip(),
                "payment": float(row.get("Payment", 0.0)),
                "principal_paid": float(row.get("Principal Paid", 0.0)),
                "interest_paid": float(row.get("Interest Paid", 0.0)),
                "remaining_balance": float(row.get("Remaining Balance", 0.0)),
                "status": status,
            }
            parse_date(record["date"], "Payment Date")
            existing = None
            for payment in self.data["payments"]:
                if (
                    str(payment.get("bank_name", "")) == record["bank_name"]
                    and str(payment.get("loan_number", "")) == record["loan_number"]
                    and int(payment.get("payment_number", 0)) == record["payment_number"]
                ):
                    existing = payment
                    break
            if existing is None:
                self.data["payments"].append(record)
            else:
                existing.update(record)
        self.save_data()
        self.refresh_autosuggestion_values()
        self.refresh_track_tree()
        self.run_filtered_report()
        self.refresh_chart_details()
        messagebox.showinfo("Import Complete", f"Imported payment rows from {file_path}")

    def export_track_rows(self) -> None:
        """Export currently filtered Track Loan rows to CSV or XLSX."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        rows = self._filtered_future_payments(
            bank_filter=self.track_bank_entry.get(),
            loan_filter=self.track_loan_entry.get(),
            status_filter=self.track_status_combo.get(),
        )
        serialized = [
            [
                row.get("bank_name", ""),
                row.get("loan_number", ""),
                row.get("payment_number", 0),
                row.get("date", ""),
                row.get("payment", 0.0),
                row.get("principal_paid", 0.0),
                row.get("interest_paid", 0.0),
                row.get("remaining_balance", 0.0),
                row.get("status", "Upcoming"),
            ]
            for row in rows
        ]
        self._write_table(file_path, PAYMENT_HEADERS, serialized)
        messagebox.showinfo("Export Complete", f"Track rows exported to {file_path}")

    def export_payment_template(self) -> None:
        """Export an empty payment template file with required headers."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        self._write_table(file_path, PAYMENT_HEADERS, [])
        messagebox.showinfo("Template Saved", f"Template saved to {file_path}")

    def _loan_status_and_remaining(self, bank_name: str, loan_number: str, loan_amount: float) -> tuple[str, float]:
        """Compute aggregate loan status and current remaining balance for a loan key."""
        today_date = datetime.today().date()
        loan_payments = [
            payment
            for payment in self.data["payments"]
            if str(payment.get("bank_name", "")).lower() == bank_name.lower()
            and str(payment.get("loan_number", "")).lower() == loan_number.lower()
        ]
        if not loan_payments:
            return "Upcoming", loan_amount
        loan_payments.sort(key=lambda item: int(item.get("payment_number", 0)))
        upcoming = [
            item
            for item in loan_payments
            if str(item.get("status", "Upcoming")) == "Upcoming"
            and parse_date(str(item.get("date", "")), "Payment Date").date() >= today_date
        ]
        if not upcoming:
            return "Paid", 0.0
        next_payment = upcoming[0]
        next_number = int(next_payment.get("payment_number", 0))
        if next_number <= 1:
            return "Upcoming", loan_amount
        previous = None
        for row in loan_payments:
            if int(row.get("payment_number", 0)) == next_number - 1:
                previous = row
                break
        if previous is None:
            return "Upcoming", loan_amount
        return "Upcoming", float(previous.get("remaining_balance", loan_amount))

    def build_report_rows(self, only_active: bool = False, apply_filter: bool = True) -> list[dict[str, object]]:
        """Build normalized report rows from loans with optional active-only and form filtering."""
        bank_filter = self.report_bank_entry.get().strip().lower() if apply_filter else ""
        loan_filter = self.report_loan_entry.get().strip().lower() if apply_filter else ""
        rows = []
        for loan in self.data["loans"]:
            bank_name = str(loan.get("bank_name", ""))
            loan_number = str(loan.get("loan_number", ""))
            if bank_filter and bank_filter not in bank_name.lower():
                continue
            if loan_filter and loan_filter not in loan_number.lower():
                continue
            status, remaining = self._loan_status_and_remaining(bank_name, loan_number, float(loan.get("loan_amount", 0.0)))
            if only_active and status != "Upcoming":
                continue
            start_date = str(loan.get("start_date", ""))
            end_date = str(loan.get("end_date", ""))
            duration_days = max((parse_date(end_date, "End Date") - parse_date(start_date, "Start Date")).days, 0)
            duration = f"{round(duration_days / 365, 1)} years" if duration_days > 365 else f"{duration_days} days"
            rows.append(
                {
                    "bank_name": bank_name,
                    "loan_number": loan_number,
                    "start_date": start_date,
                    "end_date": end_date,
                    "duration": duration,
                    "status": status,
                    "loan_amount": float(loan.get("loan_amount", 0.0)),
                    "remaining_balance": round(remaining, 2),
                }
            )
        rows.sort(key=lambda item: (item["bank_name"].lower(), item["loan_number"].lower()))
        return rows

    def _render_report_rows(self, rows: list[dict[str, object]], mode: str) -> None:
        """Render report rows into the tree and update the summary view text."""
        self.current_report_rows = rows
        self.report_tree.delete(*self.report_tree.get_children())
        for row in rows:
            self.report_tree.insert(
                "",
                "end",
                values=(
                    row["bank_name"],
                    row["loan_number"],
                    row["start_date"],
                    row["end_date"],
                    row["duration"],
                    row["status"],
                    f"{row['loan_amount']:.2f}",
                    f"{row['remaining_balance']:.2f}",
                ),
            )
        total_open = sum(1 for row in rows if row["status"] == "Upcoming")
        total_loan_amount = sum(float(row["loan_amount"]) for row in rows)
        total_remaining = sum(float(row["remaining_balance"]) for row in rows)
        row_keys = {(str(row["bank_name"]).lower(), str(row["loan_number"]).lower()) for row in rows}
        paid_count = 0
        upcoming_count = 0
        for payment in self.data["payments"]:
            key = (str(payment.get("bank_name", "")).lower(), str(payment.get("loan_number", "")).lower())
            if key not in row_keys:
                continue
            if str(payment.get("status", "Upcoming")) == "Paid":
                paid_count += 1
            else:
                upcoming_count += 1
        summary = (
            f"Data Mode: {self.data_visibility}\n"
            f"Total Open Loans: {total_open}\n"
            f"Total Loan Amount: ${total_loan_amount:,.2f}\n"
            f"Remaining Loan Balance: ${total_remaining:,.2f}\n"
            f"Payment Status Counts:\n"
            f"\tPaid: {paid_count},\n"
            f"\tUpcoming: {upcoming_count}\n"
            f"Rows in view: {len(rows)} ({mode})"
        )
        self.summary_text.configure(state="normal")
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert(tk.END, summary)
        self.summary_text.configure(state="disabled")

    def run_filtered_report(self) -> None:
        """Run and render the reports tab using current bank/loan filters."""
        rows = self.build_report_rows(only_active=False, apply_filter=True)
        self._render_report_rows(rows, "Filtered Report")

    def clear_report_filters(self) -> None:
        """Clear report filter entries and restore full loan history report view."""
        self.report_bank_entry.delete(0, tk.END)
        self.report_loan_entry.delete(0, tk.END)
        self.run_loan_history_report()

    def on_report_double_click(self, _event=None) -> None:
        """Populate report filters from a double-clicked report row."""
        selected = self.report_tree.selection()
        if not selected:
            return
        values = self.report_tree.item(selected[0], "values")
        if not values:
            return
        self.report_bank_entry.delete(0, tk.END)
        self.report_bank_entry.insert(0, str(values[0]))
        self.report_loan_entry.delete(0, tk.END)
        self.report_loan_entry.insert(0, str(values[1]))

    def run_active_loan_report(self) -> None:
        """Run and render report rows limited to active (upcoming) loans."""
        rows = self.build_report_rows(only_active=True, apply_filter=False)
        self._render_report_rows(rows, "All Active Loan")

    def run_loan_history_report(self) -> None:
        """Run and render report rows for complete loan history."""
        rows = self.build_report_rows(only_active=False, apply_filter=False)
        self._render_report_rows(rows, "Loan History")

    def export_reports(self) -> None:
        """Export the currently rendered report rows to CSV or XLSX."""
        if not self.current_report_rows:
            raise ValueError("Run a report before exporting.")
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
        )
        if not file_path:
            return
        rows = [
            [
                row["bank_name"],
                row["loan_number"],
                row["start_date"],
                row["end_date"],
                row["duration"],
                row["status"],
                row["loan_amount"],
                row["remaining_balance"],
            ]
            for row in self.current_report_rows
        ]
        self._write_table(file_path, REPORT_HEADERS, rows)
        messagebox.showinfo("Export Complete", f"Report exported to {file_path}")

    def _chart_filtered_rows(self) -> list[dict[str, object]]:
        """Return report rows filtered for charting by date range, bank/loan, and status."""
        start_date = parse_date(self.chart_start_entry.get().strip(), "Chart Start Date").date()
        end_date = parse_date(self.chart_end_entry.get().strip(), "Chart End Date").date()
        if end_date < start_date:
            raise ValueError("Chart end date must be after or equal to chart start date.")
        bank_filter = self.chart_bank_entry.get().strip().lower()
        loan_filter = self.chart_loan_entry.get().strip().lower()
        status_filter = self.chart_status_combo.get().strip()
        rows = self.build_report_rows(only_active=False, apply_filter=False)
        filtered = []
        for row in rows:
            row_start = parse_date(str(row["start_date"]), "Start Date").date()
            row_end = parse_date(str(row["end_date"]), "End Date").date()
            if row_end < start_date or row_start > end_date:
                continue
            if bank_filter and bank_filter != str(row["bank_name"]).strip().lower():
                continue
            if loan_filter and loan_filter != str(row["loan_number"]).strip().lower():
                continue
            if status_filter != "All" and row["status"] != status_filter:
                continue
            filtered.append(row)
        return filtered

    def search_chart_details(self) -> None:
        """Apply bank/loan chart search, auto-set date bounds, and refresh chart views."""
        bank_filter = self.chart_bank_entry.get().strip().lower()
        loan_filter = self.chart_loan_entry.get().strip().lower()
        matching_loans = [
            loan
            for loan in self.data["loans"]
            if (not bank_filter or bank_filter == str(loan.get("bank_name", "")).strip().lower())
            and (not loan_filter or loan_filter == str(loan.get("loan_number", "")).strip().lower())
        ]
        if matching_loans:
            start_dates = [parse_date(str(loan.get("start_date", "")), "Start Date") for loan in matching_loans]
            end_dates = [parse_date(str(loan.get("end_date", "")), "End Date") for loan in matching_loans]
            self.chart_start_entry.delete(0, tk.END)
            self.chart_start_entry.insert(0, min(start_dates).strftime(DATE_FORMAT))
            self.chart_end_entry.delete(0, tk.END)
            self.chart_end_entry.insert(0, max(end_dates).strftime(DATE_FORMAT))
        self.refresh_chart_details()

    def refresh_chart_details(self) -> None:
        """Refresh chart details table and KPI text using current chart filters."""
        rows = self._chart_filtered_rows()
        self.chart_tree.delete(*self.chart_tree.get_children())
        for row in rows:
            self.chart_tree.insert(
                "",
                "end",
                values=(
                    row["bank_name"],
                    row["loan_number"],
                    f"{float(row['loan_amount']):.2f}",
                    f"{float(row['remaining_balance']):.2f}",
                    format_duration_remaining(str(row["end_date"])),
                ),
            )
        self._update_chart_kpi(rows)

    def _update_chart_kpi(self, rows: list[dict[str, object]]) -> None:
        """Update chart KPI summary text from the supplied chart-filtered rows."""
        start_date = parse_date(self.chart_start_entry.get().strip(), "Chart Start Date").date()
        end_date = parse_date(self.chart_end_entry.get().strip(), "Chart End Date").date()
        status_filter = self.chart_status_combo.get().strip()
        payment_rows = self._chart_payment_rows(rows, start_date, end_date, status_filter)

        total_loan_amount = sum(float(row["loan_amount"]) for row in rows)
        total_remaining = sum(float(row["remaining_balance"]) for row in rows)
        total_payments = sum(float(payment.get("payment", 0.0)) for payment in payment_rows)
        total_interest_paid = sum(float(payment.get("interest_paid", 0.0)) for payment in payment_rows)
        avg_loan_amount = (total_loan_amount / len(rows)) if rows else 0.0
        avg_remaining = (total_remaining / len(rows)) if rows else 0.0
        avg_interest_rate = 0.0
        avg_payment_amount = 0.0
        if rows:
            key_set = {(str(row["bank_name"]).lower(), str(row["loan_number"]).lower()) for row in rows}
            matching = [
                loan
                for loan in self.data["loans"]
                if (str(loan.get("bank_name", "")).lower(), str(loan.get("loan_number", "")).lower()) in key_set
            ]
            if matching:
                avg_interest_rate = sum(float(loan.get("annual_interest_rate", 0.0)) for loan in matching) / len(matching)
            first_payments = []
            for key in key_set:
                key_rows = [
                    payment
                    for payment in self.data["payments"]
                    if str(payment.get("bank_name", "")).lower() == key[0]
                    and str(payment.get("loan_number", "")).lower() == key[1]
                ]
                if not key_rows:
                    continue
                key_rows.sort(key=lambda item: int(item.get("payment_number", 0)))
                first_payments.append(float(key_rows[0].get("payment", 0.0)))
            if first_payments:
                avg_payment_amount = sum(first_payments) / len(first_payments)
        kpi = (
            f"Number of Loans: {len(rows)}\n"
            f"Total Loan Amount: ${total_loan_amount:,.2f}\n"
            f"Total Remaining Balance: ${total_remaining:,.2f}\n"
            f"Total Payments: ${total_payments:,.2f}\n"
            f"Total Interest Paid: ${total_interest_paid:,.2f}\n"
            f"Average Loan Amount: ${avg_loan_amount:,.2f}\n"
            f"Average Payment Amount: ${avg_payment_amount:,.2f}\n"
            f"Average Remaining Balance: ${avg_remaining:,.2f}\n"
            f"Average Interest Rate: {avg_interest_rate:.2f}%"
        )
        self.chart_kpi_text.configure(state="normal")
        self.chart_kpi_text.delete("1.0", tk.END)
        self.chart_kpi_text.insert(tk.END, kpi)
        self.chart_kpi_text.configure(state="disabled")

    def _chart_payment_rows(
        self,
        rows: list[dict[str, object]],
        start_date,
        end_date,
        status_filter: str,
    ) -> list[dict[str, object]]:
        """Return chart payment rows filtered by selected loans, date range, and status filter."""
        selected_keys = {(str(row["bank_name"]).lower(), str(row["loan_number"]).lower()) for row in rows}
        filtered: list[dict[str, object]] = []
        for payment in self.data["payments"]:
            key = (str(payment.get("bank_name", "")).lower(), str(payment.get("loan_number", "")).lower())
            if key not in selected_keys:
                continue
            payment_date = parse_date(str(payment.get("date", "")), "Payment Date").date()
            if payment_date < start_date or payment_date > end_date:
                continue
            payment_status = str(payment.get("status", "Upcoming"))
            if status_filter != "All" and payment_status != status_filter:
                continue
            filtered.append(payment)

        def payment_sort_key(item: dict[str, object]) -> tuple[datetime, int]:
            payment_date = parse_date(str(item.get("date", "")), "Payment Date")
            try:
                payment_number = int(item.get("payment_number", 0))
            except (TypeError, ValueError):
                payment_number = 0
            return payment_date, payment_number

        filtered.sort(key=payment_sort_key)
        return filtered

    def clear_chart_filters(self) -> None:
        """Reset chart filters to defaults and refresh chart detail/KPI views."""
        self.chart_bank_entry.delete(0, tk.END)
        self.chart_loan_entry.delete(0, tk.END)
        self.chart_option_combo.set("Overview")
        self.chart_start_entry.delete(0, tk.END)
        self.chart_start_entry.insert(0, datetime.today().strftime(DATE_FORMAT))
        self.chart_end_entry.delete(0, tk.END)
        self.chart_end_entry.insert(0, (datetime.today() + timedelta(days=365)).strftime(DATE_FORMAT))
        self.chart_status_combo.set("All")
        self.refresh_chart_details()

    def display_chart(self) -> None:
        """Render and display the chart figure in a separate Matplotlib window."""
        if plt is None:
            raise ImportError("matplotlib is required for chart display. Install matplotlib to use charts.")
        rows = self._chart_filtered_rows()
        if not rows:
            raise ValueError("No chart data matches the selected filters.")
        filter_start = parse_date(self.chart_start_entry.get().strip(), "Chart Start Date")
        filter_end = parse_date(self.chart_end_entry.get().strip(), "Chart End Date")
        status_filter = self.chart_status_combo.get().strip()
        chart_option = self.chart_option_combo.get().strip() or "Overview"
        total_loan_amount = sum(float(row["loan_amount"]) for row in rows)
        total_remaining = sum(float(row["remaining_balance"]) for row in rows)
        payment_rows = self._chart_payment_rows(rows, filter_start.date(), filter_end.date(), status_filter)
        total_payments = sum(float(payment.get("payment", 0.0)) for payment in payment_rows)
        total_interest_paid = sum(float(payment.get("interest_paid", 0.0)) for payment in payment_rows)
        status_counts = {"Paid": 0, "Upcoming": 0}
        for payment in payment_rows:
            payment_status = str(payment.get("status", "Upcoming"))
            if payment_status in status_counts:
                status_counts[payment_status] += 1
        kpi_text = (
            f"Loans: {len(rows)} | Payments in range: {len(payment_rows)}\n"
            f"Total Loan Amount: ${total_loan_amount:,.2f}\n"
            f"Remaining Balance: ${total_remaining:,.2f}\n"
            f"Total Payments: ${total_payments:,.2f}\n"
            f"Total Interest Paid: ${total_interest_paid:,.2f}"
        )

        figure = plt.figure(figsize=(14, 9))
        grid = figure.add_gridspec(2, 2, height_ratios=[1, 1.2])

        if chart_option == "By Bank":
            bank_loan_totals: dict[str, float] = {}
            bank_remaining_totals: dict[str, float] = {}
            for row in rows:
                bank_name = str(row["bank_name"])
                bank_loan_totals[bank_name] = bank_loan_totals.get(bank_name, 0.0) + float(row["loan_amount"])
                bank_remaining_totals[bank_name] = bank_remaining_totals.get(bank_name, 0.0) + float(row["remaining_balance"])

            bank_payment_totals: dict[str, float] = {}
            bank_interest_totals: dict[str, float] = {}
            bank_paid_counts: dict[str, int] = {}
            bank_upcoming_counts: dict[str, int] = {}
            for payment in payment_rows:
                bank_name = str(payment.get("bank_name", "Unknown"))
                bank_payment_totals[bank_name] = bank_payment_totals.get(bank_name, 0.0) + float(payment.get("payment", 0.0))
                bank_interest_totals[bank_name] = bank_interest_totals.get(bank_name, 0.0) + float(payment.get("interest_paid", 0.0))
                if str(payment.get("status", "Upcoming")) == "Paid":
                    bank_paid_counts[bank_name] = bank_paid_counts.get(bank_name, 0) + 1
                else:
                    bank_upcoming_counts[bank_name] = bank_upcoming_counts.get(bank_name, 0) + 1

            banks = sorted(bank_loan_totals.keys(), key=str.lower)
            x_axis = list(range(len(banks)))
            ax1 = figure.add_subplot(grid[0, 0])
            bar_width = 0.4
            ax1.bar([x - bar_width / 2 for x in x_axis], [bank_loan_totals[bank] for bank in banks], width=bar_width, label="Loan Amount")
            ax1.bar([x + bar_width / 2 for x in x_axis], [bank_remaining_totals[bank] for bank in banks], width=bar_width, label="Remaining Balance")
            ax1.set_title("Exposure by Bank")
            ax1.set_xticks(x_axis, banks, rotation=25, ha="right")
            ax1.set_ylabel("Amount")
            ax1.grid(True, axis="y", alpha=0.3)
            ax1.legend(fontsize=8)

            ax2 = figure.add_subplot(grid[0, 1])
            ax2.plot(banks, [bank_payment_totals.get(bank, 0.0) for bank in banks], marker="o", label="Total Payments")
            ax2.plot(banks, [bank_interest_totals.get(bank, 0.0) for bank in banks], marker="o", label="Interest Paid")
            ax2.set_title("Payments and Interest by Bank")
            ax2.set_ylabel("Amount")
            ax2.grid(True, alpha=0.3)
            ax2.legend(fontsize=8)
            ax2.tick_params(axis="x", rotation=25)

            ax3 = figure.add_subplot(grid[1, 0])
            pie_values = [bank_remaining_totals[bank] for bank in banks]
            if sum(pie_values) > 0:
                ax3.pie(pie_values, labels=banks, autopct="%1.1f%%", startangle=90)
            else:
                ax3.text(0.5, 0.5, "No remaining balance data", ha="center", va="center")
            ax3.set_title("Remaining Balance Share by Bank")

            ax4 = figure.add_subplot(grid[1, 1])
            ax4.bar(banks, [bank_paid_counts.get(bank, 0) for bank in banks], label="Paid")
            ax4.bar(
                banks,
                [bank_upcoming_counts.get(bank, 0) for bank in banks],
                bottom=[bank_paid_counts.get(bank, 0) for bank in banks],
                label="Upcoming",
            )
            ax4.set_title("Payment Status Counts by Bank")
            ax4.set_ylabel("Payment Count")
            ax4.tick_params(axis="x", rotation=25)
            ax4.grid(True, axis="y", alpha=0.3)
            ax4.legend(fontsize=8)
        elif chart_option == "By Loan":
            loan_labels = [f"{row['bank_name']} | {row['loan_number']}" for row in rows]
            loan_remaining = {label: float(row["remaining_balance"]) for label, row in zip(loan_labels, rows)}
            loan_payments: dict[str, float] = {label: 0.0 for label in loan_labels}
            loan_interest: dict[str, float] = {label: 0.0 for label in loan_labels}
            loan_series: dict[str, list[tuple[int, float]]] = {label: [] for label in loan_labels}

            for payment in payment_rows:
                label = f"{payment.get('bank_name', '')} | {payment.get('loan_number', '')}"
                if label not in loan_payments:
                    continue
                loan_payments[label] += float(payment.get("payment", 0.0))
                loan_interest[label] += float(payment.get("interest_paid", 0.0))
                try:
                    payment_number = int(payment.get("payment_number", 0))
                except (TypeError, ValueError):
                    payment_number = 0
                loan_series[label].append((payment_number, float(payment.get("remaining_balance", 0.0))))

            selected_labels = sorted(loan_labels, key=lambda label: loan_remaining[label], reverse=True)[:8]
            x_axis = list(range(len(selected_labels)))

            ax1 = figure.add_subplot(grid[0, 0])
            plotted = False
            for label in selected_labels:
                points = sorted(loan_series.get(label, []), key=lambda item: item[0])
                if not points:
                    continue
                ax1.plot([item[0] for item in points], [item[1] for item in points], marker="o", label=label)
                plotted = True
            if plotted:
                ax1.legend(fontsize=7, loc="upper right")
            else:
                ax1.text(0.5, 0.5, "No payment trajectories in selected range", ha="center", va="center")
            ax1.set_title("Remaining Balance Trajectory by Loan")
            ax1.set_xlabel("Payment #")
            ax1.set_ylabel("Remaining Balance")
            ax1.grid(True, alpha=0.3)

            ax2 = figure.add_subplot(grid[0, 1])
            ax2.bar(x_axis, [loan_payments[label] for label in selected_labels], color="#4C78A8")
            ax2.set_title("Total Payments by Loan")
            ax2.set_xticks(x_axis, selected_labels, rotation=25, ha="right")
            ax2.set_ylabel("Amount")
            ax2.grid(True, axis="y", alpha=0.3)

            ax3 = figure.add_subplot(grid[1, 0])
            ax3.bar(x_axis, [loan_interest[label] for label in selected_labels], color="#F58518")
            ax3.set_title("Total Interest Paid by Loan")
            ax3.set_xticks(x_axis, selected_labels, rotation=25, ha="right")
            ax3.set_ylabel("Amount")
            ax3.grid(True, axis="y", alpha=0.3)

            ax4 = figure.add_subplot(grid[1, 1])
            pie_labels = selected_labels[:6]
            pie_values = [loan_remaining[label] for label in pie_labels]
            if sum(pie_values) > 0:
                ax4.pie(pie_values, labels=pie_labels, autopct="%1.1f%%", startangle=90)
            else:
                ax4.text(0.5, 0.5, "No remaining balance data", ha="center", va="center")
            ax4.set_title("Remaining Balance Share by Loan")
        else:
            sorted_rows = sorted(rows, key=lambda item: float(item["remaining_balance"]), reverse=True)[:8]
            row_labels = [f"{row['bank_name']} | {row['loan_number']}" for row in sorted_rows]
            ax1 = figure.add_subplot(grid[0, 0])
            ax1.barh(row_labels, [float(row["remaining_balance"]) for row in sorted_rows], color="#59A14F")
            ax1.set_title("Top Remaining Balances")
            ax1.set_xlabel("Remaining Balance")
            ax1.grid(True, axis="x", alpha=0.3)

            ax2 = figure.add_subplot(grid[0, 1])
            pie_labels = ["Paid", "Upcoming"]
            pie_values = [status_counts[label] for label in pie_labels]
            if sum(pie_values) > 0:
                ax2.pie(pie_values, labels=pie_labels, autopct="%1.1f%%", startangle=90)
            else:
                ax2.text(0.5, 0.5, "No payment status data", ha="center", va="center")
            ax2.set_title("Payment Status Distribution")

            ax3 = figure.add_subplot(grid[1, 0])
            if payment_rows:
                cumulative_dates = []
                cumulative_payments = []
                cumulative_interest = []
                running_payments = 0.0
                running_interest = 0.0
                for payment in payment_rows:
                    running_payments += float(payment.get("payment", 0.0))
                    running_interest += float(payment.get("interest_paid", 0.0))
                    cumulative_dates.append(parse_date(str(payment.get("date", "")), "Payment Date"))
                    cumulative_payments.append(running_payments)
                    cumulative_interest.append(running_interest)
                ax3.plot(cumulative_dates, cumulative_payments, marker="o", label="Cumulative Payments")
                ax3.plot(cumulative_dates, cumulative_interest, marker="o", label="Cumulative Interest")
                ax3.legend(fontsize=8)
            else:
                ax3.text(0.5, 0.5, "No payment history in selected range", ha="center", va="center")
            ax3.set_title("Cumulative Payment Trend")
            ax3.set_xlabel("Payment Date")
            ax3.set_ylabel("Amount")
            ax3.grid(True, alpha=0.3)

            ax4 = figure.add_subplot(grid[1, 1])
            bar_labels = row_labels
            x_axis = list(range(len(bar_labels)))
            ax4.bar([x - 0.2 for x in x_axis], [float(row["loan_amount"]) for row in sorted_rows], width=0.4, label="Loan Amount")
            ax4.bar([x + 0.2 for x in x_axis], [float(row["remaining_balance"]) for row in sorted_rows], width=0.4, label="Remaining Balance")
            ax4.set_title("Loan Amount vs Remaining")
            ax4.set_xticks(x_axis, bar_labels, rotation=25, ha="right")
            ax4.set_ylabel("Amount")
            ax4.grid(True, axis="y", alpha=0.3)
            ax4.legend(fontsize=8)

        figure.suptitle(f"{chart_option} Dashboard ({filter_start.date()} to {filter_end.date()})", fontsize=14, fontweight="bold")
        figure.text(0.01, 0.01, kpi_text, fontsize=9, va="bottom")

        figure.tight_layout(rect=(0, 0.08, 1, 0.95))
        plt.show(block=False)
        self.refresh_chart_details()

    def _resolve_ml_loan(self) -> dict[str, object]:
        """Resolve a single loan for ML analysis from bank/loan filters."""
        bank_filter = self.ml_bank_entry.get().strip().lower()
        loan_filter = self.ml_loan_entry.get().strip().lower()
        matches = [
            loan
            for loan in self.data["loans"]
            if (not bank_filter or bank_filter == str(loan.get("bank_name", "")).strip().lower())
            and (not loan_filter or loan_filter == str(loan.get("loan_number", "")).strip().lower())
        ]
        if not matches:
            raise ValueError("No matching loan found for the selected Bank Name and Loan Number.")
        if len(matches) > 1:
            raise ValueError("Multiple loans match the selection. Select both Bank Name and Loan Number.")
        return matches[0]

    def run_ml_payment_analysis(self) -> None:
        """Run payment optimization analysis and suggest best and alternate extra-payment options."""
        loan = self._resolve_ml_loan()
        principal = float(loan.get("loan_amount", 0.0))
        annual_rate = float(loan.get("annual_interest_rate", 0.0))
        years = int(loan.get("term_years", 0))
        start_date = parse_date(str(loan.get("start_date", "")), "Start Date")

        requested_extra = float(self.ml_extra_entry.get().strip())
        max_extra = float(self.ml_max_extra_entry.get().strip())
        if requested_extra < 0:
            raise ValueError("Extra / Month must be greater than or equal to 0.")
        if max_extra <= 0:
            raise ValueError("Max Extra to Evaluate must be greater than 0.")

        effective_max = max(max_extra, requested_extra)
        baseline = simulate_extra_payment_metrics(principal, annual_rate, years, start_date, 0.0)
        baseline_months = int(baseline["months"])
        baseline_interest = float(baseline["total_interest"])
        baseline_payoff_date = str(baseline["payoff_date"])

        candidate_extras = sorted(
            {
                round(float(extra), 2)
                for extra in [50.0, 100.0, requested_extra] + [float(value) for value in range(50, int(effective_max) + 1, 50)]
                if extra > 0
            }
        )
        if len(candidate_extras) < 2:
            candidate_extras = [round(max(requested_extra, 50.0), 2), round(max(requested_extra, 50.0) + 50.0, 2)]

        option_rows: list[dict[str, object]] = []
        for extra_payment in candidate_extras:
            metrics = simulate_extra_payment_metrics(principal, annual_rate, years, start_date, extra_payment)
            months_saved = baseline_months - int(metrics["months"])
            interest_saved = baseline_interest - float(metrics["total_interest"])
            option_rows.append(
                {
                    "extra_payment": round(extra_payment, 2),
                    "months_saved": months_saved,
                    "interest_saved": round(interest_saved, 2),
                    "new_payoff_date": str(metrics["payoff_date"]),
                }
            )

        features = [float(row["extra_payment"]) for row in option_rows]
        month_targets = [float(row["months_saved"]) for row in option_rows]
        interest_targets = [float(row["interest_saved"]) for row in option_rows]
        month_intercept, month_slope = _fit_linear_model(features, month_targets)
        interest_intercept, interest_slope = _fit_linear_model(features, interest_targets)

        requested_metrics = simulate_extra_payment_metrics(principal, annual_rate, years, start_date, requested_extra)
        requested_months_saved = baseline_months - int(requested_metrics["months"])
        requested_interest_saved = baseline_interest - float(requested_metrics["total_interest"])
        predicted_months_saved = max(0.0, month_intercept + month_slope * requested_extra)
        predicted_interest_saved = max(0.0, interest_intercept + interest_slope * requested_extra)

        ranked = sorted(option_rows, key=lambda row: (float(row["interest_saved"]), float(row["months_saved"])), reverse=True)
        best_option = ranked[0]
        alternate_option = ranked[1] if len(ranked) > 1 else ranked[0]

        self.ml_tree.delete(*self.ml_tree.get_children())
        for label, row in (("Best", best_option), ("Alternate", alternate_option)):
            self.ml_tree.insert(
                "",
                "end",
                values=(
                    f"{label}: ${float(row['extra_payment']):,.2f}",
                    int(row["months_saved"]),
                    f"${float(row['interest_saved']):,.2f}",
                    row["new_payoff_date"],
                ),
            )

        summary = (
            f"Loan: {loan.get('bank_name', '')} | {loan.get('loan_number', '')}\n"
            f"Base Monthly Payment: ${float(baseline['scheduled_payment']):,.2f}\n"
            f"Base Payoff Date: {baseline_payoff_date}\n"
            f"Paying an extra ${requested_extra:,.2f}/month would reduce the overall loan by {requested_months_saved} payments, "
            f"saving ${requested_interest_saved:,.2f} in interest.\n"
            f"Model estimate for ${requested_extra:,.2f}/month ({int(round(predicted_months_saved))} payments reduced and "
            f"${predicted_interest_saved:,.2f} interest saved).\n"
            f"Recommended best option: ${float(best_option['extra_payment']):,.2f}/month "
            f"({int(best_option['months_saved'])} payments saved, ${float(best_option['interest_saved']):,.2f} interest saved)."
        )
        self.ml_summary_text.configure(state="normal")
        self.ml_summary_text.delete("1.0", tk.END)
        self.ml_summary_text.insert(tk.END, summary)
        self.ml_summary_text.configure(state="disabled")

    def clear_ml_analysis(self) -> None:
        """Clear ML inputs and results, restoring defaults."""
        self.ml_bank_entry.delete(0, tk.END)
        self.ml_loan_entry.delete(0, tk.END)
        self.ml_extra_entry.delete(0, tk.END)
        self.ml_extra_entry.insert(0, "200")
        self.ml_max_extra_entry.delete(0, tk.END)
        self.ml_max_extra_entry.insert(0, "500")
        self.refresh_ml_tab()

    def _write_table(self, file_path: str, headers: list[str], rows: list[list[object]]) -> None:
        """Write tabular data to CSV or XLSX based on file extension."""
        safe_rows = [self._sanitize_export_row(row) for row in rows]
        _, ext = os.path.splitext(file_path.lower())
        if ext == ".csv":
            with open(file_path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(headers)
                writer.writerows(safe_rows)
            return
        if ext == ".xlsx":
            if Workbook is None:
                raise ImportError("openpyxl is required for .xlsx export.")
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Export"
            sheet.append(headers)
            for row in safe_rows:
                sheet.append(row)
            workbook.save(file_path)
            return
        raise ValueError("Unsupported export extension. Use .csv or .xlsx.")

    def _sanitize_export_cell(self, value: object) -> object:
        """Neutralize spreadsheet formula prefixes in untrusted string cells before export."""
        if not isinstance(value, str):
            return value
        stripped = value.lstrip()
        if stripped.startswith(("=", "+", "-", "@")):
            return f"'{value}"
        return value

    def _sanitize_export_row(self, row: list[object]) -> list[object]:
        """Return a sanitized copy of an export row with spreadsheet-safe cell values."""
        return [self._sanitize_export_cell(value) for value in row]

    def _read_table(self, file_path: str) -> list[dict[str, object]]:
        """Read CSV or XLSX tabular data into a list of row dictionaries."""
        _, ext = os.path.splitext(file_path.lower())
        if ext == ".csv":
            with open(file_path, "r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                return [dict(row) for row in reader]
        if ext == ".xlsx":
            if load_workbook is None:
                raise ImportError("openpyxl is required for .xlsx import.")
            workbook = load_workbook(file_path)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(value).strip() if value is not None else "" for value in rows[0]]
            parsed_rows: list[dict[str, object]] = []
            for raw in rows[1:]:
                mapped: dict[str, object] = {}
                for idx, header in enumerate(headers):
                    mapped[header] = raw[idx] if idx < len(raw) else None
                parsed_rows.append(mapped)
            return parsed_rows
        raise ValueError("Unsupported import extension. Use .csv or .xlsx.")


def main() -> None:
    """Create the Tk root, initialize the app, and start the main event loop."""
    root = tk.Tk()
    app = LoanScheduleApp(root)

    def _handle_tk_exception(exc_type, exc_value, exc_traceback) -> None:
        """Log uncaught Tk callback exceptions and show a sanitized user error dialog."""
        LOGGER.exception("Tk callback exception", exc_info=(exc_type, exc_value, exc_traceback))
        messagebox.showerror("Application Error", "An unexpected application error occurred.")

    root.report_callback_exception = _handle_tk_exception
    app.root.mainloop()


if __name__ == "__main__":
    main()
