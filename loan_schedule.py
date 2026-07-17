import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, timedelta

# ========================================================================== #
# ================================== INFO ================================== #
# ========================================================================== #
# - User enters loan amount, interest rate, years, and start date.
# - Click "Calculate Schedule" to see the full amortization table.
# - Click "Save to Excel" to export the schedule.
# ========================================================================== #
# ================================== TODO ================================== #
# ========================================================================== #

# ========================================================================== #

def calculate_schedule(principal, annual_rate, years, start_date):
    """
    Calculate the loan amortization schedule.

    Args:
        principal (float): The initial loan amount.
        annual_rate (float): The annual interest rate (as a percentage, e.g., 5 for 5%).
        years (int): The term of the loan in years.
        start_date (datetime): The start date of the loan.

    Returns:
        pd.DataFrame: A DataFrame containing the full amortization schedule with columns:
            'Payment #', 'Date', 'Payment', 'Principal Paid', 'Interest Paid', 
            'Remaining Balance'
    """
    monthly_rate = annual_rate / 12 / 100
    n_payments = years * 12
    # Monthly payment formula
    payment = principal * (monthly_rate * (1 + monthly_rate) ** n_payments) / ((1 + monthly_rate) ** n_payments - 1)
    schedule = []
    balance = principal
    payment_date = start_date

    for i in range(1, n_payments + 1):
        interest = balance * monthly_rate
        principal_paid = payment - interest
        balance -= principal_paid
        if balance < 1e-2:  # Adjust for floating point
            principal_paid += balance
            balance = 0
        schedule.append({
            "Payment #": i,
            "Date": payment_date.strftime("%Y-%m-%d"),
            "Payment": round(payment, 2),
            "Principal Paid": round(principal_paid, 2),
            "Interest Paid": round(interest, 2),
            "Remaining Balance": round(balance, 2)
        })
        payment_date += timedelta(days=30)  # Approximate month
        if balance <= 0:
            break
    return pd.DataFrame(schedule)

def show_schedule():
    """
    Retrieve user input, calculate the loan schedule, and display it in the text widget.

    Raises:
        Shows a messagebox if input is invalid.
    """
    try:
        principal = float(entry_principal.get())
        rate = float(entry_rate.get())
        years = int(entry_years.get())
        start = datetime.strptime(entry_start.get(), "%Y-%m-%d")
    except Exception as e:
        messagebox.showerror("Input Error", f"Invalid input: {e}")
        return

    df = calculate_schedule(principal, rate, years, start)
    text.delete("1.0", tk.END)
    text.insert(tk.END, f"{'Pmt#':<6}{'Date':<12}{'Payment':<12}{'Principal':<12}{'Interest':<12}{'Balance':<12}\n")
    text.insert(tk.END, "-"*66 + "\n")
    for _, row in df.iterrows():
        text.insert(tk.END, f"{int(row['Payment #']):<6}{row['Date']:<12}${row['Payment']:<11,.2f}${row['Principal Paid']:<11,.2f}${row['Interest Paid']:<11,.2f}${row['Remaining Balance']:<11,.2f}\n")
    # Store for saving
    show_schedule.df = df

def save_to_excel():
    """
    Save the calculated loan schedule to an Excel file chosen by the user.

    Raises:
        Shows a warning if no schedule has been calculated yet.
        Shows an info messagebox when the file is saved.
    """
    if not hasattr(show_schedule, "df"):
        messagebox.showwarning("No Data", "Please calculate the schedule first.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        show_schedule.df.to_excel(file_path, index=False)
        messagebox.showinfo("Saved", f"Schedule saved to {file_path}")

root = tk.Tk()
root.title("Loan Amortization Schedule")
root.geometry("700x600")

tk.Label(root, text="Loan Amount:").grid(row=0, column=0, sticky="e")
entry_principal = tk.Entry(root)
entry_principal.grid(row=0, column=1)

tk.Label(root, text="Annual Interest Rate (%):").grid(row=1, column=0, sticky="e")
entry_rate = tk.Entry(root)
entry_rate.grid(row=1, column=1)

tk.Label(root, text="Term (years):").grid(row=2, column=0, sticky="e")
entry_years = tk.Entry(root)
entry_years.grid(row=2, column=1)

tk.Label(root, text="Start Date (YYYY-MM-DD):").grid(row=3, column=0, sticky="e")
entry_start = tk.Entry(root)
entry_start.insert(0, datetime.today().strftime("%Y-%m-%d"))
entry_start.grid(row=3, column=1)

tk.Button(root, text="Calculate Schedule", 
          command=show_schedule).grid(row=4, column=1, columnspan=2, pady=10)
tk.Button(root, text="Save to Excel", 
          command=save_to_excel).grid(row=4, column=0, columnspan=2, pady=5)

text = tk.Text(root, width=80, height=25, font=("Consolas", 10))
text.grid(row=6, column=0, columnspan=2, padx=60, pady=10)

root.mainloop()
