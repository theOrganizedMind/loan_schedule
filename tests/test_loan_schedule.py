from datetime import datetime

import pytest

from loan_schedule import calculate_schedule


def test_calculate_schedule_has_expected_shape_and_columns():
    start_date = datetime(2026, 1, 1)
    schedule = calculate_schedule(100_000, 5.0, 30, start_date)

    assert len(schedule) == 360
    assert list(schedule.columns) == [
        "Payment #",
        "Date",
        "Payment",
        "Principal Paid",
        "Interest Paid",
        "Remaining Balance",
    ]
    assert schedule.iloc[-1]["Remaining Balance"] == pytest.approx(0.0, abs=0.01)


def test_calculate_schedule_first_row_matches_known_values():
    start_date = datetime(2026, 1, 1)
    schedule = calculate_schedule(1200, 12.0, 1, start_date)
    first = schedule.iloc[0]

    assert first["Payment #"] == 1
    assert first["Date"] == "2026-01-01"
    assert first["Payment"] == pytest.approx(106.62, abs=0.01)
    assert first["Interest Paid"] == pytest.approx(12.00, abs=0.01)
    assert first["Principal Paid"] == pytest.approx(94.62, abs=0.01)
    assert first["Remaining Balance"] == pytest.approx(1105.38, abs=0.01)


def test_calculate_schedule_handles_zero_interest():
    start_date = datetime(2026, 1, 1)
    schedule = calculate_schedule(1200, 0.0, 1, start_date)

    assert len(schedule) == 12
    assert schedule.iloc[0]["Payment"] == pytest.approx(100.00, abs=0.01)
    assert schedule.iloc[0]["Interest Paid"] == pytest.approx(0.00, abs=0.01)
    assert schedule.iloc[-1]["Remaining Balance"] == pytest.approx(0.0, abs=0.01)


def test_calculate_schedule_uses_30_day_payment_intervals():
    start_date = datetime(2026, 1, 1)
    schedule = calculate_schedule(1200, 12.0, 1, start_date)

    assert schedule.iloc[0]["Date"] == "2026-01-01"
    assert schedule.iloc[1]["Date"] == "2026-01-31"


@pytest.mark.parametrize(
    ("principal", "annual_rate", "years"),
    [
        (0, 5.0, 30),
        (-1, 5.0, 30),
        (100_000, -0.1, 30),
        (100_000, 5.0, 0),
    ],
)
def test_calculate_schedule_rejects_invalid_numeric_inputs(principal, annual_rate, years):
    start_date = datetime(2026, 1, 1)
    with pytest.raises(ValueError):
        calculate_schedule(principal, annual_rate, years, start_date)


def test_calculate_schedule_rejects_non_datetime_start_date():
    with pytest.raises(TypeError):
        calculate_schedule(100_000, 5.0, 30, "2026-01-01")
