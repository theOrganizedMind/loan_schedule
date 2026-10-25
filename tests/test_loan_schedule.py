from datetime import datetime

import pytest

from loan_schedule import (
    build_payment_rows,
    calculate_schedule,
    format_duration_remaining,
    resolve_data_directory,
)


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


def test_format_duration_remaining_switches_units():
    assert format_duration_remaining("2028-01-01", today=datetime(2026, 1, 1)) == "2.0 years"
    assert format_duration_remaining("2026-02-01", today=datetime(2026, 1, 1)) == "31 days"


def test_build_payment_rows_sets_upcoming_and_paid_by_date():
    schedule = calculate_schedule(1200, 12.0, 1, datetime(2026, 1, 1))
    rows = build_payment_rows("Bank", "LN-1", schedule, today=datetime(2026, 1, 31))

    assert rows[0]["status"] == "Paid"
    assert rows[1]["status"] == "Upcoming"
    assert rows[0]["bank_name"] == "Bank"
    assert rows[0]["loan_number"] == "LN-1"


def test_resolve_data_directory_prefers_private_directory(tmp_path):
    private_dir = tmp_path / "json_files"
    private_dir.mkdir()
    directory, mode = resolve_data_directory(str(tmp_path))

    assert directory == str(private_dir)
    assert mode == "PRIVATE"


def test_resolve_data_directory_falls_back_to_public_directory(tmp_path):
    directory, mode = resolve_data_directory(str(tmp_path))

    assert mode == "PUBLIC"
    assert directory == str(tmp_path / "data")
    assert (tmp_path / "data").exists()
