"""
Quick test script for attendance processing logic.
"""

import os
import sys
from datetime import date

# Ensure imports work
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

from utils import parse_duration_to_minutes, parse_date, normalize_email
from processor import process_attendance


def test_duration_parsing():
    print("=== Duration Parsing Tests ===")
    tests = [
        (75, 75.0),
        ("45", 45.0),
        ("1:30:00", 90.0),
        ("55", 55.0),
        ("1h 30m", 90.0),
        ("90 minutes", 90.0),
        ("2 hours", 120.0),
        (None, 0),
    ]
    for input_val, expected in tests:
        result = parse_duration_to_minutes(input_val)
        status = "PASS" if abs(result - expected) < 0.01 else "FAIL"
        print(f"  {status}: parse_duration({input_val!r}) = {result} (expected {expected})")


def test_date_parsing():
    print("\n=== Date Parsing Tests ===")
    tests = [
        ("15-01-2025", date(2025, 1, 15)),
        ("15/01/2025", date(2025, 1, 15)),
        ("2025-01-15", date(2025, 1, 15)),
        ("15-01-25", date(2025, 1, 15)),
    ]
    for input_val, expected in tests:
        result = parse_date(input_val)
        status = "PASS" if result == expected else "FAIL"
        print(f"  {status}: parse_date({input_val!r}) = {result} (expected {expected})")


def test_email_normalization():
    print("\n=== Email Normalization Tests ===")
    tests = [
        ("Alice@Example.COM", "alice@example.com"),
        ("  bob@test.com  ", "bob@test.com"),
        (None, ""),
        ("", ""),
    ]
    for input_val, expected in tests:
        result = normalize_email(input_val)
        status = "PASS" if result == expected else "FAIL"
        print(f"  {status}: normalize({input_val!r}) = {result!r} (expected {expected!r})")


def test_processing():
    print("\n=== Attendance Processing Test ===")
    test_dir = os.path.join(script_dir, "test_data")
    reg_path = os.path.join(test_dir, "sample_registration.csv")
    zoom_path = os.path.join(test_dir, "sample_zoom.csv")

    result = process_attendance(
        reg_filepath=reg_path,
        zoom_filepath=zoom_path,
        selected_date=date(2025, 1, 15),
        attendance_day=1,
        program_filter=None,
        min_duration=60,
        output_dir=test_dir,
    )

    print(f"  Success: {result.success}")
    print(f"  {result.summary()}")

    # Expected:
    # alice@example.com: 75 min -> Y
    # bob@example.com: 45+30=75 min -> Y
    # charlie@example.com: 1:30:00=90 min -> Y
    # diana@example.com: 55 min -> N (below threshold)
    # eve@example.com: 62 min -> Y
    # henry@example.com: 59+5=64 min -> Y
    # jack@example.com: 120 min -> Y
    # Total for date 15-01-2025: 7 registrations
    # Present: 6 (alice, bob, charlie, eve, henry, jack)
    # Below threshold: 1 (diana)
    # Not found: 0

    print(f"\n  Expected: 7 registrations, 6 present, 1 below threshold, 0 not found")
    print(f"  Got:      {result.total_registrations} registrations, "
          f"{result.matched_present} present, "
          f"{result.matched_below_threshold} below threshold, "
          f"{result.not_found_in_zoom} not found")
    print(f"  Unmatched Zoom emails: {result.unmatched_zoom_emails}")
    print(f"  Expected unmatched: ['unknown@random.com']")

    all_pass = (
        result.total_registrations == 7
        and result.matched_present == 6
        and result.matched_below_threshold == 1
        and result.not_found_in_zoom == 0
        and result.unmatched_zoom_emails == ['unknown@random.com']
    )
    print(f"\n  Overall: {'ALL TESTS PASSED' if all_pass else 'SOME TESTS FAILED'}")


if __name__ == '__main__':
    test_duration_parsing()
    test_date_parsing()
    test_email_normalization()
    test_processing()
