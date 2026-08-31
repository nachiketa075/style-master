"""
Automated tests for MM Size decimal preservation (Rule: never remove
decimal points from measurement values).

Run with:  venv\\Scripts\\python.exe test_mm_size.py
"""

from __future__ import annotations

import sys

import table_extractor as te
from ocr_engine import load_image_bgr, ocr_words_tiled


def _check(condition: bool, description: str, failures: list[str]) -> None:
    status = "PASS" if condition else "FAIL"
    print(f"[{status}] {description}")
    if not condition:
        failures.append(description)


def test_validate_mm_size(failures: list[str]) -> None:
    print("\n-- validate_mm_size() --")
    valid_cases = ["0.85", "0.90", "1", "1.00", "1.1", "1.8x1.8", "2.25x2.25", "3.5", "10"]
    for v in valid_cases:
        _check(te.validate_mm_size(v) is True, f'validate_mm_size({v!r}) is VALID', failures)

    invalid_cases = ["", "abc", "1.8x", "x1.8", "FALSE"]
    for v in invalid_cases:
        _check(te.validate_mm_size(v) is False, f'validate_mm_size({v!r}) is INVALID', failures)


def test_recover_missing_decimals_passthrough(failures: list[str]) -> None:
    """Cases the recovery function must NEVER touch: already has a
    decimal, isn't a plain digit/x run, or is empty. Uses a 1x1 dummy
    crop since these should all short-circuit before any image work."""
    import numpy as np

    print("\n-- _recover_missing_decimals() passthrough (no image evidence needed) --")
    dummy = np.zeros((10, 10, 3), dtype=np.uint8)
    cases = {
        "0.85": "0.85",  # already has a decimal - untouched
        "FALSE": "FALSE",  # not a digit/x run - untouched
        "6.5-7.0": "6.5-7.0",  # already has a decimal - untouched
        "": "",  # empty - untouched
    }
    for input_text, expected in cases.items():
        result = te._recover_missing_decimals(dummy, input_text)
        _check(result == expected, f"_recover_missing_decimals({input_text!r}) == {expected!r} (got {result!r})", failures)

    # A pure digit run with NO period-shaped blobs in the image (blank
    # crop) must be left exactly as OCR produced it - this is the
    # "never guess 18 -> 1.8" requirement.
    result = te._recover_missing_decimals(dummy, "18")
    _check(result == "18", f'_recover_missing_decimals("18", blank image) == "18" (no guessing) (got {result!r})', failures)
    result = te._recover_missing_decimals(dummy, "10")
    _check(result == "10", f'_recover_missing_decimals("10", blank image) == "10" (not "1.0") (got {result!r})', failures)


def test_real_pendant_image(failures: list[str]) -> None:
    """Mandatory real-image regression test (PDVWN0161(PVN0176).jpg):
    the Princess Blue-Sapphire row's MM Size must extract as exactly
    "1.8x1.8", never "18x18", and the other MM Size values in the same
    table must remain exactly as they were before this fix."""
    print("\n-- Real image test: samples/PDVWN0161(PVN0176).jpg --")
    path = "samples/PDVWN0161(PVN0176).jpg"
    try:
        img = load_image_bgr(path)
    except Exception as exc:
        _check(False, f"could not load {path}: {exc}", failures)
        return

    words = ocr_words_tiled(img, psm=11)
    result = te.extract_table(img, words)

    _check(len(result.rows) == 5, f"extracted exactly 5 stone rows (got {len(result.rows)})", failures)

    expected_mm_sizes = ["0.85", "0.9", "1", "1.1", "1.8x1.8"]
    actual_mm_sizes = [r.get("MM Size", "") for r in result.rows]
    print(f"  MM Size values: {actual_mm_sizes}")

    for i, expected in enumerate(expected_mm_sizes):
        actual = actual_mm_sizes[i] if i < len(actual_mm_sizes) else None
        _check(
            actual == expected,
            f"row {i} MM Size == {expected!r} (got {actual!r})",
            failures,
        )

    # The critical regression case from the bug report.
    _check(
        "18x18" not in actual_mm_sizes,
        'MM Size never contains "18x18" (the decimal-loss bug)',
        failures,
    )

    # Spot-check that unrelated columns were not disturbed by this fix.
    if len(result.rows) == 5:
        last = result.rows[4]
        _check(last.get("Location") == "Center", f"row 4 Location unaffected (got {last.get('Location')!r})", failures)
        _check(last.get("Sieve Size") == "FALSE", f"row 4 Sieve Size unaffected (got {last.get('Sieve Size')!r})", failures)
        _check(last.get("Setting Type") == "Prong", f"row 4 Setting Type unaffected (got {last.get('Setting Type')!r})", failures)
        # Regression: this cell's printed "0" was previously misread as
        # the letter "i" (no digits at all) by unconstrained OCR - fixed
        # via shape-based zero recovery (table_extractor._is_zero_shaped).
        _check(last.get("Total Wt") == "0", f"row 4 Total Wt == '0' (got {last.get('Total Wt')!r})", failures)

    # Table TOTAL row (Rule: read the printed total, one value per
    # image, never invented/blank when the source sheet has one).
    print(f"  total_pcs={result.total_pcs!r} total_wt={result.total_wt!r}")
    _check(result.total_pcs == "41", f"total_pcs == '41' (got {result.total_pcs!r})", failures)
    _check(result.total_wt == "0.1650", f"total_wt == '0.1650' (trailing zero preserved) (got {result.total_wt!r})", failures)


def test_real_engagement_ring_image(failures: list[str]) -> None:
    """Mandatory real-image regression test
    (MBS061780G-E(REVN1087)(RVN2106).jpg): the Center Oval row's MM
    Size must extract as exactly "4.0x2.9", never "40x?2.9" - Tesseract
    misclassifies the first decimal point as a literal "?" rather than
    dropping it, a distinct failure mode from the pendant image's
    dropped-decimal case above."""
    print("\n-- Real image test: samples/MBS061780G-E(REVN1087)(RVN2106).jpg --")
    path = "samples/MBS061780G-E(REVN1087)(RVN2106).jpg"
    try:
        img = load_image_bgr(path)
    except Exception as exc:
        _check(False, f"could not load {path}: {exc}", failures)
        return

    words = ocr_words_tiled(img, psm=11)
    result = te.extract_table(img, words)

    _check(len(result.rows) == 7, f"extracted exactly 7 stone rows (got {len(result.rows)})", failures)

    actual_mm_sizes = [r.get("MM Size", "") for r in result.rows]
    print(f"  MM Size values: {actual_mm_sizes}")

    _check(
        "40x?2.9" not in actual_mm_sizes,
        'MM Size never contains "40x?2.9" (the "?"-for-decimal-point bug)',
        failures,
    )
    if result.rows:
        _check(
            result.rows[0].get("MM Size") == "4.0x2.9",
            f"row 0 (Center Oval) MM Size == '4.0x2.9' (got {result.rows[0].get('MM Size')!r})",
            failures,
        )

    print(f"  total_pcs={result.total_pcs!r} total_wt={result.total_wt!r}")
    _check(result.total_pcs == "57", f"total_pcs == '57' (got {result.total_pcs!r})", failures)
    _check(result.total_wt == "0.4200", f"total_wt == '0.4200' (got {result.total_wt!r})", failures)


def main() -> int:
    failures: list[str] = []
    test_validate_mm_size(failures)
    test_recover_missing_decimals_passthrough(failures)
    test_real_pendant_image(failures)
    test_real_engagement_ring_image(failures)

    print(f"\n{'=' * 50}")
    if failures:
        print(f"{len(failures)} TEST(S) FAILED:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("ALL TESTS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
