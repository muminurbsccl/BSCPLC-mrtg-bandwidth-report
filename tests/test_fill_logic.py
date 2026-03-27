import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# These imports will fail until Task 2 adds the symbols — that's expected (RED phase)
from mrtg_bandwidth_report import (
    _pick_e_fill,
    _FILL_GREEN, _FILL_AMBER, _FILL_ORANGE, _FILL_RED, _FILL_BLUE,
)


def test_exactly_70_pct_is_green():
    assert _pick_e_fill(70.0, False) is _FILL_GREEN


def test_71_pct_is_amber():
    assert _pick_e_fill(71.0, False) is _FILL_AMBER


def test_exactly_90_pct_is_amber():
    assert _pick_e_fill(90.0, False) is _FILL_AMBER


def test_91_pct_is_orange():
    assert _pick_e_fill(91.0, False) is _FILL_ORANGE


def test_exactly_100_pct_is_orange():
    assert _pick_e_fill(100.0, False) is _FILL_ORANGE


def test_101_pct_is_red():
    assert _pick_e_fill(101.0, False) is _FILL_RED


def test_corrected_within_70_is_blue():
    # Auto-corrected OCR value, within safe allocation — flag blue for awareness
    assert _pick_e_fill(50.0, True) is _FILL_BLUE


def test_corrected_at_80_pct_is_amber():
    # Corrected but in amber zone — utilization band takes priority over blue
    assert _pick_e_fill(80.0, True) is _FILL_AMBER


def test_corrected_over_100_is_red():
    # Even if OCR-corrected, red if still over allocation
    assert _pick_e_fill(150.0, True) is _FILL_RED


def test_zero_pct_is_green():
    assert _pick_e_fill(0.0, False) is _FILL_GREEN
