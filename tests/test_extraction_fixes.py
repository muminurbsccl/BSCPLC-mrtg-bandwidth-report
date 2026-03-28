"""Tests for the 12 extraction pipeline bug fixes."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mrtg_bandwidth_report import (
    match_graph_to_row,
    extract_client_keyword,
    _extract_maximum,
    _direction_pattern,
    _fuzzy_match,
    parse_graphs_from_text,
    convert_to_mbps,
    EXPECTED_E_ROWS,
    _SUM_CELLS,
)


# ── Bug 1: ADN DhakaColo pattern ordering ──────────────────────────

def test_adn_dhakacolo_sec_maps_to_e32():
    row, _ = match_graph_to_row(
        "2-IPT-BSCPLC-DHK-CORE-03 - ADN-DhakaColo-SEC - TenGigE0/0/0/1"
    )
    assert row == "E32"


def test_adn_dhakacolo_pri_maps_to_e31():
    row, _ = match_graph_to_row(
        "2-IPT-BSCPLC-DHK-CORE-03 - ADN-DhakaColo - TenGigE0/0/0/2"
    )
    assert row == "E31"


# ── Bug 2: Unit regex handles full unit strings ────────────────────

def test_unit_mbps_full_string():
    val, unitless = _extract_maximum("Inbound  Maximum: 7.8 Mbps", "Inbound")
    assert abs(val - 7.8) < 0.01
    assert not unitless


def test_unit_gbps_full_string():
    val, _ = _extract_maximum("Inbound  Maximum: 2.93 Gbps", "Inbound")
    assert abs(val - 2930.0) < 1.0


def test_unit_kbps_full_string():
    val, _ = _extract_maximum("Outbound  Maximum: 456 kbps", "Outbound")
    assert abs(val - 0.456) < 0.001


def test_unit_single_letter_still_works():
    val, _ = _extract_maximum("Inbound  Maximum: 5.5 M", "Inbound")
    assert abs(val - 5.5) < 0.01


def test_unit_no_unit_treats_as_bps():
    val, unitless = _extract_maximum("Inbound  Maximum: 1000000", "Inbound")
    assert abs(val - 1.0) < 0.01
    assert unitless


# ─�� Bug 3: E40 and E51 removed from EXPECTED_E_ROWS ───────────────

def test_e40_not_in_expected():
    assert "E40" not in EXPECTED_E_ROWS


def test_e51_not_in_expected():
    assert "E51" not in EXPECTED_E_ROWS


def test_adjacent_rows_still_present():
    assert "E39" in EXPECTED_E_ROWS
    assert "E41" in EXPECTED_E_ROWS
    assert "E50" in EXPECTED_E_ROWS


# ── Bug 4: OCR-garbled direction keywords ──────────────���───────────

def test_direction_lnbound_lowercase_l():
    val, _ = _extract_maximum("lnbound  Maximum: 5.5 M", "Inbound")
    assert val is not None and abs(val - 5.5) < 0.01


def test_direction_1nbound_digit_one():
    val, _ = _extract_maximum("1nbound  Maximum: 3.0 G", "Inbound")
    assert val is not None and abs(val - 3000.0) < 1.0


def test_direction_0utbound_zero_for_o():
    val, _ = _extract_maximum("0utbound  Maximum: 3.2 G", "Outbound")
    assert val is not None and abs(val - 3200.0) < 1.0


def test_direction_normal_inbound():
    val, _ = _extract_maximum("Inbound  Maximum: 10.0 M", "Inbound")
    assert val is not None and abs(val - 10.0) < 0.01


def test_direction_normal_outbound():
    val, _ = _extract_maximum("Outbound  Maximum: 8.0 M", "Outbound")
    assert val is not None and abs(val - 8.0) < 0.01


# ── Bug 5: None vs 0.0 conflation ─────────────────────────────────

def test_no_match_returns_none():
    val, _ = _extract_maximum("no match here at all", "Inbound")
    assert val is None


def test_nan_returns_none():
    # "nan" doesn't match the numeric regex, so returns None (not found)
    val, _ = _extract_maximum("Inbound  Maximum: nan M", "Inbound")
    assert val is None


# ── Bug 6: SUM mode for cache cells ───────────────��───────────────

def test_sum_cells_contains_e54():
    assert "E54" in _SUM_CELLS


# ── Bug 7: Expanded interface name detection ───────────────────────

def test_gigabitethernet_detected():
    text = "1-IPT-BSCPLC-COX-03 - TestClient - GigabitEthernet0/0\nInbound Maximum: 5.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


def test_gi_abbreviated_detected():
    text = "1-IPT-BSCPLC-COX-03 - TestClient - Gi0/0/0/1\nInbound Maximum: 5.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


def test_te_abbreviated_detected():
    text = "1-IPT-BSCPLC-COX-03 - TestClient - Te0/0/0/1\nInbound Maximum: 5.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


def test_bundle_ether_still_detected():
    text = "1-IPT-BSCPLC-COX-03 - Coronet-IPT - Bundle-Ether172\nInbound Maximum: 5.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


# ── Bug 8: Expanded title markers ─────────────────────��───────────

def test_bdren_marker_detected():
    text = "1-IPT-BSCPLC-COX-03 - BDREN-PRI - Bundle-Ether100\nInbound Maximum: 5.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


def test_coronet_marker_detected():
    text = "some-prefix - CORONET-IPT - TenGigE0/0\nOutbound Maximum: 2.0 G"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


def test_three_part_title_format_detected():
    """Title with ' - ' delimiters should be detected even without known markers."""
    text = "UnknownDevice - SomeClient - Bundle-Ether50\nInbound Maximum: 1.0 M"
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1


# ── Bug 9: Client keyword extraction for two-part titles ──────────

def test_two_part_title_returns_client_not_interface():
    kw = extract_client_keyword("ADN-DhakaColo - TenGigE0/0/0")
    assert "ADN" in kw


def test_three_part_title_returns_middle():
    kw = extract_client_keyword("1-IPT-BSCPLC-COX-03 - Coronet-IPT - HundredGigE0/0/")
    assert "Coronet" in kw


def test_single_part_title_returns_itself():
    kw = extract_client_keyword("SomeClientName")
    assert kw == "SomeClientName"


# ── Bug 10: OCR artifact cleanup ──────────────────────────────────

def test_pipe_replaced_with_l():
    row, _ = match_graph_to_row("1-IPT-BSCPLC-COX-03 - De|ta-LD - Bundle-Ether100")
    # Delta-LD should still match even with | -> l substitution
    assert row is not None


# ── Bug 11: Expanded fuzzy token map ──────────────────────────────

def test_fuzzy_equitel():
    row, _ = _fuzzy_match("SOMETHING EQUITEL SOMETHING")
    assert row == "E5"


def test_fuzzy_bdccl():
    row, _ = _fuzzy_match("SOMETHING BDCCL SOMETHING")
    assert row == "E43"


def test_fuzzy_telnet():
    row, _ = _fuzzy_match("SOMETHING TELNET SOMETHING")
    assert row == "E50"


def test_fuzzy_exabyte_cloudflare():
    row, _ = _fuzzy_match("EXABYTE CLOUDFLARE TEJ")
    assert row == "E54"


def test_fuzzy_delta_ld():
    row, _ = _fuzzy_match("DELTA LD COX")
    assert row == "E58"


def test_fuzzy_adn_dhakacolo_sec():
    """SEC entry must beat the generic ADN DHAKACOLO entry."""
    row, _ = _fuzzy_match("ADN DHAKACOLO SEC")
    assert row == "E32"


# ── Bug 12: Search window size ───────────────────────────���─────────

def test_value_found_beyond_30_lines():
    """Value at line 40 should be found with the new 50-line window."""
    lines = ["1-IPT-BSCPLC-COX-03 - TestClient - Bundle-Ether100"]
    lines += ["noise line"] * 38
    lines += ["Inbound  Maximum: 99.0 M"]
    text = "\n".join(lines)
    graphs = parse_graphs_from_text(text)
    assert len(graphs) >= 1
    assert graphs[0]["inbound_max"] is not None
    assert abs(graphs[0]["inbound_max"] - 99.0) < 0.1
