#!/usr/bin/env python3
"""
MRTG Bandwidth Report Generator
================================
Extracts Maximum Usage (Mbps) from MRTG/Cacti graph PDFs and generates
a Bandwidth Report (MAX) xlsx file from a template.

ACCURACY NOTE:
  This tool uses OCR (Tesseract) to read text from graph images in PDFs.
  OCR accuracy on MRTG graphs is typically ~80-90%. Some values may be
  misread, especially:
    - Unit letters (M/G/k) can be missed or misread
    - Decimal points can be confused with commas
    - Graph titles with OCR artifacts may not match patterns
  ALWAYS manually verify the output xlsx against the source PDF,
  especially for critical or unusual values.

Requirements:
  Python packages: pip install openpyxl pdf2image pytesseract Pillow
  System packages: Tesseract OCR, Poppler (pdftoppm)

Usage:
  python mrtg_bandwidth_report.py                 # Launch GUI
  python mrtg_bandwidth_report.py --cli --pdf <input.pdf> --template <template.xlsx> --date "26 March 2026"

Installation (one-liner):
  macOS:   brew install tesseract poppler && pip install openpyxl pdf2image pytesseract Pillow
  Ubuntu:  sudo apt install -y tesseract-ocr poppler-utils && pip install openpyxl pdf2image pytesseract Pillow
  Windows: choco install tesseract poppler && pip install openpyxl pdf2image pytesseract Pillow
"""

import re
import os
import sys
import json
import shutil
import logging
import argparse
from datetime import datetime

# ---------------------------------------------------------------------------
# Third-party imports (checked at startup)
# ---------------------------------------------------------------------------
MISSING_DEPS = []
try:
    from openpyxl import load_workbook
except ImportError:
    MISSING_DEPS.append("openpyxl")
try:
    from pdf2image import convert_from_path
except ImportError:
    MISSING_DEPS.append("pdf2image")
try:
    import pytesseract
    from PIL import Image
except ImportError:
    MISSING_DEPS.append("pytesseract Pillow")

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("mrtg")


# =====================================================================
# SECTION 1 — GRAPH CLIENT NAME → SPREADSHEET ROW MAPPING
# =====================================================================
# The graph title from OCR typically looks like:
#   "1-IPT-BSCPLC-COX-03 - Coronet-IPT - HundredGigE0/0/"
#   "2-IPT-BSCPLC-DHK-CORE-03 - Velocity-DhakaColo - Bundle-Ether655"
#
# We extract the CLIENT KEYWORD (middle part between " - " delimiters)
# and match against these patterns.
#
# Each entry: (regex_pattern, excel_cell, description)
# Pattern is matched against the CLIENT KEYWORD (case-insensitive).
# First match wins — put more specific patterns first.
# =====================================================================

GRAPH_TO_ROW_MAP = [
    # ---- IIG Clients (rows 4-28) ----
    (r"BDHUB.*IIG|BDHUB-15G-IIG", "E4", "BDHUB DHK IIG"),
    (r"Equitel|EQUITEL", "E5", "Equitel DHK"),
    (r"Skytel.*DC|Skytel.*PRIMARY|Skytel.*IIG.*PRI", "E6", "Skytel Primary"),
    (r"Skytel.*TEJ|Skytel.*SECONDARY|Skytel.*IIG.*SEC", "E7", "Skytel Secondary"),
    (r"PEEREX.*TEJ\b|PEEREX-TEJ", "E8", "Peerex DHK"),
    (r"PEEREX.*9\.?5G|PEEREX-9", "E9", "Peerex Cox-9500"),
    (r"PEEREX.*COX.*0[2z]|PEEREX-COX-0[2z]", "E10", "Peerex Cox-3432"),
    (r"F.H.*KKT.*BE|FAH.*KKT|F@H.*KKT", "E11", "F@Home KKT"),
    (r"NOVOCOM|Novocom", "E12", "NOVOCOM DHK"),
    (r"WINDSTREAM.*IIG|Windstream.*IIG", "E13", "Windstream COX IIG"),
    (r"Velocity.*Tej|VELOCITY.*TEJ", "E14", "Velocity DHK-Tej-Primary"),
    (r"Velocity.*DHK|Velocity.*DhakaColo|Velocity.*Khaza|VELOCITY.*DHK", "E15", "Velocity DHK-Khaza-Sec"),
    (r"Virgo|VIRGO", "E16", "Virgo DHK"),
    (r"DELTA.IPT(?!.*LD)|DELTA-COX\b(?!.*LD)|Delta-COX\b(?!.*LD)", "E17", "Delta COX IIG"),
    (r"Exabyte-IPT|Exabyte.*IPT.*Bundle-Ether172", "E18", "Exabyte COX IIG"),
    (r"Coronet-IPT|Coronet.*IPT|CORONET.*IPT|C.?RONET.*IPT", "E19", "Coronet COX IIG"),
    (r"INTRAGLOBE.*IPT|Intraglobe.*IPT", "E20", "Intraglobe KKT IIG"),
    (r"GMax-IPT|GMAX.*IPT|GMax.*IPT", "E21", "Green Max COX IIG"),
    (r"BD.?LINK.*(?:IIG|DC)|BDLINK.*DC", "E22", "BD-LINK DHK"),
    (r"ADNGateway|ADN.*Gateway", "E23", "ADN-GW DHK-Primary"),
    (r"REGO.COX.IIG|REGO_COX_IIG|REGO.*COX.*IIG", "E25", "Rego COX IIG"),
    (r"REGO-IIG|REGO.*IIG(?!.*COX)", "E26", "Rego KKT IIG"),
    (r"GFCL-IPT|GFCL.*IPT", "E27", "GFCL COX IIG"),
    (r"MAXHUB.*COX|BE-MAXHUB|MAX.?HUB", "E28", "Max Hub Ltd COX"),

    # ---- ISP Clients (rows 31-51) ----
    (r"ADN.*DHK.*ISP|ADN-DhakaColo|ADN.*DhakaColo|ADN.*DC(?!.*LD)", "E31", "ADN DHK-Primary"),
    (r"TELET.LK.*MOG|Teletalk.*PRI.*DHK|TELETALK.*PRI.*DHK|Teletalk.*MOG", "E33", "Teletalk DHK-Primary"),
    (r"Telet.lk.*CTG.*Sec|TELET.LK.*CTG.*Sec|Teltatalk.*CTG.*Sec", "E35", "Teletalk CTG-Secondary"),
    (r"Teletalk.*PRI.*CTG|TELETALK.*PRI.*CTG|Teletalk.*CTG(?!.*Sec)", "E34", "Teletalk CTG-Primary"),
    (r"COL.CTG.Pri|COL.*CTG.*Pri|COL.CTG.BE|COL.*CTG.*BE", "E36", "COL CTG-Primary"),
    (r"COL.CTG.SEC|COL.*CTG.*SEC|COL.CTG.Sec", "E37", "COL CTG-Secondary"),
    (r"COXLINKT|COX.?LINK.*COX", "E38", "COX-Link COX"),
    (r"SSOnline.*Cloud|SS.?Online.*Cloud|SSOnline.*DC|SS.?Online.*DC", "E39", "SS Online DHK-Primary"),
    (r"BDREN.*PRI|BDREN.*DHK.*03.*TO.*CGS|BDREN-DHK.*CGS", "E41", "BDREN DHK-Primary"),
    (r"BDREN.*SEC(?!.*Equinix)", "E42", "BDREN DHK-Secondary"),
    (r"BDCCL|BD.?CCL", "E43", "BDCCL DHK"),
    (r"Link3.*Dhaka.*Colo|LINK3.*DC|Link3.*DC", "E44", "Link3 DHK-Primary-DhakColo"),
    (r"Link3.*Tej|LINK3.*TEJ", "E45", "Link3 DHK-Secondary-Tejgaon"),
    (r"Dhaka.?Link.*Pri|DHAKA.?LINK.*PRI", "E46", "Dhaka-Link DHK-Primary"),
    (r"Dhaka.?Link.*Sec|DHAKA.?LINK.*SEC", "E47", "Dhaka-Link DHK-Secondary"),
    (r"BDREN.*Equinix|BDREN.*CC.*Equinix", "E48", "BDREN DHK (Equinix)"),
    (r"Race.?Online|RACE.?ONLINE", "E49", "Race Online Ltd"),
    (r"Telnet.*ICT|TELNET.*ICT|Telnet.*DC|Telnet-DC", "E50", "Telnet Dhk-Colo"),
    (r"C0RONET-CTG|CORONET-CTG.*(?!IPT)", "E19", "Coronet (CTG fallback)"),

    # ---- Cache (rows 54-55) ----
    (r"Exabyte.*Cloudflare.*TEJ", "E54", "Exabyte Cache TEJ"),
    (r"EDGENEXT.*CLOUD|BE-EDGENEXT", "E54", "Exabyte Cache (EDGENEXT)"),
    (r"Exabyte.*Cloudflare.*DC", "E54", "Exabyte Cache DC"),

    # ---- LD Clients (rows 58-67) ----
    (r"DELTA.LD|Delta.LD", "E58", "Delta COX LD"),
    (r"Intraglobe.LD|INTRAGLOBE.LD", "E59", "Intraglobe KKT LD"),
    (r"Coronet.LD|CORONET.LD|C.?RONET.LD", "E60", "Coronet COX LD"),
    (r"GFCL.LD|GFCL.*LD", "E61", "GFCL COX LD"),
    (r"BDHUB.LD(?!.*15G)|BDHUB.*LD.*313", "E62", "BDHUB COX LD"),
    (r"GMAX.LD|GMax.LD|Green.?Max.*LD", "E63", "Green Max COX LD"),
    (r"Exabyte.LD|EXABYTE.LD|EXABYTE.*COX.*LD", "E64", "Exabyte COX LD"),
    (r"Windstream.LD|WINDSTREAM.LD", "E65", "Windstream KKT LD"),
    (r"COL.COX.LD|C0L.COX.LD", "E66", "COL COX LD"),
    (r"SS.?Online.LD|SSONLINE.LD", "E67", "SS Online COX LD"),
    (r"BDHUB.LD.15G|BDHUB.*LD.*15G", "E62", "BDHUB COX LD (15G)"),
    (r"Delta.LD.*5000.*Cox|Delta-LD-5000", "E58", "Delta COX LD (Cox's Bazar)"),
]

# Additional broad fallbacks — used when the full title is tried
FALLBACK_MAP = [
    (r"GREENMAX.*COX(?!.*LD)", "E21", "Green Max COX IIG (GREENMAX)"),
    (r"Windstrem.*IPT|WINDSTREAM.*IPT(?!.*IIG)", "E13", "Windstream COX (typo)"),
    (r"PEEREX.*DHKCOLO", "E8", "Peerex DHK (DHKCOLO)"),
    (r"DELTA-IPT.*Ether193", "E17", "Delta COX IIG (IPT/193)"),
    (r"BDHUB-IIG", "E4", "BDHUB DHK IIG (simple)"),
    (r"COL-CTG-Pri", "E36", "COL CTG-Primary (Pri)"),
    (r"Telt?at?alk.*CTG.*Sec|Teltatalk.*Sec", "E35", "Teletalk CTG-Sec (OCR)"),
    (r"TELET\w+.*MOG|Teletalk.*MOG", "E33", "Teletalk DHK (MOG)"),
    (r"ADN.*DhakaColo", "E31", "ADN DHK-Primary (DhakaColo)"),
    (r"BDREN-PRI|BDREN.*PRI", "E41", "BDREN DHK-Primary (PRI)"),
    (r"BDREN-SEC|BDREN.*SEC", "E42", "BDREN DHK-Secondary (SEC)"),
    (r"Telnet-DC|Telnet.*DC", "E50", "Telnet Dhk-Colo (DC)"),
    (r"SSOnline-DC|SS.?Online.*DC", "E39", "SS Online DHK (DC)"),
    (r"BDLINK|BD.LINK(?!.*LD)", "E22", "BD-LINK DHK (simple)"),
]


# =====================================================================
# SECTION 2 — OCR + PARSING ENGINE
# =====================================================================

def pdf_to_images(pdf_path: str, dpi: int = 250) -> list:
    """Convert PDF pages to PIL Image list."""
    log.info(f"Converting PDF to images at {dpi} DPI ...")
    images = convert_from_path(pdf_path, dpi=dpi)
    log.info(f"  -> {len(images)} pages converted.")
    return images


def ocr_full_page(page_img: Image.Image) -> str:
    """Run Tesseract OCR on a full page image."""
    try:
        return pytesseract.image_to_string(page_img, config="--psm 6")
    except Exception as e:
        log.warning(f"OCR failed: {e}")
        return ""


def extract_client_keyword(title: str) -> str:
    """
    Extract the client keyword from a graph title.
    Input:  "1-IPT-BSCPLC-COX-03 - Coronet-IPT - HundredGigE0/0/"
    Output: "Coronet-IPT"

    The client name is typically the 2nd segment when split by " - ".
    """
    parts = re.split(r"\s+-\s+", title)
    if len(parts) >= 2:
        return parts[1].strip()
    return title.strip()


def parse_graphs_from_text(text: str) -> list:
    """
    Parse full-page OCR text to extract multiple graph data blocks.

    Each MRTG graph produces text roughly like:
      <title line with interface name>
      ...graph area noise...
      Inbound   Current: X   Average: Y   Maximum: Z
      Outbound  Current: X   Average: Y   Maximum: Z
      Total: ...

    Returns list of dicts with keys: title, inbound_max, outbound_max
    """
    graphs = []
    lines = text.split("\n")

    # Strategy: find title lines (containing interface names like Bundle-Ether, TenGigE, HundredGigE)
    # then find the Inbound/Outbound stats that follow each title
    title_indices = []
    for i, line in enumerate(lines):
        if re.search(r"(Bundle-Ether|TenGigE|HundredGigE|GigE\d|Ether\d{2,})", line, re.I):
            # Make sure it looks like a title (has BSCPLC or client name pattern)
            if re.search(r"(BSCPLC|IPT|IPBW|LD|CORE)", line, re.I) or re.search(r"\d-\w+-\w+", line):
                title_indices.append(i)

    for ti_idx, ti in enumerate(title_indices):
        title = lines[ti].strip()

        # Search for Inbound/Outbound Maximum in lines after this title
        # (up to the next title or 30 lines, whichever comes first)
        next_ti = title_indices[ti_idx + 1] if ti_idx + 1 < len(title_indices) else len(lines)
        search_end = min(ti + 30, next_ti)
        search_block = "\n".join(lines[ti:search_end])

        in_max = _extract_maximum(search_block, "Inbound")
        out_max = _extract_maximum(search_block, "Outbound")

        graphs.append({
            "title": title,
            "inbound_max": in_max,
            "outbound_max": out_max,
        })

    return graphs


def _extract_maximum(text_block: str, direction: str) -> float:
    """
    Extract the Maximum value for Inbound or Outbound from a text block.

    Handles OCR quirks: "Max1mum", "Maximun", "Maxinum" etc.
    Handles units: G, M, k, or plain bps.
    Handles -nan values.
    """
    # Find lines containing the direction keyword
    for line in text_block.split("\n"):
        if not re.search(rf"\b{direction}\b", line, re.I):
            continue

        # Sometimes stats wrap to next line, so we also grab surrounding text
        # Look for Maximum: <value> [unit]
        # OCR can produce: "Max1mum:", "Maximum:", "Maximun:", etc.
        pat = r"Max\w*[:\s]+\s*([-\d.,]+)\s*([GgMmKk]?)\b"
        match = re.search(pat, line, re.I)
        if match:
            val_str = match.group(1).strip()
            unit = match.group(2).strip()

            if "nan" in val_str.lower():
                return 0.0

            try:
                val = float(val_str.replace(",", ""))
            except ValueError:
                continue

            return convert_to_mbps(val, unit)

    # Try a broader search across all lines with Maximum
    lines = text_block.split("\n")
    for i, line in enumerate(lines):
        if re.search(rf"\b{direction}\b", line, re.I):
            # Check this line and the next 2 combined
            combined = line
            if i + 1 < len(lines):
                combined += " " + lines[i + 1]
            if i + 2 < len(lines):
                combined += " " + lines[i + 2]

            pat = r"Max\w*[:\s]+\s*([-\d.,]+)\s*([GgMmKk]?)\b"
            match = re.search(pat, combined, re.I)
            if match:
                val_str = match.group(1).strip()
                unit = match.group(2).strip()
                if "nan" in val_str.lower():
                    return 0.0
                try:
                    val = float(val_str.replace(",", ""))
                except ValueError:
                    continue
                return convert_to_mbps(val, unit)

    return None


def convert_to_mbps(value: float, unit: str) -> float:
    """
    Convert MRTG value+unit to Megabits per second.

    MRTG graph units:
      G  = Gbps  -> multiply by 1000
      M  = Mbps  -> as-is
      k  = kbps  -> divide by 1000
      (none) = bps -> divide by 1,000,000
    """
    u = unit.upper() if unit else ""
    if u == "G":
        return round(value * 1000, 2)
    elif u == "M":
        return round(value, 2)
    elif u == "K":
        return round(value / 1000, 4)
    else:
        # Plain bps — if value is already large (>1000), assume it's already in Mbps-scale
        # This handles OCR edge cases where the unit letter wasn't captured
        if value > 100000:
            return round(value / 1_000_000, 2)
        elif value > 1000:
            # Ambiguous — could be bps or Mbps without unit
            # Heuristic: if value fits expected Mbps range, keep as-is
            return round(value, 2)
        else:
            # Small number without unit — likely already correct or bps
            return round(value, 4)


def match_graph_to_row(title: str) -> tuple:
    """
    Given a full graph title, find which spreadsheet row it maps to.
    Returns (cell_ref, description) or (None, None).

    Strategy:
    1. Extract the client keyword from the title
    2. Match against GRAPH_TO_ROW_MAP using the client keyword
    3. If no match, try matching full title against GRAPH_TO_ROW_MAP
    4. If still no match, try FALLBACK_MAP against full title
    """
    if not title:
        return None, None

    # Clean OCR artifacts
    clean_title = re.sub(r"[|!]", "l", title)
    clean_title = re.sub(r"\s+", " ", clean_title)

    # Extract client keyword
    client = extract_client_keyword(clean_title)

    # Try matching client keyword
    for pattern, row, desc in GRAPH_TO_ROW_MAP:
        if re.search(pattern, client, re.I):
            return row, desc

    # Try matching full title
    for pattern, row, desc in GRAPH_TO_ROW_MAP:
        if re.search(pattern, clean_title, re.I):
            return row, desc

    # Try fallback patterns on full title
    for pattern, row, desc in FALLBACK_MAP:
        if re.search(pattern, clean_title, re.I):
            return row, desc

    return None, None


# =====================================================================
# SECTION 3 — FULL EXTRACTION PIPELINE
# =====================================================================

def extract_all_graphs(pdf_path: str, dpi: int = 250, progress_cb=None) -> dict:
    """
    Main pipeline: PDF → images → OCR → parsed stats → mapped to rows.

    Returns dict with keys:
      results: { "E4": {"mbps": 13550, "title": "...", "desc": "..."}, ... }
      unmatched: list of unmatched graph dicts
      could_not_open: count of "Could not open!" entries
      all_graphs: all parsed graph info for debugging
      total_pages: number of pages
    """
    images = pdf_to_images(pdf_path, dpi)
    total_pages = len(images)

    results = {}
    unmatched = []
    could_not_open = 0
    all_graphs = []

    for page_idx, page_img in enumerate(images):
        page_num = page_idx + 1
        if progress_cb:
            progress_cb(page_num, total_pages, f"Processing page {page_num}/{total_pages}")
        log.info(f"Processing page {page_num}/{total_pages} ...")

        # OCR the full page
        full_text = ocr_full_page(page_img)

        # Count "Could not open!" entries
        cno_count = len(re.findall(r"Could not open", full_text, re.I))
        could_not_open += cno_count

        # Parse graph blocks from the text
        graphs = parse_graphs_from_text(full_text)
        log.info(f"  Found {len(graphs)} graph(s) on page {page_num}")

        for g in graphs:
            in_max = g["inbound_max"] or 0.0
            out_max = g["outbound_max"] or 0.0
            max_mbps = max(in_max, out_max)

            row_ref, desc = match_graph_to_row(g["title"])

            info = {
                "page": page_num,
                "title": g["title"],
                "inbound_max": in_max,
                "outbound_max": out_max,
                "max_mbps": max_mbps,
                "row_ref": row_ref,
                "desc": desc,
            }
            all_graphs.append(info)

            if row_ref:
                # Keep the larger value if duplicate
                if row_ref not in results or max_mbps > results[row_ref]["mbps"]:
                    results[row_ref] = {"mbps": max_mbps, "title": g["title"], "desc": desc}
                log.info(f"  MATCH: {desc} -> {row_ref} = {max_mbps:.2f} Mbps")
            else:
                unmatched.append(info)
                log.warning(f"  UNMATCHED: {g['title'][:70]}")

    return {
        "results": results,
        "unmatched": unmatched,
        "could_not_open": could_not_open,
        "all_graphs": all_graphs,
        "total_pages": total_pages,
    }


# =====================================================================
# SECTION 4 — XLSX GENERATION
# =====================================================================

def generate_report(template_path: str, extraction_data: dict, output_path: str, report_date: str = None):
    """
    Load template xlsx, fill in E-column values, update title date, save output.
    """
    log.info(f"Loading template: {template_path}")
    wb = load_workbook(template_path)
    ws = wb.active

    if report_date:
        ws["A1"] = f"Daily Usage Report ({report_date})"

    filled = 0
    for row_ref, data in extraction_data.items():
        mbps = data["mbps"]
        try:
            cell = ws[row_ref]
            if mbps is not None and mbps > 0:
                cell.value = round(mbps, 2) if mbps < 100 else round(mbps)
            elif mbps == 0:
                cell.value = 0
            filled += 1
            log.info(f"  {row_ref} = {cell.value}  ({data.get('desc', '')})")
        except Exception as e:
            log.error(f"  Failed to write {row_ref}: {e}")

    wb.save(output_path)
    log.info(f"Report saved: {output_path} ({filled} cells filled)")
    return output_path


# =====================================================================
# SECTION 5 — CLI MODE
# =====================================================================

def run_cli():
    parser = argparse.ArgumentParser(description="MRTG Bandwidth Report Generator (CLI)")
    parser.add_argument("--pdf", required=True, help="Input PDF with MRTG graphs")
    parser.add_argument("--template", required=True, help="Template xlsx file")
    parser.add_argument("--output", help="Output xlsx path (auto-generated if omitted)")
    parser.add_argument("--date", help="Report date, e.g. '26 March 2026'")
    parser.add_argument("--dpi", type=int, default=250, help="PDF render DPI (default: 250)")
    parser.add_argument("--debug-json", help="Save debug info to JSON file")
    args = parser.parse_args()

    if not os.path.isfile(args.pdf):
        print(f"ERROR: PDF not found: {args.pdf}"); sys.exit(1)
    if not os.path.isfile(args.template):
        print(f"ERROR: Template not found: {args.template}"); sys.exit(1)

    if not args.output:
        date_str = args.date or datetime.now().strftime("%d %B %Y")
        args.output = f"Bandwidth Report (MAX) For {date_str}.xlsx"

    print(f"Extracting from: {args.pdf}")
    data = extract_all_graphs(args.pdf, dpi=args.dpi)

    print(f"\n{'='*50}")
    print(f"Extraction Summary:")
    print(f"  Pages:           {data['total_pages']}")
    print(f"  Graphs matched:  {len(data['results'])}")
    print(f"  Unmatched:       {len(data['unmatched'])}")
    print(f"  Could not open:  {data['could_not_open']}")

    if data["unmatched"]:
        print(f"\nUnmatched graphs (may need mapping updates):")
        for g in data["unmatched"]:
            print(f"  Page {g['page']}: {g['title'][:70]}  (max={g['max_mbps']:.2f})")

    print(f"\nMatched values:")
    for ref in sorted(data["results"].keys(), key=lambda x: (x[0], int(re.search(r'\d+', x).group()))):
        info = data["results"][ref]
        print(f"  {ref} = {info['mbps']:.2f} Mbps  ({info['desc']})")

    generate_report(args.template, data["results"], args.output, args.date)
    print(f"\nOutput: {args.output}")

    if args.debug_json:
        debug = {
            "results": {k: {"mbps": v["mbps"], "desc": v["desc"]} for k, v in data["results"].items()},
            "unmatched": [{"page": g["page"], "title": g["title"], "max_mbps": g["max_mbps"]} for g in data["unmatched"]],
            "could_not_open": data["could_not_open"],
        }
        with open(args.debug_json, "w") as f:
            json.dump(debug, f, indent=2)
        print(f"Debug JSON: {args.debug_json}")


# =====================================================================
# SECTION 6 — GUI (tkinter)
# =====================================================================

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    import threading

    class MRTGApp:
        def __init__(self, root):
            self.root = root
            self.root.title("MRTG Bandwidth Report Generator")
            self.root.geometry("850x720")
            self.root.minsize(700, 550)
            self.pdf_path = tk.StringVar()
            self.template_path = tk.StringVar()
            self.output_path = tk.StringVar()
            self.report_date = tk.StringVar(value=datetime.now().strftime("%d %B %Y"))
            self.dpi = tk.IntVar(value=250)
            self.running = False
            self._build_ui()

        def _build_ui(self):
            main = ttk.Frame(self.root, padding=12)
            main.pack(fill=tk.BOTH, expand=True)

            ttk.Label(main, text="MRTG Bandwidth Report Generator",
                      font=("Helvetica", 16, "bold")).pack(pady=(0, 10))

            # Input section
            inp = ttk.LabelFrame(main, text="Input Files", padding=8)
            inp.pack(fill=tk.X, pady=4)

            r1 = ttk.Frame(inp); r1.pack(fill=tk.X, pady=2)
            ttk.Label(r1, text="Input PDF:", width=14, anchor="e").pack(side=tk.LEFT)
            ttk.Entry(r1, textvariable=self.pdf_path, width=50).pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
            ttk.Button(r1, text="Browse...", command=self._browse_pdf).pack(side=tk.LEFT)

            r2 = ttk.Frame(inp); r2.pack(fill=tk.X, pady=2)
            ttk.Label(r2, text="Template XLSX:", width=14, anchor="e").pack(side=tk.LEFT)
            ttk.Entry(r2, textvariable=self.template_path, width=50).pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
            ttk.Button(r2, text="Browse...", command=self._browse_template).pack(side=tk.LEFT)

            # Output section
            out = ttk.LabelFrame(main, text="Output Settings", padding=8)
            out.pack(fill=tk.X, pady=4)

            r3 = ttk.Frame(out); r3.pack(fill=tk.X, pady=2)
            ttk.Label(r3, text="Report Date:", width=14, anchor="e").pack(side=tk.LEFT)
            ttk.Entry(r3, textvariable=self.report_date, width=30).pack(side=tk.LEFT, padx=4)

            r4 = ttk.Frame(out); r4.pack(fill=tk.X, pady=2)
            ttk.Label(r4, text="Output File:", width=14, anchor="e").pack(side=tk.LEFT)
            ttk.Entry(r4, textvariable=self.output_path, width=50).pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
            ttk.Button(r4, text="Browse...", command=self._browse_output).pack(side=tk.LEFT)

            r5 = ttk.Frame(out); r5.pack(fill=tk.X, pady=2)
            ttk.Label(r5, text="DPI:", width=14, anchor="e").pack(side=tk.LEFT)
            ttk.Spinbox(r5, from_=100, to=400, textvariable=self.dpi, width=8).pack(side=tk.LEFT, padx=4)
            ttk.Label(r5, text="(higher = better OCR, slower)").pack(side=tk.LEFT, padx=4)

            # Progress
            pf = ttk.Frame(main); pf.pack(fill=tk.X, pady=6)
            self.progress_var = tk.DoubleVar()
            ttk.Progressbar(pf, variable=self.progress_var, maximum=100).pack(
                fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 8))
            self.progress_label = ttk.Label(pf, text="Ready", width=30)
            self.progress_label.pack(side=tk.LEFT)

            # Buttons
            bf = ttk.Frame(main); bf.pack(pady=6)
            self.run_btn = ttk.Button(bf, text="Generate Report", command=self._run)
            self.run_btn.pack(side=tk.LEFT, padx=4)
            ttk.Button(bf, text="View Mapping", command=self._show_mapping).pack(side=tk.LEFT, padx=4)
            ttk.Button(bf, text="Help", command=self._show_help).pack(side=tk.LEFT, padx=4)

            # Log
            lf = ttk.LabelFrame(main, text="Processing Log", padding=4)
            lf.pack(fill=tk.BOTH, expand=True, pady=4)
            self.log_text = scrolledtext.ScrolledText(lf, height=14, font=("Courier", 9))
            self.log_text.pack(fill=tk.BOTH, expand=True)

        def _browse_pdf(self):
            p = filedialog.askopenfilename(title="Select Input PDF",
                                           filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
            if p:
                self.pdf_path.set(p)
                if not self.output_path.get():
                    d = self.report_date.get()
                    self.output_path.set(os.path.join(os.path.dirname(p),
                                                      f"Bandwidth Report (MAX) For {d}.xlsx"))

        def _browse_template(self):
            p = filedialog.askopenfilename(title="Select Template XLSX",
                                           filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
            if p:
                self.template_path.set(p)

        def _browse_output(self):
            p = filedialog.asksaveasfilename(title="Save Output As", defaultextension=".xlsx",
                                             filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
            if p:
                self.output_path.set(p)

        def _log(self, msg):
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)

        def _progress(self, cur, total, msg):
            self.progress_var.set((cur / total) * 100 if total else 0)
            self.progress_label.config(text=msg)
            self.root.update_idletasks()

        def _run(self):
            if self.running:
                return
            pdf = self.pdf_path.get()
            tpl = self.template_path.get()
            out = self.output_path.get()
            date = self.report_date.get()
            dpi = self.dpi.get()

            if not pdf or not os.path.isfile(pdf):
                messagebox.showerror("Error", "Please select a valid input PDF."); return
            if not tpl or not os.path.isfile(tpl):
                messagebox.showerror("Error", "Please select a valid template XLSX."); return
            if not out:
                messagebox.showerror("Error", "Please set an output file path."); return

            self.running = True
            self.run_btn.config(state="disabled")
            self.log_text.delete("1.0", tk.END)
            self._log(f"PDF: {pdf}")
            self._log(f"Template: {tpl}")
            self._log(f"DPI: {dpi}\n")

            def worker():
                try:
                    data = extract_all_graphs(pdf, dpi=dpi,
                        progress_cb=lambda c, t, m: self.root.after(0, self._progress, c, t, m))

                    self.root.after(0, self._log, f"\n--- Extraction Complete ---")
                    self.root.after(0, self._log, f"Matched:  {len(data['results'])}")
                    self.root.after(0, self._log, f"Unmatched: {len(data['unmatched'])}")
                    self.root.after(0, self._log, f"Could not open: {data['could_not_open']}")

                    if data["unmatched"]:
                        self.root.after(0, self._log, "\nUnmatched (may need mapping):")
                        for g in data["unmatched"]:
                            self.root.after(0, self._log,
                                f"  Pg{g['page']}: {g['title'][:65]} (max={g['max_mbps']:.2f})")

                    self.root.after(0, self._log, "\nMatched values:")
                    for ref in sorted(data["results"]):
                        info = data["results"][ref]
                        self.root.after(0, self._log,
                            f"  {ref} = {info['mbps']:.2f} Mbps  ({info['desc']})")

                    self.root.after(0, self._log, "\nGenerating report ...")
                    generate_report(tpl, data["results"], out, date)
                    self.root.after(0, self._log, f"\nSaved: {out}")
                    self.root.after(0, self._progress, 100, 100, "Done!")
                    self.root.after(0, lambda: messagebox.showinfo("Success",
                        f"Report generated!\n\nMatched: {len(data['results'])}\n"
                        f"Unmatched: {len(data['unmatched'])}\nSaved: {out}"))
                except Exception as e:
                    self.root.after(0, self._log, f"\nERROR: {e}")
                    self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
                finally:
                    self.running = False
                    self.root.after(0, lambda: self.run_btn.config(state="normal"))

            threading.Thread(target=worker, daemon=True).start()

        def _show_mapping(self):
            win = tk.Toplevel(self.root)
            win.title("Graph -> Row Mapping")
            win.geometry("750x500")
            txt = scrolledtext.ScrolledText(win, font=("Courier", 9))
            txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
            txt.insert(tk.END, "# Edit GRAPH_TO_ROW_MAP in the script to customize.\n\n")
            txt.insert(tk.END, f"{'PATTERN':<55} {'CELL':<8} DESCRIPTION\n")
            txt.insert(tk.END, "-" * 100 + "\n")
            for pat, row, desc in GRAPH_TO_ROW_MAP:
                txt.insert(tk.END, f"{pat:<55} {row:<8} {desc}\n")
            txt.insert(tk.END, "\n--- FALLBACK ---\n")
            for pat, row, desc in FALLBACK_MAP:
                txt.insert(tk.END, f"{pat:<55} {row:<8} {desc}\n")

        def _show_help(self):
            messagebox.showinfo("Help",
                "MRTG Bandwidth Report Generator\n\n"
                "1. Select the input PDF containing MRTG/Cacti graphs\n"
                "2. Select the template XLSX (previous day's report)\n"
                "3. Set the report date\n"
                "4. Click 'Generate Report'\n\n"
                "The tool uses OCR (Tesseract) to read graph statistics\n"
                "and maps them to spreadsheet rows.\n\n"
                "To customize graph-to-row mapping, edit the\n"
                "GRAPH_TO_ROW_MAP list in the Python script.\n\n"
                "CLI mode: python mrtg_bandwidth_report.py --cli --help")

    root = tk.Tk()
    MRTGApp(root)
    root.mainloop()


# =====================================================================
# SECTION 7 — DEPENDENCY CHECK + ENTRY POINT
# =====================================================================

def check_dependencies():
    errors = []
    if MISSING_DEPS:
        errors.append(f"Missing Python packages: {', '.join(MISSING_DEPS)}")
        errors.append("Install with: pip install openpyxl pdf2image pytesseract Pillow")

    if "pytesseract" not in " ".join(MISSING_DEPS):
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            errors.append("Tesseract OCR not found.")
            errors.append("  macOS:   brew install tesseract")
            errors.append("  Ubuntu:  sudo apt install tesseract-ocr")
            errors.append("  Windows: choco install tesseract")

    if "pdf2image" not in " ".join(MISSING_DEPS):
        if not shutil.which("pdftoppm"):
            errors.append("Poppler not found (needed by pdf2image).")
            errors.append("  macOS:   brew install poppler")
            errors.append("  Ubuntu:  sudo apt install poppler-utils")
            errors.append("  Windows: choco install poppler")

    return errors


def main():
    if "--cli" in sys.argv:
        errs = check_dependencies()
        if errs:
            print("DEPENDENCY ERRORS:\n  " + "\n  ".join(errs)); sys.exit(1)
        sys.argv.remove("--cli")
        run_cli()
    else:
        errs = check_dependencies()
        if errs:
            try:
                import tkinter as tk
                from tkinter import messagebox
                r = tk.Tk(); r.withdraw()
                messagebox.showerror("Missing Dependencies", "\n".join(errs))
                r.destroy()
            except Exception:
                print("DEPENDENCY ERRORS:\n  " + "\n  ".join(errs))
            sys.exit(1)
        run_gui()


if __name__ == "__main__":
    main()
