#!/usr/bin/env python3
"""
Automated MRTG Bandwidth Report Pipeline
=========================================
1. Opens Outlook web via Playwright Chromium
2. Logs in and finds the latest "BSCPLC MRTG Report" email
3. Opens Outlook's Print preview (three-dot → Print) and saves as PDF
4. Runs the OCR report generator on the PDF
5. Emails the .xlsx report to the recipient

Requires: pip install playwright python-dotenv openpyxl pdf2image pytesseract Pillow
Then run: playwright install chromium
"""

import os
import sys
import time
import smtplib
import logging
import subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
load_dotenv(Path(__file__).parent / ".env")

OUTLOOK_EMAIL = os.environ["OUTLOOK_EMAIL"]
OUTLOOK_PASSWORD = os.environ["OUTLOOK_PASSWORD"]
REPORT_RECIPIENT = os.environ["REPORT_RECIPIENT"]
TEMPLATE_PATH = os.environ["TEMPLATE_PATH"]

OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/"
EMAIL_SUBJECT = "BSCPLC MRTG Report"

SCRIPT_DIR = Path(__file__).parent
PDF_DIR = SCRIPT_DIR / "pdfs"
REPORT_DIR = SCRIPT_DIR / "reports"
BROWSER_DATA_DIR = SCRIPT_DIR / ".browser_data"  # persistent session/cookies

# Ensure Tesseract and Poppler are on PATH
_TESSERACT_DIR = r"C:\Program Files\Tesseract-OCR"
_POPPLER_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", ""),
    r"Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe",
    r"poppler-25.07.0\Library\bin",
)
for _dir in [_TESSERACT_DIR, _POPPLER_DIR]:
    if os.path.isdir(_dir) and _dir not in os.environ.get("PATH", ""):
        os.environ["PATH"] = _dir + os.pathsep + os.environ.get("PATH", "")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("auto_report")


# ---------------------------------------------------------------------------
# Step 1: Login to Outlook Web
# ---------------------------------------------------------------------------
def login_outlook(page):
    """Log into Outlook web with email and password.

    Session cookies are saved/restored via Playwright storage state,
    so the login page may auto-redirect without needing credentials.
    """
    log.info("Navigating to Outlook...")
    page.goto("https://login.microsoftonline.com/")
    page.wait_for_load_state("load")
    time.sleep(2)

    # Enter email (skip if session auto-redirected past this step)
    email_field = page.locator('input[type="email"]')
    if email_field.is_visible(timeout=5000):
        log.info("Entering email...")
        page.fill('input[type="email"]', OUTLOOK_EMAIL)
        page.click('input[type="submit"]')
        page.wait_for_load_state("load")
        time.sleep(3)
    else:
        log.info("Email field not shown (session may be restored).")

    # Enter password (skip if SSO or session handles it)
    password_field = page.locator('input[type="password"]')
    if password_field.is_visible(timeout=5000):
        log.info("Entering password...")
        page.fill('input[type="password"]', OUTLOOK_PASSWORD)
        page.click('input[type="submit"]')
        page.wait_for_load_state("load")
        time.sleep(3)
    else:
        log.info("Password field not shown (SSO or session restored).")

    # "Stay signed in?" prompt — click Yes (keeps session for future runs)
    try:
        page.click('input[id="idSIButton9"]', timeout=5000)
    except Exception:
        try:
            page.click('input[id="idBtn_Back"]', timeout=3000)
        except Exception:
            pass

    # Wait for M365 landing page
    log.info("Waiting for Microsoft 365 to load...")
    page.wait_for_load_state("load", timeout=60000)
    time.sleep(5)

    # Navigate to Outlook Mail
    log.info("Navigating to Outlook Mail...")
    page.goto(OUTLOOK_URL, wait_until="domcontentloaded", timeout=60000)
    time.sleep(10)
    log.info(f"Current URL: {page.url}")
    log.info("Outlook loaded successfully.")


# ---------------------------------------------------------------------------
# Step 2: Find and open the latest MRTG report email
# ---------------------------------------------------------------------------
def find_and_open_email(page):
    """Search for the latest BSCPLC MRTG Report email and open it."""
    log.info(f"Searching for email with subject: {EMAIL_SUBJECT}")

    # Debug: save screenshot to diagnose search box issues
    page.screenshot(path=str(Path(__file__).parent / "debug_before_search.png"))

    # Use Outlook search — try multiple selectors for different Outlook versions
    search_box = None
    for selector in [
        'input[aria-label="Search"]',
        '[id="topSearchInput"]',
        'input[placeholder*="Search"]',
        '[role="search"] input',
        'button[aria-label="Search"]',
    ]:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=5000):
                search_box = loc
                log.info(f"Found search box: {selector}")
                break
        except Exception:
            continue

    if not search_box:
        # Try clicking a search button/icon that expands into a search input
        try:
            page.locator('button[aria-label="Search"]').first.click(timeout=5000)
            time.sleep(2)
            search_box = page.locator('input[aria-label="Search"], input[placeholder*="Search"]').first
        except Exception:
            pass

    if not search_box:
        raise RuntimeError("Could not find Outlook search box. See debug_before_search.png")

    search_box.click()
    search_box.fill(EMAIL_SUBJECT)
    page.keyboard.press("Enter")
    time.sleep(5)

    # Click the first (most recent) email in the results
    log.info("Opening the latest matching email...")
    first_email = page.locator(
        f'[aria-label*="{EMAIL_SUBJECT}"], '
        f'span:has-text("{EMAIL_SUBJECT}")'
    ).first
    first_email.click()
    time.sleep(5)
    log.info("Email opened.")


# ---------------------------------------------------------------------------
# Step 3: Print email to PDF via Outlook's Print preview
# ---------------------------------------------------------------------------
def print_email_to_pdf(page, context, pdf_path: Path):
    """Click three-dot menu → Print in Outlook, then save the print preview as PDF."""
    log.info("Opening Print preview...")

    # Scroll the email body to force-load ALL embedded graph images.
    # Outlook web lazy-loads images — we must scroll through the entire
    # email content to trigger loading of every graph.
    # Use JavaScript to scroll the email reading pane container directly,
    # which is more reliable than keyboard PageDown on the wrong element.
    log.info("Scrolling email to load all graph images...")
    page.evaluate("""() => {
        // Find the email body scrollable container
        const containers = [
            document.querySelector('[role="main"] [data-app-section="ConversationContainer"]'),
            document.querySelector('[role="main"] .customScrollBar'),
            document.querySelector('[role="main"]'),
            document.querySelector('.ReadingPaneContainerV2'),
        ].filter(Boolean);
        const el = containers[0] || document.scrollingElement || document.documentElement;
        // Scroll to bottom in steps to trigger lazy image loading
        const totalHeight = el.scrollHeight;
        const step = 500;
        for (let y = 0; y <= totalHeight; y += step) {
            el.scrollTop = y;
        }
        // Back to top
        el.scrollTop = 0;
    }""")
    time.sleep(3)

    # Also do keyboard scrolling as a fallback (covers cases where
    # the JS scroll target wasn't the right container)
    for _ in range(30):
        page.keyboard.press("PageDown")
        time.sleep(0.3)
    page.keyboard.press("Home")
    time.sleep(2)

    # Wait for all images to finish loading
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        time.sleep(3)  # fallback wait

    # Click three-dot "More actions" button near Reply/Reply all/Forward
    # Try multiple selectors
    clicked = False
    for selector in [
        'button[aria-label="More items"]',
        'button[aria-label="More actions"]',
        'button[aria-label="More mail actions"]',
        'button[title="More actions"]',
    ]:
        try:
            btn = page.locator(selector).first
            if btn.is_visible(timeout=2000):
                btn.click()
                clicked = True
                log.info(f"Clicked: {selector}")
                break
        except Exception:
            continue

    if not clicked:
        page.screenshot(path=str(SCRIPT_DIR / "debug_no_menu.png"))
        log.warning("Could not find three-dot menu. Will save page directly as PDF.")

    time.sleep(1)

    # Click "Print" in the dropdown menu — listen for popup before clicking
    popup_page = None
    if clicked:
        time.sleep(1)
        try:
            with context.expect_page(timeout=10000) as new_page_info:
                print_item = page.locator(
                    '[role="menuitem"]:has-text("Print"), '
                    'button:has-text("Print")'
                ).first
                print_item.click(timeout=5000)
                log.info("Clicked Print menu item.")
            popup_page = new_page_info.value
            popup_page.wait_for_load_state("domcontentloaded", timeout=30000)
            log.info(f"Print preview popup opened: {popup_page.url}")
        except Exception as e:
            log.info(f"No popup detected ({e}). Will use current page.")

    time.sleep(3)

    # Determine which page to save as PDF
    print_page = popup_page
    if not print_page:
        all_pages = context.pages
        if len(all_pages) > 1:
            print_page = all_pages[-1]

    if print_page and print_page != page:
        print_page.wait_for_load_state("domcontentloaded", timeout=30000)
        time.sleep(5)

        # Scroll to load all content in print preview
        for _ in range(30):
            print_page.keyboard.press("PageDown")
            time.sleep(0.3)
        print_page.keyboard.press("Home")
        time.sleep(2)

        log.info(f"Print preview URL: {print_page.url}")
        print_page.screenshot(path=str(SCRIPT_DIR / "debug_print_preview.png"), full_page=True)

        # Use Chromium's page.pdf() on the print preview — clean email content only
        print_page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            margin={"top": "0.4in", "bottom": "0.4in", "left": "0.4in", "right": "0.4in"},
        )
        log.info(f"PDF saved from print preview: {pdf_path}")
        print_page.close()
    else:
        # Fallback: save the current page as PDF directly
        log.info("No print preview popup detected. Saving current page as PDF...")
        page.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            margin={"top": "0.4in", "bottom": "0.4in", "left": "0.4in", "right": "0.4in"},
        )
        log.info(f"PDF saved from main page: {pdf_path}")

    return pdf_path


# ---------------------------------------------------------------------------
# Step 4: Run the OCR report generator
# ---------------------------------------------------------------------------
def generate_report(pdf_path: Path, date_str: str):
    """Run the MRTG bandwidth report generator on the saved PDF."""
    log.info("Running OCR report generator...")
    output_name = f"Bandwidth Report (MAX) For {date_str}.xlsx"
    output_path = REPORT_DIR / output_name

    # Remove Excel lock files that block writing (left by crashed Excel or previous runs)
    lock_file = REPORT_DIR / f"~${output_name}"
    if lock_file.exists():
        log.info(f"Removing stale lock file: {lock_file}")
        lock_file.unlink()

    # Run OCR subprocess. Use subprocess.DEVNULL for stdin and pipe stderr only.
    # stdout is inherited (printed to console) to avoid pipe buffer deadlock
    # that occurs with capture_output=True on large OCR output (~25 pages).
    result = subprocess.run(
        [
            sys.executable,
            str(SCRIPT_DIR / "mrtg_bandwidth_report.py"),
            "--cli",
            "--pdf", str(pdf_path),
            "--template", TEMPLATE_PATH,
            "--date", date_str,
            "--output", str(output_path),
        ],
        stdin=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
        timeout=240,  # 4 minutes max for OCR processing
    )

    if result.returncode != 0:
        log.error(f"Report generation failed:\nSTDERR: {result.stderr}")
        raise RuntimeError(f"Report generation failed: {result.stderr}")

    log.info(f"Report generated: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# Step 5: Email the report
# ---------------------------------------------------------------------------
def email_report(report_path: Path, date_str: str):
    """Send the generated report via Outlook SMTP to the recipient."""
    log.info(f"Emailing report to {REPORT_RECIPIENT}...")

    msg = MIMEMultipart()
    msg["From"] = OUTLOOK_EMAIL
    msg["To"] = REPORT_RECIPIENT
    msg["Subject"] = f"Bandwidth Report (MAX) For {date_str}"

    body = (
        f"Hi,\n\n"
        f"Please find the attached Bandwidth Report (MAX) for {date_str}.\n\n"
        f"This report was generated automatically.\n\n"
        f"Regards,\n"
        f"BSCPLC IIG NOC"
    )
    msg.attach(MIMEText(body, "plain"))

    # Attach the xlsx file
    with open(report_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={report_path.name}",
    )
    msg.attach(part)

    # Send via Outlook SMTP
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
        server.send_message(msg)

    log.info("Report emailed successfully.")


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def main():
    # The MRTG email contains graphs for the PREVIOUS day (24h ending ~midnight).
    # Use yesterday's date for the PDF filename and report title.
    yesterday = datetime.now() - timedelta(days=1)
    report_date = yesterday.strftime("%d %B %Y")
    log.info(f"=== MRTG Auto Report Pipeline — report for {report_date} ===")

    # Create output directories
    PDF_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)

    pdf_path = PDF_DIR / f"MRTG_Report_{yesterday.strftime('%Y-%m-%d')}.pdf"

    # --- Browser automation using Playwright Chromium ---
    # Save/restore session cookies so login is skipped on subsequent runs.
    storage_state_path = BROWSER_DATA_DIR / "storage_state.json"
    BROWSER_DATA_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx_opts = dict(
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"
            ),
        )
        # Restore previous session if available
        if storage_state_path.exists():
            ctx_opts["storage_state"] = str(storage_state_path)
        context = browser.new_context(**ctx_opts)
        page = context.new_page()

        login_outlook(page)
        find_and_open_email(page)
        print_email_to_pdf(page, context, pdf_path)

        # Save session cookies for next run, then close browser quickly
        try:
            context.storage_state(path=str(storage_state_path))
        except Exception as e:
            log.warning(f"Could not save session state: {e}")
        # Force-close all pages first to speed up browser.close()
        for p_page in context.pages:
            try:
                p_page.close()
            except Exception:
                pass
        browser.close()
        log.info("Browser closed.")

    # --- Report generation ---
    report_path = generate_report(pdf_path, report_date)

    # --- Email delivery ---
    email_report(report_path, report_date)

    log.info("=== Pipeline complete ===")


if __name__ == "__main__":
    main()
