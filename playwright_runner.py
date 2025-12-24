"""
Lightweight runner to perform the Shopify Channel Performance export using Playwright.

Usage:
  1. Ensure `.env` contains SHOPIFY_STORE_SLUG, SHOPIFY_EMAIL, SHOPIFY_PASSWORD.
  2. Install deps: `python -m pip install -r requirements.txt`
  3. Install browsers: `python -m playwright install chromium`
  4. Run: `python playwright_runner.py`

This script uses the synchronous Playwright API and saves the CSV to DOWNLOAD_DIR.
"""
from pathlib import Path
import os
import time
import shutil
import json
import logging
import smtplib
import csv
from email.message import EmailMessage
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

load_dotenv()
STORE_SLUG = os.getenv('SHOPIFY_STORE_SLUG')
EMAIL = os.getenv('SHOPIFY_EMAIL')
PASSWORD = os.getenv('SHOPIFY_PASSWORD')
DOWNLOAD_DIR = Path(os.getenv('DOWNLOAD_DIR', './downloads')).resolve()
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
DEBUG_DIR = DOWNLOAD_DIR / 'debug'
DEBUG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = DEBUG_DIR / 'playwright_runner.log'

# Setup logging
logger = logging.getLogger('playwright_runner')
logger.setLevel(logging.INFO)
fmt = logging.Formatter('%(asctime)s %(levelname)s: %(message)s')
fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
fh.setFormatter(fmt)
logger.addHandler(fh)
sh = logging.StreamHandler()
sh.setFormatter(fmt)
logger.addHandler(sh)
SINCE = os.getenv('SINCE', '2025-10-01')
UNTIL = os.getenv('UNTIL', '2025-12-19')
COUNTRY = os.getenv('COUNTRY', 'US')
REPORT_URL = (
    f"https://admin.shopify.com/store/{STORE_SLUG}/marketing/reports/channels"
    f"?attributionModel=last_click_non_direct"
    f"&since={SINCE}&until={UNTIL}"
    f"&sortColumn=sessions&sortDirection=desc"
    f"&country={COUNTRY}"
)

# Storage state file (will be created after an interactive login)
STATE_FILE = Path(os.getenv('PLAYWRIGHT_STATE_FILE', 'playwright_storage_state.json'))


def send_alert_if_configured(subject: str, body: str):
    """Send an email alert if SMTP env vars are set; otherwise just log."""
    smtp_host = os.getenv('ALERT_SMTP_HOST')
    smtp_port = int(os.getenv('ALERT_SMTP_PORT', '0')) if os.getenv('ALERT_SMTP_PORT') else None
    to_addr = os.getenv('ALERT_TO')
    user = os.getenv('ALERT_SMTP_USER')
    password = os.getenv('ALERT_SMTP_PASS')
    use_tls = os.getenv('ALERT_SMTP_TLS', '1') not in ('0', 'false', 'no')

    if not smtp_host or not to_addr:
        logger.warning('Alert not configured (ALERT_SMTP_HOST or ALERT_TO missing). Skipping email alert.')
        return

    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = user or f'noreply@{smtp_host}'
        msg['To'] = to_addr
        msg.set_content(body)

        if smtp_port and smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_host, smtp_port) as s:
                if user and password:
                    s.login(user, password)
                s.send_message(msg)
        else:
            with smtplib.SMTP(smtp_host, smtp_port or 25) as s:
                if use_tls:
                    s.starttls()
                if user and password:
                    s.login(user, password)
                s.send_message(msg)
        logger.info('Sent alert email to %s via %s', to_addr, smtp_host)
    except Exception as e:
        logger.exception('Failed to send alert email: %s', e)


def run_export(headless=True, timeout_seconds=180):
    if not STORE_SLUG:
        logger.error('Missing SHOPIFY_STORE_SLUG in .env')
        return 1

    with sync_playwright() as p:
        logger.info('Launching browser (headless=%s)', headless)
        browser = p.chromium.launch(headless=headless)

        # If we have a saved storage state, load it to reuse an authenticated session
        if STATE_FILE.exists():
            try:
                with open(STATE_FILE, 'r', encoding='utf-8') as f:
                    storage = json.load(f)
                context = browser.new_context(storage_state=storage, accept_downloads=True)
                print(f'Loaded storage state from {STATE_FILE}')
            except Exception as e:
                print('Failed to load storage state, starting fresh context:', e)
                context = browser.new_context(accept_downloads=True)
        else:
            # No saved state — start an interactive context and prompt user to login
            context = browser.new_context(accept_downloads=True)

        page = context.new_page()
        try:
            page.goto(f'https://admin.shopify.com/store/{STORE_SLUG}', timeout=60000)
            logger.info('Navigated to admin start URL')

            # If no saved state, let the user perform manual login in the opened browser
            if not STATE_FILE.exists():
                logger.info('No saved Playwright storage state found; performing automated login')
                # Automated login using .env credentials
                try:
                    # Wait for email field and fill it
                    logger.info('Waiting for email input field...')
                    page.wait_for_selector('input[name="account_email"], input[id="account_email"], input[placeholder*="email" i]', timeout=30000)
                    page.fill('input[name="account_email"], input[id="account_email"], input[placeholder*="email" i]', EMAIL)
                    logger.info('Filled email: %s', EMAIL)
                    
                    # Click continue/next button
                    logger.info('Clicking continue button...')
                    try:
                        page.click('button[type="submit"], button:has-text("Continue"), button:has-text("Next")')
                    except Exception:
                        page.press('input[name="account_email"], input[id="account_email"]', 'Enter')
                    
                    # Wait for password field and fill it
                    logger.info('Waiting for password input field...')
                    page.wait_for_selector('input[name="account_password"], input[id="account_password"], input[type="password"]', timeout=30000)
                    page.fill('input[name="account_password"], input[id="account_password"], input[type="password"]', PASSWORD)
                    logger.info('Filled password')
                    
                    # Click login button
                    logger.info('Clicking login button...')
                    try:
                        page.click('button[type="submit"], button:has-text("Log in"), button:has-text("Sign in")')
                    except Exception:
                        page.press('input[type="password"]', 'Enter')
                    
                    # Check for 2FA prompt
                    logger.info('Checking for 2FA prompt...')
                    try:
                        page.wait_for_selector('input[name="two_factor_code"], input[name="otp"], input[placeholder*="code" i], input[placeholder*="2fa" i]', timeout=10000)
                        logger.info('2FA prompt detected; requesting OTP from user')
                        otp = input('Enter your Shopify 2FA code: ').strip()
                        page.fill('input[name="two_factor_code"], input[name="otp"], input[placeholder*="code" i], input[placeholder*="2fa" i]', otp)
                        page.click('button[type="submit"], button:has-text("Verify")')
                        logger.info('Submitted 2FA code')
                    except Exception as e:
                        logger.info('No 2FA prompt detected or timeout: %s', e)
                    
                    logger.info('Login automation complete; storage state will be saved after admin UI loads')
                except Exception as e:
                    logger.exception('Automated login failed: %s', e)
                    logger.warning('Falling back to manual login prompt')
                    print('\nAutomated login failed. Please log in manually in the opened browser.')
                    print('After successful login, return here and press Enter to continue.')
                    input('Press Enter after you have logged in...')
                
                # After login (automated or manual), save the storage state for reuse
                try:
                    context.storage_state(path=str(STATE_FILE))
                    logger.info('Saved storage state to %s', STATE_FILE)
                except Exception as e:
                    logger.exception('Failed to save storage state: %s', e)

            # Wait for admin UI to appear (with fallback to just waiting a bit)
            ui_loaded = False
            try:
                page.wait_for_selector("nav[aria-label='Primary'], #AppFrameMain", timeout=30000)
                ui_loaded = True
                logger.info('Admin UI detected via selector')
            except PlaywrightTimeout:
                logger.info('Admin UI selector timeout; checking if page is already loaded')
                # Sometimes the admin loads but selectors don't match; just wait and check
                try:
                    page.wait_for_url('**/admin**', timeout=30000)
                    ui_loaded = True
                    logger.info('Admin UI detected via URL')
                except Exception as e2:
                    logger.warning('Admin URL check also failed: %s; proceeding anyway', e2)
                    page.wait_for_load_state('networkidle', timeout=15000)
                    ui_loaded = True

            if not ui_loaded:
                logger.exception('Admin UI did not load within timeout (%s seconds)', timeout_seconds)
                # save debug artifacts
                png = DEBUG_DIR / 'admin_ui_timeout.png'
                html = DEBUG_DIR / 'admin_ui_timeout.html'
                try:
                    page.screenshot(path=str(png))
                    with open(html, 'w', encoding='utf-8') as f:
                        f.write(page.content())
                    logger.info('Saved debug screenshot to %s and page HTML to %s', png, html)
                except Exception as e:
                    logger.exception('Failed to save debug artifacts: %s', e)
                raise

            print('Admin UI loaded; navigating to report...')
            page.goto(REPORT_URL, timeout=60000)

            # Trigger export and wait for download
            try:
                with page.expect_download(timeout=timeout_seconds*1000) as dl_info:
                    try:
                        page.click("button:has-text('Export')")
                    except Exception:
                        page.click("button[aria-label='More actions'], button[aria-haspopup='menu']")
                        page.click("button:has-text('Export')")
                download = dl_info.value
                stamp = time.strftime('%Y%m%d_%H%M%S')
                dest = DOWNLOAD_DIR / f'shopify_channel_performance_{COUNTRY}_{SINCE}_{UNTIL}_{stamp}.csv'
                download.save_as(str(dest))
                logger.info('Downloaded and saved to %s', dest)
                # Optionally upload to Google Sheets if configured
                try:
                    if os.getenv('UPLOAD_TO_SHEET', '0').lower() in ('1', 'true', 'yes'):
                        logger.info('UPLOAD_TO_SHEET enabled; attempting to upload %s to Google Sheets', dest)
                        upload_csv_to_sheet(dest)
                except Exception as e:
                    logger.exception('Upload to Google Sheets failed: %s', e)
                    send_alert_if_configured('Shopify export upload failed', f'CSV downloaded but upload failed: {e} -- see {LOG_FILE}')
            except Exception as e:
                logger.exception('Export or download failed: %s', e)
                send_alert_if_configured('Shopify export failed', f'Export failed: {e} -- see logs: {LOG_FILE}')
                return 2
        finally:
            try:
                context.close()
                browser.close()
            except Exception:
                pass
    return 0


def upload_csv_to_sheet(csv_path: Path):
    """Upload CSV to Google Sheets using OAuth 2.0.

    Required env vars:
      - SHEET_ID : Google Sheet ID
      - SHEET_NAME : optional (default 'Sheet1')

    First run: opens browser for OAuth consent; saves token locally.
    Subsequent runs: uses saved token (no browser needed).
    """
    sheet_id = os.getenv('SHEET_ID')
    sheet_name = os.getenv('SHEET_NAME', 'Sheet1')
    token_file = DEBUG_DIR / 'token.json'

    if not sheet_id:
        logger.error('Google Sheets upload requested but SHEET_ID not set')
        raise RuntimeError('Missing SHEET_ID in .env')

    # Import heavy Google libs lazily
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except Exception as e:
        logger.exception('Google OAuth libraries not installed: %s', e)
        raise

    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None

    # Load existing token if available
    if token_file.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_file), scopes)
            logger.info('Loaded cached OAuth token from %s', token_file)
        except Exception as e:
            logger.warning('Failed to load token, will re-authenticate: %s', e)
            creds = None

    # If no valid token, perform OAuth flow (opens browser on first run)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logger.info('Token expired; refreshing...')
            creds.refresh(Request())
        else:
            logger.info('No valid token found; starting OAuth flow (browser will open)...')
            # Use a generic installed app flow; Google will prompt user to select/authenticate
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', scopes, redirect_uri='http://localhost:8080'
            )
            creds = flow.run_local_server(port=8080, open_browser=True)
            logger.info('OAuth authentication complete; saving token...')

        # Save token for future runs
        try:
            with open(token_file, 'w', encoding='utf-8') as f:
                f.write(creds.to_json())
            logger.info('Saved OAuth token to %s', token_file)
        except Exception as e:
            logger.exception('Failed to save token: %s', e)

    # Build Sheets API service and upload
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Read CSV into list of rows
        rows = []
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row_idx, r in enumerate(reader):
                if row_idx == 0:
                    # Header row: add "Upload_Date" column
                    r = list(r) + ['Upload_Date']
                else:
                    # Data rows: add today's date
                    r = list(r) + [time.strftime('%Y-%m-%d')]
                rows.append(r)

        # If SHEET_NAME is explicitly configured (not default), use it directly
        # Otherwise, get the sheet list to find the first sheet
        target_sheet = sheet_name
        if sheet_name == 'Sheet1':  # Default value, try to find actual sheet
            try:
                sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
                sheets = sheet_metadata.get('sheets', [])
                if sheets:
                    target_sheet = sheets[0]['properties']['title']
                    logger.info('Found sheet: %s (using instead of default Sheet1)', target_sheet)
            except Exception as e:
                logger.warning('Could not fetch sheet list; using configured name %s: %s', sheet_name, e)
        else:
            logger.info('Using explicitly configured sheet name: %s', sheet_name)

        # Write values starting at A1 (clears existing and writes fresh)
        range_spec = f'{target_sheet}!A1'
        body = {'values': rows}
        result = service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=range_spec, valueInputOption='RAW', body=body
        ).execute()
        logger.info('Uploaded %s rows to sheet %s (%s)', result.get('updatedRows', len(rows)), target_sheet, sheet_id)
    except Exception as e:
        logger.exception('Failed to upload CSV to Google Sheets: %s', e)
        raise


if __name__ == '__main__':
    headless_flag = os.getenv('CHROME_HEADLESS', '1')
    is_headless = not (headless_flag.lower() in ('0', 'false', 'no'))
    try:
        exit_code = run_export(headless=is_headless, timeout_seconds=int(os.getenv('LONG_TIMEOUT', '180')))
    except Exception as e:
        logger.exception('Runner crashed with exception: %s', e)
        send_alert_if_configured('Shopify export runner crashed', f'Exception: {e} -- check {LOG_FILE}')
        raise
    finally:
        logger.info('Run finished with exit code %s', locals().get('exit_code', 'unknown'))
    raise SystemExit(exit_code)
