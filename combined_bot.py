"""
Combined Bot
============
Runs both workflows sequentially using a single .env file:

  1. IREC Holdings Bot  — logs into evident.app, downloads CSVs for each account
  2. TIGR Registry Bot  — logs into tigrsregistry.apx.com, downloads Excel sub-account reports
  3. Combine Output     — merges all downloaded files into a single Excel using the template

.env keys used
--------------
  IREC_EMAIL          Email for evident.app
  IREC_PASSWORD       Password for evident.app
  HEADLESS            (optional) true/false — default false

  TIGR_MYUSERNAME     Username for TIGR "MY" account
  TIGR_MYPASSWORD     Password for TIGR "MY" account
  TIGR_SGUSERNAME     (optional) Username for TIGR "SG" account
  TIGR_SGPASSWORD     (optional) Password for TIGR "SG" account

Requirements:
  pip install playwright python-dotenv openpyxl pandas
  python -m playwright install chromium

Output Template Columns (hardcoded):
  Registry | Asset ID | Asset | Country | Fuel | Tech | TIGR Vintage |
  Period Start | Period End | Quantity | Transferor | Feed-in Tariff |
  Sub-Account | Sub-Account ID
"""

import os
import sys
import time
import logging
import traceback
import calendar
from pathlib import Path
from datetime import datetime, date

# ── Keep window open on crash ─────────────────────────────────────────────────
def _fatal(msg=""):
    if msg:
        print(msg)
    print("\n" + "=" * 60)
    _pause("  Press Enter to close this window...")
    sys.exit(1)

sys.excepthook = lambda t, v, tb: (
    print("\n[CRASH] An unexpected error occurred:\n"),
    traceback.print_exception(t, v, tb),
    _pause("\n  Press Enter to close this window..."),
)

# ── Imports ───────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
except ImportError:
    _fatal(
        "[ERROR] 'python-dotenv' is not installed.\n"
        "  Run:  pip install playwright python-dotenv openpyxl pandas\n"
        "  Then: python -m playwright install chromium"
    )

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    _fatal(
        "[ERROR] 'playwright' is not installed.\n"
        "  Run:  pip install playwright python-dotenv openpyxl pandas\n"
        "  Then: python -m playwright install chromium"
    )

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    _fatal(
        "[ERROR] 'openpyxl' or 'pandas' is not installed.\n"
        "  Run:  pip install openpyxl pandas"
    )

# ── Shared config ─────────────────────────────────────────────────────────────
load_dotenv()

DOWNLOAD_DIR   = Path("downloads")
LOG_DIR        = Path("logs")
SCREENSHOT_DIR = Path("screenshots")
OUTPUT_DIR     = Path("output")
for d in (DOWNLOAD_DIR, LOG_DIR, SCREENSHOT_DIR, OUTPUT_DIR):
    d.mkdir(exist_ok=True)

HEADLESS = os.getenv("HEADLESS", "true").lower() == "true"
CI      = os.getenv("CI", "false").lower() == "true"  # set automatically by GitHub Actions

def _pause(msg="  Press Enter to close this window..."):
    """Pause for user input on desktop; skip silently in CI."""
    if not CI:
        input(msg)

# ── Hardcoded Output Template Columns ─────────────────────────────────────────
TEMPLATE_COLUMNS = [
    "Registry",
    "Asset ID",
    "Asset",
    "Country",
    "Fuel",
    "Tech",
    "Period Start",
    "Period End",
    "Quantity",
    "Transferor",
    "Feed-in Tariff",
    "Sub-Account",
    "Sub-Account ID",
]

# ── IREC column mapping ───────────────────────────────────────────────────────
# IREC CSV columns → Template columns
# Device       → Asset ID
# Device Name  → Asset
# Fuel type    → Fuel
# Technology   → Tech
# Country      → Country (first 2 chars)
# Volume       → Quantity
IREC_COLUMN_MAP = {
    "Device":       "Asset ID",
    "Device Name":  "Asset",
    "Fuel type":    "Fuel",
    "Technology":   "Tech",
    "Country":      "Country",
    "Volume":       "Quantity",
}

# ── TIGR country name → ISO 2-letter code ────────────────────────────────────
TIGR_COUNTRY_MAP = {
    "viet nam":   "VN",
    "vietnam":    "VN",
    "malaysia":   "MY",
    "singapore":  "SG",
}

# ─────────────────────────────────────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────────────────────────────────────
def setup_logging() -> logging.Logger:
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"combined_bot_{ts}.log"
    fmt      = "%(asctime)s  [%(levelname)-8s]  %(message)s"
    logging.basicConfig(
        level=logging.DEBUG,
        format=fmt,
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger = logging.getLogger("combined_bot")
    logger.info("Log file: %s", log_file)
    return logger


# ─────────────────────────────────────────────────────────────────────────────
#  SHARED HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def banner(msg):
    print(f"\n{'═' * 60}\n  {msg}\n{'═' * 60}")

def step(msg):
    print(f"\n{'─' * 60}\n  {msg}\n{'─' * 60}")

def screenshot(page, name: str, logger: logging.Logger):
    ts   = datetime.now().strftime("%H%M%S")
    path = SCREENSHOT_DIR / f"{ts}_{name}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        logger.debug("Screenshot → %s", path)
    except Exception as exc:
        logger.warning("Screenshot failed '%s': %s", name, exc)

def all_frames(page):
    return [page.main_frame] + [f for f in page.frames if f != page.main_frame]

def click_in_any_frame(page, selectors, label, logger, timeout=6_000):
    frames = all_frames(page)
    logger.debug("Searching %d frame(s) for: %s", len(frames), label)
    for sel in selectors:
        for frame in frames:
            try:
                el = frame.wait_for_selector(sel, timeout=timeout, state="visible")
                if el:
                    logger.debug("  ✓ '%s' found in frame '%s' via: %s", label, frame.url[:80], sel)
                    el.click()
                    return frame, el
            except PWTimeout:
                continue
            except Exception as exc:
                logger.debug("  Error sel='%s' frame='%s': %s", sel, frame.url[:60], exc)
    return None, None


# ═════════════════════════════════════════════════════════════════════════════
#  BOT 1 — IREC Holdings (evident.app)
# ═════════════════════════════════════════════════════════════════════════════

IREC_LOGIN_URL = "https://evident.app/login"

# (search_term, account_code, account_name_link, file_label, click_name_header)
IREC_TARGET_ACCOUNTS = [
    ("SAXONTRADE01",  "SAXONTRADE01", "SaxonTrade01",             "SaxonTrade01",                  False),
    ("FRG",           "T0LFRGI4",     "Saxon Renewables Pte Ltd", "SaxonRenewablesPteLtd_T0LFRGI4", False),
    ("t9mq",          "T9MQ5BR4",     "SAXONTRADE02",              "TADAU_T9MQ",                     False),
    ("T0CV1BYF",      "T0CV1BYF",     "IS Energy Sdn Bhd",         "ISENERGY_T0CV1BYF",              False),
]

# File label prefix that identifies the TADAU account download
TADAU_LABEL_PREFIX = "TADAU_"

# File label prefix that identifies the IS Energy account download
ISENERGY_LABEL_PREFIX = "ISENERGY_"


def irec_validate_env():
    email    = os.getenv("IREC_EMAIL")
    password = os.getenv("IREC_PASSWORD")
    missing  = [k for k, v in {"IREC_EMAIL": email, "IREC_PASSWORD": password}.items() if not v]
    if missing:
        _fatal(
            f"\n[ERROR] Missing IREC credential(s) in .env: {', '.join(missing)}\n"
            "  Add to your .env file:\n"
            "    IREC_EMAIL=you@example.com\n"
            "    IREC_PASSWORD=your_password\n"
        )
    return email, password


def irec_login(page, email, password):
    """
    New multi-step Xpansiv/Auth0 login for IREC (evident.app).

    The login page (auth.xpansiv.com) now uses a stepped flow:
      Step 1 — Enter email in  input[name="username"]  then click Continue
      Step 2 — Enter password  in  input[type="password"]  then click Continue
      Step 3 — Optional authorise/confirm page → click Continue

    The email field is  name="username" type="text"  (NOT type="email"),
    which is why the old selector failed.
    """
    step("IREC — Logging in (new Xpansiv multi-step) ...")
    page.goto(IREC_LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(2)
    print(f"  [i] Login page loaded → {page.url}")

    # ── Step 1: enter email / username ────────────────────────────────────────
    # Auth0 ULP uses name="username" with type="text" — NOT type="email"
    email_sel = (
        'input[name="username"],'
        'input[type="email"],'
        'input[name="email"],'
        'input[placeholder*="email" i],'
        'input[placeholder*="Email" i]'
    )
    try:
        ef = page.locator(email_sel).first
        ef.wait_for(state="visible", timeout=15_000)
        ef.click()
        ef.fill("")
        ef.type(email, delay=60)
        print("  [✓] Email entered")
    except PWTimeout:
        _fatal(
            "[ERROR] IREC email field not found.\n"
            "  The login page may have changed — check the URL and page source.\n"
            f"  Current URL: {page.url}"
        )

    # Click the Continue button
    # The actual button is: <button type="submit" name="action" value="default"
    #   data-action-button-primary="true">Continue</button>
    try:
        cont = page.locator(
            'button[data-action-button-primary="true"],'
            'button[name="action"][value="default"],'
            'button._button-login-id,'
            'button[type="submit"]'
        ).first
        cont.wait_for(state="visible", timeout=8_000)
        cont.click()
        print("  [✓] Continue clicked (after email)")
    except PWTimeout:
        print("  [!] Continue button not found — pressing Enter")
        ef.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass
    time.sleep(1.5)
    print(f"  [i] After email step → {page.url}")

    # ── Step 2: enter password ────────────────────────────────────────────────
    try:
        pf = page.locator('input[type="password"]').first
        pf.wait_for(state="visible", timeout=15_000)
        pf.click()
        pf.fill("")
        pf.type(password, delay=60)
        print("  [✓] Password entered")
    except PWTimeout:
        _fatal(
            "[ERROR] IREC password field not found after email step.\n"
            f"  Current URL: {page.url}"
        )

    # Click Continue / Log In if available — otherwise just press Enter
    try:
        cont2 = page.locator(
            'button[data-action-button-primary="true"],'
            'button[name="action"][value="default"],'
            'button._button-login-id,'
            'button[type="submit"]'
        ).first
        cont2.wait_for(state="visible", timeout=3_000)
        cont2.click()
        print("  [✓] Continue clicked (after password)")
    except PWTimeout:
        print("  [i] No Continue button on password page — pressing Enter")
        pf.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=20_000)
    except PWTimeout:
        pass
    time.sleep(1.5)
    print(f"  [i] After password step → {page.url}")

    # ── Step 3: optional authorise / confirm page ─────────────────────────────
    try:
        body_text = page.inner_text("body").lower()
    except Exception:
        body_text = ""

    confirm_words = ("accept", "authorize", "authorise", "allow", "confirm")
    if any(w in body_text for w in confirm_words) and "password" not in body_text:
        print("  [i] Confirmation page detected — clicking Continue")
        try:
            conf_btn = page.locator(
                'button:has-text("Continue"),'
                'button:has-text("Accept"),'
                'button:has-text("Authorize"),'
                'button:has-text("Authorise"),'
                'button:has-text("Allow"),'
                'button[type="submit"],'
                'input[type="submit"]'
            ).first
            conf_btn.wait_for(state="visible", timeout=8_000)
            conf_btn.click()
            print("  [✓] Confirmation accepted")
            try:
                page.wait_for_load_state("networkidle", timeout=20_000)
            except PWTimeout:
                pass
            time.sleep(1.5)
        except PWTimeout:
            print("  [!] Confirmation button not found — continuing anyway")

    # ── Verify login succeeded ────────────────────────────────────────────────
    # Check for error messages before declaring success
    try:
        body_text = page.inner_text("body").lower()
    except Exception:
        body_text = ""

    for bad in ("wrong email or password", "invalid credentials", "login failed",
                "access denied", "authentication failed"):
        if bad in body_text:
            _fatal(
                f"[FAILED] IREC login rejected: page says '{bad}'\n"
                "  Check IREC_EMAIL / IREC_PASSWORD in .env."
            )

    # If still on an auth/login URL after all steps, something went wrong
    current = page.url
    if "auth." in current and "/login" in current:
        err_el = page.locator('[class*="error"],[class*="alert"],[role="alert"]')
        msg = err_el.first.text_content().strip() if err_el.count() > 0 else "unknown error"
        _fatal(
            f"[FAILED] IREC login did not redirect away from auth page.\n"
            f"  URL: {current}\n  Message: {msg}\n"
            "  Check IREC_EMAIL / IREC_PASSWORD in .env."
        )

    print(f"  [✓] Logged in → {page.url}")


def irec_go_to_accounts(page):
    step("IREC — Navigating to Accounts ...")
    try:
        link = page.locator('nav a:has-text("Accounts"), aside a:has-text("Accounts"), a:has-text("Accounts")').first
        link.wait_for(state="visible", timeout=10_000)
        link.click()
        page.wait_for_selector('text="Account Code"', timeout=15_000)
        print("  [✓] Accounts page loaded")
    except PWTimeout:
        _fatal("[ERROR] Could not load the IREC Accounts page.")


def irec_search_accounts(page, search_term, account_code):
    step(f"IREC — Searching for: '{search_term}'")
    try:
        sb = page.locator('input[placeholder*="Search" i], input[type="search"]').first
        sb.wait_for(state="visible", timeout=10_000)
        sb.click()
        sb.press("Control+a")
        sb.press("Backspace")
        sb.type(search_term, delay=80)
        print(f"  [✓] Typed '{search_term}'")
    except PWTimeout:
        _fatal("[ERROR] IREC search box not found.")

    try:
        page.wait_for_selector(f'text="{account_code}"', timeout=15_000)
        page.wait_for_timeout(1_000)
        print(f"  [✓] Account code '{account_code}' visible")
    except PWTimeout:
        _fatal(f"[ERROR] Account '{account_code}' not found after searching '{search_term}'.")


def irec_click_account_name(page, account_name, click_name_header=False):
    step(f"IREC — Clicking account: '{account_name}'")
    if click_name_header:
        try:
            lnk = page.locator('a:has-text("Account Name"), span:has-text("Account Name"), th:has-text("Account Name")').first
            lnk.wait_for(state="visible", timeout=10_000)
            lnk.click()
            page.wait_for_load_state("networkidle", timeout=15_000)
            print("  [✓] Clicked 'Account Name' header")
        except PWTimeout:
            _fatal("[ERROR] 'Account Name' link/element not found.")

    try:
        nl = page.locator(f'a:has-text("{account_name}")').first
        nl.wait_for(state="visible", timeout=10_000)
        nl.click()
        print(f"  [✓] Clicked '{account_name}'")
    except PWTimeout:
        _fatal(f"[ERROR] Account name link '{account_name}' not found.")

    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
        print(f"  [✓] Account detail loaded → {page.url}")
    except PWTimeout:
        _fatal(f"[ERROR] Account detail page did not load for '{account_name}'.")


def irec_click_view_holdings(page, account_name):
    step("IREC — Clicking 'View Account Holdings' ...")
    try:
        btn = page.locator(
            'a:has-text("View Account Holdings"), button:has-text("View Account Holdings"), '
            'a:has-text("Account Holdings"), button:has-text("Account Holdings")'
        ).first
        btn.wait_for(state="visible", timeout=15_000)
        btn.click()
        print("  [✓] 'View Account Holdings' clicked")
    except PWTimeout:
        _fatal(f"[ERROR] 'View Account Holdings' not found for '{account_name}'.")

    try:
        page.wait_for_load_state("networkidle", timeout=15_000)
        print(f"  [✓] Holdings page loaded → {page.url}")
    except PWTimeout:
        _fatal(f"[ERROR] Holdings page did not load for '{account_name}'.")


def irec_download_csv(page, search_term, account_code, account_name, label, click_name_header=False):
    irec_go_to_accounts(page)
    irec_search_accounts(page, search_term, account_code)
    irec_click_account_name(page, account_name, click_name_header=click_name_header)
    irec_click_view_holdings(page, account_name)

    step("IREC — Waiting for Account Holdings table ...")
    try:
        page.wait_for_selector('text="Account Holdings"', timeout=15_000)
        page.wait_for_load_state("networkidle", timeout=15_000)
        page.wait_for_timeout(1_500)
        print("  [✓] Holdings table ready")
    except PWTimeout:
        _fatal(f"[ERROR] Holdings table did not load for '{account_name}'.")

    step(f"IREC — Downloading CSV for '{account_name}' ...")
    try:
        csv_btn = page.locator('button:has-text("CSV"), a:has-text("CSV")').first
        csv_btn.wait_for(state="visible", timeout=15_000)
        with page.expect_download(timeout=30_000) as dl_info:
            csv_btn.click()
        save_path = DOWNLOAD_DIR / f"{label}_holdings.csv"
        if save_path.exists():
            save_path.unlink()
            print(f"  [i] Replaced existing file: {save_path.name}")
        dl_info.value.save_as(str(save_path))
        print(f"  [✓] CSV saved → {save_path}")
        return save_path
    except PWTimeout:
        _fatal(f"[ERROR] CSV button not found or download timed out for '{account_name}'.")


def irec_logout(page):
    step("IREC — Logging out ...")
    try:
        btn = page.locator('a:has-text("Log Out"), button:has-text("Log Out")').first
        btn.wait_for(state="visible", timeout=10_000)
        btn.click()
        page.wait_for_url(lambda url: "/login" in url or url.endswith("/"), timeout=10_000)
        print("  [✓] Logged out")
    except PWTimeout:
        print("  [!] Log Out button not found — skipping.")


def _browser_args() -> list:
    """
    Return Chromium launch args that work on both a local desktop and a
    headless CI/CD runner (GitHub Actions, etc.).

    Key flags
    ---------
    --no-sandbox              Required on Linux CI (no user namespace support).
    --disable-setuid-sandbox  Companion to --no-sandbox.
    --disable-dev-shm-usage   /dev/shm is tiny on many CI runners; use /tmp instead.
    --disable-gpu             No GPU on headless servers.
    --password-store=basic    Tells Chromium NOT to use the OS keychain / kwallet /
                              gnome-keyring / macOS Keychain for storing passwords.
                              This is what causes the "confirm password" popup on
                              laptops and crashes on CI.
    --use-mock-keychain       macOS: use an in-memory keychain instead of the
                              system keychain — eliminates the confirmation dialog.
    --disable-features=...    Disable PasswordManager & AutofillServerCommunication
                              so Chromium never tries to save/sync credentials.
    --start-maximized         Only useful in headed mode; harmless in headless mode.
    """
    return [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--password-store=basic",
        "--use-mock-keychain",
        "--disable-features=PasswordManager,AutofillServerCommunication",
        "--start-maximized",
    ]


def run_irec(logger):
    banner("BOT 1 — IREC Holdings (evident.app)")
    email, password = irec_validate_env()

    saved_files = []
    _ci_args = _browser_args()
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=_ci_args,
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        )
        page    = context.new_page()

        irec_login(page, email, password)

        for search_term, account_code, account_name, label, click_name_header in IREC_TARGET_ACCOUNTS:
            try:
                path = irec_download_csv(
                    page, search_term, account_code, account_name, label, click_name_header
                )
                saved_files.append((account_name, path, None))
            except SystemExit:
                saved_files.append((account_name, None, "fatal error (see above)"))

        irec_logout(page)
        browser.close()

    return saved_files


# ═════════════════════════════════════════════════════════════════════════════
#  BOT 2 — TIGR Registry (tigrsregistry.apx.com)
# ═════════════════════════════════════════════════════════════════════════════

TIGR_URL          = "https://tigrsregistry.apx.com/mymodule/mypage.asp"
TIGR_LOGIN_BUTTON = "https://tigrsregistry.apx.com"   # landing page with "LOGIN TO REGISTRY"
TIGR_TIMEOUT      = 30_000
TIGR_NAV_TO       = 60_000
TIGR_SLOW_MO      = 300


def tigr_load_env(logger):
    account_defs = [
        ("MY Account", "TIGR_MYUSERNAME", "TIGR_MYPASSWORD"),
        ("SG Account", "TIGR_SGUSERNAME", "TIGR_SGPASSWORD"),
    ]
    accounts = []
    for label, u_var, p_var in account_defs:
        u = os.getenv(u_var, "").strip()
        p = os.getenv(p_var, "").strip()
        if u and p:
            accounts.append((label, u, p))
            logger.info("TIGR credentials loaded: [%s] user='%s'", label, u)
        else:
            logger.warning("TIGR: skipping [%s] — %s or %s not set.", label, u_var, p_var)

    if not accounts:
        logger.error(
            "No TIGR credentials found in .env.\n"
            "Add at least one pair:\n"
            "  TIGR_MYUSERNAME=...  TIGR_MYPASSWORD=...\n"
            "  TIGR_SGUSERNAME=...  TIGR_SGPASSWORD=..."
        )
        return []
    logger.info("TIGR accounts to process: %d", len(accounts))
    return accounts


def tigr_find_login_frame(page, logger):
    """
    Locate the frame that contains the TIGR login form.

    Search order (most-specific → most-generic):
      1. Frame has #myuserid  (TIGR's exact username field id)
      2. Frame has input[name="myuserid"]
      3. Frame has input[type="password"]   (generic fallback)

    Username is checked first because the TIGR login page requires
    the username to be filled before the password field is activated.
    """
    username_selectors = [
        '#myuserid',
        'input[name="myuserid"]',
        'input[type="password"]',   # fallback — at minimum a pw field must exist
    ]
    deadline = time.time() + 15
    while time.time() < deadline:
        for frame in all_frames(page):
            for sel in username_selectors:
                try:
                    if frame.query_selector(sel):
                        logger.info("TIGR login frame found via '%s' in: %s", sel, frame.url[:80])
                        return frame
                except Exception:
                    pass
        logger.debug("TIGR: login form not found yet, waiting 1 s…")
        time.sleep(1)

    logger.error("TIGR: login form not found after 15 s. Frames present:")
    for f in all_frames(page):
        logger.error("  %s", f.url)
    return None


def _tigr_click_button(page, selectors, label, logger, timeout=10_000):
    """Click the first matching visible button/link across all frames."""
    for sel in selectors:
        for frame in all_frames(page):
            try:
                el = frame.wait_for_selector(sel, timeout=timeout, state="visible")
                if el:
                    el.click()
                    logger.info("TIGR: clicked '%s' via selector '%s'", label, sel)
                    return True
            except PWTimeout:
                continue
            except Exception as exc:
                logger.debug("TIGR: sel='%s' frame='%s' error: %s", sel, frame.url[:60], exc)
    return False


def tigr_login(page, username, password, logger):
    """
    New multi-step Xpansiv/Auth0 login flow:
      1. Open the TIGR landing page and click "LOGIN TO REGISTRY"
      2. Enter email address and click Continue
      3. Enter password and click Continue
      4. If a "Continue" confirmation page appears, click Continue again
    Falls back to legacy single-page login if the new flow is not detected.
    """
    # ── Step 1: open landing page ─────────────────────────────────────────────
    logger.info("TIGR STEP 1 — Opening landing page: %s", TIGR_LOGIN_BUTTON)
    page.goto(TIGR_LOGIN_BUTTON, timeout=TIGR_NAV_TO, wait_until="domcontentloaded")
    time.sleep(2)
    logger.info("TIGR page title: %s | URL: %s", page.title(), page.url)
    screenshot(page, "tigr_01_landing", logger)

    screenshot(page, "tigr_02_auth_page", logger)
    logger.info("TIGR: auth page URL: %s", page.url)

    # ── Detect: TIGR single-page login vs new multi-step (Auth0) login ────────
    #
    # TIGR registry (tigrsregistry.apx.com) uses a classic single-page form:
    #   <input id="myuserid"   name="myuserid"   type="text">    ← username
    #   <input id="mypassword" name="mypassword"  type="password"> ← password
    #
    # The new Xpansiv/Auth0 flow only shows the email field on the first screen
    # (no password visible until after clicking Continue).
    #
    # Detection priority:
    #   1. If #myuserid is present  → TIGR single-page form  → legacy path
    #   2. If both a text/email AND a password field exist    → legacy path
    #   3. Otherwise                                          → stepped path
    # ─────────────────────────────────────────────────────────────────────────
    has_tigr_form   = False
    has_both_fields = False
    has_email_only  = False

    for frame in all_frames(page):
        try:
            # Check for TIGR-specific username field first
            if frame.query_selector('#myuserid, input[name="myuserid"]'):
                has_tigr_form = True
                logger.info("TIGR: detected #myuserid — classic TIGR single-page form.")
                break

            pw_inputs = frame.query_selector_all('input[type="password"]')
            em_inputs = frame.query_selector_all(
                'input[type="email"], input[name="username"], input[name="email"],'
                'input[placeholder*="email" i]'
            )
            if pw_inputs and em_inputs:
                has_both_fields = True
            elif em_inputs and not pw_inputs:
                has_email_only = True
        except Exception:
            pass

    if has_tigr_form or has_both_fields:
        # ── Single-page login (TIGR classic UI) ───────────────────────────────
        logger.info("TIGR: using legacy single-page login path (username → password → submit).")
        _tigr_legacy_fill(page, username, password, logger)
    else:
        # ── New multi-step login (Auth0 / new Xpansiv UI) ─────────────────────
        logger.info("TIGR: using stepped multi-step login path (email → continue → password).")
        _tigr_stepped_login(page, username, password, logger)

    # ── Verify login success ───────────────────────────────────────────────────
    logger.info("TIGR: post-login URL: %s", page.url)
    screenshot(page, "tigr_05_after_login", logger)
    try:
        body = page.inner_text("body").lower()
    except Exception:
        body = ""
    for phrase in ("invalid username", "invalid password", "incorrect password",
                   "login failed", "authentication failed", "access denied",
                   "wrong email or password", "wrong password"):
        if phrase in body:
            screenshot(page, "tigr_error_login_failed", logger)
            raise RuntimeError(
                f"TIGR login failed — page says '{phrase}'. "
                "Check TIGR_MYUSERNAME / TIGR_MYPASSWORD in .env."
            )
    logger.info("TIGR: login successful.")


def _tigr_stepped_login(page, username, password, logger):
    """Handle new Auth0-style multi-step login: email → password → optional continue."""

    # ── Step A: enter email ────────────────────────────────────────────────────
    logger.info("TIGR AUTH: Step A — entering email/username")
    email_selectors = [
        'input[type="email"]',
        'input[name="username"]',
        'input[name="email"]',
        'input[placeholder*="email" i]',
        'input[placeholder*="username" i]',
    ]
    email_field = None
    for sel in email_selectors:
        for frame in all_frames(page):
            try:
                el = frame.wait_for_selector(sel, timeout=5_000, state="visible")
                if el:
                    email_field = el
                    logger.info("TIGR: email field found via '%s'", sel)
                    break
            except PWTimeout:
                continue
        if email_field:
            break

    if not email_field:
        screenshot(page, "tigr_error_no_email_field", logger)
        raise RuntimeError("TIGR: could not find the email/username input on the login page.")

    email_field.click()
    email_field.fill("")
    email_field.type(username, delay=60)
    screenshot(page, "tigr_03a_email_entered", logger)

    # Click the Continue/Next button after email
    continue_selectors = [
        'button[type="submit"]',
        'button:has-text("Continue")',
        'button:has-text("Next")',
        'input[type="submit"]',
        'input[value="Continue"]',
        'input[value="Next"]',
    ]
    if not _tigr_click_button(page, continue_selectors, "Continue (after email)", logger):
        logger.warning("TIGR: no Continue button — pressing Enter.")
        email_field.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
    except PWTimeout:
        logger.warning("TIGR: networkidle timeout after email step.")
    time.sleep(1.5)
    screenshot(page, "tigr_03b_after_email_continue", logger)
    logger.info("TIGR: after email step URL: %s", page.url)

    # ── Step B: enter password ─────────────────────────────────────────────────
    logger.info("TIGR AUTH: Step B — entering password")
    pw_field = None
    deadline = time.time() + 15
    while time.time() < deadline and pw_field is None:
        for frame in all_frames(page):
            try:
                el = frame.wait_for_selector('input[type="password"]', timeout=3_000, state="visible")
                if el:
                    pw_field = el
                    logger.info("TIGR: password field found in frame: %s", frame.url[:80])
                    break
            except PWTimeout:
                continue
        if not pw_field:
            time.sleep(1)

    if not pw_field:
        screenshot(page, "tigr_error_no_password_field", logger)
        raise RuntimeError("TIGR: password input not found after email step.")

    pw_field.click()
    pw_field.fill("")
    pw_field.type(password, delay=60)
    screenshot(page, "tigr_04a_password_entered", logger)

    # Click Continue/Login after password
    if not _tigr_click_button(page, continue_selectors + [
        'button:has-text("Login")', 'button:has-text("Log In")',
        'button:has-text("Sign In")', 'input[value="Login"]',
    ], "Continue (after password)", logger):
        logger.warning("TIGR: no Continue/Login button — pressing Enter.")
        pw_field.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
    except PWTimeout:
        logger.warning("TIGR: networkidle timeout after password step.")
    time.sleep(1.5)
    screenshot(page, "tigr_04b_after_password_continue", logger)
    logger.info("TIGR: after password step URL: %s", page.url)

    # ── Step C: optional "Continue" confirmation page ─────────────────────────
    # Auth0 sometimes shows a page asking to confirm/authorise the app
    try:
        body = page.inner_text("body").lower()
    except Exception:
        body = ""

    confirm_phrases = ("accept", "authorize", "authorise", "allow", "continue", "confirm")
    if any(ph in body for ph in confirm_phrases) and "password" not in body:
        logger.info("TIGR: detected confirmation/authorise page — clicking Continue.")
        confirm_selectors = [
            'button:has-text("Continue")',
            'button:has-text("Accept")',
            'button:has-text("Authorize")',
            'button:has-text("Authorise")',
            'button:has-text("Allow")',
            'button[type="submit"]',
            'input[type="submit"]',
        ]
        if _tigr_click_button(page, confirm_selectors, "Confirm/Continue", logger):
            try:
                page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
            except PWTimeout:
                logger.warning("TIGR: networkidle timeout after confirmation step.")
            time.sleep(1.5)
            screenshot(page, "tigr_04c_after_confirmation", logger)
            logger.info("TIGR: after confirmation URL: %s", page.url)
        else:
            logger.warning("TIGR: confirmation page detected but no button found — continuing.")


def _tigr_legacy_fill(page, username, password, logger):
    """
    Single-page TIGR login (tigrsregistry.apx.com).

    The TIGR login form uses named fields:
      <input id="myuserid"   name="myuserid"   type="text">
      <input id="mypassword" name="mypassword"  type="password">
      <button type="submit">Login</button>

    The username MUST be filled and confirmed before the password field
    is touched — matching the page's natural tab order and preventing
    auto-clear behaviour observed on some browsers.

    Selector priority (username):
      1. #myuserid           — exact TIGR field id
      2. input[name="myuserid"]
      3. input[type="text"]  — generic fallback
    Selector priority (password):
      1. #mypassword         — exact TIGR field id
      2. input[name="mypassword"]
      3. input[type="password"] — generic fallback
    """
    frame = tigr_find_login_frame(page, logger)
    if frame is None:
        screenshot(page, "tigr_error_no_login_form", logger)
        raise RuntimeError("TIGR: could not find login form (no password input).")

    # ── Step 1: locate and fill USERNAME field ────────────────────────────────
    u_field = None
    for u_sel in (
        '#myuserid',
        'input[name="myuserid"]',
        'input[type="text"]',
        'input[type="email"]',
        'input[name="username"]',
        'input[name="email"]',
    ):
        try:
            el = frame.wait_for_selector(u_sel, timeout=5_000, state="visible")
            if el:
                u_field = el
                logger.info("TIGR: username field found via '%s'", u_sel)
                break
        except PWTimeout:
            continue

    if not u_field:
        screenshot(page, "tigr_error_no_username_field", logger)
        raise RuntimeError("TIGR: no username input found on login page.")

    u_field.click()
    u_field.fill("")
    u_field.type(username, delay=60)
    logger.info("TIGR: username entered: '%s'", username)
    screenshot(page, "tigr_legacy_01_username_entered", logger)

    # Small pause so the page registers the username before moving to password
    time.sleep(0.5)

    # ── Step 2: locate and fill PASSWORD field ────────────────────────────────
    p_field = None
    for p_sel in (
        '#mypassword',
        'input[name="mypassword"]',
        'input[type="password"]',
    ):
        try:
            el = frame.wait_for_selector(p_sel, timeout=5_000, state="visible")
            if el:
                p_field = el
                logger.info("TIGR: password field found via '%s'", p_sel)
                break
        except PWTimeout:
            continue

    if not p_field:
        screenshot(page, "tigr_error_no_password_field", logger)
        raise RuntimeError("TIGR: no password input found on login page.")

    p_field.click()
    p_field.fill("")
    p_field.type(password, delay=60)
    logger.info("TIGR: password entered.")
    screenshot(page, "tigr_legacy_02_credentials_filled", logger)

    # ── Step 3: click Login button / submit ───────────────────────────────────
    login_selectors = [
        'button[type="submit"]',        # <button type="submit">Login</button>
        'button.btn:has-text("Login")',
        'button:has-text("Login")',
        'input[value="Login"]',
        'button:has-text("Log In")',
        'input[value="Log In"]',
        'button:has-text("Sign In")',
        'input[type="submit"]',
        'button',
    ]
    submitted = False
    for sel in login_selectors:
        try:
            btn = frame.wait_for_selector(sel, timeout=3_000, state="visible")
            if btn:
                btn.click()
                submitted = True
                logger.info("TIGR: clicked login button via selector '%s'", sel)
                break
        except PWTimeout:
            continue
        except Exception as exc:
            logger.warning("TIGR: could not click '%s': %s", sel, exc)

    if not submitted:
        logger.warning("TIGR: no login button found — pressing Enter on password field.")
        p_field.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
    except PWTimeout:
        logger.warning("TIGR: networkidle timeout after login, continuing…")


def tigr_click_reports(page, logger):
    logger.info("TIGR STEP 3 — Clicking 'Reports'")
    selectors = [
        'a:has-text("Reports")', 'button:has-text("Reports")', 'input[value="Reports"]',
        'a:has-text("Report")', 'button:has-text("Report")',
        '[id*="report" i]', '[class*="report" i]', 'a[href*="report" i]',
    ]
    _, el = click_in_any_frame(page, selectors, "Reports", logger)
    if not el:
        screenshot(page, "tigr_error_no_reports", logger)
        raise RuntimeError("TIGR: 'Reports' button not found.")
    time.sleep(1.5)
    screenshot(page, "tigr_04_reports_open", logger)
    logger.info("TIGR: Reports menu opened.")


def tigr_click_sub_accounts(page, logger):
    logger.info("TIGR STEP 4 — Clicking 'My Sub-Accounts'")
    selectors = [
        'a:has-text("My Sub-Accounts")', 'li:has-text("My Sub-Accounts")',
        'button:has-text("My Sub-Accounts")', 'td:has-text("My Sub-Accounts")',
        'a:has-text("Sub-Accounts")', 'a:has-text("Sub-Account")',
        ':has-text("My Sub-Accounts")',
    ]
    _, el = click_in_any_frame(page, selectors, "My Sub-Accounts", logger)
    if not el:
        screenshot(page, "tigr_error_no_sub_accounts", logger)
        raise RuntimeError("TIGR: 'My Sub-Accounts' not found.")

    try:
        page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
    except PWTimeout:
        logger.warning("TIGR: networkidle timeout, continuing…")
    time.sleep(1)
    screenshot(page, "tigr_05_sub_accounts_loaded", logger)
    logger.info("TIGR: Sub-Accounts loaded. URL: %s", page.url)


def tigr_click_active(page, logger):
    logger.info("TIGR STEP 5 — Clicking 'Active' filter")
    selectors = [
        'input[value="Active"]', 'button:has-text("Active")',
        'a:has-text("Active")', 'td:has-text("Active")',
        'label:has-text("Active")', '[id*="active" i]',
    ]
    _, el = click_in_any_frame(page, selectors, "Active", logger)
    if not el:
        screenshot(page, "tigr_error_no_active", logger)
        raise RuntimeError("TIGR: 'Active' filter not found.")
    time.sleep(2)
    screenshot(page, "tigr_06_active_applied", logger)
    logger.info("TIGR: 'Active' filter applied.")


def tigr_download(page, label, logger) -> Path:
    logger.info("TIGR STEP 6 — Downloading Excel file")
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    selectors_ordered = [
        'a[href*=".xls"]', 'a[href*="excel" i]', 'a[href*="Export" i]', 'a[href*="download" i]',
        'a img[alt*="excel" i]', 'a img[alt*="download" i]', 'a img[alt*="export" i]',
        'a img[src*="excel" i]', 'a img[src*="xls" i]', 'a img[src*="download" i]',
        '[id*="excel" i]', '[id*="export" i]', '[id*="download" i]',
        '[class*="excel" i]', '[class*="export" i]',
        '[title*="excel" i]', '[title*="download" i]', '[title*="export" i]',
    ]

    found_frame = None
    found_el    = None
    frames      = all_frames(page)

    for sel in selectors_ordered:
        for frame in frames:
            try:
                el = frame.wait_for_selector(sel, timeout=3_000, state="visible")
                if el:
                    found_frame = frame
                    found_el    = el
                    logger.info("TIGR: download icon found via selector: %s", sel)
                    break
            except PWTimeout:
                continue
            except Exception as exc:
                logger.debug("TIGR: sel='%s' error: %s", sel, exc)
        if found_el:
            break

    if not found_el:
        logger.warning("TIGR: keyword selectors failed — trying positional fallback.")
        for frame in frames:
            try:
                anchors      = frame.query_selector_all("a")
                icon_anchors = [
                    a for a in anchors
                    if not (a.inner_text() or "").strip() and a.query_selector("img")
                ]
                logger.info("TIGR: icon-style anchors found: %d", len(icon_anchors))
                if len(icon_anchors) >= 2:
                    found_el    = icon_anchors[1]
                    found_frame = frame
                    logger.info("TIGR: using positional fallback icon[1]")
                    break
                elif len(icon_anchors) == 1:
                    found_el    = icon_anchors[0]
                    found_frame = frame
                    logger.warning("TIGR: only 1 icon anchor found — clicking it.")
                    break
            except Exception as exc:
                logger.debug("TIGR: positional fallback error: %s", exc)

    if not found_el:
        screenshot(page, "tigr_error_no_download_button", logger)
        raise RuntimeError("TIGR: could not find the Excel download icon.")

    logger.info("TIGR: triggering download…")
    with page.expect_download(timeout=60_000) as dl_info:
        found_el.click()

    dl        = dl_info.value
    save_path = DOWNLOAD_DIR / f"temp_{label.replace(' ', '_')}.csv"
    if save_path.exists():
        save_path.unlink()
        logger.info("TIGR: removed old file: %s", save_path)
    dl.save_as(str(save_path))

    logger.info("TIGR ✅ File downloaded → %s", save_path.resolve())
    screenshot(page, "tigr_07_download_complete", logger)
    return save_path


def tigr_logout(page, logger):
    logger.info("TIGR STEP 7 — Logging out")
    selectors = [
        'a:has-text("Logout")', 'a:has-text("Log Out")', 'a:has-text("Log off")',
        'a:has-text("Sign Out")', 'button:has-text("Logout")', 'button:has-text("Log Out")',
        '[id*="logout" i]', '[href*="logout" i]', '[href*="logoff" i]', '[href*="signout" i]',
    ]
    _, el = click_in_any_frame(page, selectors, "Logout", logger)
    if not el:
        screenshot(page, "tigr_error_no_logout", logger)
        raise RuntimeError("TIGR: Logout button not found.")

    try:
        page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
    except PWTimeout:
        logger.warning("TIGR: networkidle timeout after logout, continuing…")

    logger.info("TIGR: logged out. URL: %s", page.url)
    screenshot(page, "tigr_08_logged_out", logger)
    time.sleep(1)


def tigr_run_account(page, label, username, password, logger):
    logger.info("━" * 60)
    logger.info("  TIGR ACCOUNT: %s  (user: %s)", label, username)
    logger.info("━" * 60)

    tigr_login(page, username, password, logger)

    # After Auth0 login the browser may land on the Xpansiv home page rather
    # than the TIGR app page.  Navigate there explicitly if needed.
    if TIGR_URL not in page.url:
        logger.info("TIGR: navigating to app page: %s", TIGR_URL)
        page.goto(TIGR_URL, timeout=TIGR_NAV_TO, wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=TIGR_NAV_TO)
        except PWTimeout:
            logger.warning("TIGR: networkidle timeout navigating to app page.")
        time.sleep(1)
        logger.info("TIGR: app page loaded — URL: %s", page.url)
        screenshot(page, "tigr_06_app_page", logger)

    tigr_click_reports(page, logger)
    tigr_click_sub_accounts(page, logger)
    tigr_click_active(page, logger)
    saved = tigr_download(page, label, logger)
    tigr_logout(page, logger)

    logger.info("TIGR ✅  [%s] Done. File: %s", label, saved.resolve())
    return saved


def run_tigr(logger):
    banner("BOT 2 — TIGR Registry (tigrsregistry.apx.com)")
    accounts = tigr_load_env(logger)
    if not accounts:
        logger.error("No TIGR accounts configured — skipping TIGR bot.")
        return []

    results = []
    _ci_args = _browser_args()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=HEADLESS,
            slow_mo=TIGR_SLOW_MO,
            downloads_path=str(DOWNLOAD_DIR.resolve()),
            args=_ci_args,
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080} if HEADLESS else None,
        )
        page = context.new_page()
        page.on("console",   lambda m: logger.debug("[browser] %s: %s", m.type, m.text))
        page.on("pageerror", lambda e: logger.error("[page error] %s", e))

        for idx, (label, username, password) in enumerate(accounts, start=1):
            logger.info("TIGR: processing account %d/%d: %s", idx, len(accounts), label)
            try:
                saved = tigr_run_account(page, label, username, password, logger)
                results.append((label, saved, None))
            except (RuntimeError, PWTimeout, Exception) as exc:
                logger.error("TIGR ❌  [%s] FAILED: %s", label, exc)
                screenshot(page, f"tigr_error_{label.replace(' ', '_')}", logger)
                results.append((label, None, str(exc)))
                try:
                    logger.info("TIGR: recovering — navigating to login page…")
                    page.goto(TIGR_URL, timeout=TIGR_NAV_TO, wait_until="domcontentloaded")
                    time.sleep(2)
                except Exception as nav_exc:
                    logger.error("TIGR: recovery failed: %s", nav_exc)

        context.close()
        browser.close()

    return results


# ═════════════════════════════════════════════════════════════════════════════
#  BOT 3 — COMBINE OUTPUT
# ═════════════════════════════════════════════════════════════════════════════

def swap_month_day(date_val):
    """
    Swap month and day in a date value.
    Input:  6/1/2026  (month=6, day=1)  → Output: 1/6/2026  (month=1, day=6)
    Handles datetime objects, date objects, and string representations.
    Returns a date object with month and day swapped.
    """
    if date_val is None:
        return None
    # Normalise to a date/datetime
    if isinstance(date_val, (datetime, date)):
        d = date_val if isinstance(date_val, date) else date_val.date()
        # Swap: new month = old day, new day = old month
        try:
            return date(d.year, d.day, d.month)
        except ValueError:
            return None
    # Try to parse string
    val = str(date_val).strip()
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            d = datetime.strptime(val, fmt).date()
            return date(d.year, d.day, d.month)
        except ValueError:
            continue
    return None


def last_day_of_month(d):
    """Return a date for the last day of the month that d falls in."""
    if d is None:
        return None
    _, last_day = calendar.monthrange(d.year, d.month)
    return date(d.year, d.month, last_day)


def map_tigr_country(country_val):
    """Map TIGR country names to 2-letter codes."""
    if country_val is None:
        return None
    key = str(country_val).strip().lower()
    return TIGR_COUNTRY_MAP.get(key, str(country_val).strip())


def map_fuel_code(fuel_val, tech_val=None):
    """
    Map the Fuel column value to a short code for column E.

    Rules (case-insensitive substring match):
        Solar        → SLR
        Biogas       → BIO
        Biomass      → BMS
        Geothermal   → GEO
        Wind         → WND
        Hydro-electric + Tech contains "Dam"          → LHY
        Hydro-electric + Tech contains "Run of river" → SHYD

    Returns the original value unchanged if no rule matches.
    """
    if fuel_val is None:
        return None
    fuel_lower = str(fuel_val).strip().lower()
    tech_lower = str(tech_val).strip().lower() if tech_val else ""

    if "solar" in fuel_lower:
        return "SLR"
    if "biogas" in fuel_lower:
        return "BIO"
    if "biomass" in fuel_lower:
        return "BMS"
    if "geothermal" in fuel_lower:
        return "GEO"
    if "wind" in fuel_lower:
        return "WND"
    if "hydro" in fuel_lower:
        if "dam" in tech_lower:
            return "LHY"
        if "run of river" in tech_lower:
            return "SHYD"
    return str(fuel_val).strip()


def irec_country_to_code(country_val):
    """Take the first 2 characters of the IREC country field."""
    if country_val is None:
        return None
    return str(country_val).strip()[:2].upper()


def _make_getter(row, columns):
    """
    Return a case-insensitive column getter for a single row.
    Defined outside the loop to avoid the Python closure-capture bug
    where inner functions in loops capture the loop variable by reference.
    """
    col_map = {c.strip().lower(): c for c in columns}
    def get(col):
        actual = col_map.get(col.strip().lower())
        if actual is None:
            return None
        v = row[actual]
        return None if pd.isna(v) else str(v).strip()
    return get


def _find_header_row(path, required_keywords, max_scan=10):
    """
    Scan the first max_scan rows of an Excel file to find the real header row.
    Returns the 0-based row index where the header lives, or 0 as fallback.
    required_keywords: list of lowercase strings that should appear in the header row.
    """
    try:
        raw = pd.read_excel(path, header=None, nrows=max_scan, dtype=str)
        for i, row in raw.iterrows():
            row_lower = [str(v).strip().lower() for v in row if pd.notna(v)]
            if any(kw in row_lower for kw in required_keywords):
                return i
    except Exception:
        pass
    return 0


def process_irec_files(irec_file_paths, logger):
    """
    Read all IREC CSV files and return a list of dicts matching TEMPLATE_COLUMNS.
    IREC column mapping:
        Device         → Asset ID
        Device Name    → Asset
        Fuel type      → Fuel
        Technology     → Tech
        Country        → Country (first 2 chars)
        Volume         → Quantity
        Period / Start → Period Start  (tries multiple common column names)
        End / To       → Period End    (tries multiple common column names)
    Registry = "IREC"
    """
    # Possible IREC column names for Period Start
    IREC_START_ALIASES = [
        "period start", "start date", "from", "date from", "reporting start",
        "issue date", "issuance date", "period from", "valid from", "start",
    ]
    # Possible IREC column names for Period End
    IREC_END_ALIASES = [
        "period end", "end date", "to", "date to", "reporting end",
        "expiry date", "period to", "valid to", "end",
    ]
    # Possible IREC column names for Fuel
    IREC_FUEL_ALIASES = [
        "fuel", "fuel type", "energy type", "source",
    ]

    def reformat_date(val):
        """Convert any recognisable date string to DD/MM/YYYY."""
        if not val:
            return val
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        return val  # return as-is if format unrecognised

    rows = []
    for path in irec_file_paths:
        if path is None or not Path(path).exists():
            logger.warning("IREC file not found, skipping: %s", path)
            continue
        try:
            # utf-8-sig strips the UTF-8 BOM (\ufeff) that evident.app prepends to
            # the first column header — without this "Device" becomes "\ufeffDevice"
            # and _make_getter can never find it, silently producing empty rows.
            df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
            # Also strip any stray whitespace from column names
            df.columns = [c.strip() for c in df.columns]
            logger.info("IREC file '%s': %d rows, columns: %s", path, len(df), list(df.columns))
        except Exception as exc:
            logger.error("Failed to read IREC file '%s': %s", path, exc)
            continue

        for row_num, (_, row) in enumerate(df.iterrows(), start=2):
            try:
                get = _make_getter(row, df.columns)

                period_start = next((get(a) for a in IREC_START_ALIASES if get(a)), None)
                period_end   = next((get(a) for a in IREC_END_ALIASES   if get(a)), None)

                out = {col: None for col in TEMPLATE_COLUMNS}
                out["Registry"]     = "IREC"
                out["Asset ID"]     = get("Device")
                out["Asset"]        = get("Device Name")
                out["Fuel"]         = next((get(a) for a in IREC_FUEL_ALIASES if get(a)), None)
                out["Tech"]         = get("Technology")
                out["Fuel"]         = map_fuel_code(out["Fuel"], out["Tech"])
                out["Country"]      = irec_country_to_code(get("Country"))
                raw_qty             = get("Volume")
                out["Quantity"]     = float(raw_qty) if raw_qty else None
                out["Period Start"] = reformat_date(period_start)
                out["Period End"]   = reformat_date(period_end)
                rows.append(out)
            except Exception as exc:
                logger.warning("IREC: skipping row %d in '%s' due to error: %s", row_num, path, exc)

    logger.info("IREC: total rows processed: %d", len(rows))
    return rows


def process_tigr_files(tigr_file_paths, logger):
    """
    Read all TIGR CSV files and return a list of dicts matching TEMPLATE_COLUMNS.
    TIGR transformations:
        Fuel Type    → Tech
        TIGR Vintage → swap month/day → Period Start
        Period End   = last day of Period Start month
        Country      → map Viet Nam→VN, Malaysia→MY, Singapore→SG
    Registry = "TIGR"
    """
    # All possible column name aliases for the Vintage field
    VINTAGE_ALIASES = ["tigr vintage", "vintage", "period", "month", "issuance date", "issue date"]
    # All possible column name aliases for the Quantity/Credits field
    QUANTITY_ALIASES = ["quantity", "volume", "credits", "mwh", "amount", "certificates"]
    # All possible column name aliases for the Fuel/Tech field
    FUEL_ALIASES = ["fuel type", "fuel", "technology", "tech", "energy type"]

    rows = []
    for path in tigr_file_paths:
        if path is None or not Path(path).exists():
            logger.warning("TIGR file not found, skipping: %s", path)
            continue

        try:
            ext = Path(path).suffix.lower()
            if ext in (".xlsx", ".xls"):
                # Find the real header row (TIGR Excel files sometimes have
                # a title / logo row above the actual column headers)
                header_row = _find_header_row(
                    path,
                    required_keywords=["vintage", "sub-account", "quantity", "country", "fuel"],
                )
                df = pd.read_excel(path, header=header_row, dtype=str)
            else:
                df = pd.read_csv(path, dtype=str)
            df.dropna(how="all", inplace=True)
            df.columns = [str(c).strip() for c in df.columns]   # tidy whitespace
            df.reset_index(drop=True, inplace=True)
            logger.info("TIGR file '%s': %d rows, columns: %s", path, len(df), list(df.columns))
        except Exception as exc:
            logger.error("Failed to read TIGR file '%s': %s", path, exc)
            continue

        for _, row in df.iterrows():
            get = _make_getter(row, df.columns)

            # Vintage: try multiple alias names
            vintage_raw = None
            for alias in VINTAGE_ALIASES:
                vintage_raw = get(alias)
                if vintage_raw:
                    break

            # Quantity: try multiple alias names
            quantity_val = None
            for alias in QUANTITY_ALIASES:
                quantity_val = get(alias)
                if quantity_val:
                    break

            # Fuel/Tech: read from file, fall back to "Solar Photovoltaics"
            tech_val = None
            for alias in FUEL_ALIASES:
                tech_val = get(alias)
                if tech_val:
                    break
            if not tech_val:
                tech_val = "Solar Photovoltaics"

            # Swap month↔day for Period Start, derive Period End
            period_start = swap_month_day(vintage_raw)
            period_end   = last_day_of_month(period_start)

            def fmt_date(d):
                if d is None:
                    return None
                return f"{d.month}/{d.day}/{d.year}"

            out = {col: None for col in TEMPLATE_COLUMNS}
            out["Registry"]     = "TIGR"
            out["Fuel"]         = map_fuel_code(tech_val, tech_val)
            out["Tech"]         = tech_val
            out["Period Start"] = fmt_date(period_start)
            out["Period End"]   = fmt_date(period_end)
            out["Country"]      = map_tigr_country(get("Country"))
            out["Quantity"]     = float(quantity_val) if quantity_val else None

            # Pass through any remaining template columns that map directly
            for col in TEMPLATE_COLUMNS:
                if out[col] is None:
                    out[col] = get(col)

            rows.append(out)

    logger.info("TIGR: total rows processed: %d", len(rows))
    return rows


def _apply_sheet_styles(ws, rows_data, col_widths, logger, is_previous=False):
    """
    Helper: write header + data rows with styling into a given worksheet.
    `rows_data` is a list of dicts keyed by TEMPLATE_COLUMNS.
    If `is_previous=True`, header uses a grey tone to visually distinguish it.
    """
    # ── Header styling ────────────────────────────────────────────────────────
    if is_previous:
        header_fill = PatternFill("solid", fgColor="5A5A5A")   # dark grey for Previous
    else:
        header_fill = PatternFill("solid", fgColor="1F4E79")   # dark blue for Current

    header_font  = Font(bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border  = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = thin_border

    # ── Row styling ───────────────────────────────────────────────────────────
    if is_previous:
        irec_fill = PatternFill("solid", fgColor="C8D8E8")   # muted blue for IREC (previous)
        tigr_fill = PatternFill("solid", fgColor="C8E8C8")   # muted green for TIGR (previous)
    else:
        irec_fill = PatternFill("solid", fgColor="DDEEFF")   # light blue for IREC (current)
        tigr_fill = PatternFill("solid", fgColor="DFFFDF")   # light green for TIGR (current)

    row_align = Alignment(horizontal="left", vertical="center")

    for row_idx, row_data in enumerate(rows_data, start=2):
        registry = row_data.get("Registry", "")
        fill     = irec_fill if registry == "IREC" else tigr_fill

        for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, start=1):
            value = row_data.get(col_name)
            cell  = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill   = fill
            cell.border = thin_border
            if col_name == "Quantity":
                cell.alignment    = Alignment(horizontal="right", vertical="center")
                cell.number_format = "#,##0.000000"
            else:
                cell.alignment = row_align

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(TEMPLATE_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(col_name, 15)

    # Freeze header row
    ws.freeze_panes = "A2"


def _load_previous_rows(prev_path, logger):
    """
    Read the 'Current' sheet from the previous combined_output.xlsx and return
    its rows as a list of dicts keyed by TEMPLATE_COLUMNS.
    Falls back to reading whichever sheet exists if 'Current' is absent.
    """
    try:
        import openpyxl
        wb_prev = openpyxl.load_workbook(str(prev_path), read_only=True, data_only=True)

        # Prefer "Current", then "Combined" (legacy name), then first sheet
        sheet_name = None
        for candidate in ("Current", "Combined"):
            if candidate in wb_prev.sheetnames:
                sheet_name = candidate
                break
        if sheet_name is None:
            sheet_name = wb_prev.sheetnames[0]

        ws_prev = wb_prev[sheet_name]
        rows_iter = ws_prev.iter_rows(values_only=True)

        header = next(rows_iter, None)
        if header is None:
            logger.warning("Previous output sheet '%s' is empty.", sheet_name)
            wb_prev.close()
            return []

        prev_rows = []
        for row in rows_iter:
            row_dict = {col: row[i] if i < len(row) else None
                        for i, col in enumerate(header)}
            # Re-key to TEMPLATE_COLUMNS so unknown columns are silently dropped
            prev_rows.append({col: row_dict.get(col) for col in TEMPLATE_COLUMNS})

        wb_prev.close()
        logger.info("Loaded %d previous row(s) from '%s' (sheet: %s)",
                    len(prev_rows), prev_path.name, sheet_name)
        return prev_rows

    except Exception as exc:
        logger.warning("Could not read previous output '%s': %s — skipping comparison.", prev_path, exc)
        return []


def _parse_date_for_summary(val):
    """
    Parse a date value (string, date, datetime) and return a Python date object.
    Tries common formats used in the output rows.  Returns None on failure.
    """
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _quarter_date_range(year, quarter):
    """
    Return (period_start_str, period_end_str) for a given year and quarter.
    e.g. year=2024, quarter=1  →  ("01/01/2024", "31/03/2024")
    """
    q_start_month = (quarter - 1) * 3 + 1
    q_end_month   = quarter * 3
    _, last_day   = calendar.monthrange(year, q_end_month)
    start = date(year, q_start_month, 1)
    end   = date(year, q_end_month, last_day)
    return start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")


def _period_range_label(start_date, end_date):
    """
    Format a human-readable period range label.
    e.g. date(2026,1,1), date(2026,6,30) → "Jan 26 - Jun 26"
    Accepts date objects or (year, month) tuples.
    """
    if isinstance(start_date, tuple):
        start_date = date(start_date[0], start_date[1], 1)
    if isinstance(end_date, tuple):
        end_date = date(end_date[0], end_date[1], 1)
    s = start_date.strftime("%b %y")   # e.g. "Jan 26"
    e = end_date.strftime("%b %y")     # e.g. "Jun 26"
    return f"{s} - {e}"


def _write_summary_sheet(wb, all_rows, logger):
    """
    Create a 'Summary' sheet that aggregates Quantity by:
        Registry | Country | Fuel | Production Start Month | Production Start Year |
        Production End Month | Production End Year | Period Start | Period End | Quantity

    Rules:
    - Each unique (Registry, Country, Fuel, Year, Quarter) combination gets ONE row.
    - If all rows in a group share the same Registry/Country/Fuel, a TOTAL row is appended.
    - NO cell is ever left empty — every field is populated with a value (0 for missing qty).
    - Production Start Month/Year and Production End Month/Year are derived from
      that quarter's Period Start and Period End dates (e.g. Period Start = 01/01/2026
      → Production Start Month = "Jan", Production Start Year = 2026).
    - Period Start / Period End show the actual start and end dates of that quarter.
    - Rows whose Period Start cannot be parsed are logged and skipped.
    """
    from collections import defaultdict

    SUMMARY_HEADERS = [
        "Registry", "Country", "Fuel",
        "Production Start Month", "Production Start Year",
        "Production End Month", "Production End Year",
        "Period Start", "Period End",
        "Quantity",
    ]

    # ── Build aggregation dict ────────────────────────────────────────────────
    # key: (Registry, Country, Fuel, Year, Quarter) → total qty
    # Production Start/End Month & Year are derived from the quarter's
    # Period Start / Period End dates (e.g. Period Start = 01/01/2026 →
    # Production Start Month = "Jan", Production Start Year = 2026).
    agg = defaultdict(float)

    skipped = 0
    for row in all_rows:
        d = _parse_date_for_summary(row.get("Period Start"))
        if d is None:
            skipped += 1
            continue

        quarter  = (d.month - 1) // 3 + 1
        year     = d.year
        registry = str(row.get("Registry") or "").strip() or "Unknown"
        country  = str(row.get("Country")  or "").strip() or "Unknown"
        fuel     = str(row.get("Fuel")     or "").strip() or "Unknown"

        try:
            qty = float(row.get("Quantity") or 0)
        except (TypeError, ValueError):
            qty = 0.0

        agg[(registry, country, fuel, year, quarter)] += qty

    if skipped:
        logger.warning("Summary sheet: %d row(s) skipped (unparseable Period Start).", skipped)

    # ── Sort: Registry → Country → Fuel → Year → Quarter ─────────────────────
    sorted_keys = sorted(agg.keys(), key=lambda k: (k[0], k[1], k[2], k[3], k[4]))

    # ── Build flat list of display rows ──────────────────────────────────────
    # Each detail row + a subtotal row per (Registry, Country, Fuel) group
    display_rows = []   # list of (row_values_list, is_total_row)

    # Group by (Registry, Country, Fuel)
    from itertools import groupby
    group_key_fn = lambda k: (k[0], k[1], k[2])

    for (registry, country, fuel), group_iter in groupby(sorted_keys, key=group_key_fn):
        group_keys = list(group_iter)
        group_total = 0.0
        for key in group_keys:
            _, _, _, year, quarter = key
            qty = agg[key]
            group_total += qty
            ps_str, pe_str = _quarter_date_range(year, quarter)
            q_start_month  = (quarter - 1) * 3 + 1
            q_end_month    = quarter * 3
            start_date_obj = date(year, q_start_month, 1)
            end_date_obj   = date(year, q_end_month,   1)

            prod_start_month = start_date_obj.strftime("%b")
            prod_start_year  = start_date_obj.year
            prod_end_month   = end_date_obj.strftime("%b")
            prod_end_year    = end_date_obj.year

            display_rows.append((
                [registry, country, fuel,
                 prod_start_month, prod_start_year,
                 prod_end_month, prod_end_year,
                 ps_str, pe_str, qty],
                False,
            ))



    # ── Create worksheet ──────────────────────────────────────────────────────
    ws = wb.create_sheet(title="Summary")

    # ── Styles ────────────────────────────────────────────────────────────────
    header_fill   = PatternFill("solid", fgColor="1F4E79")
    header_font   = Font(bold=True, color="FFFFFF", size=11, name="Arial")
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)

    total_fill    = PatternFill("solid", fgColor="D9E1F2")   # soft blue for subtotal rows
    total_font    = Font(bold=True, color="1F4E79", size=10, name="Arial")

    fill_a        = PatternFill("solid", fgColor="EAF2FB")   # light blue-grey (even rows)
    fill_b        = PatternFill("solid", fgColor="FFFFFF")   # white (odd rows)

    data_font     = Font(name="Arial", size=10)
    num_fmt       = "#,##0.000000"
    right_align   = Alignment(horizontal="right",  vertical="center")
    left_align    = Alignment(horizontal="left",   vertical="center")
    center_align  = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="thin"),
    )
    thick_bottom = Border(
        left=Side(style="thin"),  right=Side(style="thin"),
        top=Side(style="thin"),   bottom=Side(style="medium"),
    )

    # ── Write header row ──────────────────────────────────────────────────────
    for col_idx, h in enumerate(SUMMARY_HEADERS, start=1):
        cell           = ws.cell(row=1, column=col_idx, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align
        cell.border    = thin_border

    # ── Write data rows ───────────────────────────────────────────────────────
    for row_idx, (values, is_total) in enumerate(display_rows, start=2):
        if is_total:
            fill   = total_fill
            font   = total_font
            border = thick_bottom
        else:
            fill   = fill_a if row_idx % 2 == 0 else fill_b
            font   = data_font
            border = thin_border

        for col_idx, (h, val) in enumerate(zip(SUMMARY_HEADERS, values), start=1):
            # Guarantee no empty cell — fallback to 0 for numbers, "—" for text
            if val is None or val == "":
                val = 0.0 if h == "Quantity" else "—"

            cell           = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = font
            cell.fill      = fill
            cell.border    = border

            if h == "Quantity":
                cell.alignment    = right_align
                cell.number_format = num_fmt
            elif h in ("Period Start", "Period End"):
                cell.alignment = center_align
            elif h in ("Production Start Month", "Production Start Year",
                       "Production End Month", "Production End Year"):
                cell.alignment = center_align
                if is_total:
                    cell.font = Font(bold=True, color="1F4E79", size=10,
                                     name="Arial", italic=True)
            else:
                cell.alignment = left_align

    # ── Column widths ─────────────────────────────────────────────────────────
    summary_widths = {
        "Registry":                14,
        "Country":                 10,
        "Fuel":                    14,
        "Production Start Month":  16,
        "Production Start Year":   16,
        "Production End Month":    16,
        "Production End Year":     16,
        "Period Start":            16,
        "Period End":              16,
        "Quantity":                20,
    }
    for col_idx, h in enumerate(SUMMARY_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = summary_widths.get(h, 14)

    # ── Row height for header ─────────────────────────────────────────────────
    ws.row_dimensions[1].height = 30

    ws.freeze_panes = "A2"
    logger.info("'Summary' sheet written: %d row(s)", len(display_rows))
    print(f"  [i] 'Summary' sheet added: {len(display_rows)} row(s).")
    return ws


def write_combined_excel(all_rows, logger, tadau_rows=None, isenergy_rows=None):
    """
    Write all rows into a styled Excel file using the hardcoded template columns.

    Behaviour
    ---------
    * The output is always saved as  output/combined_output.xlsx  (fixed name).
    * The workbook contains:
        - "Current"   — main combined data (IREC + TIGR, excluding TADAU / IS Energy)
        - "TADAU"     — data from the T9MQ/TADAU account (if any)
        - "IS Energy" — data from the IS Energy Sdn Bhd (T0CV1BYF) account (if any)
        - "Summary"   — aggregates Quantity by Registry/Country/Fuel/Year/Quarter
    * Any existing combined_output.xlsx is simply overwritten.

    Columns not available in a source remain empty (None → blank cell).
    """
    output_path = OUTPUT_DIR / "combined_output.xlsx"

    col_widths = {
        "Registry":       12,
        "Asset ID":       18,
        "Asset":          30,
        "Country":        10,
        "Fuel":           18,
        "Tech":           22,
        "Period Start":   14,
        "Period End":     14,
        "Quantity":       14,
        "Transferor":     20,
        "Feed-in Tariff": 15,
        "Sub-Account":    18,
        "Sub-Account ID": 16,
    }

    # ── Build workbook ────────────────────────────────────────────────────────
    wb = Workbook()

    # "Current" sheet — always first (active by default)
    ws_current = wb.active
    ws_current.title = "Current"
    _apply_sheet_styles(ws_current, all_rows, col_widths, logger, is_previous=False)
    logger.info("'Current' sheet written: %d row(s)", len(all_rows))

    # "TADAU" sheet — data from the T9MQ account (separate, not merged into Current)
    if tadau_rows:
        ws_tadau = wb.create_sheet(title="TADAU")
        _apply_sheet_styles(ws_tadau, tadau_rows, col_widths, logger, is_previous=False)
        logger.info("'TADAU' sheet written: %d row(s)", len(tadau_rows))
        print(f"  [i] 'TADAU' sheet added: {len(tadau_rows)} row(s).")
    else:
        logger.info("'TADAU' sheet skipped — no TADAU rows found.")

    # "IS Energy" sheet — data from the IS Energy Sdn Bhd (T0CV1BYF) account (separate, not merged into Current)
    if isenergy_rows:
        ws_isenergy = wb.create_sheet(title="IS Energy")
        _apply_sheet_styles(ws_isenergy, isenergy_rows, col_widths, logger, is_previous=False)
        logger.info("'IS Energy' sheet written: %d row(s)", len(isenergy_rows))
        print(f"  [i] 'IS Energy' sheet added: {len(isenergy_rows)} row(s).")
    else:
        logger.info("'IS Energy' sheet skipped — no IS Energy rows found.")

    # "Summary" sheet — always last; aggregates Quantity by Registry/Country/Fuel/Year/Quarter
    _write_summary_sheet(wb, all_rows, logger)

    # ── Save (overwrites the fixed filename) ─────────────────────────────────
    wb.save(str(output_path))
    logger.info("✅ Combined output saved → %s", output_path.resolve())
    print(f"\n  ✅ Combined Excel saved → {output_path.resolve()}")
    return output_path


def detect_file_types(csv_paths, logger):
    """
    Open each CSV and inspect its headers to decide if it is an IREC or TIGR file.
    IREC signature: contains 'device' and 'volume' column headers.
    TIGR signature: contains 'vintage' or 'sub-account' column headers.
    Unrecognised files are logged and skipped.

    Robustness notes
    ----------------
    * Read with encoding='utf-8-sig' to strip the UTF-8 BOM (\ufeff) that many
      web-app CSV exports prepend to the first column header.  Without this,
      "Device" becomes "\ufeffDevice" and never matches "device".
    * Headers are stripped of surrounding whitespace AND any leading/trailing
      non-alphanumeric characters before comparison, as an extra safety net.
    * IREC_SIGNATURES is expanded to cover 'device name' in case the export
      omits the plain 'Device' column but keeps 'Device Name'.
    * The filename itself is used as a last-resort fallback: IREC CSVs are
      saved as  <label>_holdings.csv  by irec_download_csv().
    """
    # IREC: exported with "Device" and/or "Device Name" columns
    IREC_SIGNATURES = {"device", "device name"}
    # TIGR: exported with "TIGR Vintage" or "Sub-Account" columns
    TIGR_SIGNATURES = {"tigr vintage", "sub-account", "vintage", "sub account", "subaccount"}

    irec_paths = []
    tigr_paths = []

    for path in csv_paths:
        try:
            # utf-8-sig strips the BOM that evident.app / many web apps prepend
            df_head = pd.read_csv(path, nrows=0, dtype=str, encoding="utf-8-sig")
            # Strip whitespace AND any leading/trailing non-word chars (e.g. BOM remnants)
            import re as _re
            headers = {_re.sub(r'^\W+|\W+$', '', c).strip().lower() for c in df_head.columns}
        except Exception as exc:
            logger.warning("detect_file_types: could not read '%s': %s", path, exc)
            continue

        logger.debug("detect_file_types: '%s' cleaned headers → %s", path, sorted(headers))

        if IREC_SIGNATURES & headers:
            logger.info("  → IREC : %s  (headers: %s)", path, list(df_head.columns))
            irec_paths.append(path)
        elif TIGR_SIGNATURES & headers:
            logger.info("  → TIGR : %s  (headers: %s)", path, list(df_head.columns))
            tigr_paths.append(path)
        elif "_holdings.csv" in str(path).lower():
            # Filename fallback: irec_download_csv() always saves as <label>_holdings.csv
            logger.warning(
                "  → IREC (filename fallback): %s  (headers didn't match known signatures: %s)",
                path, sorted(headers),
            )
            irec_paths.append(path)
        else:
            logger.warning("  → UNKNOWN (skipped): %s  (headers: %s)", path, sorted(headers))

    logger.info("Detected %d IREC and %d TIGR file(s)", len(irec_paths), len(tigr_paths))
    return irec_paths, tigr_paths


def run_combine(irec_results, tigr_results, logger):
    banner("BOT 3 — Combine Output")

    # ── Scan the downloads folder directly ─────────────────────────────────────
    # IREC downloads are CSV files.
    # TIGR downloads are Excel files (.xlsx / .xls).
    # Scan both independently so one missing source does not block the other.
    all_csvs   = sorted([str(p) for p in DOWNLOAD_DIR.glob("*.csv")  if p.is_file()])
    all_excels = sorted([str(p) for p in DOWNLOAD_DIR.glob("*.xls*") if p.is_file()])
    logger.info("Downloads folder scan → %d CSV file(s) + %d Excel file(s)",
                len(all_csvs), len(all_excels))

    if not all_csvs and not all_excels:
        logger.warning("No downloadable files in: %s", DOWNLOAD_DIR.resolve())
        print(f"  [!] No files found in {DOWNLOAD_DIR.resolve()} — nothing to combine.")
        return None

    # IREC: classify CSVs by header content (device column = IREC signature)
    irec_paths_all, _extra_tigr_csvs = detect_file_types(all_csvs, logger)

    # Split IREC paths into regular vs TADAU vs IS Energy by filename prefix
    tadau_paths    = [p for p in irec_paths_all if Path(p).name.startswith(TADAU_LABEL_PREFIX)]
    isenergy_paths = [p for p in irec_paths_all if Path(p).name.startswith(ISENERGY_LABEL_PREFIX)]
    irec_paths     = [p for p in irec_paths_all
                       if not Path(p).name.startswith(TADAU_LABEL_PREFIX)
                       and not Path(p).name.startswith(ISENERGY_LABEL_PREFIX)]
    logger.info("IREC paths split → %d regular, %d TADAU, %d IS Energy",
                len(irec_paths), len(tadau_paths), len(isenergy_paths))

    # TIGR: Excel files from the downloads folder
    tigr_paths = all_excels
    if _extra_tigr_csvs:
        # Unexpected TIGR-looking CSVs — include them too (belt-and-braces)
        tigr_paths = _extra_tigr_csvs + tigr_paths
        logger.info("TIGR: also including %d CSV file(s) with TIGR signatures", len(_extra_tigr_csvs))

    logger.info("Processing %d IREC CSV(s), %d TADAU CSV(s), %d IS Energy CSV(s) and %d TIGR Excel(s)",
                len(irec_paths), len(tadau_paths), len(isenergy_paths), len(tigr_paths))

    irec_rows     = process_irec_files(irec_paths,     logger)
    tadau_rows    = process_irec_files(tadau_paths,    logger)
    isenergy_rows = process_irec_files(isenergy_paths, logger)
    tigr_rows     = process_tigr_files(tigr_paths,     logger)

    all_rows = irec_rows + tigr_rows
    logger.info("Total rows to write: %d (%d IREC + %d TIGR) + %d TADAU (separate sheet) + %d IS Energy (separate sheet)",
                len(all_rows), len(irec_rows), len(tigr_rows), len(tadau_rows), len(isenergy_rows))

    if not all_rows and not tadau_rows and not isenergy_rows:
        logger.warning("No data rows to write — skipping Excel output.")
        print("  [!] No data to combine. Check that downloads succeeded.")
        return None

    return write_combined_excel(
        all_rows, logger,
        tadau_rows=tadau_rows or None,
        isenergy_rows=isenergy_rows or None,
    )


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═════════════════════════════════════════════════════════════════════════════
def main():
    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("Combined Bot — %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)

    # ── Run IREC ─────────────────────────────────────────────────────────────
    irec_results = []
    try:
        irec_results = run_irec(logger)
    except SystemExit:
        logger.error("IREC bot exited early (fatal credential error).")
    except Exception as exc:
        logger.error("IREC bot crashed: %s", exc)

    # ── Run TIGR ─────────────────────────────────────────────────────────────
    tigr_results = []
    try:
        tigr_results = run_tigr(logger)
    except SystemExit:
        logger.error("TIGR bot exited early.")
    except Exception as exc:
        logger.error("TIGR bot crashed: %s", exc)

    # ── Combine Output ────────────────────────────────────────────────────────
    combined_path = None
    try:
        combined_path = run_combine(irec_results, tigr_results, logger)
    except Exception as exc:
        logger.error("Combine step crashed: %s", exc)
        traceback.print_exc()

    # ── Summary ───────────────────────────────────────────────────────────────
    banner("FINAL SUMMARY")
    all_ok = True

    print("\n  ── IREC Holdings ──")
    if not irec_results:
        print("  [!] No IREC results (check credentials or errors above)")
        all_ok = False
    for account_name, path, err in irec_results:
        if err:
            print(f"  ❌  [{account_name}]  FAILED: {err}")
            all_ok = False
        else:
            print(f"  ✅  [{account_name}]  → {path}")

    print("\n  ── TIGR Registry ──")
    if not tigr_results:
        print("  [!] No TIGR results (check credentials or errors above)")
        all_ok = False
    for label, path, err in tigr_results:
        if err:
            print(f"  ❌  [{label}]  FAILED: {err}")
            all_ok = False
        else:
            print(f"  ✅  [{label}]  → {path}")

    print("\n  ── Combined Output ──")
    if combined_path:
        print(f"  ✅  Combined Excel → {combined_path}")
    else:
        print("  [!] Combined output not generated (no data or error above)")
        all_ok = False

    print(f"\n  Downloads in : {DOWNLOAD_DIR.resolve()}")
    print(f"  Output in    : {OUTPUT_DIR.resolve()}")
    print("=" * 60)
    _pause("\n  Press Enter to close this window...")
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
