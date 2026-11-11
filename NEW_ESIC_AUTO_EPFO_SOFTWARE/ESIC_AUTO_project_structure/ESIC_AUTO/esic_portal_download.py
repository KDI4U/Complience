import asyncio
import os
import re
import sys
import time
import logging
from datetime import datetime
from pathlib import Path
import ollama
from openpyxl import load_workbook
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ============================================================
# PATHS - FULL ABSOLUTE PATHS
# ============================================================

PROJECT_FOLDER = Path(__file__).resolve().parent
EXCEL_FILE = PROJECT_FOLDER.parent / "esic_users.xlsx"
if not EXCEL_FILE.exists():
    _fallback_excel = PROJECT_FOLDER / "esic_users.xlsx"
    if _fallback_excel.exists():
        EXCEL_FILE = _fallback_excel

DATA_FOLDER = PROJECT_FOLDER / "data"

SHEET_NAME = None  # None = active sheet
OLLAMA_MODEL = "gemma4:12b"

# RETRY RULES
MAX_CAPTCHA_RETRIES = 4       # Captcha issue me 4 baar retry krna hai
MAX_SERVER_ERROR_RETRIES = 1  # Server error aaye usko 1 attempt krna hai (no further retry)
MAX_RETRIES_PER_USER = MAX_CAPTCHA_RETRIES

LOGIN_URL = "https://portal.esic.gov.in/ESICInsurance1/ESICInsurancePortal/Portal_Loginnew.aspx"

DATE_FOLDER_NAME = datetime.now().strftime("%d%m%Y")
TARGET_DOWNLOAD_DIR = DATA_FOLDER / DATE_FOLDER_NAME / "data_downloaded"
CAPTCHA_DIR = DATA_FOLDER / "captcha"

# Required folders automatically created
CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
TARGET_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("esic_portal")
if not logger.handlers:
    _sh = logging.StreamHandler(sys.stdout)
    _sh.setLevel(logging.INFO)
    _sh.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] [esic_portal]: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(_sh)
    logger.setLevel(logging.INFO)


# ============================================================
# CUSTOM EXCEPTIONS FOR RETRY POLICY
# ============================================================
class ESICServerError(Exception):
    """Raised when ESIC portal returns 500 / Server Internal Error."""
    pass


class ESICCaptchaError(Exception):
    """Raised when ESIC login fails due to incorrect captcha or remained on login screen."""
    pass


def is_server_error(url_str: str = "", content_str: str = "") -> bool:
    """Detects if ESIC portal returned a 500 or Server Internal Error."""
    u = (url_str or "").lower()
    c = (content_str or "").lower()
    # Check URL patterns
    if any(k in u for k in ["errorpage500", "errorpage", "aspxerrorpath", "error.aspx", "500.html", "500.htm"]):
        return True
    # Server / Internal Error strings in page content or error message
    server_error_phrases = [
        "internal server error",
        "server internal error",
        "server error in '/' application",
        "server error in",
        "runtime error",
        "500 - internal server error",
        "500 internal server error",
        "http error 500",
        "error 500",
        "http 500",
        "unexpected condition that prevented it from fulfilling your request",
        "service unavailable",
        "503 service unavailable",
        "an error occurred while processing your request",
        "the page cannot be displayed",
    ]
    if any(phrase in c for phrase in server_error_phrases):
        return True
    return False



def clean_excel_val(val):
    """Converts numbers cleanly to integer string format avoiding scientific float notation."""
    if val is None:
        return ""
    if isinstance(val, float):
        return f"{int(val)}"
    if isinstance(val, int):
        return str(val)

    val_str = str(val).strip()

    if val_str.endswith(".0"):
        val_str = val_str[:-2]

    if re.search(r"^\d+\.?\d*[eE]\+\d+$", val_str):
        try:
            return f"{int(float(val_str))}"
        except Exception:
            pass

    return val_str


def ensure_ollama_running(timeout: int = 15) -> bool:
    """Checks if Ollama service is responsive; if not, attempts to launch it automatically."""
    import urllib.request
    import subprocess
    import shutil

    url = "http://127.0.0.1:11434/api/tags"
    try:
        with urllib.request.urlopen(url, timeout=2) as resp:
            if resp.status == 200:
                return True
    except Exception:
        pass

    logger.info("Ollama is not running. Attempting to start Ollama automatically...")

    ollama_app_path = Path(os.path.expanduser(r"~\AppData\Local\Programs\Ollama\ollama app.exe"))
    try:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        if ollama_app_path.exists():
            subprocess.Popen([str(ollama_app_path)], creationflags=flags)
        elif shutil.which("ollama"):
            subprocess.Popen(["ollama", "serve"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
    except Exception as e:
        logger.warning(f"Could not automatically launch Ollama: {e}")

    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            with urllib.request.urlopen(url, timeout=2) as resp:
                if resp.status == 200:
                    logger.info("Ollama is now online and ready.")
                    return True
        except Exception:
            time.sleep(1)

    logger.error("Failed to connect to Ollama after waiting.")
    return False


def _get_captcha_sync(image_path):
    """Synchronous call to Ollama AI model to read ESIC CAPTCHA image with fast token limits."""
    for attempt in range(2):
        try:
            kwargs = {
                "model": OLLAMA_MODEL,
                "messages": [{
                    "role": "user",
                    "content": (
                        "Read all characters from this CAPTCHA image accurately. "
                        "Output ONLY the exact alphanumeric characters shown. "
                        "Do not add spaces, punctuation, or any other text."
                    ),
                    "images": [str(image_path)]
                }],
                "options": {
                    "num_predict": 25,
                    "temperature": 0.0
                }
            }
            try:
                response = ollama.chat(**kwargs, think=False)
            except TypeError:
                response = ollama.chat(**kwargs)

            raw_text = ""
            if hasattr(response, "message") and response.message:
                raw_text = (getattr(response.message, "content", "") or "").strip()
                if not raw_text and getattr(response.message, "thinking", None):
                    candidates = re.findall(r"['\"]([a-zA-Z0-9]{4,7})['\"]", response.message.thinking)
                    if candidates:
                        raw_text = candidates[-1]
            elif isinstance(response, dict):
                msg = response.get("message", {})
                raw_text = (msg.get("content") or "").strip()
                if not raw_text and msg.get("thinking"):
                    candidates = re.findall(r"['\"]([a-zA-Z0-9]{4,7})['\"]", msg.get("thinking"))
                    if candidates:
                        raw_text = candidates[-1]

            cleaned = re.sub(r"[^a-zA-Z0-9]", "", raw_text)
            if cleaned:
                logger.info(f"Ollama CAPTCHA extracted: {cleaned}")
                return cleaned
            logger.warning(f"Ollama returned empty or non-alphanumeric text ('{raw_text}'). Retrying...")
        except Exception as e:
            logger.error(f"Ollama CAPTCHA extraction attempt {attempt + 1} failed: {e}")
            ensure_ollama_running(timeout=10)
    return ""


async def get_captcha_from_ollama(image_path):
    """Async wrapper around Ollama sync chat call."""
    return await asyncio.to_thread(_get_captcha_sync, image_path)


async def dismiss_all_popups(page):
    """Dismisses all ESIC pop-up dialogs using exact element IDs quickly."""
    popup_selectors = [
        "input#div1_close",
        "input#btnpnlSanjeevanicolse",
        "input#btnpnlcolse",
        "input#btnpnlPwdcolse",
        'input[name="div1_close"]',
        'input[name="btnpnlSanjeevanicolse"]',
        'input[name="btnpnlcolse"]',
        'input[name="btnpnlPwdcolse"]',
        'input[value="I Agree"]',
        'input[value="Close"]',
        'input[value="X"]'
    ]

    for _ in range(3):
        clicked_any = False
        for selector in popup_selectors:
            try:
                el = page.locator(selector).first
                if await el.count() > 0 and await el.is_visible():
                    btn_id = await el.get_attribute("id") or selector
                    await el.click(timeout=800)
                    logger.info(f"Dismissed Popup ({btn_id})")
                    await page.wait_for_timeout(300)
                    clicked_any = True
            except Exception:
                pass
        if not clicked_any:
            break


async def fill_login_and_captcha(page, loginid, password, captcha_path, auto_mode=True):
    username_input = page.locator(
        'input#txtUserName, input[name*="UserName"], input[type="text"]:visible'
    ).first
    if await username_input.count() == 0:
        raise RuntimeError("Login ID field not found.")

    await username_input.fill(str(loginid))

    password_input = page.locator(
        'input#txtPassword, input[name*="Password"], input[type="password"]:visible'
    ).first
    if await password_input.count() == 0:
        raise RuntimeError("Password field not found.")

    await password_input.fill(str(password))

    captcha_img = page.locator(
        'img#img1, img[src*="ChallanHandler"], img[id*="Captcha"], img[id*="captcha"]'
    ).first

    if await captcha_img.count() == 0:
        raise RuntimeError("ESIC CAPTCHA image element not found.")

    await captcha_img.screenshot(path=str(captcha_path))

    captcha_text = await get_captcha_from_ollama(captcha_path)

    if not captcha_text:
        raise ESICCaptchaError("Ollama failed to extract CAPTCHA (empty text). Reloading for retry.")

    if len(captcha_text) < 4 or len(captcha_text) > 10:
        if auto_mode:
            logger.warning(f"Ollama CAPTCHA result suspicious ('{captcha_text}'). Proceeding with attempt...")
        else:
            captcha_text = input("Enter CAPTCHA manually: ").strip()

    captcha_input = page.locator(
        'input#txtChallanCaptcha, input[name="txtChallanCaptcha"], input[placeholder*="Captcha"]'
    ).first

    if await captcha_input.count() == 0:
        raise RuntimeError("ESIC CAPTCHA input field not found.")

    await captcha_input.click()
    await captcha_input.fill("")
    await captcha_input.fill(captcha_text)
    logger.info(f"CAPTCHA filled: {captcha_text}")
    await page.wait_for_timeout(500)


async def login(page, loginid, password, safe_loginid, auto_mode=True):
    response = await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
    if response and response.status >= 500:
        raise ESICServerError(f"ESIC Portal returned HTTP {response.status} on Login URL")

    # Fast wait after load
    await page.wait_for_timeout(500)

    # Check for early server error
    if is_server_error(page.url):
        raise ESICServerError("ESIC Portal returned Server 500 Error on Login URL")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    captcha_path = CAPTCHA_DIR / f"captcha_{safe_loginid}_{timestamp}.png"

    await fill_login_and_captcha(page, loginid, password, captcha_path, auto_mode=auto_mode)

    # Wait time reduced to 2s (fast and reliable)
    logger.info("Waiting 2s before clicking Login...")
    await page.wait_for_timeout(2000)

    login_button = page.locator(
        'input#btnLogin, input[name="btnLogin"], input[value="Login"], input.login100-form-btn-gou'
    ).first

    if await login_button.count() == 0:
        login_button = page.get_by_role("button", name="Login", exact=False)

    if await login_button.count() == 0:
        raise RuntimeError("LOGIN button not found.")

    await login_button.click()
    logger.info("Login button clicked.")

    try:
        await page.wait_for_load_state("domcontentloaded", timeout=25000)
    except PlaywrightTimeoutError:
        pass

    # Quick settle wait
    await page.wait_for_timeout(1500)

    # Check if login immediately redirected to ErrorPage500
    if is_server_error(page.url):
        raise ESICServerError("ESIC Server 500 Error occurred immediately after Login submit.")


async def open_target_and_download(page, loginid, target_dir):
    # Check if page is currently on Server Error
    page_content = ""
    try:
        page_content = await page.content()
    except Exception:
        pass

    if is_server_error(page.url, page_content):
        raise ESICServerError("ESIC Server Internal Error (500) on page.")

    # Check if still on Login screen (meaning login failed)
    if "Portal_Loginnew.aspx" in page.url:
        error_msg_el = page.locator('#lblMsg, .error, span[style*="color:Red"], span[style*="color:red"]').first
        err_text = ""
        if await error_msg_el.count() > 0:
            try:
                err_text = (await error_msg_el.inner_text()).strip()
            except Exception:
                pass

        if is_server_error(page.url, err_text or page_content):
            raise ESICServerError(f"ESIC Server Error on Login: {err_text or 'Server Internal Error'}")

        if err_text:
            if "password" in err_text.lower() and "captcha" not in err_text.lower():
                raise RuntimeError(f"Login Failed (Password issue): {err_text}")
            raise ESICCaptchaError(f"Login Failed (Captcha issue): {err_text}")
        raise ESICCaptchaError("Login Failed (Invalid CAPTCHA / Webpage remained on Login screen)")

    # Dashboard loaded - quick popups dismissal
    await page.wait_for_timeout(800)
    await dismiss_all_popups(page)

    await page.keyboard.press("Enter")
    await page.wait_for_timeout(600)

    target = page.locator(
        'a#lnkInsertIPDetails, '
        'a[id*="lnkInsertIPDetails"], '
        'a:has-text("IP Mapping with Contractor/Principal Employer"), '
        'a:has-text("IP Mapping")'
    ).first

    try:
        await target.wait_for(state="visible", timeout=12000)
    except PlaywrightTimeoutError:
        try:
            curr_content = await page.content()
        except Exception:
            curr_content = ""
        if is_server_error(page.url, curr_content):
            raise ESICServerError("ESIC Server 500 Error occurred on Dashboard.")
        raise RuntimeError("Dashboard option 'IP Mapping' link not found.")

    logger.info("Opening IP Mapping window...")

    async with page.expect_popup(timeout=25000) as popup_info:
        await target.click()

    popup_page = await popup_info.value
    logger.info("Popup window opened.")

    # Active response listener on popup and main page to detect 500 / server errors immediately
    server_error_state = {"error": False, "msg": ""}

    def on_response_check(resp):
        try:
            if resp.status >= 500:
                server_error_state["error"] = True
                server_error_state["msg"] = f"HTTP {resp.status} on {resp.url}"
            elif any(k in resp.url.lower() for k in ["errorpage", "aspxerrorpath"]):
                server_error_state["error"] = True
                server_error_state["msg"] = f"Redirected to error page: {resp.url}"
        except Exception:
            pass

    popup_page.on("response", on_response_check)
    page.on("response", on_response_check)

    try:
        await popup_page.wait_for_load_state("domcontentloaded", timeout=15000)
    except PlaywrightTimeoutError:
        pass

    await popup_page.wait_for_timeout(500)

    # Check for Server Error on popup page immediately
    popup_content = ""
    try:
        popup_content = await popup_page.content()
    except Exception:
        pass

    if server_error_state["error"] or is_server_error(popup_page.url, popup_content) or is_server_error(page.url):
        raise ESICServerError(f"ESIC Server Error on IP Mapping window: {server_error_state['msg'] or popup_page.url}")

    download_btn = popup_page.locator(
        'input#ctl00_HomePageContent_download, '
        'input[name*="HomePageContent$download"], '
        'input[title*="Download Existing IPDetails"], '
        'input[value="Download"]'
    ).first

    try:
        await download_btn.wait_for(state="visible", timeout=6000)
    except PlaywrightTimeoutError:
        popup_content = ""
        try:
            popup_content = await popup_page.content()
        except Exception:
            pass
        if server_error_state["error"] or is_server_error(popup_page.url, popup_content) or is_server_error(page.url):
            raise ESICServerError("ESIC Server Internal Error on IP Details window")
        raise RuntimeError("Download button not found on popup page.")

    safe_loginid = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(loginid))

    # Trigger download with active polling so we abort instantly if server error occurs
    download_task = asyncio.create_task(popup_page.wait_for_event("download", timeout=25000))
    await download_btn.click()

    start_poll = asyncio.get_event_loop().time()
    while not download_task.done():
        if server_error_state["error"]:
            download_task.cancel()
            raise ESICServerError(f"ESIC Server Error during download: {server_error_state['msg']}")

        if is_server_error(popup_page.url):
            download_task.cancel()
            raise ESICServerError(f"ESIC Server Error page reached during download: {popup_page.url}")

        try:
            p_content = await popup_page.content()
            if is_server_error(popup_page.url, p_content):
                download_task.cancel()
                raise ESICServerError("ESIC Internal Server Error detected in popup content during download")
        except Exception:
            pass

        if is_server_error(page.url):
            download_task.cancel()
            raise ESICServerError(f"ESIC Server Error detected on main portal page: {page.url}")

        await asyncio.sleep(0.3)

        if asyncio.get_event_loop().time() - start_poll > 25:
            break

    if not download_task.done():
        download_task.cancel()
        curr_c = ""
        try:
            curr_c = await popup_page.content()
        except Exception:
            pass
        if server_error_state["error"] or is_server_error(popup_page.url, curr_c):
            raise ESICServerError("ESIC Server Error on download timeout")
        raise ESICServerError("Download timed out (Server unresponsive, skipping to next)")

    try:
        download = await download_task
    except Exception as exc:
        curr_c = ""
        try:
            curr_c = await popup_page.content()
        except Exception:
            pass
        if server_error_state["error"] or is_server_error(popup_page.url, curr_c):
            raise ESICServerError(f"ESIC Server Error on download: {exc}")
        raise ESICServerError(f"Download event failed: {exc}")

    original_filename = download.suggested_filename or "Employer_IP_Details.xls"
    ext = os.path.splitext(original_filename)[1] or ".xls"

    destination = target_dir / f"{safe_loginid}{ext}"
    await download.save_as(destination)
    logger.info(f"File downloaded -> {destination.name}")

    try:
        await popup_page.close()
    except Exception:
        pass

    return destination


def read_users():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME else wb.active

    users = []
    for row in range(2, ws.max_row + 1):
        raw_loginid = ws.cell(row, 1).value
        raw_password = ws.cell(row, 2).value

        loginid = clean_excel_val(raw_loginid)
        password = clean_excel_val(raw_password)

        if not loginid:
            continue

        if not password:
            ws.cell(row, 3).value = "PASSWORD_MISSING"
            continue

        users.append((row, loginid, password))

    try:
        wb.save(EXCEL_FILE)
    except Exception as e:
        logger.warning(f"Could not save Excel file changes in read_users: {e}")
    return users


def write_response(row, response):
    if not row or row <= 0:
        return
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb[SHEET_NAME] if SHEET_NAME else wb.active
        ws.cell(row, 3).value = response
        wb.save(EXCEL_FILE)
    except Exception as e:
        logger.warning(f"Could not write status to Excel (file may be open): {e}")


async def run_portal_download(headless=False, auto_mode=True, single_user=None, single_password=None, target_codes=None, only_failed=False):
    logger.info("Starting ESIC Portal Download Automation")
    logger.info(f"Excel File: {EXCEL_FILE.name}")
    logger.info(f"Download Dir: {TARGET_DOWNLOAD_DIR}")
    logger.info(f"Policy: Server Error -> 1 attempt only | Captcha issues -> {MAX_CAPTCHA_RETRIES} attempts")

    # Ensure Ollama AI service is running before attempting logins
    if not ensure_ollama_running(timeout=15):
        logger.error("Ollama AI Service is offline and could not be started automatically. Aborting ESIC download.")
        return False

    # Ensure directories exist
    CAPTCHA_DIR.mkdir(parents=True, exist_ok=True)
    TARGET_DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    all_users = read_users()

    if single_user:
        # If single_password missing, search in all_users
        if not single_password:
            found = next((u for u in all_users if str(u[1]).strip() == str(single_user).strip()), None)
            if found:
                users = [found]
                logger.info(f"Running for Single User: {single_user} (Password loaded from Excel)")
            else:
                logger.error(f"User {single_user} password not found in Excel and not provided.")
                return False
        else:
            found = next((u for u in all_users if str(u[1]).strip() == str(single_user).strip()), None)
            row_idx = found[0] if found else 0
            users = [(row_idx, str(single_user).strip(), str(single_password).strip())]
            logger.info(f"Running for Single User: {single_user}")
    elif target_codes:
        codes_list = [c.strip() for c in str(target_codes).split(",") if c.strip()]
        users = [u for u in all_users if str(u[1]).strip() in codes_list]
        logger.info(f"Target Codes Filter applied: found {len(users)} of {len(codes_list)} requested code(s): {', '.join(codes_list)}")
        if not users:
            logger.warning(f"None of the target codes {codes_list} matched accounts in Excel.")
            return False
    elif only_failed:
        wb_check = load_workbook(EXCEL_FILE, data_only=True)
        ws_check = wb_check[SHEET_NAME] if SHEET_NAME else wb_check.active
        failed_users = []
        for u in all_users:
            status_val = str(ws_check.cell(u[0], 3).value or "").upper()
            if "FAIL" in status_val or "ERROR" in status_val or "PENDING" in status_val or not status_val:
                failed_users.append(u)
        users = failed_users
        logger.info(f"Only-Failed Filter applied: {len(users)} account(s) pending/failed to process.")
        if not users:
            logger.info("No failed/pending accounts found to retry! All accounts completed successfully.")
            return True
    else:
        users = all_users
        if not users:
            logger.warning("No valid users found in Excel.")
            return True

    logger.info(f"Total Users to Process: {len(users)}")

    success_count = 0
    failed_count = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=headless,
            slow_mo=0,
        )

        for index, (row, loginid, password) in enumerate(users, start=1):
            logger.info(f"Processing User {index}/{len(users)} | Login ID: {loginid}")

            user_success = False
            for attempt in range(1, MAX_CAPTCHA_RETRIES + 1):
                if attempt > 1:
                    logger.info(f"Retry attempt {attempt}/{MAX_CAPTCHA_RETRIES} for User: {loginid} (Captcha Retry)")

                context = await browser.new_context(
                    accept_downloads=True,
                    viewport={"width": 1600, "height": 900},
                )
                page = await context.new_page()

                try:
                    safe_loginid = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in str(loginid))
                    await login(page, loginid, password, safe_loginid, auto_mode=auto_mode)
                    downloaded = await open_target_and_download(page, loginid, TARGET_DOWNLOAD_DIR)

                    write_response(row, f"SUCCESS - {downloaded.name}")
                    logger.info(f"User {loginid} SUCCESS: {downloaded.name}")
                    user_success = True
                    success_count += 1
                    break

                except ESICServerError as srv_err:
                    err_msg = str(srv_err)
                    logger.error(f"User {loginid} Attempt {attempt} [SERVER ERROR]: {err_msg}")
                    logger.warning(
                        f"⚡ Internal Server Error on code {loginid}. Closing immediately and moving to next user without waiting."
                    )
                    write_response(row, f"FAILED - SERVER_ERROR ({err_msg})")
                    failed_count += 1
                    # Server error is limited to 1 attempt - break immediately to next user!
                    break

                except ESICCaptchaError as cap_err:
                    err_msg = str(cap_err)
                    logger.warning(f"User {loginid} Attempt {attempt}/{MAX_CAPTCHA_RETRIES} [CAPTCHA ISSUE]: {err_msg}")

                    if attempt < MAX_CAPTCHA_RETRIES:
                        logger.info(f"Retrying User {loginid} due to Captcha issue ({attempt + 1}/{MAX_CAPTCHA_RETRIES})...")
                        await asyncio.sleep(1.5)
                    else:
                        logger.error(f"User {loginid} failed after {MAX_CAPTCHA_RETRIES} Captcha attempts.")
                        write_response(row, f"FAILED - CAPTCHA_ERROR (Tried {MAX_CAPTCHA_RETRIES} times: {err_msg})")
                        failed_count += 1

                except Exception as exc:
                    err_str = str(exc)

                    # If this generic exception is actually a server error or timeout
                    if (
                        is_server_error(err_str)
                        or "500" in err_str
                        or "server error" in err_str.lower()
                        or "internal server error" in err_str.lower()
                        or "download timed out" in err_str.lower()
                    ):
                        logger.error(f"User {loginid} Attempt {attempt} [SERVER/PORTAL ERROR detected]: {err_str}")
                        logger.warning(f"Server/Portal Error detected for code {loginid}. Closing immediately and moving to next user.")
                        write_response(row, f"FAILED - SERVER_ERROR ({err_str})")
                        failed_count += 1
                        break

                    # Password error -> stop immediately
                    if "password" in err_str.lower() and "captcha" not in err_str.lower():
                        logger.error(f"User {loginid} Password Error: {err_str}. Stopping.")
                        write_response(row, f"FAILED - {err_str}")
                        failed_count += 1
                        break

                    # Other error: retry up to MAX_CAPTCHA_RETRIES
                    logger.error(f"User {loginid} Attempt {attempt}/{MAX_CAPTCHA_RETRIES} Error: {err_str}")
                    if attempt < MAX_CAPTCHA_RETRIES:
                        logger.info(f"Retrying User {loginid} (Attempt {attempt + 1}/{MAX_CAPTCHA_RETRIES})...")
                        await asyncio.sleep(1.5)
                    else:
                        write_response(row, f"FAILED - {err_str}")
                        logger.error(f"User {loginid} failed after {MAX_CAPTCHA_RETRIES} attempts.")
                        failed_count += 1

                finally:
                    await context.close()

        await browser.close()
        logger.info(f"ESIC Portal Download Summary: {success_count} Succeeded, {failed_count} Failed.")

        # Return True if at least 1 file was successfully downloaded so downstream pipeline can process
        if success_count > 0:
            if failed_count > 0:
                logger.warning(f"Proceeding to Merge/Compare pipeline with {success_count} downloaded file(s) ({failed_count} user(s) failed).")
            return True
        else:
            logger.error("0 files downloaded. Cannot proceed to downstream processing.")
            return False


def run(headless=False, auto_mode=True):
    """Entry point used by n8n or external caller."""
    return asyncio.run(
        run_portal_download(
            headless=headless,
            auto_mode=auto_mode,
        )
    )


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="ESIC Portal Download")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--user", type=str, default=None, help="Specific user ID / code")
    parser.add_argument("--password", type=str, default=None, help="Password for specific user")
    parser.add_argument("--codes", type=str, default=None, help="Comma-separated target codes")
    parser.add_argument("--only-failed", action="store_true", help="Process only failed accounts")
    cli_args = parser.parse_args()

    ok = asyncio.run(
        run_portal_download(
            headless=cli_args.headless,
            auto_mode=True,
            single_user=cli_args.user,
            single_password=cli_args.password,
            target_codes=cli_args.codes,
            only_failed=cli_args.only_failed
        )
    )
    if not ok:
        sys.exit(1)
