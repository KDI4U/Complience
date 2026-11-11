import os
import sys
import re
import csv
import zipfile
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
import ollama

# ============================================================
# PROJECT BASE & CONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

try:
    from config import EXCEL_PATH, CAPTCHA_FOLDER, TARGET_DOWNLOAD_FOLDER, EPFO_DOWNLOAD_DIR, DATE_FOLDER_NAME
    BASE_DOWNLOAD_FOLDER = EPFO_DOWNLOAD_DIR
except ImportError:
    BASE_DIR = Path(__file__).resolve().parent.parent
    EXCEL_PATH = BASE_DIR / "data" / "users.xlsx"
    CAPTCHA_FOLDER = BASE_DIR / "data" / "captcha_images"
    BASE_DOWNLOAD_FOLDER = BASE_DIR / "data" / "pf_downloads"
    DATE_FOLDER_NAME = datetime.now().strftime("%d%m%Y")
    TARGET_DOWNLOAD_FOLDER = BASE_DOWNLOAD_FOLDER / DATE_FOLDER_NAME

OLLAMA_MODEL = "gemma4:12b"
MAX_RETRIES_PER_USER = 3
LOGIN_URL = "https://unifiedportal-emp.epfindia.gov.in/epfo/"

# Create target folders if they don't exist
CAPTCHA_FOLDER.mkdir(parents=True, exist_ok=True)
TARGET_DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)


def clean_excel_val(val):
    if val is None:
        return ""
    if isinstance(val, float):
        return f"{int(val)}"
    if isinstance(val, int):
        return str(val)

    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    if re.search(r'^\d+\.?\d*[eE]\+\d+$', val_str):
        try:
            return f"{int(float(val_str))}"
        except:
            pass
    return val_str


def clean_employee_name(name):
    """
    Removes titles/honorifics (Mr., Ms., Mrs., Shri., Smt., Miss., Dr., Master) at start of name,
    removes extra spaces, BUT strictly keeps 'Md', 'Md.', or 'MD' intact.
    Example: 'Ms. MADHUSMITA PRIYADARSHINI BISWAL' -> 'MADHUSMITA PRIYADARSHINI BISWAL'
    Example: 'MD SHOAIB' -> 'MD SHOAIB' (Kept intact)
    """
    if not name:
        return ""

    cleaned_name = str(name).strip()
    pattern = r'^(?:Mr\.|Mr|Ms\.|Ms|Mrs\.|Mrs|Shri\.|Shri|Smt\.|Smt|Miss\.|Miss|Dr\.|Dr|Master\.|Master)\s+'
    cleaned_name = re.sub(pattern, '', cleaned_name, flags=re.IGNORECASE)
    cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
    return cleaned_name


def fix_captcha_pattern(raw_text):
    """
    Cleans and formats CAPTCHA:
    - 1st & 2nd characters: 2 UPPERCASE letters (e.g. AB)
    - 3rd & 4th characters: 2 LOWERCASE letters (e.g. cd)
    - 5th character: 1 NUMERIC DIGIT (e.g. 1)
    """
    if not raw_text:
        return ""

    cleaned = re.sub(r'[^a-zA-Z0-9]', '', raw_text)
    if len(cleaned) < 5:
        return cleaned

    chars = list(cleaned[:5])

    letter_to_digit = {
        'O': '0', 'o': '0', 'Q': '0',
        'I': '1', 'l': '1', 'i': '1', 't': '1', 'J': '1',
        'Z': '2', 'z': '2',
        'E': '3', 'e': '3',
        'A': '4', 'a': '4',
        'S': '5', 's': '5',
        'G': '6', 'b': '6',
        'T': '7', 'j': '7',
        'B': '8',
        'g': '9', 'q': '9'
    }

    digit_to_letter = {
        '0': 'O', '1': 'L', '2': 'Z', '3': 'E', '4': 'A',
        '5': 'S', '6': 'G', '7': 'T', '8': 'B', '9': 'G'
    }

    # Position 1 & 2 (index 0, 1) -> Must be UPPERCASE letter
    for i in (0, 1):
        if chars[i].isdigit():
            chars[i] = digit_to_letter.get(chars[i], 'A').upper()
        else:
            chars[i] = chars[i].upper()

    # Position 3 & 4 (index 2, 3) -> Must be LOWERCASE letter
    for i in (2, 3):
        if chars[i].isdigit():
            chars[i] = digit_to_letter.get(chars[i], 'a').lower()
        else:
            chars[i] = chars[i].lower()

    # Position 5 (index 4) -> Must be NUMERIC DIGIT
    if not chars[4].isdigit():
        chars[4] = letter_to_digit.get(chars[4], '1')

    return "".join(chars)


def get_captcha_from_ollama(image_path):
    try:
        # Restored earlier high-accuracy prompt
        prompt_content = (
            "Read the CAPTCHA text from the image. The CAPTCHA strictly contains 5 characters in total with this exact structure:\n"
            "1. First 2 characters are UPPERCASE letters (e.g. AB)\n"
            "2. Next 2 characters are LOWERCASE letters (e.g. cd)\n"
            "3. The 5th character is 1 NUMERIC DIGIT (e.g. 1)\n"
            "Example pattern: ABcd1\n"
            "Respond ONLY with these 5 characters. Do NOT add any extra words, spaces, explanations, or punctuation."
        )

        response = ollama.chat(
            model=OLLAMA_MODEL,
            messages=[{
                'role': 'user',
                'content': prompt_content,
                'images': [str(image_path)]
            }]
        )

        raw_text = response['message']['content'].strip()
        print(f"Ollama Raw Output: {raw_text}")
        
        fixed_captcha = fix_captcha_pattern(raw_text)
        print(f"Cleaned CAPTCHA: {fixed_captcha}")
        return fixed_captcha

    except Exception as e:
        print("Ollama Reading Error:", e)
        return ""


def extract_and_clean_downloaded_file(downloaded_file_path, userid, target_folder):
    r"""
    Unzips downloaded EPFO ZIP file, formats columns into 15 exact headers,
    cleans Employee Name (keeps Md.), and saves as data\pf_downloads\DDMMYYYY\{userid}.csv
    """
    safe_userid = re.sub(r'[^\w\-]', '_', str(userid))
    target_csv_path = target_folder / f"{safe_userid}.csv"
    raw_content = ""

    if zipfile.is_zipfile(downloaded_file_path):
        print(f"Detected ZIP archive: {downloaded_file_path}")
        try:
            with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                print(f"Files inside ZIP: {file_list}")
                txt_csv_files = [f for f in file_list if f.endswith('.csv') or f.endswith('.txt') or '.' not in f]
                target_inner_file = txt_csv_files[0] if txt_csv_files else file_list[0]
                with zip_ref.open(target_inner_file) as inner_file:
                    raw_content = inner_file.read().decode('utf-8-sig', errors='ignore')
            try:
                os.remove(downloaded_file_path)
            except:
                pass
        except Exception as zip_err:
            print(f"Error unzipping file: {zip_err}")
            return None
    else:
        # Plain text / CSV file
        try:
            with open(downloaded_file_path, 'r', encoding='utf-8-sig', errors='ignore') as infile:
                raw_content = infile.read()
            try:
                os.remove(downloaded_file_path)
            except:
                pass
        except Exception as read_err:
            print(f"Error reading file: {read_err}")
            return None

    if not raw_content.strip():
        print("Downloaded file content is empty!")
        return None

    lines = [line for line in raw_content.splitlines() if line.strip()]
    if not lines:
        return None

    header_line = lines[0]
    delimiter = ','
    if '#~#' in header_line:
        delimiter = '#~#'
    elif '~!' in header_line:
        delimiter = '~!'
    elif '\t' in header_line:
        delimiter = '\t'
    elif ',' in header_line:
        delimiter = ','

    reader = csv.DictReader(lines, delimiter=delimiter)
    target_headers = [
        'UAN', 'PF NO', 'Employee Name', 'Gender', 'DoB', 'DoJ',
        "Father's/Husband's Name", 'NameRelation', 'Marital Status',
        'Mobile', 'Email ID', 'EPFO Employer Code',
        'Nomination Filed', 'Is AADHAAR Verified', 'Face Auth Status'
    ]

    rows = []
    for row in reader:
        cleaned_row = {str(k).strip(): str(v).strip() for k, v in row.items() if k is not None}
        pf_no = cleaned_row.get('Member ID', cleaned_row.get('MemberID', cleaned_row.get('PF NO', ''))).strip()
        employer_code = pf_no[:15] if len(pf_no) >= 15 else pf_no

        raw_emp_name = cleaned_row.get('Name', cleaned_row.get('Employee Name', '')).strip()
        cleaned_emp_name = clean_employee_name(raw_emp_name)

        raw_father_name = cleaned_row.get("Father's/Husband's Name", cleaned_row.get("Father/Husband Name", '')).strip()
        cleaned_father_name = clean_employee_name(raw_father_name)

        new_row = {
            'UAN': cleaned_row.get('UAN', '').strip(),
            'PF NO': pf_no,
            'Employee Name': cleaned_emp_name,
            'Gender': cleaned_row.get('Gender', '').strip(),
            'DoB': cleaned_row.get('DoB', '').strip(),
            'DoJ': cleaned_row.get('DoJ', '').strip(),
            "Father's/Husband's Name": cleaned_father_name,
            'NameRelation': cleaned_row.get('Relation', cleaned_row.get('NameRelation', '')).strip(),
            'Marital Status': cleaned_row.get('Marital Status', cleaned_row.get('MaritalStatus', '')).strip(),
            'Mobile': cleaned_row.get('Mobile', '').strip(),
            'Email ID': cleaned_row.get('Email ID', cleaned_row.get('EmailID', '')).strip(),
            'EPFO Employer Code': cleaned_row.get('EPFO Employer Code', cleaned_row.get('Est ID', employer_code)).strip(),
            'Nomination Filed': cleaned_row.get('Nomination Filed', '').strip(),
            'Is AADHAAR Verified': cleaned_row.get('Is AADHAAR Verified', '').strip(),
            'Face Auth Status': cleaned_row.get('Face Auth Status', '').strip(),
        }
        rows.append(new_row)

    with open(target_csv_path, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=target_headers)
        writer.writeheader()
        writer.writerows(rows)

    print(f"SUCCESS: Extracted & Cleaned CSV saved ({len(rows)} rows) → {target_csv_path}")
    return str(target_csv_path)


def dismiss_all_popups(page):
    """
    Guarantees that any open alert modal or popup (initial alert or post-login notice)
    is completely closed before continuing. Uses multi-method click + DOM removal JS.
    """
    print("Checking and closing all active alert popups...")
    page.wait_for_timeout(1000)

    # 1. Playwright visible button clicks
    try:
        modal_btns = page.locator("#btnCloseModal, .modal.show button, .modal button.btn-danger, .modal .btn-close").all()
        for btn in modal_btns:
            if btn.is_visible():
                try:
                    btn.click(force=True, timeout=2000)
                    print("Modal button clicked via Playwright.")
                except:
                    pass
    except Exception as err:
        pass

    # 2. Hardened JS dismissal & removal to prevent any lingering backdrop overlay
    try:
        page.evaluate("""
            () => {
                const btns = document.querySelectorAll('#btnCloseModal, .modal button, .btn-close, button.btn-danger, [data-bs-dismiss="modal"]');
                btns.forEach(b => {
                    if (b.offsetWidth > 0 && b.offsetHeight > 0) {
                        b.click();
                    }
                });
                document.querySelectorAll('.modal, .modal-backdrop').forEach(el => {
                    el.style.display = 'none';
                    el.classList.remove('show');
                    if (el.classList.contains('modal-backdrop')) {
                        el.remove();
                    }
                });
                document.body.classList.remove('modal-open');
            }
        """)
    except Exception as js_err:
        pass

    # 3. Press Escape as final guarantee
    try:
        page.keyboard.press("Escape")
    except:
        pass

    page.wait_for_timeout(1000)


def process_single_user(page, userid, password):
    safe_userid = re.sub(r'[^\w\-]', '_', str(userid))

    # Auto accept browser dialogs
    try:
        page.on("dialog", lambda dialog: dialog.accept())
    except:
        pass

    # 1. Open Site
    print(f"Opening EPFO site: {LOGIN_URL}")
    page.goto(LOGIN_URL, wait_until="networkidle", timeout=45000)
    page.wait_for_timeout(2500)

    # 2. GUARANTEED Initial Alert Dismissal
    dismiss_all_popups(page)

    # 3. Username
    username_elem = page.locator("#username1, input[type='text']").first
    username_elem.wait_for(state="visible", timeout=10000)
    username_elem.fill(userid)
    print(f"Username filled: {userid}")
    page.wait_for_timeout(2000)

    # 4. Password
    password_elem = page.locator('[name="password"], input[type="password"]').first
    password_elem.wait_for(state="visible", timeout=10000)
    password_elem.fill(password)
    print("Password filled.")
    page.wait_for_timeout(2000)

    # Make sure popup didn't re-appear
    dismiss_all_popups(page)

    # 5. CAPTCHA Field & Screenshot
    captcha_elem = page.locator("#captcha, input[placeholder*='captcha' i]").first
    captcha_elem.wait_for(state="visible", timeout=10000)
    captcha_elem.click()

    captcha_img = page.locator("div#ToggleCaptcha img, #ToggleCaptcha img, img.captcha-img, img[src*='captcha' i]").first
    captcha_img.wait_for(state="visible", timeout=10000)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    image_path = CAPTCHA_FOLDER / f"captcha_{safe_userid}_{timestamp}.png"
    captcha_img.screenshot(path=str(image_path))
    print(f"CAPTCHA image saved → {image_path}")

    # 6. CAPTCHA Auto Fill via Ollama
    captcha_text = get_captcha_from_ollama(image_path)
    if not captcha_text:
        raise RuntimeError("Captcha reading failed (empty response from Ollama). Retrying...")
    captcha_elem.fill(captcha_text)
    print(f"CAPTCHA auto filled: {captcha_text}")

    # Increased Wait Time After CAPTCHA Fill (3.5 seconds) as requested
    page.wait_for_timeout(3500)

    # 7. Sign In
    print("Clicking Sign In button...")
    try:
        signIn_btn = page.get_by_role("button", name="Sign In").first
        signIn_btn.click(force=True, timeout=5000)
    except:
        signIn_btn = page.locator("button.btn-logging, button.btn-success, button[type='submit']").first
        signIn_btn.click(force=True)

    print("Sign In clicked.")

    # 8. Login ke baad wait (Increased wait time to 8 seconds as requested)
    page.wait_for_timeout(8000)

    # Verify login success
    dashboards_elem = page.get_by_role("button", name="Dashboards").first
    try:
        is_logged_in = dashboards_elem.is_visible(timeout=6000)
    except:
        is_logged_in = False

    if not is_logged_in:
        raise RuntimeError("Login failed (Invalid CAPTCHA or Password).")

    print("Login SUCCESSFUL!")

    # 9. GUARANTEED Post-login Notice Dismissal
    dismiss_all_popups(page)
    page.wait_for_timeout(2000)

    # 10. Dashboard
    print("Clicking Dashboards dropdown...")
    dashboards_elem.click(force=True)
    page.wait_for_timeout(3500)

    # 11. Active Member
    print("Clicking ACTIVE MEMBER link...")
    active_elem = page.get_by_role("link", name="ACTIVE MEMBER").first
    active_elem.wait_for(state="visible", timeout=10000)
    active_elem.click(force=True)
    page.wait_for_timeout(4500)

    # 12. Export / Download CSV (Robust expect_download handler)
    print("Triggering Active Members Download...")
    download_btn = page.locator('[title="Download all active members"], a[href*="activeMembersCsv"], a:has-text("Export")').first
    download_btn.wait_for(state="visible", timeout=10000)

    try:
        with page.expect_download(timeout=40000) as download_info:
            download_btn.click(force=True)
        download = download_info.value
    except Exception as d_err:
        print(f"Standard download click warning: {d_err}. Trying JS click...")
        with page.expect_download(timeout=40000) as download_info:
            download_btn.evaluate("el => el.click()")
        download = download_info.value

    temp_download_path = TARGET_DOWNLOAD_FOLDER / f"raw_{download.suggested_filename}"
    download.save_as(temp_download_path)
    print(f"Downloaded raw file captured → {temp_download_path}")

    # Extract ZIP & clean CSV into target date folder: {userid}.csv
    final_csv = extract_and_clean_downloaded_file(temp_download_path, userid, TARGET_DOWNLOAD_FOLDER)
    if not final_csv:
        raise RuntimeError("Failed to extract & format CSV file.")

    page.wait_for_timeout(2500)

    # 13. Logout
    try:
        logout_btn = page.locator("a[href*='logout'], //a[contains(text(),'Logout')]").first
        if logout_btn.is_visible(timeout=3000):
            logout_btn.click(force=True)
            print("Logged out successfully.")
            page.wait_for_timeout(3000)
    except:
        pass

    return "SUCCESS"


def main(single_user=None, single_password=None, headless=False, target_codes=None, only_failed=False):
    wb = None
    if not EXCEL_PATH.exists() and not (single_user and single_password):
        print(f"Excel file not found at: {EXCEL_PATH}")
        return

    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb.active
        total_rows = ws.max_row
        all_users = []
        for row in range(2, total_rows + 1):
            userid = ws[f"A{row}"].value
            password = ws[f"B{row}"].value
            if not userid or not password:
                continue
            all_users.append((row, clean_excel_val(userid), clean_excel_val(password)))
    else:
        all_users = []

    if single_user:
        if not single_password:
            found = next((u for u in all_users if str(u[1]).strip() == str(single_user).strip()), None)
            if found:
                user_list = [found]
                print(f"Running for Single User: {single_user} (Password loaded from Excel)")
            else:
                print(f"Password for user {single_user} not found in Excel.")
                return
        else:
            found = next((u for u in all_users if str(u[1]).strip() == str(single_user).strip()), None)
            row_idx = found[0] if found else 0
            user_list = [(row_idx, str(single_user).strip(), str(single_password).strip())]
            print(f"Running for Single User: {single_user}")
    elif target_codes:
        codes_list = [c.strip() for c in str(target_codes).split(",") if c.strip()]
        user_list = [u for u in all_users if str(u[1]).strip() in codes_list]
        print(f"Filtered for {len(user_list)} of {len(codes_list)} target code(s): {', '.join(codes_list)}")
        if not user_list:
            print("No matching codes found in Excel.")
            return
    elif only_failed:
        failed_users = []
        for u in all_users:
            status_val = str(ws[f"C{u[0]}"].value or "").upper()
            if "FAIL" in status_val or "ERROR" in status_val or "PENDING" in status_val or not status_val:
                failed_users.append(u)
        user_list = failed_users
        print(f"Filtered {len(user_list)} failed/pending code(s) for retry.")
        if not user_list:
            print("No failed/pending accounts found to retry!")
            return
    else:
        user_list = all_users
        print(f"Total Users in Excel: {len(user_list)}")

    print(f"Target Download Folder (Date-wise): {TARGET_DOWNLOAD_FOLDER}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=["--start-maximized", "--disable-popup-blocking"]
        )
        context = browser.new_context(
            no_viewport=True,
            accept_downloads=True
        )
        page = context.new_page()

        for row, userid_str, password_str in user_list:
            row_label = f"Row {row}" if row > 0 else "Single Run"
            print(f"\n{'='*60}")
            print(f"Processing {row_label} → User: {userid_str}")
            print('='*60)

            for attempt in range(1, MAX_RETRIES_PER_USER + 1):
                if attempt > 1:
                    print(f"\n--- RETRY ATTEMPT {attempt}/{MAX_RETRIES_PER_USER} for User: {userid_str} ---")

                try:
                    result = process_single_user(page, userid_str, password_str)
                    if result == "SUCCESS":
                        print(f"{row_label} Result: SUCCESS")
                        if wb and row > 0:
                            ws[f"C{row}"] = "SUCCESS"
                        break
                except Exception as err:
                    print(f"Attempt {attempt} Error for user {userid_str}: {err}")
                    if attempt < MAX_RETRIES_PER_USER:
                        print("Refreshing page & retrying user...")
                        try:
                            page.goto(LOGIN_URL, wait_until="networkidle", timeout=20000)
                        except:
                            pass
                        page.wait_for_timeout(3000)
                    else:
                        if wb and row > 0:
                            ws[f"C{row}"] = f"ERROR - {str(err)}"
                        print(f"{row_label} failed after {MAX_RETRIES_PER_USER} attempts.")

            if wb and row > 0:
                try:
                    wb.save(EXCEL_PATH)
                    print(f"Excel updated for user: {userid_str}")
                except Exception as save_err:
                    print(f"Error saving Excel file: {save_err}")

        browser.close()
        print("\nAll users processed successfully!")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="EPFO Download")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--user", type=str, default=None, help="Specific user ID / code")
    parser.add_argument("--password", type=str, default=None, help="Password for specific user")
    parser.add_argument("--codes", type=str, default=None, help="Comma-separated list of target codes")
    parser.add_argument("--only-failed", action="store_true", help="Process only failed accounts")
    cli_args = parser.parse_args()

    main(
        single_user=cli_args.user,
        single_password=cli_args.password,
        headless=cli_args.headless,
        target_codes=cli_args.codes,
        only_failed=cli_args.only_failed
    )
