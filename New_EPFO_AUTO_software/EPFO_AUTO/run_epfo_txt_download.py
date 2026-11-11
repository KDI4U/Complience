import os
import re
import csv
import zipfile
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
import ollama

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "data" / "users.xlsx"
CAPTCHA_FOLDER = BASE_DIR / "data" / "captcha_images"
EPFO_DOWNLOADS_TXT_DIR = Path(r"D:\epfo_downloads_txt")

OLLAMA_MODEL = "gemma3:4b"
MAX_RETRIES_PER_USER = 3
LOGIN_URL = "https://unifiedportal-emp.epfindia.gov.in/epfo/"

# Create required folders
CAPTCHA_FOLDER.mkdir(parents=True, exist_ok=True)
EPFO_DOWNLOADS_TXT_DIR.mkdir(parents=True, exist_ok=True)


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


def fix_captcha_pattern(raw_text):
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

    # Pos 1 & 2 -> Uppercase
    for i in (0, 1):
        if chars[i].isdigit():
            chars[i] = digit_to_letter.get(chars[i], 'A').upper()
        else:
            chars[i] = chars[i].upper()

    # Pos 3 & 4 -> Lowercase
    for i in (2, 3):
        if chars[i].isdigit():
            chars[i] = digit_to_letter.get(chars[i], 'a').lower()
        else:
            chars[i] = chars[i].lower()

    # Pos 5 -> Digit
    if not chars[4].isdigit():
        chars[4] = letter_to_digit.get(chars[4], '1')

    return "".join(chars)


def get_captcha_from_ollama(image_path):
    try:
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


def dismiss_all_popups(page):
    print("Checking and closing all active alert popups...")
    page.wait_for_timeout(1000)

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
    print(f"\nOpening EPFO site: {LOGIN_URL}")
    page.goto(LOGIN_URL, wait_until="networkidle", timeout=45000)
    page.wait_for_timeout(2500)

    # 2. Initial Alert Dismissal
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
    page.wait_for_timeout(8000)

    # Verify login success
    dashboards_elem = page.get_by_role("button", name="Dashboards").first
    try:
        is_logged_in = dashboards_elem.is_visible(timeout=6000)
    except:
        is_logged_in = False

    if not is_logged_in:
        raise RuntimeError("Login failed (Invalid CAPTCHA or Password).")

    print("\n>>> Login SUCCESSFUL! <<<")
    dismiss_all_popups(page)
    page.wait_for_timeout(2000)

    # Create User folder in D:\epfo_downloads_txt\<userid>
    user_txt_folder = EPFO_DOWNLOADS_TXT_DIR / safe_userid
    user_txt_folder.mkdir(parents=True, exist_ok=True)
    print(f"Created user folder → {user_txt_folder}")

    # 8. Click Payments dropdown menu
    print("Clicking Payments dropdown...")
    dismiss_all_popups(page)
    try:
        payments_dropdown = page.locator("a.dropdown-toggle:has-text('Payments'), li.dropdown:has-text('Payments') > a, a[role='button']:has-text('Payments')").first
        payments_dropdown.wait_for(state="visible", timeout=10000)
        payments_dropdown.click(force=True)
    except Exception as pe:
        print(f"Warning clicking Payments dropdown: {pe}")
        page.locator("text=Payments").first.click(force=True)

    page.wait_for_timeout(3000)

    # 9. Click 1st dropdown option: Return Filing (Quick Links)
    print("Clicking 1st dropdown option: Return Filing (Quick Links)...")
    try:
        quick_links_opt = page.locator("a[href*='returnEstHomePage'], a:has-text('RETURN FILING (QUICK LINKS)'), a:has-text('Return Filing (Quick Links)')").first
        quick_links_opt.wait_for(state="visible", timeout=10000)
        quick_links_opt.click(force=True)
    except Exception as qe:
        print(f"Warning clicking Return Filing Quick Links: {qe}")
        page.locator("ul.dropdown-menu a.dropdown-item").first.click(force=True)

    page.wait_for_timeout(4000)
    dismiss_all_popups(page)

    # 10. Click Return Filing History link
    print("Clicking Return Filing History link...")
    try:
        history_link = page.locator("a[href*='loadReturnFilingHistoryPage'], a:has-text('Return Filing History'), a[title*='return filing history']").first
        history_link.wait_for(state="visible", timeout=10000)
        history_link.click(force=True)
    except Exception as he:
        print(f"Warning clicking Return Filing History link: {he}")
        page.locator("text=Return Filing History").first.click(force=True)

    page.wait_for_timeout(4000)
    dismiss_all_popups(page)

    # 11. 10 time Tab click
    print("Pressing Tab key 10 times with wait time...")
    for i in range(1, 11):
        page.keyboard.press("Tab")
        print(f"  -> Tab {i}/10 pressed")
        page.wait_for_timeout(500)

    # 12. Wait 2 seconds
    print("Waiting 2 seconds...")
    page.wait_for_timeout(2000)

    # 13. Enter press
    print("Pressing Enter key...")
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)

    # 14. Search pe click
    print("Clicking Search button...")
    try:
        search_btn = page.locator("#searchReturnFileBtn, button[value='Submit']:has-text('Search'), button:has-text('Search')").first
        if search_btn.is_visible(timeout=5000):
            search_btn.click(force=True)
        else:
            page.keyboard.press("Enter")
    except Exception as se:
        print(f"Warning clicking Search button: {se}")
        page.keyboard.press("Enter")

    page.wait_for_timeout(4000)
    dismiss_all_popups(page)

    # Download any resulting TXT / ECR / summary files into user_txt_folder
    try:
        txt_download_btns = page.locator("a[href*='txt' i], a[href*='download' i], a[href*='ecr' i], button:has-text('Download')").all()
        for idx, btn in enumerate(txt_download_btns):
            if btn.is_visible():
                try:
                    with page.expect_download(timeout=15000) as txt_d_info:
                        btn.click(force=True)
                    txt_download = txt_d_info.value
                    txt_file_path = user_txt_folder / txt_download.suggested_filename
                    txt_download.save_as(txt_file_path)
                    print(f"Saved downloaded file ({idx+1}) → {txt_file_path}")
                except Exception as dl_err:
                    print(f"Could not download file from button {idx+1}: {dl_err}")
    except Exception as ecr_err:
        print(f"No additional text files downloaded: {ecr_err}")

    page.wait_for_timeout(2500)
    return "SUCCESS"


def main():
    if not EXCEL_PATH.exists():
        print(f"Excel file not found at: {EXCEL_PATH}")
        return

    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active

    total_rows = ws.max_row
    print(f"Total Rows in Excel: {total_rows - 1}")
    print(f"Target Download Folder: {EPFO_DOWNLOADS_TXT_DIR}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-popup-blocking"]
        )
        context = browser.new_context(
            no_viewport=True,
            accept_downloads=True
        )
        page = context.new_page()

        for row in range(2, total_rows + 1):
            userid = ws[f"A{row}"].value
            password = ws[f"B{row}"].value

            if not userid or not password:
                continue

            userid_str = clean_excel_val(userid)
            password_str = clean_excel_val(password)

            print(f"\n{'='*60}")
            print(f"Processing Row {row}/{total_rows} → User: {userid_str}")
            print('='*60)

            for attempt in range(1, MAX_RETRIES_PER_USER + 1):
                if attempt > 1:
                    print(f"\n--- RETRY ATTEMPT {attempt}/{MAX_RETRIES_PER_USER} for User: {userid_str} ---")

                try:
                    result = process_single_user(page, userid_str, password_str)
                    if result == "SUCCESS":
                        print(f"Row {row} Result: SUCCESS")
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

        browser.close()
        print("\nProcess finished!")


if __name__ == "__main__":
    main()
