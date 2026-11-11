import sys
import logging
from pathlib import Path
from playwright.sync_api import sync_playwright

# D DRIVE DYNAMIC ABSOLUTE PATHS
PROJECT_FOLDER = Path(__file__).resolve().parent
UPLOAD_FILE = PROJECT_FOLDER / "compaire" / "mainupload.csv"
SCREENSHOTS_FOLDER = PROJECT_FOLDER / "screenshots"

USERNAME = "admin"
PASSWORD = "Super@2047"
URL = "https://hrms.uistech.in/login"

logger = logging.getLogger("hrms_upload")
if not logger.hasHandlers():
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(ch)
    logger.setLevel(logging.INFO)


def upload_to_esic(headless=False, auto_mode=True):
    logger.info("Starting HRMS ESIC Master Upload Automation...")

    if not UPLOAD_FILE.exists():
        raise FileNotFoundError(f"ESIC upload CSV file not found: {UPLOAD_FILE}")

    file_size = UPLOAD_FILE.stat().st_size
    if file_size == 0:
        logger.warning("ESIC upload CSV file is empty. Skipping HRMS upload.")
        return True

    logger.info(f"Uploading File: {UPLOAD_FILE.name} ({round(file_size / 1024, 2)} KB)")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page(viewport={"width": 1920, "height": 1080})

        page.set_default_timeout(30000)
        page.set_default_navigation_timeout(90000)

        try:
            logger.info("Navigating to HRMS login...")
            page.goto(URL, wait_until="commit", timeout=90000)

            page.locator("#username").wait_for(state="visible", timeout=60000)
            page.locator("#username").fill(USERNAME)
            page.locator("#password").fill(PASSWORD)
            page.get_by_role("button", name="Sign In").click()

            page.wait_for_timeout(4000)
            logger.info("Login successful. Navigating directly to ESIC & PF Master...")

            # Direct navigation avoids flaky collapsed sidebar
            page.goto("https://hrms.uistech.in/master/epfo-esic-master", wait_until="commit", timeout=60000)
            page.wait_for_timeout(3000)

            # Fallback if direct routing didn't switch URL immediately
            if "master/epfo-esic-master" not in page.url:
                try:
                    page.mouse.move(50, 300)
                    page.wait_for_timeout(500)
                    master_menu = page.locator("p.sidebarmenu-name, span.menu-title", has_text="Master").first
                    if master_menu.is_visible():
                        master_menu.click()
                        page.wait_for_timeout(1000)
                    esic_pf_master = page.locator("span.menu-title", has_text="ESIC & PF Master").first
                    if esic_pf_master.is_visible():
                        esic_pf_master.click()
                except Exception:
                    pass

            page.wait_for_url("**/master/epfo-esic-master", timeout=60000)
            page.wait_for_timeout(3000)

            esic_tab = page.locator("button.tab-item", has_text="ESIC Master").first
            esic_tab.wait_for(state="visible", timeout=30000)
            esic_tab.click()
            page.wait_for_timeout(3000)

            file_input = page.locator('input[type="file"]').first
            file_input.wait_for(state="attached", timeout=30000)
            file_input.set_input_files(str(UPLOAD_FILE))
            logger.info(f"Selected file: {UPLOAD_FILE.name}")

            page.wait_for_timeout(3000)

            upload_button = page.get_by_role("button", name="Upload CSV")
            upload_button.wait_for(state="visible", timeout=30000)

            # Wait for button to become enabled
            for _ in range(20):
                if upload_button.is_enabled():
                    break
                page.wait_for_timeout(500)

            upload_button.click()
            logger.info("Upload CSV button clicked.")

            page.wait_for_timeout(20000)
            logger.info("HRMS ESIC Master Upload process completed successfully.")

            if not auto_mode and sys.stdin.isatty():
                try:
                    input("Press Enter after checking upload result...")
                except Exception:
                    pass

            return True

        except Exception as exc:
            logger.error(f"HRMS ESIC Master Upload failed: {exc}")
            try:
                SCREENSHOTS_FOLDER.mkdir(parents=True, exist_ok=True)
                screenshot_path = SCREENSHOTS_FOLDER / "esic_upload_error.png"
                page.screenshot(path=str(screenshot_path), full_page=True)
                logger.info(f"Saved error screenshot -> {screenshot_path}")
            except Exception:
                pass
            raise
        finally:
            browser.close()


if __name__ == "__main__":
    headless_flag = "--headless" in sys.argv
    upload_to_esic(headless=headless_flag, auto_mode=True)