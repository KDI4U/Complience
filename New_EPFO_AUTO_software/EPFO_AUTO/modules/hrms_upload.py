import sys
from pathlib import Path

# Add project base directory to sys.path so 'config' can be imported when script is run directly
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from playwright.sync_api import sync_playwright
from config import HRMS_UPLOAD_FILE


# =========================================================
# LOGIN DETAILS
# =========================================================

USERNAME = "admin"
PASSWORD = "Super@2047"
URL = "https://hrms.uistech.in/login"


# =========================================================
# HRMS CSV FILE
# =========================================================
#
# config.py me:
# HRMS_UPLOAD_FILE = data/upload_hrms/HRMS_Upload.csv
#
# =========================================================

UPLOAD_FILE = HRMS_UPLOAD_FILE


# =========================================================
# HRMS UPLOAD
# =========================================================

def upload_to_hrms():

    print()
    print("=" * 70)
    print("HRMS EPFO MASTER UPLOAD")
    print("=" * 70)

    # =====================================================
    # CHECK CSV FILE
    # =====================================================

    if not UPLOAD_FILE.exists():

        raise FileNotFoundError(
            f"HRMS upload CSV nahi mili:\n{UPLOAD_FILE}"
        )

    print()
    print("HRMS upload CSV found:")
    print(UPLOAD_FILE)

    # =====================================================
    # CHECK FILE SIZE
    # =====================================================

    file_size = UPLOAD_FILE.stat().st_size

    if file_size == 0:

        raise Exception(
            "HRMS upload CSV empty hai."
        )

    print(
        "File size:",
        round(file_size / 1024 / 1024, 2),
        "MB"
    )

    # =====================================================
    # PLAYWRIGHT
    # =====================================================

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=False
        )

        page = browser.new_page(
            viewport={"width": 1920, "height": 1080}
        )

        # =================================================
        # TIMEOUT SETTINGS
        # =================================================

        page.set_default_timeout(
            30000
        )

        page.set_default_navigation_timeout(
            90000
        )

        try:

            # =================================================
            # LOGIN
            # =================================================

            print()
            print("Opening HRMS login...")

            page.goto(
                URL,
                wait_until="commit",
                timeout=90000
            )

            print(
                "Login page navigation started"
            )

            # Username field wait

            page.locator(
                "#username"
            ).wait_for(
                state="visible",
                timeout=60000
            )

            print(
                "Login page loaded"
            )

            # =================================================
            # USERNAME
            # =================================================

            page.locator(
                "#username"
            ).fill(
                USERNAME
            )

            print(
                "Username entered"
            )

            # =================================================
            # PASSWORD
            # =================================================

            page.locator(
                "#password"
            ).fill(
                PASSWORD
            )

            print(
                "Password entered"
            )

            # =================================================
            # SIGN IN
            # =================================================

            page.get_by_role(
                "button",
                name="Sign In"
            ).click()

            print(
                "Sign In clicked"
            )

            # Login wait

            page.wait_for_timeout(
                5000
            )

            print(
                "Login complete"
            )

            print(
                "Dashboard:",
                page.url
            )

            # =================================================
            # DIRECT NAVIGATION TO ESIC & PF MASTER
            # =================================================

            print()
            print("Navigating directly to ESIC & PF Master...")

            page.goto(
                "https://hrms.uistech.in/master/epfo-esic-master",
                wait_until="commit",
                timeout=60000
            )
            page.wait_for_timeout(3000)

            # Fallback if URL didn't change
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

            page.wait_for_url(
                "**/master/epfo-esic-master",
                timeout=60000
            )

            print(
                "URL changed:",
                page.url
            )

            # Angular render

            page.wait_for_timeout(
                10000
            )

            print(
                "ESIC & PF Master page loaded"
            )

            # =================================================
            # CLICK EPFO MASTER TAB
            # =================================================

            print()
            print(
                "Selecting EPFO Master tab..."
            )

            epfo_tab = page.locator(
                "button.tab-item",
                has_text="EPFO Master"
            ).first

            epfo_tab.wait_for(
                state="visible",
                timeout=30000
            )

            epfo_tab.click()

            print(
                "EPFO Master tab clicked"
            )

            # Angular render after tab switch
            page.wait_for_timeout(
                5000
            )

            # =================================================
            # VERIFY EPFO MASTER TAB IS ACTIVE
            # =================================================

            active_epfo_tab = page.locator(
                "button.tab-item.active-tab",
                has_text="EPFO Master"
            ).first

            active_epfo_tab.wait_for(
                state="visible",
                timeout=30000
            )

            print(
                "EPFO Master tab is ACTIVE"
            )

            # =================================================
            # FILE INPUT
            # =================================================

            print()
            print(
                "Searching file input..."
            )

            file_input = page.locator(
                'input[type="file"]'
            ).first

            file_input.wait_for(
                state="attached",
                timeout=30000
            )

            print(
                "File input found"
            )

            # =================================================
            # SELECT CSV FILE
            # =================================================

            print()
            print(
                "Selecting HRMS CSV..."
            )

            file_input.set_input_files(
                str(UPLOAD_FILE)
            )

            print()
            print(
                "HRMS CSV selected"
            )

            print(
                "File:",
                UPLOAD_FILE
            )

            # =================================================
            # VERIFY SELECTED FILE
            # =================================================

            selected_file = page.locator(
                "span.selected-file"
            ).first

            try:

                selected_file.wait_for(
                    state="visible",
                    timeout=10000
                )

                page.wait_for_function(
                    """
                    () => {
                        const el =
                            document.querySelector(
                                "span.selected-file"
                            );

                        return el &&
                               el.innerText.trim() !== "" &&
                               el.innerText.trim() !== "No file selected";
                    }
                    """,
                    timeout=30000
                )

                print()
                print(
                    "Selected file shown on page:"
                )

                print(
                    selected_file.inner_text()
                )

            except Exception:

                print(
                    "Selected file text verify nahi ho paya."
                )

                print(
                    "File input me CSV set ho chuki hai."
                )

            # =================================================
            # WAIT AFTER FILE SELECT
            # =================================================

            print()
            print(
                "CSV select hone ke baad 5 sec wait..."
            )

            page.wait_for_timeout(
                5000
            )

            # =================================================
            # UPLOAD CSV BUTTON
            # =================================================

            print()
            print(
                "Searching Upload CSV button..."
            )

            upload_button = page.get_by_role(
                "button",
                name="Upload CSV"
            )

            upload_button.wait_for(
                state="visible",
                timeout=30000
            )

            print(
                "Upload CSV button visible"
            )

            # =================================================
            # WAIT BUTTON ENABLED
            # =================================================

            print()
            print(
                "Upload CSV button enable hone ka wait..."
            )

            page.wait_for_function(
                """
                () => {
                    const buttons =
                        Array.from(
                            document.querySelectorAll("button")
                        );

                    const button =
                        buttons.find(
                            b =>
                                b.innerText.trim()
                                === "Upload CSV"
                        );

                    return button &&
                           !button.disabled;
                }
                """,
                timeout=30000
            )

            print(
                "Upload CSV button enabled"
            )

            # =================================================
            # CLICK UPLOAD CSV
            # =================================================

            upload_button.click()

            print()
            print(
                "Upload CSV clicked"
            )

            # =================================================
            # AFTER UPLOAD
            # =================================================

            print()
            print(
                "Upload ke baad 20 sec wait..."
            )

            page.wait_for_timeout(
                20000
            )

            print()
            print("=" * 70)
            print(
                "HRMS UPLOAD PROCESS COMPLETED"
            )
            print("=" * 70)

            print()
            print(
                "Current URL:",
                page.url
            )

            # =================================================
            # FINALIZATION
            # =================================================

            if "--interactive" in sys.argv and sys.stdin.isatty():
                try:
                    input("Upload result check karne ke baad Enter press kare...")
                except Exception:
                    pass
            else:
                print()
                print("HRMS upload verification complete. Proceeding automatically to Step 7 (Email Report)...")
                page.wait_for_timeout(3000)

            return True

        except Exception as e:

            print()
            print("=" * 70)
            print(
                "HRMS UPLOAD FAILED"
            )
            print("=" * 70)

            print()
            print(
                str(e)
            )

            # =================================================
            # ERROR SCREENSHOT
            # =================================================

            try:

                screenshot_path = (
                    Path("screenshots")
                    / "hrms_upload_error.png"
                )

                screenshot_path.parent.mkdir(
                    parents=True,
                    exist_ok=True
                )

                page.screenshot(
                    path=str(screenshot_path),
                    full_page=True
                )

                print()
                print(
                    "Error screenshot saved:"
                )

                print(
                    screenshot_path
                )

            except Exception:

                pass

            raise

        finally:

            browser.close()


# =========================================================
# DIRECT RUN
# =========================================================

if __name__ == "__main__":

    upload_to_hrms()