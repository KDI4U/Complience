import os
import sys
import time
import argparse
import logging
from pathlib import Path
from datetime import datetime

# Configure Windows console encoding safely
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ============================================================
# ESKO PRODUCTION AUTOMATION PIPELINE - ESIC AUTO MAIN
# ============================================================

# D DRIVE DYNAMIC ABSOLUTE PATHS
PROJECT_FOLDER = Path(__file__).resolve().parent
if str(PROJECT_FOLDER) not in sys.path:
    sys.path.insert(0, str(PROJECT_FOLDER))

LOGS_FOLDER = PROJECT_FOLDER / "logs"
LOGS_FOLDER.mkdir(parents=True, exist_ok=True)

# Required Script Files
CLEAN_FILE = PROJECT_FOLDER / "clean_old_esicdata.py"
PORTAL_FILE = PROJECT_FOLDER / "esic_portal_download.py"
MERGE_FILE = PROJECT_FOLDER / "merge.py"
DB_FILE = PROJECT_FOLDER / "dbdownload.py"
COMPARE_FILE = PROJECT_FOLDER / "Compare.py"
HRMS_FILE = PROJECT_FOLDER / "hrms_esiclogin.py"
MAIL_FILE = PROJECT_FOLDER / "send_esic_email.py"

HRMS_UPLOAD_FILE = PROJECT_FOLDER / "compaire" / "mainupload.csv"


def setup_logger(verbose=False):
    log_filename = LOGS_FOLDER / f"esic_auto_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    root_logger.handlers.clear()

    formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S")

    # Console Handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG if verbose else logging.INFO)
    ch.setFormatter(formatter)
    root_logger.addHandler(ch)

    # File Handler
    fh = logging.FileHandler(log_filename, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] [%(name)s:%(lineno)d]: %(message)s"))
    root_logger.addHandler(fh)

    logger = logging.getLogger("esic_auto_main")
    return logger, log_filename


def check_prerequisites(logger):
    required_files = [
        ("Cleanup Script", CLEAN_FILE),
        ("Portal Download Script", PORTAL_FILE),
        ("Merge Script", MERGE_FILE),
        ("DB Download Script", DB_FILE),
        ("Compare Script", COMPARE_FILE),
        ("HRMS Upload Script", HRMS_FILE),
        ("Email Notification Script", MAIL_FILE),
    ]

    missing = False
    for name, path in required_files:
        if path.exists():
            logger.debug(f"Found required file: {name} -> {path.name}")
        else:
            logger.error(f"Missing required file: {name} ({path})")
            missing = True

    try:
        from esic_portal_download import ensure_ollama_running
        if not ensure_ollama_running(timeout=10):
            logger.warning("Ollama AI Service is not responding. Captcha reading may fail if Ollama is not started.")
    except Exception as e:
        logger.debug(f"Ollama check note: {e}")

    return not missing


def run_step_module(step_no, total_steps, step_name, action_func, logger):
    logger.info("=" * 70)
    logger.info(f"STEP {step_no}/{total_steps} : {step_name}")
    logger.info("=" * 70)

    start_time = time.time()
    try:
        success = action_func()
        elapsed = time.time() - start_time
        if success is False:
            logger.error(f"Step {step_no} ({step_name}) FAILED after {elapsed:.2f}s")
            return False, elapsed
        logger.info(f"Step {step_no} ({step_name}) SUCCESS ({elapsed:.2f}s)")
        return True, elapsed
    except Exception as exc:
        elapsed = time.time() - start_time
        logger.error(f"Step {step_no} ({step_name}) EXCEPTION: {exc}", exc_info=True)
        return False, elapsed


def main():
    parser = argparse.ArgumentParser(
        description="ESKO Production Automation Pipeline - ESIC AUTO",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--auto", "-a", action="store_true", default=True, help="Run full pipeline automatically without blocking prompts")
    parser.add_argument("--interactive", "-i", action="store_true", help="Prompt for confirmation between steps")
    parser.add_argument("--headless", "-H", action="store_true", help="Run Playwright browser automation in headless mode")
    parser.add_argument("--step", "-s", type=int, choices=range(1, 8), help="Run a specific single step (1-7)")
    parser.add_argument("--from-step", "-f", type=int, choices=range(1, 8), help="Start pipeline from a specific step (e.g. 5)")
    parser.add_argument("--codes", "-c", type=str, default=None, help="Comma-separated list of Employer Codes to process")
    parser.add_argument("--only-failed", action="store_true", help="Only process accounts that previously failed or had errors")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose debug logging")

    args = parser.parse_args()
    auto_mode = not args.interactive

    logger, log_file = setup_logger(verbose=args.verbose)

    logger.info("=" * 70)
    logger.info("          ESKO PRODUCTION AUTOMATION PIPELINE - ESIC AUTO")
    logger.info("=" * 70)
    logger.info(f"Project Folder : {PROJECT_FOLDER}")
    logger.info(f"Execution Log  : {log_file.name}")
    logger.info(f"Execution Mode : {'AUTO (Unattended)' if auto_mode else 'INTERACTIVE'}")
    logger.info(f"Browser Mode   : {'HEADLESS' if args.headless else 'VISIBLE'}")
    if args.codes:
        logger.info(f"Target Filter  : Specific Codes ({args.codes})")
    elif args.only_failed:
        logger.info("Target Filter  : Only Failed Accounts")
    logger.info("=" * 70)

    if not check_prerequisites(logger):
        logger.error("Prerequisite check failed. Aborting pipeline.")
        sys.exit(1)

    # Dynamic Step Definitions
    def step1_action():
        from clean_old_esicdata import clean_old_data
        return clean_old_data()

    def step2_action():
        import asyncio
        from esic_portal_download import run_portal_download
        return asyncio.run(run_portal_download(
            headless=args.headless,
            auto_mode=auto_mode,
            target_codes=args.codes,
            only_failed=args.only_failed
        ))

    def step3_action():
        from merge import merge_esic_files
        return merge_esic_files()

    def step4_action():
        from dbdownload import download_db_data
        return bool(download_db_data())

    def step5_action():
        from Compare import run_comparison
        return run_comparison()

    def step6_action():
        from hrms_esiclogin import upload_to_esic
        return upload_to_esic(headless=args.headless, auto_mode=auto_mode)

    def step7_action():
        from send_esic_email import send_email_report
        return send_email_report()

    steps = [
        (1, "Clean Old Data", step1_action),
        (2, "ESIC Portal Download", step2_action),
        (3, "Merge & Column Rename", step3_action),
        (4, "PostgreSQL DB Download", step4_action),
        (5, "Data Comparison", step5_action),
        (6, "HRMS ESIC Master Upload", step6_action),
        (7, "Send Email Status Report", step7_action),
    ]

    if args.step:
        steps = [s for s in steps if s[0] == args.step]
    elif args.from_step:
        steps = [s for s in steps if s[0] >= args.from_step]

    results = []
    total_start_time = time.time()

    for step_no, step_name, step_func in steps:
        success, elapsed = run_step_module(step_no, 7, step_name, step_func, logger)
        results.append((step_no, step_name, "SUCCESS" if success else "FAILED", elapsed))

        if not success:
            if step_no == 6:
                logger.warning(f"Step 6 ({step_name}) had an issue, but proceeding to Step 7 (Send Email Status Report) so report is delivered.")
                continue
            logger.error(f"Pipeline stopped at Step {step_no}: {step_name}")
            break

        if not auto_mode and step_no < len(steps):
            if sys.stdin.isatty():
                try:
                    input(f"\nPress Enter to proceed to Step {step_no + 1}...")
                except Exception:
                    pass

    total_elapsed = time.time() - total_start_time
    all_passed = all(status == "SUCCESS" for _, _, status, _ in results)

    # Final Summary Report
    print("\n")
    logger.info("=" * 70)
    logger.info("                    PIPELINE EXECUTION SUMMARY")
    logger.info("=" * 70)
    for s_no, s_name, status, s_time in results:
        status_tag = "[ OK ]" if status == "SUCCESS" else "[ FAIL ]"
        logger.info(f"Step {s_no}: {s_name:<28} {status_tag:<8} ({s_time:.2f}s)")

    logger.info("-" * 70)
    logger.info(f"Overall Status  : {'SUCCESS' if all_passed else 'FAILED'}")
    logger.info(f"Total Time      : {total_elapsed:.2f} seconds")
    logger.info(f"Log File Saved  : {log_file}")
    logger.info("=" * 70)

    if not auto_mode and sys.stdin.isatty():
        try:
            input("\nExecution completed. Press Enter to exit...")
        except Exception:
            pass

    if not all_passed:
        sys.exit(1)


if __name__ == "__main__":
    main()
