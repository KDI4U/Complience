import sys
import subprocess
from pathlib import Path


# =========================================================
# PROJECT BASE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent


# =========================================================
# RUN MODULE
# =========================================================

def run_module(module_name, step_name):

    print()
    print("=" * 70)
    print(step_name)
    print("=" * 70)

    try:

        cmd = [sys.executable, "-m"] + module_name.split()
        result = subprocess.run(cmd, cwd=BASE_DIR)

        if result.returncode != 0:

            print()
            print("=" * 70)
            print(f"{step_name} FAILED")
            print("=" * 70)

            return False

        print()
        print("=" * 70)
        print(f"{step_name} COMPLETED")
        print("=" * 70)

        return True

    except Exception as e:

        print()
        print("=" * 70)
        print(f"{step_name} ERROR")
        print("=" * 70)

        print()
        print(str(e))

        return False


# =========================================================
# MAIN AUTOMATION
# =========================================================

def main(step_to_run=None, from_step=None, headless=False, target_codes=None, only_failed=False):

    print()
    print("=" * 70)
    print("        EPFO AUTOMATION SYSTEM")
    print("=" * 70)

    print()
    print("Automation start ho rahi hai...")

    pipeline = [
        (1, "modules.cleanup", "STEP 1 : OLD DATA CLEANUP"),
        (2, "modules.epfo_download", "STEP 2 : EPFO DOWNLOAD"),
        (3, "modules.excel_merge", "STEP 3 : EXCEL MERGE"),
        (4, "modules.db_download", "STEP 4 : DATABASE DOWNLOAD"),
        (5, "modules.compare_data", "STEP 5 : DATA COMPARISON"),
        (6, "modules.hrms_upload", "STEP 6 : HRMS UPLOAD"),
        (7, "modules.send_email", "STEP 7 : SEND STATUS & EMAIL REPORT"),
    ]

    if step_to_run:
        pipeline = [s for s in pipeline if s[0] == step_to_run]
        if not pipeline:
            print(f"Invalid step: {step_to_run}. Valid steps are 1-7.")
            return
    elif from_step:
        pipeline = [s for s in pipeline if s[0] >= from_step]
        if not pipeline:
            print(f"Invalid from-step: {from_step}. Valid steps are 1-7.")
            return

    for step_num, mod_name, step_title in pipeline:
        if step_num == 2:
            flags = []
            if headless:
                flags.append("--headless")
            if target_codes:
                flags.extend(["--codes", str(target_codes)])
            if only_failed:
                flags.append("--only-failed")
            mod_cmd = f"{mod_name} {' '.join(flags)}".strip()
            success = run_module(mod_cmd, step_title)
        else:
            success = run_module(mod_name, step_title)

        if not success:
            if step_num == 6:
                print(f"[WARNING] {step_title} had an issue, but proceeding to Step 7 to send email status report.")
                continue
            print(f"{step_title} fail hua. Process STOP.")
            return

    print()
    print()
    print("=" * 70)
    print("        EPFO AUTOMATION FINISHED")
    print("=" * 70)
    print()
    print("Executed step(s) completed successfully.")
    print("=" * 70)


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="EPFO Automation Pipeline")
    parser.add_argument("--step", "-s", type=int, choices=range(1, 8), help="Run a specific single step (1-7)")
    parser.add_argument("--from-step", "-f", type=int, choices=range(1, 8), help="Start pipeline from a specific step (e.g. 5)")
    parser.add_argument("--codes", "-c", type=str, default=None, help="Comma-separated list of User IDs to process")
    parser.add_argument("--only-failed", action="store_true", help="Only process accounts that previously failed or had errors")
    parser.add_argument("--headless", "-H", action="store_true", help="Run Playwright browser automation in headless mode")
    args = parser.parse_args()

    try:
        main(
            step_to_run=args.step,
            from_step=args.from_step,
            headless=args.headless,
            target_codes=args.codes,
            only_failed=args.only_failed
        )
    except KeyboardInterrupt:
        print()
        print()
        print("Process user ne stop kar diya.")
    except Exception as e:
        print()
        print("=" * 70)
        print("MAIN SYSTEM ERROR")
        print("=" * 70)
        print()
        print(str(e))