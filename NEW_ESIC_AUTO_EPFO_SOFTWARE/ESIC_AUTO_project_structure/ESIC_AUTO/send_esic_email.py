import os
import sys
import logging
from pathlib import Path

# Configure paths
CURRENT_DIR = Path(__file__).resolve().parent
COMPLIANCE_ROOT = CURRENT_DIR.parents[2]

if str(COMPLIANCE_ROOT) not in sys.path:
    sys.path.insert(0, str(COMPLIANCE_ROOT))

logger = logging.getLogger("esic_auto_main")

def send_email_report():
    """
    Step 7 action for ESIC pipeline:
    Generates the latest ESIC Excel status report and emails it via Microsoft 365 Outlook (MS Graph).
    """
    try:
        import mailer
        print("[INFO] STEP 7: Generating ESIC Status & Failure Report Excel...")
        logger.info("Generating ESIC Status & Failure Report Excel...")

        print("[INFO] Sending ESIC Report email via Microsoft 365 Outlook (MS Graph)...")
        logger.info("Sending ESIC Report email via Microsoft 365 Outlook (MS Graph)...")

        result = mailer.send_failure_email_via_ms_graph(service="esic")

        success_msg = (
            f"[SUCCESS] ESIC Report email dispatched successfully! "
            f"Attached: {result.get('report_name')} | Subject: {result.get('subject')}"
        )
        print(success_msg)
        logger.info(success_msg)
        return True

    except Exception as e:
        err_msg = f"[ERROR] Step 7 (ESIC Email Notification) failed: {e}"
        print(err_msg)
        logger.error(err_msg, exc_info=True)
        return False

if __name__ == "__main__":
    success = send_email_report()
    sys.exit(0 if success else 1)
