import sys
from pathlib import Path

# Configure paths
CURRENT_DIR = Path(__file__).resolve().parent
COMPLIANCE_ROOT = CURRENT_DIR.parents[2]

if str(COMPLIANCE_ROOT) not in sys.path:
    sys.path.insert(0, str(COMPLIANCE_ROOT))

def send_email_report():
    """
    Step 7 action for EPFO pipeline:
    Generates the latest EPFO Excel status report and emails it via Microsoft 365 Outlook (MS Graph).
    """
    print()
    print("=" * 70)
    print("STEP 7 : SEND STATUS & EMAIL REPORT (EPFO)")
    print("=" * 70)

    try:
        import mailer
        print("[INFO] STEP 7: Generating EPFO Compliance Report Excel...")
        print("[INFO] Sending EPFO Report email via Microsoft 365 Outlook (MS Graph)...")

        result = mailer.send_failure_email_via_ms_graph(service="epfo")

        print()
        print("=" * 70)
        print("EPFO EMAIL SENT SUCCESSFULLY")
        print(f"Attached Report : {result.get('report_name')}")
        print(f"Subject         : {result.get('subject')}")
        print("=" * 70)
        return True

    except Exception as e:
        print()
        print("=" * 70)
        print(f"STEP 7 ERROR: {e}")
        print("=" * 70)
        return False

if __name__ == "__main__":
    success = send_email_report()
    sys.exit(0 if success else 1)
