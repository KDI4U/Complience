import os
import sys
import json
import base64
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Tuple

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import msal
import requests

from dotenv import load_dotenv

# Base paths
COMPLIANCE_ROOT = Path(r"D:\Complience")
ENV_PATH = COMPLIANCE_ROOT / ".env"
if ENV_PATH.exists():
    load_dotenv(ENV_PATH)

CONFIG_PATH = COMPLIANCE_ROOT / "ms_graph_config.json"
RECIPIENTS_EXCEL = COMPLIANCE_ROOT / "email_recipients.xlsx"
REPORTS_DIR = COMPLIANCE_ROOT / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

ESIC_EXCEL = COMPLIANCE_ROOT / "NEW_ESIC_AUTO_EPFO_SOFTWARE" / "ESIC_AUTO_project_structure" / "esic_users.xlsx"
if not ESIC_EXCEL.exists():
    _alt = COMPLIANCE_ROOT / "NEW_ESIC_AUTO_EPFO_SOFTWARE" / "ESIC_AUTO_project_structure" / "ESIC_AUTO" / "esic_users.xlsx"
    if _alt.exists():
        ESIC_EXCEL = _alt

EPFO_EXCEL = COMPLIANCE_ROOT / "New_EPFO_AUTO_software" / "EPFO_AUTO" / "data" / "users.xlsx"

logger = logging.getLogger("mailer")


def get_recipients() -> Dict[str, List[str]]:
    """Reads recipients from email_recipients.xlsx with fallback to config json or .env."""
    to_list = []
    cc_list = []

    if RECIPIENTS_EXCEL.exists():
        try:
            wb = openpyxl.load_workbook(RECIPIENTS_EXCEL, data_only=True)
            ws = wb.active
            for row in range(2, ws.max_row + 1):
                rec_type = str(ws.cell(row, 1).value or "").strip().upper()
                email = str(ws.cell(row, 2).value or "").strip()
                active = str(ws.cell(row, 4).value or "YES").strip().upper()

                if not email or "@" not in email:
                    continue
                if active == "NO":
                    continue

                if rec_type == "TO":
                    to_list.append(email)
                elif rec_type == "CC":
                    cc_list.append(email)
        except Exception as e:
            logger.warning(f"Error reading recipients from Excel: {e}")

    # Fallback if Excel was empty
    if not to_list:
        to_list = ["arun_mishra@uistech.in"]
    if not cc_list:
        cc_list = ["kundan_it@uistech.in", "Shiv.mangal@uistech.in", "Amresh_Sinha@uistech.in"]

    return {"to": list(dict.fromkeys(to_list)), "cc": list(dict.fromkeys(cc_list))}


def save_recipients_to_excel(to_list: List[str], cc_list: List[str]):
    """Saves updated recipients back to email_recipients.xlsx."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Recipients"
    ws.append(["Type", "Email", "Name", "Active"])

    for em in to_list:
        ws.append(["TO", em.strip(), "", "YES"])
    for em in cc_list:
        ws.append(["CC", em.strip(), "", "YES"])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(RECIPIENTS_EXCEL)


def get_ms_graph_token() -> str:
    """
    Acquires a valid MS Graph access token.
    Prioritizes .env environment variables; falls back to ms_graph_config.json.
    """
    client_id = os.getenv("MS_GRAPH_CLIENT_ID")
    tenant_id = os.getenv("MS_GRAPH_TENANT_ID")
    env_cache_raw = os.getenv("MS_GRAPH_TOKEN_CACHE_JSON")

    token_cache_dict = None

    if client_id and tenant_id and env_cache_raw:
        try:
            cleaned_cache = env_cache_raw.strip()
            if (cleaned_cache.startswith("'") and cleaned_cache.endswith("'")) or \
               (cleaned_cache.startswith('"') and cleaned_cache.endswith('"')):
                cleaned_cache = cleaned_cache[1:-1]
            token_cache_dict = json.loads(cleaned_cache)
        except Exception as e:
            logger.warning(f"Could not parse MS_GRAPH_TOKEN_CACHE_JSON from .env: {e}")

    # Fallback to json if not found in .env
    if not token_cache_dict and CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        client_id = client_id or cfg.get("client_id")
        tenant_id = tenant_id or cfg.get("tenant_id")
        token_cache_dict = cfg.get("token_cache", {})

    if not client_id or not tenant_id or not token_cache_dict:
        raise RuntimeError("MS Graph configuration not found in .env or ms_graph_config.json")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["Mail.Send", "User.Read"]

    cache = msal.SerializableTokenCache()
    cache.deserialize(json.dumps(token_cache_dict))

    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)
    accounts = app.get_accounts()

    token_res = None
    if accounts:
        token_res = app.acquire_token_silent(scopes, account=accounts[0])

    if not token_res or "access_token" not in token_res:
        rf_items = list(token_cache_dict.get("RefreshToken", {}).values())
        if rf_items:
            rf = rf_items[0].get("secret")
            if rf:
                token_res = app.acquire_token_by_refresh_token(rf, scopes=scopes)

    if not token_res or "access_token" not in token_res:
        raise RuntimeError(f"Could not acquire MS Graph token: {token_res}")

    # If cache state changed, update json config cache
    if cache.has_state_changed and CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                c_json = json.load(f)
            c_json["token_cache"] = json.loads(cache.serialize())
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(c_json, f, indent=2)
        except Exception:
            pass

    return token_res["access_token"]


def collect_service_accounts(service: str = "esic") -> Tuple[List[Dict[str, Any]], int, int, int]:
    """
    Scans specified service Excel (ESIC or EPFO).
    Returns (failed_and_pending_records, total_count, success_count, fail_count).
    """
    records = []
    excel_path = ESIC_EXCEL if service.lower() == "esic" else EPFO_EXCEL
    svc_label = "ESIC" if service.lower() == "esic" else "EPFO"

    total_count = 0
    success_count = 0
    fail_count = 0

    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        for r in range(2, ws.max_row + 1):
            raw_id = ws.cell(r, 1).value
            raw_resp = ws.cell(r, 3).value
            if raw_id is None:
                continue
            total_count += 1
            loginid = str(int(raw_id)) if isinstance(raw_id, float) and raw_id.is_integer() else str(raw_id).strip()
            resp = str(raw_resp or "").strip()

            is_success = "SUCCESS" in resp.upper()
            if is_success:
                success_count += 1
            else:
                fail_count += 1
                status_label = "FAILED" if ("FAIL" in resp.upper() or "ERROR" in resp.upper()) else "NOT DOWNLOADED"
                records.append({
                    "service": svc_label,
                    "row": r,
                    "loginid": loginid,
                    "status": status_label,
                    "details": resp if resp else "Not processed / Pending download"
                })

    return records, total_count, success_count, fail_count


def collect_failures_and_pending() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Scans ESIC and EPFO user excels for any failed or non-downloaded accounts."""
    esic_records, _, _, _ = collect_service_accounts("esic")
    epfo_records, _, _, _ = collect_service_accounts("epfo")
    return esic_records, epfo_records


def generate_service_report_excel(service: str = "all") -> Path:
    """
    Generates a styled Excel report for ESIC, EPFO, or both (all).
    """
    timestamp_str = datetime.now().strftime("%d%m%Y_%H%M%S")
    now_fmt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_fail = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_pending = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_fail = Font(color="991B1B", bold=True)
    font_pending = Font(color="92400E", bold=True)
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    svc_lower = service.lower()

    if svc_lower in ("esic", "epfo"):
        records, total_c, succ_c, fail_c = collect_service_accounts(svc_lower)
        svc_name = "ESIC" if svc_lower == "esic" else "EPFO"
        theme_color = "0E7490" if svc_lower == "esic" else "6D28D9"
        report_filename = f"{svc_name}_Compliance_Report_{timestamp_str}.xlsx"
        report_path = REPORTS_DIR / report_filename

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"{svc_name}_Fail_Pending"

        headers = ["S.No", f"{svc_name} Code / ID", "Download Status", "Failure / Error Details", "Report Timestamp"]
        ws.append(headers)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = font_header
            cell.fill = PatternFill(start_color=theme_color, end_color=theme_color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_thin

        for idx, rec in enumerate(records, start=1):
            row_num = idx + 1
            ws.append([idx, rec["loginid"], rec["status"], rec["details"], now_fmt])

            status_cell = ws.cell(row_num, 3)
            if rec["status"] == "FAILED":
                status_cell.fill = fill_fail
                status_cell.font = font_fail
            else:
                status_cell.fill = fill_pending
                status_cell.font = font_pending

            for c in range(1, 6):
                ws.cell(row_num, c).border = border_thin
                if c in (1, 3):
                    ws.cell(row_num, c).alignment = Alignment(horizontal="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        wb.save(report_path)
        return report_path

    # Otherwise ALL / COMBINED
    return generate_failure_report_excel()


def generate_failure_report_excel() -> Path:
    """Generates a fresh styled Excel workbook with all failed and not-downloaded records."""
    esic_records, epfo_records = collect_failures_and_pending()
    all_records = esic_records + epfo_records

    timestamp_str = datetime.now().strftime("%d%m%Y_%H%M%S")
    report_filename = f"Failed_And_Not_Downloaded_Report_{timestamp_str}.xlsx"
    report_path = REPORTS_DIR / report_filename

    wb = openpyxl.Workbook()
    ws_all = wb.active
    ws_all.title = "All_Failed_Accounts"

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_fail = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_pending = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_fail = Font(color="991B1B", bold=True)
    font_pending = Font(color="92400E", bold=True)
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = [
        "S.No",
        "Compliance Service",
        "Employer Code / User ID",
        "Download Status",
        "Failure / Error Details",
        "Report Generated At"
    ]
    ws_all.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws_all.cell(1, col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    now_fmt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for idx, rec in enumerate(all_records, start=1):
        row_num = idx + 1
        ws_all.append([
            idx,
            rec["service"],
            rec["loginid"],
            rec["status"],
            rec["details"],
            now_fmt
        ])

        status_cell = ws_all.cell(row_num, 4)
        if rec["status"] == "FAILED":
            status_cell.fill = fill_fail
            status_cell.font = font_fail
        else:
            status_cell.fill = fill_pending
            status_cell.font = font_pending

        for c in range(1, 7):
            cell = ws_all.cell(row_num, c)
            cell.border = border_thin
            if c in (1, 2, 4):
                cell.alignment = Alignment(horizontal="center")

    # Sheet 2: ESIC Specific
    ws_esic = wb.create_sheet(title="ESIC_Failed")
    ws_esic.append(["S.No", "ESIC Employer Code", "Status", "Details", "Timestamp"])
    for col_idx in range(1, 6):
        c = ws_esic.cell(1, col_idx)
        c.font = font_header
        c.fill = PatternFill(start_color="0E7490", end_color="0E7490", fill_type="solid")
        c.border = border_thin

    for idx, rec in enumerate(esic_records, start=1):
        ws_esic.append([idx, rec["loginid"], rec["status"], rec["details"], now_fmt])

    # Sheet 3: EPFO Specific
    ws_epfo = wb.create_sheet(title="EPFO_Failed")
    ws_epfo.append(["S.No", "EPFO User ID", "Status", "Details", "Timestamp"])
    for col_idx in range(1, 6):
        c = ws_epfo.cell(1, col_idx)
        c.font = font_header
        c.fill = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")
        c.border = border_thin

    for idx, rec in enumerate(epfo_records, start=1):
        ws_epfo.append([idx, rec["loginid"], rec["status"], rec["details"], now_fmt])

    for ws in [ws_all, ws_esic, ws_epfo]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(report_path)

    latest_path = REPORTS_DIR / "Latest_Failed_Report.xlsx"
    try:
        wb.save(latest_path)
    except Exception:
        pass

    return report_path


def send_failure_email_via_ms_graph(service: str = "all", custom_report_path: Path = None) -> Dict[str, Any]:
    """
    Sends compliance automation report via MS Graph Outlook.
    Supports service = 'esic' (ESIC only), 'epfo' (EPFO only), or 'all' (combined).
    """
    recipients = get_recipients()
    to_list = recipients["to"]
    cc_list = recipients["cc"]

    if not to_list:
        raise ValueError("No recipient email found in configuration.")

    svc_lower = service.lower()
    now_str = datetime.now().strftime("%d-%b-%Y %I:%M %p")
    today_date = datetime.now().strftime("%d-%m-%Y")

    # Generate or get report path
    if custom_report_path and custom_report_path.exists():
        report_path = custom_report_path
    else:
        report_path = generate_service_report_excel(svc_lower)

    # Read attachment
    with open(report_path, "rb") as f:
        file_b64 = base64.b64encode(f.read()).decode("utf-8")

    token = get_ms_graph_token()

    if svc_lower == "esic":
        records, total_c, succ_c, fail_c = collect_service_accounts("esic")
        accent_color = "#0891b2"
        badge_bg = "#ecfeff"
        badge_border = "#a5f3fc"
        service_label = "ESIC"

        if fail_c == 0:
            subject = f"✅ [Success] ESIC Automation Completed & HRMS Upload Finished ({today_date})"
        else:
            subject = f"⚠️ [Alert] ESIC Automation: {fail_c} Accounts Failed / Not Downloaded ({today_date})"

        rows_html = ""
        for idx, rec in enumerate(records[:30], start=1):
            badge = "badge-fail" if rec["status"] == "FAILED" else "badge-pending"
            rows_html += f"""
            <tr>
              <td>{idx}</td>
              <td><strong>{rec['loginid']}</strong></td>
              <td><span class="{badge}">{rec['status']}</span></td>
              <td style="color: #64748b; font-size: 12px;">{rec['details']}</td>
            </tr>
            """

        status_msg = f"<p style='color: #047857; font-weight: 600;'>🎉 All {total_c} ESIC accounts downloaded and processed successfully. Data comparison & HRMS upload completed!</p>" if fail_c == 0 else f"<p>Please find attached the fresh report of <strong>{fail_c}</strong> ESIC account(s) that require attention or re-running.</p>"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1e293b; background-color: #f8fafc; margin: 0; padding: 20px; }}
            .card {{ background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; max-width: 680px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }}
            .header {{ background-color: #0e7490; background: #0e7490 linear-gradient(135deg, #0e7490, #155e75); color: #ffffff; padding: 24px 20px; text-align: center; }}
            .header h2 {{ margin: 0; font-size: 20px; font-weight: 700; letter-spacing: 0.5px; color: #ffffff; }}
            .header p {{ margin: 8px 0 0; font-size: 12px; color: #ffffff; }}
            .timestamp-badge {{ display: inline-block; background-color: #083344; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #164e63; }}
            .body {{ padding: 24px; }}
            .table {{ width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 13px; }}
            .table th {{ background: #f8fafc; padding: 10px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #475569; }}
            .table td {{ padding: 10px; border-bottom: 1px solid #f1f5f9; }}
            .badge-fail {{ background: #fee2e2; color: #991b1b; padding: 3px 8px; border-radius: 9999px; font-weight: 600; font-size: 11px; }}
            .badge-pending {{ background: #fef3c7; color: #92400e; padding: 3px 8px; border-radius: 9999px; font-weight: 600; font-size: 11px; }}
            .footer {{ padding: 16px 24px; background: #f8fafc; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center; }}
          </style>
        </head>
        <body>
          <div class="card">
            <div class="header" style="background-color: #0e7490; background: #0e7490 linear-gradient(135deg, #0e7490, #155e75); color: #ffffff; padding: 24px 20px; text-align: center; border-radius: 12px 12px 0 0;">
              <h2 style="margin: 0 0 10px 0; font-size: 20px; font-weight: 700; color: #ffffff; letter-spacing: 0.5px;">🏢 ESIC Automation &amp; HRMS Compliance Report</h2>
              <div style="margin-top: 8px;">
                <span class="timestamp-badge" style="display: inline-block; background-color: #083344; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #164e63;">
                  Generated at: {now_str} &bull; Pipeline Run Complete
                </span>
              </div>
            </div>
            <div class="body">
              <p>Dear Team,</p>
              {status_msg}

              <table style="width: 100%; margin: 18px 0; border-collapse: collapse;">
                <tr>
                  <td style="padding: 12px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #15803d;">{succ_c}</div>
                    <div style="font-size: 11px; color: #166534; font-weight: 600; text-transform: uppercase;">Successful / Done</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 12px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #b91c1c;">{fail_c}</div>
                    <div style="font-size: 11px; color: #991b1b; font-weight: 600; text-transform: uppercase;">Failed / Pending</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 12px; background: #ecfeff; border: 1px solid #a5f3fc; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #0891b2;">{total_c}</div>
                    <div style="font-size: 11px; color: #0e7490; font-weight: 600; text-transform: uppercase;">Total ESIC Accounts</div>
                  </td>
                </tr>
              </table>

              {"<table class='table'><thead><tr><th>#</th><th>ESIC Code</th><th>Status</th><th>Details</th></tr></thead><tbody>" + rows_html + "</tbody></table>" if rows_html else ""}

              <p style="margin-top: 18px;"><strong>Attachment:</strong> <code>{report_path.name}</code> contains complete ESIC account log breakdown.</p>
              <p style="font-size: 12px; color: #64748b;">Dashboard URL: <code>http://localhost:8008</code> (Select ESIC tab to view live database).</p>
            </div>
            <div class="footer">
              Automated ESIC Pipeline Report &bull; Sent via Microsoft 365 Outlook
            </div>
          </div>
        </body>
        </html>
        """

    elif svc_lower == "epfo":
        records, total_c, succ_c, fail_c = collect_service_accounts("epfo")
        service_label = "EPFO"

        if fail_c == 0:
            subject = f"✅ [Success] EPFO Automation Completed & HRMS Upload Finished ({today_date})"
        else:
            subject = f"⚠️ [Alert] EPFO Automation: {fail_c} Accounts Failed / Not Downloaded ({today_date})"

        rows_html = ""
        for idx, rec in enumerate(records[:30], start=1):
            badge = "badge-fail" if rec["status"] == "FAILED" else "badge-pending"
            rows_html += f"""
            <tr>
              <td>{idx}</td>
              <td><strong>{rec['loginid']}</strong></td>
              <td><span class="{badge}">{rec['status']}</span></td>
              <td style="color: #64748b; font-size: 12px;">{rec['details']}</td>
            </tr>
            """

        status_msg = f"<p style='color: #047857; font-weight: 600;'>🎉 All {total_c} EPFO accounts downloaded and processed successfully. Data comparison & HRMS upload completed!</p>" if fail_c == 0 else f"<p>Please find attached the fresh report of <strong>{fail_c}</strong> EPFO account(s) that require attention or re-running.</p>"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1e293b; background-color: #f8fafc; margin: 0; padding: 20px; }}
            .card {{ background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; max-width: 680px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }}
            .header {{ background-color: #5b21b6; background: #5b21b6 linear-gradient(135deg, #6d28d9, #5b21b6); color: #ffffff; padding: 24px 20px; text-align: center; }}
            .header h2 {{ margin: 0; font-size: 20px; font-weight: 700; letter-spacing: 0.5px; color: #ffffff; }}
            .header p {{ margin: 8px 0 0; font-size: 12px; color: #ffffff; }}
            .timestamp-badge {{ display: inline-block; background-color: #3b0764; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #6b21a8; }}
            .body {{ padding: 24px; }}
            .table {{ width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 13px; }}
            .table th {{ background: #f8fafc; padding: 10px; border-bottom: 2px solid #e2e8f0; text-align: left; color: #475569; }}
            .table td {{ padding: 10px; border-bottom: 1px solid #f1f5f9; }}
            .badge-fail {{ background: #fee2e2; color: #991b1b; padding: 3px 8px; border-radius: 9999px; font-weight: 600; font-size: 11px; }}
            .badge-pending {{ background: #fef3c7; color: #92400e; padding: 3px 8px; border-radius: 9999px; font-weight: 600; font-size: 11px; }}
            .footer {{ padding: 16px 24px; background: #f8fafc; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center; }}
          </style>
        </head>
        <body>
          <div class="card">
            <div class="header" style="background-color: #5b21b6; background: #5b21b6 linear-gradient(135deg, #6d28d9, #5b21b6); color: #ffffff; padding: 24px 20px; text-align: center; border-radius: 12px 12px 0 0;">
              <h2 style="margin: 0 0 10px 0; font-size: 20px; font-weight: 700; color: #ffffff; letter-spacing: 0.5px;">🏦 EPFO Automation &amp; HRMS Compliance Report</h2>
              <div style="margin-top: 8px;">
                <span class="timestamp-badge" style="display: inline-block; background-color: #3b0764; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #6b21a8;">
                  Generated at: {now_str} &bull; Pipeline Run Complete
                </span>
              </div>
            </div>
            <div class="body">
              <p>Dear Team,</p>
              {status_msg}

              <table style="width: 100%; margin: 18px 0; border-collapse: collapse;">
                <tr>
                  <td style="padding: 12px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #15803d;">{succ_c}</div>
                    <div style="font-size: 11px; color: #166534; font-weight: 600; text-transform: uppercase;">Successful / Done</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 12px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #b91c1c;">{fail_c}</div>
                    <div style="font-size: 11px; color: #991b1b; font-weight: 600; text-transform: uppercase;">Failed / Pending</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 12px; background: #faf5ff; border: 1px solid #e9d5ff; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #7c3aed;">{total_c}</div>
                    <div style="font-size: 11px; color: #6d28d9; font-weight: 600; text-transform: uppercase;">Total EPFO Accounts</div>
                  </td>
                </tr>
              </table>

              {"<table class='table'><thead><tr><th>#</th><th>EPFO User ID</th><th>Status</th><th>Details</th></tr></thead><tbody>" + rows_html + "</tbody></table>" if rows_html else ""}

              <p style="margin-top: 18px;"><strong>Attachment:</strong> <code>{report_path.name}</code> contains complete EPFO account log breakdown.</p>
              <p style="font-size: 12px; color: #64748b;">Dashboard URL: <code>http://localhost:8008</code> (Select EPFO tab to view live database).</p>
            </div>
            <div class="footer">
              Automated EPFO Pipeline Report &bull; Sent via Microsoft 365 Outlook
            </div>
          </div>
        </body>
        </html>
        """

    else:
        # COMBINED ALL
        esic_records, epfo_records = collect_failures_and_pending()
        total_failures = len(esic_records) + len(epfo_records)
        service_label = "ALL"
        subject = f"⚠️ [Alert] Compliance Auto: {total_failures} Accounts Failed / Not Downloaded ({today_date})"

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
          <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1e293b; background-color: #f8fafc; margin: 0; padding: 20px; }}
            .card {{ background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; max-width: 650px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }}
            .header {{ background-color: #0f172a; background: #0f172a linear-gradient(135deg, #0f172a, #1e293b); color: #ffffff; padding: 24px 20px; text-align: center; }}
            .header h2 {{ margin: 0; font-size: 20px; font-weight: 700; color: #ffffff; }}
            .header p {{ margin: 8px 0 0; font-size: 12px; color: #ffffff; }}
            .timestamp-badge {{ display: inline-block; background-color: #1e293b; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #334155; }}
            .body {{ padding: 24px; }}
            .footer {{ padding: 16px 24px; background: #f8fafc; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center; }}
          </style>
        </head>
        <body>
          <div class="card">
            <div class="header" style="background-color: #0f172a; background: #0f172a linear-gradient(135deg, #0f172a, #1e293b); color: #ffffff; padding: 24px 20px; text-align: center; border-radius: 12px 12px 0 0;">
              <h2 style="margin: 0 0 10px 0; font-size: 20px; font-weight: 700; color: #ffffff; letter-spacing: 0.5px;">⚖️ ESIC &amp; EPFO Automation Status Report</h2>
              <div style="margin-top: 8px;">
                <span class="timestamp-badge" style="display: inline-block; background-color: #1e293b; color: #ffffff; padding: 6px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; letter-spacing: 0.3px; border: 1px solid #334155;">
                  Generated at: {now_str} &bull; Drive: D:\\Complience
                </span>
              </div>
            </div>
            <div class="body">
              <p>Dear Team,</p>
              <p>Please find attached the fresh Excel report of all accounts that <strong>FAILED</strong> or were <strong>NOT DOWNLOADED</strong> during the latest compliance automation run.</p>
              
              <table style="width: 100%; margin: 16px 0;">
                <tr>
                  <td style="padding: 10px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #b91c1c;">{len(esic_records)}</div>
                    <div style="font-size: 11px; color: #7f1d1d; font-weight: 600;">ESIC Issues</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 10px; background: #faf5ff; border: 1px solid #e9d5ff; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #6d28d9;">{len(epfo_records)}</div>
                    <div style="font-size: 11px; color: #581c87; font-weight: 600;">EPFO Issues</div>
                  </td>
                  <td style="width: 10px;"></td>
                  <td style="padding: 10px; background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 8px; text-align: center; width: 33%;">
                    <div style="font-size: 22px; font-weight: bold; color: #0f172a;">{total_failures}</div>
                    <div style="font-size: 11px; color: #334155; font-weight: 600;">Total Unprocessed</div>
                  </td>
                </tr>
              </table>

              <p><strong>Attachment:</strong> <code>{report_path.name}</code> contains complete account-wise breakdown and error logs.</p>
            </div>
            <div class="footer">
              Automated Alert by Compliance Auto Suite &bull; Sent via Microsoft 365 Outlook
            </div>
          </div>
        </body>
        </html>
        """

    message_payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_content
            },
            "toRecipients": [{"emailAddress": {"address": email}} for email in to_list],
            "ccRecipients": [{"emailAddress": {"address": email}} for email in cc_list],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": report_path.name,
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": file_b64
                }
            ]
        },
        "saveToSentItems": "true"
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers=headers,
        json=message_payload,
        timeout=35
    )

    if response.status_code in (200, 202):
        logger.info(f"[{service_label}] Report email sent successfully to {to_list} (CC: {cc_list})")
        return {
            "success": True,
            "service": service_label,
            "message": f"Email successfully sent to {', '.join(to_list)} (CC: {', '.join(cc_list)})",
            "report_file": str(report_path),
            "report_name": report_path.name,
            "subject": subject
        }
    else:
        err_msg = f"MS Graph sendMail failed (Status {response.status_code}): {response.text}"
        logger.error(err_msg)
        raise RuntimeError(err_msg)


if __name__ == "__main__":
    print("Testing Separate Service Report Generation...")
    rep_esic = generate_service_report_excel("esic")
    print("ESIC Report generated at:", rep_esic)
    rep_epfo = generate_service_report_excel("epfo")
    print("EPFO Report generated at:", rep_epfo)

