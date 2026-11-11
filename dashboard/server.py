import os
import sys
import time
import json
import asyncio
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any

import openpyxl
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Add parent directory to sys.path to import mailer
COMPLIANCE_ROOT = Path(r"D:\Complience")
if str(COMPLIANCE_ROOT) not in sys.path:
    sys.path.insert(0, str(COMPLIANCE_ROOT))

try:
    import mailer
except ImportError:
    from mailer import (
        get_recipients, save_recipients_to_excel, generate_failure_report_excel,
        send_failure_email_via_ms_graph, collect_failures_and_pending
    )

# ESIC PATHS
ESIC_ROOT = COMPLIANCE_ROOT / "NEW_ESIC_AUTO_EPFO_SOFTWARE" / "ESIC_AUTO_project_structure"
ESIC_PROJECT = ESIC_ROOT / "ESIC_AUTO"
ESIC_EXCEL = ESIC_ROOT / "esic_users.xlsx"
if not ESIC_EXCEL.exists():
    _alt_esic = ESIC_PROJECT / "esic_users.xlsx"
    if _alt_esic.exists():
        ESIC_EXCEL = _alt_esic
ESIC_DATA_DIR = ESIC_PROJECT / "data"

# EPFO PATHS
EPFO_ROOT = COMPLIANCE_ROOT / "New_EPFO_AUTO_software" / "EPFO_AUTO"
EPFO_PROJECT = EPFO_ROOT
EPFO_EXCEL = EPFO_PROJECT / "data" / "users.xlsx"
EPFO_DATA_DIR = EPFO_PROJECT / "data" / "pf_downloads"
EPFO_TXT_DIR = Path(r"D:\epfo_downloads_txt")
REPORTS_DIR = COMPLIANCE_ROOT / "reports"
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

STATIC_DIR = Path(__file__).resolve().parent / "static"
STATIC_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Compliance Automation Suite", version="2.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================================
# PROCESS MANAGEMENT STATE (WITH REAL-TIME STAGE TRACKER)
# ============================================================
class ProcessState:
    def __init__(self):
        self.process: Optional[asyncio.subprocess.Process] = None
        self.running: bool = False
        self.service: Optional[str] = None
        self.mode: Optional[str] = None
        self.current_stage: str = "Ready / Idle"
        self.start_time: float = 0.0
        self.downloaded_count: int = 0
        self.failed_count: int = 0
        self.downloaded_users: set = set()
        self.failed_users: set = set()
        self.current_user: str = ""
        self.log_subscribers: List[asyncio.Queue] = []
        self.log_history: List[str] = []

    def reset_counts(self):
        self.downloaded_count = 0
        self.failed_count = 0
        self.downloaded_users = set()
        self.failed_users = set()
        self.current_user = ""
        self.current_stage = "Ready / Idle"

    def broadcast(self, line: str, extra: Optional[Dict[str, Any]] = None):
        self.log_history.append(line)
        if len(self.log_history) > 2500:
            self.log_history.pop(0)

        payload = {
            "line": line,
            "downloaded": self.downloaded_count,
            "failed": self.failed_count,
            "currentUser": self.current_user,
            "stage": self.current_stage,
            "running": self.running
        }
        if extra:
            payload.update(extra)

        msg = json.dumps(payload)
        for queue in list(self.log_subscribers):
            try:
                queue.put_nowait(msg)
            except Exception:
                pass

state = ProcessState()

# ============================================================
# PYDANTIC SCHEMAS
# ============================================================
class AccountItem(BaseModel):
    service: str
    loginid: str
    password: str

class AccountUpdate(BaseModel):
    service: str
    loginid: str
    password: str

class RunPipelineRequest(BaseModel):
    service: str  # "esic" or "epfo"
    step: Optional[int] = None
    from_step: Optional[int] = None
    headless: bool = False
    codes: Optional[str] = None
    only_failed: bool = False

class RunSingleRequest(BaseModel):
    service: str  # "esic" or "epfo"
    loginid: str
    password: Optional[str] = None
    headless: bool = False

class OpenPathRequest(BaseModel):
    target: str

class RecipientsUpdate(BaseModel):
    to: List[str]
    cc: List[str]

# ============================================================
# EXCEL HELPERS
# ============================================================
def get_excel_path(service: str) -> Path:
    if service.lower() == "esic":
        return ESIC_EXCEL
    elif service.lower() in ("epfo", "epdf"):
        return EPFO_EXCEL
    else:
        raise HTTPException(status_code=400, detail=f"Invalid service: {service}")

def read_accounts_from_excel(excel_path: Path) -> List[Dict[str, Any]]:
    if not excel_path.exists():
        return []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    accounts = []
    for r in range(2, ws.max_row + 1):
        raw_id = ws.cell(r, 1).value
        raw_pwd = ws.cell(r, 2).value
        raw_resp = ws.cell(r, 3).value

        if raw_id is None and raw_pwd is None:
            continue

        id_str = str(int(raw_id)) if isinstance(raw_id, float) and raw_id.is_integer() else str(raw_id or "").strip()
        pwd_str = str(int(raw_pwd)) if isinstance(raw_pwd, float) and raw_pwd.is_integer() else str(raw_pwd or "").strip()
        resp_str = str(raw_resp or "").strip()

        accounts.append({
            "row": r,
            "loginid": id_str,
            "password": pwd_str,
            "response": resp_str
        })
    return accounts

def add_account_to_excel(excel_path: Path, loginid: str, password: str) -> Dict[str, Any]:
    if not excel_path.exists():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["loginid", "password", "response"])
    else:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

    next_row = ws.max_row + 1
    if ws.max_row == 1 and (ws.cell(1, 1).value is None or ws.cell(1, 1).value == ""):
        ws.cell(1, 1, "loginid")
        ws.cell(1, 2, "password")
        ws.cell(1, 3, "response")
        next_row = 2

    ws.cell(next_row, 1, str(loginid).strip())
    ws.cell(next_row, 2, str(password).strip())
    ws.cell(next_row, 3, "PENDING")
    wb.save(excel_path)

    return {"row": next_row, "loginid": loginid, "password": password, "response": "PENDING"}

def update_account_in_excel(excel_path: Path, row: int, loginid: str, password: str) -> Dict[str, Any]:
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excel file not found")
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    if row < 2 or row > ws.max_row:
        raise HTTPException(status_code=400, detail="Invalid row number")

    ws.cell(row, 1, str(loginid).strip())
    ws.cell(row, 2, str(password).strip())
    existing_resp = ws.cell(row, 3).value or ""
    wb.save(excel_path)

    return {"row": row, "loginid": loginid, "password": password, "response": existing_resp}

def delete_account_from_excel(excel_path: Path, row: int) -> bool:
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excel file not found")
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    if row < 2 or row > ws.max_row:
        raise HTTPException(status_code=400, detail="Invalid row number")

    ws.delete_rows(row)
    wb.save(excel_path)
    return True

# ============================================================
# BACKGROUND SUBPROCESS RUNNER WITH REAL-TIME DOWNLOAD PARSER
# ============================================================
async def run_command_in_background(cmd: List[str], cwd: Path, service_name: str, mode_name: str):
    state.running = True
    state.service = service_name
    state.mode = mode_name
    state.start_time = time.time()
    state.reset_counts()
    state.current_stage = "STARTING PIPELINE..."

    state.broadcast(f"[SYSTEM] Starting process: {' '.join(cmd)}")
    state.broadcast(f"[SYSTEM] Working Directory: {cwd}")
    state.broadcast(f"[SYSTEM] Target Drive: D:\\")

    try:
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            cwd=str(cwd),
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.STDOUT,
            env={**os.environ, "PYTHONUNBUFFERED": "1", "PYTHONIOENCODING": "utf-8"}
        )
        state.process = proc

        while True:
            line_bytes = await proc.stdout.readline()
            if not line_bytes:
                break
            line_str = line_bytes.decode("utf-8", errors="replace").rstrip()
            line_upper = line_str.upper()

            # Dynamic Pipeline Stage Parser
            if "STEP 1" in line_upper or "CLEAN OLD" in line_upper or "CLEANUP" in line_upper:
                state.current_stage = "STEP 1: CLEANING OLD DATA"
            elif "STEP 2" in line_upper or "PORTAL DOWNLOAD" in line_upper or "EPFO DOWNLOAD" in line_upper:
                state.current_stage = "STEP 2: PORTAL DOWNLOADING"
            elif "STEP 3" in line_upper or "MERGE" in line_upper:
                state.current_stage = "STEP 3: MERGING EXCEL FILES"
            elif "STEP 4" in line_upper or "DB DOWNLOAD" in line_upper or "DATABASE DOWNLOAD" in line_upper or "POSTGRESQL" in line_upper:
                state.current_stage = "STEP 4: DB DOWNLOADING"
            elif "STEP 5" in line_upper or "COMPARISON" in line_upper or "COMPARE" in line_upper:
                state.current_stage = "STEP 5: DATA COMPARISON"
            elif "STEP 6" in line_upper or "HRMS" in line_upper or "UPLOAD" in line_upper:
                state.current_stage = "STEP 6: HRMS DATA UPLOADING"
            elif "STEP 7" in line_upper or "SEND STATUS & EMAIL REPORT" in line_upper or "SEND EMAIL STATUS REPORT" in line_upper or "SENDING EMAIL" in line_upper:
                state.current_stage = "STEP 7: SENDING EMAIL REPORT"

            # 1. Detect current user
            if "Processing User" in line_str or "Processing Row" in line_str or "User:" in line_str or "for user:" in line_str:
                parts = line_str.split("User:") if "User:" in line_str else (line_str.split("user:") if "user:" in line_str else line_str.split("ID:"))
                if len(parts) > 1:
                    raw_u = parts[1].strip().split()[0].strip()
                    if raw_u:
                        state.current_user = raw_u
                    if "STEP 2" in state.current_stage or "DOWNLOAD" in state.current_stage:
                        state.current_stage = f"STEP 2: DOWNLOADING ({state.current_user})"

            # 2. Detect download success (de-duplicated per user)
            is_success_line = (
                "SUCCESS -" in line_str or
                "Result: SUCCESS" in line_str or
                "File downloaded ->" in line_str or
                "Extracted & Cleaned CSV saved" in line_str
            )
            if is_success_line:
                user_key = state.current_user or f"user_{state.downloaded_count + 1}"
                if user_key not in state.downloaded_users:
                    state.downloaded_users.add(user_key)
                    state.downloaded_count = len(state.downloaded_users)
                    state.broadcast(f"🎉 [REAL-TIME] {state.downloaded_count} Downloaded! ({user_key})")

            # 3. Detect failure (de-duplicated per user)
            is_fail_line = (
                "FAILED -" in line_str or
                "Result: ERROR" in line_str or
                "failed after" in line_str
            )
            if is_fail_line:
                user_key = state.current_user or f"user_fail_{state.failed_count + 1}"
                if user_key not in state.failed_users:
                    state.failed_users.add(user_key)
                    state.failed_count = len(state.failed_users)

            state.broadcast(line_str)

        await proc.wait()
        elapsed = time.time() - state.start_time
        if proc.returncode == 0:
            state.broadcast(f"[SYSTEM] Process FINISHED SUCCESSFULLY in {elapsed:.2f}s (Total Downloaded: {state.downloaded_count})")
            
            # Check if email was already dispatched inside Step 7 of the pipeline script
            email_already_dispatched = any(
                ("dispatched successfully" in h.lower() or "email sent successfully" in h.lower())
                for h in state.log_history[-40:]
            )

            # Check if this was a full pipeline run (all steps)
            is_full_pipeline = (
                ("pipeline" in mode_name.lower() or "full" in mode_name.lower())
                and "all steps" in mode_name.lower()
                and "single" not in mode_name.lower()
                and "codes" not in mode_name.lower()
            )
            
            if is_full_pipeline and not email_already_dispatched and "step 7" not in mode_name.lower():
                # Full pipeline finished without step 7 dispatching
                state.current_stage = "SENDING EMAIL REPORT..."
                state.broadcast(f"[SYSTEM] 📧 Pipeline complete. Automatically generating & sending {service_name.upper()} status email via MS Graph Outlook...")
                try:
                    email_res = mailer.send_failure_email_via_ms_graph(service=service_name)
                    state.broadcast(f"[SYSTEM] 📧 {service_name.upper()} Email sent successfully! Attached: {email_res.get('report_name')} | Subject: {email_res.get('subject')}")
                    state.current_stage = "COMPLETED & EMAIL SENT"
                except Exception as em_err:
                    state.broadcast(f"[SYSTEM] ⚠️ Automatic Email Dispatch Error: {em_err}")
                    state.current_stage = "COMPLETED (EMAIL ERROR)"
                
                state.broadcast(f"[SYSTEM] ✅ Full cycle finished and stopped. Process complete.")
            else:
                state.current_stage = "COMPLETED"
                state.broadcast(f"[SYSTEM] ✅ Process complete.")
        else:
            state.broadcast(f"[SYSTEM] Process EXITED WITH STATUS {proc.returncode} after {elapsed:.2f}s (Downloaded: {state.downloaded_count}, Failed: {state.failed_count})")
            state.current_stage = "STOPPED (ERROR)"

    except asyncio.CancelledError:
        state.current_stage = "STOPPED (CANCELLED)"
        state.broadcast("[SYSTEM] Process CANCELLED by user.")
    except Exception as e:
        state.current_stage = "STOPPED (EXCEPTION)"
        state.broadcast(f"[SYSTEM] Process Execution Exception: {e}")
    finally:
        state.running = False
        state.process = None

# ============================================================
# API ENDPOINTS
# ============================================================

@app.get("/api/config")
def get_config():
    recipients = mailer.get_recipients()
    return {
        "esic": {
            "root": str(ESIC_ROOT),
            "project": str(ESIC_PROJECT),
            "excel": str(ESIC_EXCEL),
            "data_dir": str(ESIC_DATA_DIR),
            "excel_exists": ESIC_EXCEL.exists()
        },
        "epfo": {
            "root": str(EPFO_ROOT),
            "project": str(EPFO_PROJECT),
            "excel": str(EPFO_EXCEL),
            "data_dir": str(EPFO_DATA_DIR),
            "txt_dir": str(EPFO_TXT_DIR),
            "excel_exists": EPFO_EXCEL.exists()
        },
        "recipients": recipients
    }

@app.get("/api/accounts")
def list_accounts(service: str = Query("esic")):
    excel_path = get_excel_path(service)
    accounts = read_accounts_from_excel(excel_path)
    total = len(accounts)
    success_count = sum(1 for a in accounts if "SUCCESS" in a.get("response", "").upper())
    error_count = sum(1 for a in accounts if "ERROR" in a.get("response", "").upper() or "FAIL" in a.get("response", "").upper())
    pending_count = total - success_count - error_count
    return {
        "service": service,
        "excel_path": str(excel_path),
        "total": total,
        "success": success_count,
        "failed": error_count,
        "pending": pending_count,
        "accounts": accounts,
        "live_downloaded": state.downloaded_count if state.running else success_count,
        "live_failed": state.failed_count if state.running else error_count
    }

@app.post("/api/accounts")
def create_account(item: AccountItem):
    if not item.loginid or not item.password:
        raise HTTPException(status_code=400, detail="User ID and Password are required")
    excel_path = get_excel_path(item.service)
    created = add_account_to_excel(excel_path, item.loginid, item.password)
    return {"message": "Account added successfully", "account": created}

@app.put("/api/accounts/{row}")
def edit_account(row: int, item: AccountUpdate):
    if not item.loginid or not item.password:
        raise HTTPException(status_code=400, detail="User ID and Password are required")
    excel_path = get_excel_path(item.service)
    updated = update_account_in_excel(excel_path, row, item.loginid, item.password)
    return {"message": "Account updated successfully", "account": updated}

@app.delete("/api/accounts/{row}")
def remove_account(row: int, service: str = Query(...)):
    excel_path = get_excel_path(service)
    delete_account_from_excel(excel_path, row)
    return {"message": "Account deleted successfully"}

@app.get("/api/status")
def get_status():
    elapsed = time.time() - state.start_time if state.running else 0
    return {
        "running": state.running,
        "service": state.service,
        "mode": state.mode,
        "stage": state.current_stage,
        "elapsed": round(elapsed, 1),
        "downloaded_count": state.downloaded_count,
        "failed_count": state.failed_count,
        "current_user": state.current_user
    }

@app.get("/api/failed-accounts")
def get_failed_accounts(service: str = Query("esic")):
    excel_path = get_excel_path(service)
    accounts = read_accounts_from_excel(excel_path)
    failed = [
        a for a in accounts
        if "FAIL" in str(a.get("response", "")).upper() or "ERROR" in str(a.get("response", "")).upper()
    ]
    return {
        "service": service,
        "count": len(failed),
        "failed_codes": [a["loginid"] for a in failed],
        "accounts": failed
    }

@app.post("/api/run-pipeline")
async def run_pipeline(req: RunPipelineRequest):
    if state.running:
        raise HTTPException(status_code=409, detail="A process is already running. Please stop it first.")

    service = req.service.lower()
    py_exec = sys.executable

    step_desc = f"Step {req.step}" if req.step else (f"From Step {req.from_step}" if req.from_step else "All Steps")
    if req.codes:
        step_desc += f" [Codes: {req.codes}]"
    elif req.only_failed:
        step_desc += " [Retry Failed]"

    if service == "esic":
        cmd = [py_exec, "Main.py"]
        if req.step:
            cmd.extend(["--step", str(req.step)])
        elif req.from_step:
            cmd.extend(["--from-step", str(req.from_step)])
        if req.codes:
            cmd.extend(["--codes", str(req.codes)])
        if req.only_failed:
            cmd.append("--only-failed")
        if req.headless:
            cmd.append("--headless")
        cwd = ESIC_PROJECT
        mode_name = f"ESIC Pipeline ({step_desc})"
    elif service in ("epfo", "epdf"):
        cmd = [py_exec, "main.py"]
        if req.step:
            cmd.extend(["--step", str(req.step)])
        elif req.from_step:
            cmd.extend(["--from-step", str(req.from_step)])
        if req.codes:
            cmd.extend(["--codes", str(req.codes)])
        if req.only_failed:
            cmd.append("--only-failed")
        if req.headless:
            cmd.append("--headless")
        cwd = EPFO_PROJECT
        mode_name = f"EPFO Pipeline ({step_desc})"
    else:
        raise HTTPException(status_code=400, detail=f"Invalid service: {req.service}")

    state.running = True
    state.service = service
    state.mode = mode_name
    state.start_time = time.time()
    state.current_stage = f"STARTING: {mode_name}"

    asyncio.create_task(run_command_in_background(cmd, cwd, service, mode_name))
    return {"message": f"{mode_name} started successfully", "command": cmd}

@app.post("/api/run-single")
async def run_single(req: RunSingleRequest):
    if state.running:
        raise HTTPException(status_code=409, detail="A process is already running. Please stop it first.")

    if not req.loginid:
        raise HTTPException(status_code=400, detail="User ID / Employer Code is required.")

    service = req.service.lower()

    # Single code: if password is provided, sync/update in Excel
    pwd = str(req.password or "").strip()
    if pwd:
        try:
            excel_path = get_excel_path(service)
            accounts = read_accounts_from_excel(excel_path)
            found = next((a for a in accounts if str(a.get("loginid", "")).strip() == str(req.loginid).strip()), None)
            if found:
                if str(found.get("password", "")).strip() != pwd:
                    update_account_in_excel(excel_path, found["row"], str(req.loginid).strip(), pwd)
            else:
                add_account_to_excel(excel_path, str(req.loginid).strip(), pwd)
        except Exception as ex:
            print(f"Warning: could not sync password to Excel: {ex}")

    # Run the full pipeline (Step 1 to 7: Cleanup, Portal Download, Merge, DB, Comparison, HRMS Upload, Email Report)
    return await run_pipeline(RunPipelineRequest(
        service=service,
        codes=str(req.loginid).strip(),
        headless=req.headless
    ))

@app.post("/api/stop")
def stop_process():
    if not state.running or not state.process:
        return {"message": "No active process running"}

    try:
        state.process.kill()
        state.broadcast("[SYSTEM] Kill signal sent to process.")
        state.running = False
        state.process = None
        return {"message": "Process terminated successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Could not terminate process: {e}")

@app.get("/api/logs/stream")
async def stream_logs():
    queue = asyncio.Queue()
    # Send recent history first
    for old_line in state.log_history[-100:]:
        queue.put_nowait(json.dumps({
            "line": old_line,
            "downloaded": state.downloaded_count,
            "failed": state.failed_count,
            "currentUser": state.current_user,
            "running": state.running
        }))

    state.log_subscribers.append(queue)

    async def event_generator():
        try:
            while True:
                msg = await queue.get()
                yield f"data: {msg}\n\n"
        except asyncio.CancelledError:
            pass
        finally:
            if queue in state.log_subscribers:
                state.log_subscribers.remove(queue)

    return StreamingResponse(event_generator(), media_type="text/event-stream")

# ============================================================
# REPORTS & OUTLOOK EMAIL ENDPOINTS
# ============================================================

@app.get("/api/reports/recipients")
def get_email_recipients():
    return mailer.get_recipients()

@app.post("/api/reports/recipients")
def update_email_recipients(req: RecipientsUpdate):
    mailer.save_recipients_to_excel(req.to, req.cc)
    return {"message": "Recipients saved to Excel successfully", "recipients": {"to": req.to, "cc": req.cc}}

@app.post("/api/reports/generate-excel")
def generate_excel_report(service: Optional[str] = Query(None)):
    try:
        svc = (service or state.service or "all").lower()
        report_path = mailer.generate_service_report_excel(service=svc)
        if svc in ("esic", "epfo"):
            records, total_c, succ_c, fail_c = mailer.collect_service_accounts(svc)
            return {
                "message": f"{svc.upper()} Report generated successfully",
                "service": svc.upper(),
                "report_file": str(report_path),
                "report_name": report_path.name,
                "total_count": total_c,
                "success_count": succ_c,
                "failed_count": fail_c
            }
        else:
            esic_recs, epfo_recs = mailer.collect_failures_and_pending()
            return {
                "message": "Combined Failure Excel report generated successfully",
                "service": "ALL",
                "report_file": str(report_path),
                "report_name": report_path.name,
                "esic_count": len(esic_recs),
                "epfo_count": len(epfo_recs),
                "total_count": len(esic_recs) + len(epfo_recs)
            }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Report generation error: {e}")

@app.post("/api/reports/send-email")
def send_failure_email(service: Optional[str] = Query(None)):
    try:
        svc = (service or state.service or "esic").lower()
        res = mailer.send_failure_email_via_ms_graph(service=svc)
        state.broadcast(f"[SYSTEM] 📧 Outlook Email sent successfully for {svc.upper()}! Attached: {res.get('report_name')} | Subject: {res.get('subject')}")
        return res
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to send email: {e}")

@app.post("/api/reports/open-folder")
def open_reports_folder():
    try:
        os.startfile(str(REPORTS_DIR))
        return {"message": "Reports folder opened"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Could not open reports folder: {e}")

@app.post("/api/open-path")
def open_path(req: OpenPathRequest):
    target = req.target.lower()
    path_to_open = None

    if target == "esic_folder":
        path_to_open = ESIC_DATA_DIR
    elif target == "esic_excel":
        path_to_open = ESIC_EXCEL
    elif target == "epfo_folder":
        path_to_open = EPFO_DATA_DIR
    elif target == "epfo_excel":
        path_to_open = EPFO_EXCEL
    elif target == "epfo_txt":
        path_to_open = EPFO_TXT_DIR
    elif target == "recipients_excel":
        path_to_open = mailer.RECIPIENTS_EXCEL
    elif target == "reports_dir":
        path_to_open = REPORTS_DIR
    elif target == "compliance_root":
        path_to_open = COMPLIANCE_ROOT
    else:
        raise HTTPException(status_code=400, detail="Invalid target path requested")

    if path_to_open and path_to_open.exists():
        try:
            os.startfile(str(path_to_open))
            return {"message": f"Opened {path_to_open}"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Could not open path: {e}")
    else:
        raise HTTPException(status_code=404, detail=f"Path does not exist: {path_to_open}")

@app.post("/api/open-step-target")
def open_step_target(service: str = Query("esic"), step: int = Query(...)):
    service = service.lower()
    base_dir = ESIC_DATA_DIR if service == "esic" else (EPFO_PROJECT / "data")
    
    target_path = None
    if step in (1, 2):
        target_path = base_dir
    elif step == 3:
        name = "ESIC_Consolidated_Report.xlsx" if service == "esic" else "EPFO_Consolidated_Report.xlsx"
        p = base_dir / name
        target_path = p if p.exists() else base_dir
    elif step == 4:
        p = base_dir / "PostgreSQL_Data.xlsx"
        target_path = p if p.exists() else base_dir
    elif step == 5:
        p = base_dir / "Comparison_Output.xlsx"
        target_path = p if p.exists() else base_dir
    elif step == 6:
        p = base_dir / "HRMS_Upload.csv"
        target_path = p if p.exists() else base_dir
    elif step == 7:
        target_path = REPORTS_DIR

    if target_path and target_path.exists():
        try:
            os.startfile(str(target_path))
            return {"message": f"Opened {target_path.name}"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Could not open {target_path}: {e}")
    else:
        if base_dir.exists():
            os.startfile(str(base_dir))
            return {"message": "Opened data folder"}
        raise HTTPException(status_code=404, detail="Path not found")

@app.post("/api/open-url-host")
def open_url_host(url: str = Query(...)):
    import webbrowser
    webbrowser.open(url)
    return {"message": f"Opened {url} in browser"}

# Mount static folder
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

@app.get("/", response_class=HTMLResponse)
def index_page():
    index_file = STATIC_DIR / "index.html"
    if index_file.exists():
        return HTMLResponse(content=index_file.read_text(encoding="utf-8"))
    return HTMLResponse("<h1>Compliance Suite Dashboard</h1><p>Static files loading...</p>")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=8080, reload=False)
