import sys
import re
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# D DRIVE DYNAMIC ABSOLUTE PATHS
PROJECT_FOLDER = Path(__file__).resolve().parent
DB_FILE = PROJECT_FOLDER / "dbdata" / "dbdata.xlsx"
PORTAL_FILE = PROJECT_FOLDER / "Merge_data" / "merge sheet.xlsx"
COMPARE_FOLDER = PROJECT_FOLDER / "compaire"

OUTPUT_FILE = COMPARE_FOLDER / "compare sheet.xlsx"
MAIN_XLSX_FILE = COMPARE_FOLDER / "main.xlsx"
MAIN_CSV_FILE = COMPARE_FOLDER / "main.csv"
MAINUPLOAD_XLSX_FILE = COMPARE_FOLDER / "mainupload.xlsx"
MAINUPLOAD_CSV_FILE = COMPARE_FOLDER / "mainupload.csv"

COMPARE_FOLDER.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("compare")
if not logger.hasHandlers():
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(ch)
    logger.setLevel(logging.INFO)

COMPARE_COLUMNS = [
    "IPNumber",
    "EmployeeName",
    "DateofAppointment",
    "DateofBirth",
    "Gender",
    "UANNumber",
    "ESICEmployerCode",
]

DATE_COLUMNS = ["DateofAppointment", "DateofBirth"]

PORTAL_RENAME = {
    "IP Number": "IPNumber",
    "Name": "EmployeeName",
    "Date of Appointment": "DateofAppointment",
    "Date of Birth": "DateofBirth",
    "Gender": "Gender",
    "UAN Number": "UANNumber",
    "Employer Code": "ESICEmployerCode",
}


def normalize_value(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    val_str = str(value).strip()
    if val_str.lower() in ["nan", "none", "null", "nat"]:
        return ""

    val_str = " ".join(val_str.split())
    if re.fullmatch(r"-?\d+\.0", val_str):
        val_str = val_str[:-2]

    return val_str


def clean_employee_name(value):
    val_norm = normalize_value(value)
    if not val_norm:
        return ""
    cleaned = re.sub(r"(?i)\bmd\s*\.\s*", "", val_norm)
    return " ".join(cleaned.split())



def normalize_gender(value):
    val_norm = normalize_value(value).upper()
    if val_norm in ["M", "MALE"]:
        return "M"
    if val_norm in ["F", "FEMALE"]:
        return "F"
    return val_norm


def normalize_date(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value).normalize()

    value_str = str(value).strip()
    if not value_str:
        return ""

    if " " in value_str:
        first_part = value_str.split(" ")[0]
        if re.search(r"[\./\-]", first_part):
            value_str = first_part

    formats = [
        "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
        "%Y.%m.%d", "%Y/%m/%d", "%Y-%m-%d",
        "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
        "%m/%d/%Y", "%m-%d-%Y"
    ]

    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(value_str, fmt)).normalize()
        except Exception:
            pass

    try:
        parsed = pd.to_datetime(value_str, dayfirst=True, errors="coerce")
        if not pd.isna(parsed):
            return pd.Timestamp(parsed).normalize()
    except Exception:
        pass

    try:
        if re.fullmatch(r"\d+(\.0+)?", value_str):
            number = float(value_str)
            if 1 <= number <= 60000:
                return pd.Timestamp("1899-12-30") + pd.to_timedelta(number, unit="D")
    except Exception:
        pass

    return value_str


def comparison_value(column, value):
    if column in DATE_COLUMNS:
        normalized = normalize_date(value)
        if isinstance(normalized, (pd.Timestamp, datetime)):
            return normalized.strftime("%Y-%m-%d")
        return str(normalized).strip()

    if column == "Gender":
        return normalize_gender(value)

    return normalize_value(value).casefold()


def prepare_date_for_excel(value):
    normalized = normalize_date(value)
    if isinstance(normalized, (pd.Timestamp, datetime)):
        return normalized.to_pydatetime()
    return normalized


def run_comparison():
    logger.info("Starting ESIC Data Comparison Process...")

    if not DB_FILE.exists():
        raise FileNotFoundError(f"DB File not found: {DB_FILE}")
    if not PORTAL_FILE.exists():
        raise FileNotFoundError(f"Merged Portal File not found: {PORTAL_FILE}")

    db_df = pd.read_excel(DB_FILE, dtype=str, engine="openpyxl")
    db_df.columns = [str(col).strip() for col in db_df.columns]

    portal_df = pd.read_excel(PORTAL_FILE, dtype=str, engine="openpyxl")
    portal_df.columns = [str(col).strip() for col in portal_df.columns]
    portal_df.rename(columns=PORTAL_RENAME, inplace=True)

    missing_db = [col for col in COMPARE_COLUMNS if col not in db_df.columns]
    if missing_db:
        raise ValueError(f"DB file missing required columns: {missing_db}")

    missing_portal = [col for col in COMPARE_COLUMNS if col not in portal_df.columns]
    if missing_portal:
        raise ValueError(f"Portal file missing required columns: {missing_portal}")

    db_df = db_df[COMPARE_COLUMNS].copy()
    portal_df = portal_df[COMPARE_COLUMNS].copy()

    for col in COMPARE_COLUMNS:
        if col not in DATE_COLUMNS:
            db_df[col] = db_df[col].apply(normalize_value)
            portal_df[col] = portal_df[col].apply(normalize_value)

    if "EmployeeName" in db_df.columns:
        db_df["EmployeeName"] = db_df["EmployeeName"].apply(clean_employee_name)
    if "EmployeeName" in portal_df.columns:
        portal_df["EmployeeName"] = portal_df["EmployeeName"].apply(clean_employee_name)


    for col in DATE_COLUMNS:
        db_df[col] = db_df[col].apply(prepare_date_for_excel)
        portal_df[col] = portal_df[col].apply(prepare_date_for_excel)

    db_df["IPNumber"] = db_df["IPNumber"].apply(normalize_value)
    portal_df["IPNumber"] = portal_df["IPNumber"].apply(normalize_value)

    db_df = db_df[db_df["IPNumber"] != ""].copy()
    portal_df = portal_df[portal_df["IPNumber"] != ""].copy()

    db_duplicates = db_df[db_df["IPNumber"].duplicated(keep=False)].copy()
    portal_duplicates = portal_df[portal_df["IPNumber"].duplicated(keep=False)].copy()

    duplicate_frames = []
    if not db_duplicates.empty:
        temp = db_duplicates.copy()
        temp.insert(0, "Source", "DB")
        duplicate_frames.append(temp)
    if not portal_duplicates.empty:
        temp = portal_duplicates.copy()
        temp.insert(0, "Source", "ESIC Portal")
        duplicate_frames.append(temp)

    duplicate_df = (
        pd.concat(duplicate_frames, ignore_index=True)
        if duplicate_frames
        else pd.DataFrame(columns=["Source"] + COMPARE_COLUMNS)
    )

    db_lookup = db_df.drop_duplicates(subset=["IPNumber"], keep="first").set_index("IPNumber")
    portal_lookup = portal_df.drop_duplicates(subset=["IPNumber"], keep="first").set_index("IPNumber")

    db_ips = set(db_lookup.index)
    portal_ips = set(portal_lookup.index)

    new_ips = portal_ips - db_ips
    new_rows = []
    for ip in new_ips:
        portal_row = portal_lookup.loc[ip]
        row = {"IPNumber": ip, "Status": "NEW", "ChangedFields": "NEW RECORD"}
        for col in COMPARE_COLUMNS:
            if col != "IPNumber":
                row[col] = portal_row[col]
        new_rows.append(row)

    new_df = pd.DataFrame(new_rows)

    common_ips = portal_ips & db_ips
    comparison_rows = []

    for ip in common_ips:
        db_row = db_lookup.loc[ip]
        portal_row = portal_lookup.loc[ip]
        changed_fields = []

        for col in COMPARE_COLUMNS:
            if col == "IPNumber":
                continue
            if comparison_value(col, db_row[col]) != comparison_value(col, portal_row[col]):
                changed_fields.append(col)

        status = "UPDATE" if changed_fields else "SAME"
        row = {
            "IPNumber": ip,
            "Status": status,
            "ChangedFields": ", ".join(changed_fields) if changed_fields else ""
        }
        for col in COMPARE_COLUMNS:
            if col != "IPNumber":
                row[f"{col}_DB"] = db_row[col]
                row[f"{col}_Portal"] = portal_row[col]

        comparison_rows.append(row)

    comparison_df = pd.DataFrame(comparison_rows)
    update_df = (
        comparison_df[comparison_df["Status"] == "UPDATE"].copy()
        if not comparison_df.empty
        else pd.DataFrame()
    )

    main_frames = []
    if not new_df.empty:
        main_frames.append(new_df)
    if not update_df.empty:
        main_frames.append(update_df)

    if main_frames:
        main_df = pd.concat(main_frames, ignore_index=True, sort=False)
    else:
        main_df = pd.DataFrame(
            columns=["IPNumber", "Status", "ChangedFields"] + [c for c in COMPARE_COLUMNS if c != "IPNumber"]
        )

    # Export HRMS Main Upload Files
    HRMS_SOURCE_COLUMNS = ["IPNumber", "EmployeeName", "DateofAppointment", "DateofBirth", "Gender", "UANNumber", "ESICEmployerCode"]
    hrms_upload_ips = main_df["IPNumber"].dropna().astype(str).map(str.strip).tolist() if not main_df.empty and "IPNumber" in main_df.columns else []

    hrms_upload_rows = []
    for ip in hrms_upload_ips:
        if ip in portal_lookup.index:
            p_row = portal_lookup.loc[ip]
            hrms_upload_rows.append({
                "IPNumber": ip,
                "EmployeeName": p_row["EmployeeName"],
                "DateofAppointment": prepare_date_for_excel(p_row["DateofAppointment"]),
                "DateofBirth": prepare_date_for_excel(p_row["DateofBirth"]),
                "Gender": normalize_gender(p_row["Gender"]),
                "UANNumber": normalize_value(p_row["UANNumber"]),
                "ESICEmployerCode": normalize_value(p_row["ESICEmployerCode"]),
            })

    hrms_upload_df_standard = pd.DataFrame(hrms_upload_rows, columns=HRMS_SOURCE_COLUMNS)
    hrms_upload_df = hrms_upload_df_standard.rename(
        columns={
            "IPNumber": "IP Number",
            "EmployeeName": "Name",
            "DateofAppointment": "Date of Appointment",
            "DateofBirth": "Date of Birth",
            "Gender": "Gender",
            "UANNumber": "UAN Number",
            "ESICEmployerCode": "Employer Code",
        }
    )

    hrms_upload_df.to_excel(MAINUPLOAD_XLSX_FILE, index=False, engine="openpyxl")
    hrms_upload_df.to_csv(MAINUPLOAD_CSV_FILE, index=False, encoding="utf-8-sig")

    hrms_wb = load_workbook(MAINUPLOAD_XLSX_FILE)
    hrms_ws = hrms_wb.active
    hrms_ws.freeze_panes = "A2"
    for cell in hrms_ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column_cells in hrms_ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in column_cells)
        col_letter = get_column_letter(column_cells[0].column)
        hrms_ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 35)
    hrms_wb.save(MAINUPLOAD_XLSX_FILE)

    db_only_ips = db_ips - portal_ips
    db_only_rows = []
    for ip in db_only_ips:
        d_row = db_lookup.loc[ip]
        row = {"IPNumber": ip, "Status": "DB ONLY"}
        for col in COMPARE_COLUMNS:
            if col != "IPNumber":
                row[col] = d_row[col]
        db_only_rows.append(row)

    db_only_df = pd.DataFrame(db_only_rows)

    same_count = len(comparison_df[comparison_df["Status"] == "SAME"]) if not comparison_df.empty else 0
    update_count = len(update_df)
    new_count = len(new_df)
    db_only_count = len(db_only_df)
    duplicate_count = len(duplicate_df)

    summary_df = pd.DataFrame({
        "Metric": [
            "DB Total Rows", "Portal Total Rows", "DB Unique IPNumber",
            "Portal Unique IPNumber", "NEW", "UPDATE", "SAME", "DB ONLY", "DUPLICATE Rows"
        ],
        "Count": [
            len(db_df), len(portal_df), len(db_ips), len(portal_ips),
            new_count, update_count, same_count, db_only_count, duplicate_count
        ]
    })

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        main_df.to_excel(writer, sheet_name="MAIN", index=False)
        comparison_df.to_excel(writer, sheet_name="COMPARE", index=False)
        duplicate_df.to_excel(writer, sheet_name="DUPLICATE", index=False)
        db_only_df.to_excel(writer, sheet_name="DB ONLY", index=False)

    wb = load_workbook(OUTPUT_FILE)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        if ws.max_row >= 2:
            ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column_cells)
            col_letter = get_column_letter(column_cells[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 35)

    wb.save(OUTPUT_FILE)

    logger.info(f"Comparison Summary: NEW={new_count}, UPDATE={update_count}, SAME={same_count}")
    logger.info(f"Generated Comparison Sheet -> {OUTPUT_FILE.name}")
    logger.info(f"Generated HRMS Upload CSV -> {MAINUPLOAD_CSV_FILE.name} ({len(hrms_upload_df)} records)")
    return True


if __name__ == "__main__":
    try:
        run_comparison()
    except Exception as exc:
        logger.error(f"Comparison failed: {exc}")
        sys.exit(1)