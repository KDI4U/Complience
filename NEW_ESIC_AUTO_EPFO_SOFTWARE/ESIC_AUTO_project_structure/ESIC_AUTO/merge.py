import sys
import re
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd

# =========================================================
# PATH SETTINGS - FULL ABSOLUTE PATHS
# =========================================================

PROJECT_FOLDER = Path(__file__).resolve().parent
DATA_FOLDER = PROJECT_FOLDER / "data"
FALLBACK_DATA_FOLDER = PROJECT_FOLDER.parent / "data"
MERGE_FOLDER = PROJECT_FOLDER / "Merge_data"

OUTPUT_XLSX = MERGE_FOLDER / "merge sheet.xlsx"
OUTPUT_CSV = MERGE_FOLDER / "merge sheet.csv"

logger = logging.getLogger("merge_data")
if not logger.hasHandlers():
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(ch)
    logger.setLevel(logging.INFO)

COLUMN_RENAME = {
    "IP Number (10 Digits Only)": "IPNumber",
    "Name (Alphabets Only)": "EmployeeName",
    "Date of Appointment (DD/MM/YYYY)Eg : 01/01/1970": "DateofAppointment",
    "Date of Birth (DD/MM/YYYY) Eg : 01/01/1970": "DateofBirth",
    "Gender (M/F)": "Gender",
    "UANNumber": "UANNumber",
    "UAN Number": "UANNumber",
    "ESICEmployerCode": "ESICEmployerCode",
}

EXPECTED_COLUMNS = [
    "IPNumber",
    "EmployeeName",
    "DateofAppointment",
    "DateofBirth",
    "Gender",
    "UANNumber",
    "ESICEmployerCode",
]


def find_input_folder():
    # 1. Check primary data folder for today's date
    today_folder = DATA_FOLDER / datetime.now().strftime("%d%m%Y") / "data_downloaded"
    if today_folder.exists() and any(today_folder.glob("*.xls")):
        return today_folder

    # 2. Check primary data folder for any recent date
    if DATA_FOLDER.exists():
        for date_dir in sorted(DATA_FOLDER.iterdir(), reverse=True):
            if date_dir.is_dir():
                download_dir = date_dir / "data_downloaded"
                if download_dir.exists() and any(download_dir.glob("*.xls")):
                    return download_dir
                if any(date_dir.glob("*.xls")):
                    return date_dir

    # 3. Check fallback data folder in parent directory
    if FALLBACK_DATA_FOLDER.exists():
        fallback_today = FALLBACK_DATA_FOLDER / datetime.now().strftime("%d%m%Y") / "data_downloaded"
        if fallback_today.exists() and any(fallback_today.glob("*.xls")):
            return fallback_today
        for date_dir in sorted(FALLBACK_DATA_FOLDER.iterdir(), reverse=True):
            if date_dir.is_dir():
                download_dir = date_dir / "data_downloaded"
                if download_dir.exists() and any(download_dir.glob("*.xls")):
                    return download_dir
                if any(date_dir.glob("*.xls")):
                    return date_dir

    return today_folder


def read_esic_xls(file):
    try:
        tables = pd.read_html(str(file))
        if not tables:
            raise ValueError("No HTML tables found in file.")

        df = max(tables, key=lambda x: x.shape[0] * x.shape[1])
        df = df.astype(str)
        df = df.replace(["nan", "NaN", "None"], "")
        df = df.dropna(how="all").dropna(axis=1, how="all")
        return df
    except Exception as e:
        logger.error(f"Error reading {file.name}: {e}")
        return None


def clean_employee_name(val):
    """
    Removes honorific prefixes like 'Md.', 'md.', 'MD.', 'Md .', 'MD. ', etc.
    from employee names whenever 'md' / 'MD' is followed by a dot.
    """
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "null"]:
        return ""

    # Remove 'md' or 'MD' followed by optional spaces and a dot '.'
    cleaned = re.sub(r"(?i)\bmd\s*\.\s*", "", val_str)
    # Replace multiple spaces with a single space
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def merge_esic_files():
    logger.info("Starting ESIC Data Merge & Conversion Process...")

    input_folder = find_input_folder()
    if not input_folder.exists():
        raise FileNotFoundError(f"Input download folder not found: {input_folder}")

    MERGE_FOLDER.mkdir(parents=True, exist_ok=True)

    excel_files = list(input_folder.glob("*.xls"))
    if not excel_files:
        raise FileNotFoundError(f"No .xls files found in {input_folder}")

    logger.info(f"Input Directory: {input_folder.name} | Total Files: {len(excel_files)}")

    all_data = []
    successful_files = 0

    for file in excel_files:
        df = read_esic_xls(file)
        if df is None:
            continue

        df.columns = [str(col).strip() for col in df.columns]
        df = df.dropna(how="all")
        df["ESICEmployerCode"] = file.stem
        all_data.append(df)
        successful_files += 1

    if not all_data:
        raise ValueError("No data could be extracted from files.")

    merged_df = pd.concat(all_data, ignore_index=True, sort=False)
    merged_df.columns = [str(col).strip() for col in merged_df.columns]
    merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]

    merged_df.rename(columns=COLUMN_RENAME, inplace=True)
    merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]

    # Clean EmployeeName column (Remove Md., md., MD., md ., etc.)
    if "EmployeeName" in merged_df.columns:
        merged_df["EmployeeName"] = merged_df["EmployeeName"].apply(clean_employee_name)

    missing_cols = [col for col in EXPECTED_COLUMNS if col not in merged_df.columns]

    if missing_cols:
        raise ValueError(f"Merged file missing required columns: {missing_cols}")

    extra_cols = [col for col in merged_df.columns if col not in EXPECTED_COLUMNS]
    merged_df = merged_df[EXPECTED_COLUMNS + extra_cols]

    merged_df.to_excel(OUTPUT_XLSX, index=False, engine="openpyxl")
    merged_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    logger.info(f"Merged {successful_files}/{len(excel_files)} files cleanly into {len(merged_df)} rows.")
    logger.info(f"Generated: {OUTPUT_XLSX.name} and {OUTPUT_CSV.name}")
    return True


if __name__ == "__main__":
    try:
        merge_esic_files()
    except Exception as exc:
        logger.error(f"Merge failed: {exc}")
        sys.exit(1)
