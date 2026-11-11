from pathlib import Path
from datetime import datetime


# =========================================================
# PROJECT BASE FOLDER
# =========================================================

BASE_DIR = Path(__file__).resolve().parent


# =========================================================
# PF / EPFO INPUT EXCEL
# =========================================================

EXCEL_PATH = BASE_DIR / "data" / "users.xlsx"


# =========================================================
# CAPTCHA IMAGES
# =========================================================

CAPTCHA_FOLDER = BASE_DIR / "data" / "captcha_images"


# =========================================================
# PF DOWNLOAD
# =========================================================

EPFO_DOWNLOAD_DIR = BASE_DIR / "data" / "pf_downloads"


# =========================================================
# DATE-WISE PF DOWNLOAD FOLDER
# =========================================================

DATE_FOLDER_NAME = datetime.now().strftime("%d%m%Y")

TARGET_DOWNLOAD_FOLDER = (
    EPFO_DOWNLOAD_DIR / DATE_FOLDER_NAME
)



# =========================================================
# DATABASE
# =========================================================

DB_DIR = BASE_DIR / "data" / "db"

DB_OUTPUT = (
    DB_DIR / "DB_Data.xlsx"
)

EPFO_MERGED_FILE = (
    DB_DIR / "All_Merge.xlsx"
)


# =========================================================
# ESIC
# =========================================================

ESIC_DOWNLOAD_DIR = (
    BASE_DIR / "data" / "esic" / "downloads"
)

ESIC_MERGED_FILE = (
    BASE_DIR / "data" / "esic" / "ESIC_Merged.xlsx"
)


# =========================================================
# HRMS COMPARISON / REPORT
# =========================================================

HRMS_UPLOAD_FOLDER = (
    BASE_DIR / "data" / "Hrms_upload_sheet"
)

COMPARISON_OUTPUT = (
    HRMS_UPLOAD_FOLDER / "EPFO_Comparison_Result.xlsx"
)


# =========================================================
# HRMS ACTUAL UPLOAD FILE
# =========================================================
# Ye file HRMS portal par upload hogi.
#
# Location:
# data\upload_hrms\HRMS_Upload.xlsx
#
# Isme sirf MAIN sheet hogi
# (NEW + UPDATED records)
# =========================================================

HRMS_UPLOAD_DIR = (
    BASE_DIR / "data" / "upload_hrms"
)

HRMS_UPLOAD_FILE = (
    HRMS_UPLOAD_DIR / "HRMS_Upload.csv"
)


# =========================================================
# BACKUP
# =========================================================

BACKUP_DIR = (
    BASE_DIR / "data" / "backup"
)


# =========================================================
# RESULT
# =========================================================

RESULT_DIR = (
    BASE_DIR / "data" / "result"
)


# =========================================================
# LOGS
# =========================================================

LOG_DIR = (
    BASE_DIR / "logs"
)


# =========================================================
# SCREENSHOTS
# =========================================================

SCREENSHOT_DIR = (
    BASE_DIR / "screenshots"
)


# =========================================================
# CREATE REQUIRED FOLDERS
# =========================================================

CAPTCHA_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

EPFO_DOWNLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)

TARGET_DOWNLOAD_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

DB_DIR.mkdir(
    parents=True,
    exist_ok=True
)

ESIC_DOWNLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)

HRMS_UPLOAD_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

HRMS_UPLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)

BACKUP_DIR.mkdir(
    parents=True,
    exist_ok=True
)

RESULT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

LOG_DIR.mkdir(
    parents=True,
    exist_ok=True
)

SCREENSHOT_DIR.mkdir(
    parents=True,
    exist_ok=True
)