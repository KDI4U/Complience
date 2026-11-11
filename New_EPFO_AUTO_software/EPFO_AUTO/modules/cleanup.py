import sys
from pathlib import Path
from datetime import datetime
import shutil

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from config import BASE_DIR as CONFIG_BASE_DIR


# =========================================================
# BASE DATA FOLDER
# =========================================================

DATA_DIR = BASE_DIR / "data"


# =========================================================
# WORKING FOLDERS
# =========================================================

PF_DOWNLOAD_DIR = DATA_DIR / "pf_downloads"

CAPTCHA_DIR = DATA_DIR / "captcha_images"

DB_DIR = DATA_DIR / "db"

HRMS_UPLOAD_DIR = DATA_DIR / "Hrms_upload_sheet"


# =========================================================
# BACKUP FOLDER
# =========================================================

BACKUP_DIR = DATA_DIR / "backup"


# =========================================================
# CREATE BACKUP FOLDER
# =========================================================

def create_backup_folder():

    timestamp = datetime.now().strftime(
        "%d%m%Y_%H%M%S"
    )

    backup_folder = BACKUP_DIR / timestamp

    backup_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    return backup_folder


# =========================================================
# DELETE FOLDER CONTENT
# =========================================================

def delete_folder_content(folder):

    if not folder.exists():

        folder.mkdir(
            parents=True,
            exist_ok=True
        )

        print(f"[CREATE] {folder}")

        return

    print()
    print(f"[DELETE] {folder}")

    for item in folder.iterdir():

        try:

            if item.is_dir():

                shutil.rmtree(item)

            else:

                item.unlink()

            print(
                f"   Deleted : {item.name}"
            )

        except Exception as e:

            print(
                f"   ERROR : {item.name}"
            )

            print(
                f"   {e}"
            )


# =========================================================
# BACKUP FOLDER CONTENT
# =========================================================

def backup_folder(folder, backup_root):

    if not folder.exists():

        return

    # Check empty
    if not any(folder.iterdir()):

        print(
            f"[EMPTY] {folder}"
        )

        return

    destination = backup_root / folder.name

    destination.mkdir(
        parents=True,
        exist_ok=True
    )

    print()
    print(
        f"[BACKUP] {folder}"
    )

    for item in folder.iterdir():

        target = destination / item.name

        try:

            if item.is_dir():

                shutil.copytree(
                    item,
                    target,
                    dirs_exist_ok=True
                )

            else:

                shutil.copy2(
                    item,
                    target
                )

            print(
                f"   Copied : {item.name}"
            )

        except Exception as e:

            print(
                f"   ERROR : {item.name}"
            )

            print(
                f"   {e}"
            )


# =========================================================
# CLEANUP OLD DATA
# =========================================================

def cleanup_old_data():

    print()
    print("=" * 70)
    print("OLD DATA CLEANUP")
    print("=" * 70)

    # =====================================================
    # CREATE BACKUP ONLY IF REQUIRED
    # =====================================================

    db_has_data = (
        DB_DIR.exists()
        and any(DB_DIR.iterdir())
    )

    hrms_has_data = (
        HRMS_UPLOAD_DIR.exists()
        and any(HRMS_UPLOAD_DIR.iterdir())
    )

    backup_root = None

    if db_has_data or hrms_has_data:

        backup_root = create_backup_folder()

        print()
        print(
            f"Backup location:\n{backup_root}"
        )

    # =====================================================
    # BACKUP DB
    # =====================================================

    if db_has_data:

        backup_folder(
            DB_DIR,
            backup_root
        )

    # =====================================================
    # BACKUP HRMS
    # =====================================================

    if hrms_has_data:

        backup_folder(
            HRMS_UPLOAD_DIR,
            backup_root
        )

    # =====================================================
    # DELETE PF DOWNLOADS
    # =====================================================

    delete_folder_content(
        PF_DOWNLOAD_DIR
    )

    # =====================================================
    # DELETE CAPTCHA IMAGES
    # =====================================================

    delete_folder_content(
        CAPTCHA_DIR
    )

    # =====================================================
    # DELETE DB WORKING DATA
    # =====================================================

    delete_folder_content(
        DB_DIR
    )

    # =====================================================
    # DELETE HRMS WORKING DATA
    # =====================================================

    delete_folder_content(
        HRMS_UPLOAD_DIR
    )

    # =====================================================
    # FINAL
    # =====================================================

    print()
    print("=" * 70)
    print("CLEANUP COMPLETED")
    print("=" * 70)

    if backup_root:

        print()
        print(
            f"DB + HRMS backup saved at:"
        )

        print(
            backup_root
        )

    else:

        print()
        print(
            "DB ya HRMS me purana data nahi tha."
        )

    print()
    print(
        "PF downloads      : DELETED"
    )

    print(
        "Captcha images    : DELETED"
    )

    print(
        "DB working data   : CLEAN"
    )

    print(
        "HRMS working data : CLEAN"
    )

    print()
    print(
        "Fresh run ke liye system ready hai."
    )


# =========================================================
# DIRECT RUN
# =========================================================

if __name__ == "__main__":

    try:

        cleanup_old_data()

    except Exception as e:

        print()
        print("=" * 70)
        print("CLEANUP ERROR")
        print("=" * 70)

        print()
        print(str(e))

        raise