import sys
import shutil
import logging
from pathlib import Path
from datetime import datetime

# D DRIVE DYNAMIC ABSOLUTE PATHS
PROJECT_FOLDER = Path(__file__).resolve().parent

DATA_FOLDER = PROJECT_FOLDER / "data"
MERGE_FOLDER = PROJECT_FOLDER / "Merge_data"
DB_FOLDER = PROJECT_FOLDER / "dbdata"
COMPARE_FOLDER = PROJECT_FOLDER / "compaire"
ARCHIVE_FOLDER = PROJECT_FOLDER / "old_data"


logger = logging.getLogger("clean_data")
if not logger.hasHandlers():
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(ch)
    logger.setLevel(logging.INFO)


def delete_captcha_folders():
    if not DATA_FOLDER.exists():
        logger.info("data folder does not exist. Skipping CAPTCHA cleanup.")
        return

    deleted_count = 0
    # Check top-level captcha folder
    direct_captcha = DATA_FOLDER / "captcha"
    if direct_captcha.exists():
        try:
            shutil.rmtree(direct_captcha)
            deleted_count += 1
        except Exception as e:
            logger.error(f"Failed to delete CAPTCHA folder ({direct_captcha}): {e}")

    # Check subfolders for captcha
    for date_folder in DATA_FOLDER.iterdir():
        if not date_folder.is_dir():
            continue

        captcha_folder = date_folder / "captcha"
        if captcha_folder.exists():
            try:
                shutil.rmtree(captcha_folder)
                deleted_count += 1
            except Exception as e:
                logger.error(f"Failed to delete CAPTCHA folder ({captcha_folder}): {e}")

    logger.info(f"CAPTCHA folders deleted: {deleted_count}")


def move_folder_to_archive(source_folder, archive_root):
    if not source_folder.exists():
        return

    try:
        has_content = any(source_folder.iterdir())
    except Exception:
        has_content = True

    if not has_content:
        return

    destination = archive_root / source_folder.name
    if destination.exists():
        timestamp = datetime.now().strftime("%H%M%S")
        destination = archive_root / f"{source_folder.name}_{timestamp}"

    try:
        shutil.move(str(source_folder), str(destination))
        logger.info(f"Archived {source_folder.name} -> {destination}")
    except Exception as e:
        logger.error(f"Archive error for {source_folder.name}: {e}")


def clean_old_data():
    logger.info("Starting ESIC Old Data Cleanup...")

    if not PROJECT_FOLDER.exists():
        raise FileNotFoundError(f"Project folder not found: {PROJECT_FOLDER}")

    timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
    archive_root = ARCHIVE_FOLDER / timestamp
    archive_root.mkdir(parents=True, exist_ok=True)

    delete_captcha_folders()

    folders_to_archive = [
        DATA_FOLDER,
        MERGE_FOLDER,
        DB_FOLDER,
        COMPARE_FOLDER,
    ]

    for folder in folders_to_archive:
        move_folder_to_archive(folder, archive_root)

    DATA_FOLDER.mkdir(parents=True, exist_ok=True)
    MERGE_FOLDER.mkdir(parents=True, exist_ok=True)
    DB_FOLDER.mkdir(parents=True, exist_ok=True)
    COMPARE_FOLDER.mkdir(parents=True, exist_ok=True)

    logger.info("Fresh directories ready: data, Merge_data, dbdata, compaire")
    logger.info("Cleanup completed successfully.")
    return True


if __name__ == "__main__":
    try:
        clean_old_data()
    except Exception as exc:
        logger.error(f"Cleanup failed: {exc}")
        sys.exit(1)
