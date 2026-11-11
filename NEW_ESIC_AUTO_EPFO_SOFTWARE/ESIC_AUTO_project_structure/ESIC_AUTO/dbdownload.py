import os
import sys
import logging
from pathlib import Path
import pandas as pd
import psycopg2
from dotenv import load_dotenv

# D DRIVE DYNAMIC ABSOLUTE PATHS
PROJECT_FOLDER = Path(__file__).resolve().parent
DB_FOLDER = PROJECT_FOLDER / "dbdata"
DB_OUTPUT = DB_FOLDER / "dbdata.xlsx"
TEMP_OUTPUT = DB_FOLDER / "dbdata_temp.xlsx"
ENV_FILE = PROJECT_FOLDER / ".env"

load_dotenv(ENV_FILE)

logger = logging.getLogger("db_download")
if not logger.hasHandlers():
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(ch)
    logger.setLevel(logging.INFO)

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "port": os.getenv("DB_PORT", "5432"),
    "dbname": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
}

SQL_QUERY = 'SELECT * FROM "ESICMaster" e;'


def download_db_data():
    logger.info("Starting PostgreSQL ESIC Master Download...")

    missing_env = [key for key, val in DB_CONFIG.items() if not val]
    if missing_env:
        raise ValueError(f"Missing database configuration in .env: {', '.join(missing_env)}")

    DB_FOLDER.mkdir(parents=True, exist_ok=True)

    logger.info(f"Connecting to PostgreSQL ({DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['dbname']})...")
    conn = psycopg2.connect(**DB_CONFIG)

    try:
        cursor = conn.cursor()
        cursor.execute(SQL_QUERY)
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        cursor.close()

        df = pd.DataFrame(rows, columns=columns)
        logger.info(f"Fetched {len(df):,} records and {len(df.columns)} columns from DB.")

        if TEMP_OUTPUT.exists():
            try:
                TEMP_OUTPUT.unlink()
            except PermissionError:
                raise PermissionError(f"Temporary file locked: {TEMP_OUTPUT}")

        df.to_excel(TEMP_OUTPUT, index=False, engine="openpyxl")

        if DB_OUTPUT.exists():
            try:
                os.replace(TEMP_OUTPUT, DB_OUTPUT)
            except PermissionError:
                if TEMP_OUTPUT.exists():
                    try:
                        TEMP_OUTPUT.unlink()
                    except Exception:
                        pass
                raise PermissionError(f"{DB_OUTPUT.name} is open or locked by another program.")
        else:
            os.replace(TEMP_OUTPUT, DB_OUTPUT)

        logger.info(f"Database data successfully saved -> {DB_OUTPUT}")
        return DB_OUTPUT
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        download_db_data()
    except Exception as exc:
        logger.error(f"DB Download failed: {exc}")
        sys.exit(1)
