import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd
import psycopg2
from dotenv import load_dotenv

from config import DB_OUTPUT

load_dotenv()

# =========================
# POSTGRESQL CONFIG
# =========================

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "port": os.getenv("DB_PORT", "5432"),
    "dbname": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
}

# =========================
# FULL EPFOMaster DATA
# =========================

SQL_QUERY = '''
SELECT *
FROM "EPFOMaster" e;
'''


def download_db_data():
    missing = [key for key, value in DB_CONFIG.items() if not value]

    if missing:
        raise ValueError(
            "Missing database settings in .env: "
            + ", ".join(missing)
        )

    DB_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 50)
    print("PostgreSQL DB DATA DOWNLOAD")
    print("=" * 50)
    print("Connecting to PostgreSQL...")

    conn = psycopg2.connect(**DB_CONFIG)

    try:
        print("Connected successfully.")
        print('Fetching complete data from "EPFOMaster"...')

        df = pd.read_sql_query(SQL_QUERY, conn)

        # Always create a fresh Excel file.
        if DB_OUTPUT.exists():
            DB_OUTPUT.unlink()

        df.to_excel(DB_OUTPUT, index=False, engine="openpyxl")

        print(f"Records fetched : {len(df):,}")
        print(f"Columns         : {len(df.columns):,}")
        print(f"Saved file      : {DB_OUTPUT}")
        print("DB download completed successfully.")

        return DB_OUTPUT

    finally:
        conn.close()
        print("PostgreSQL connection closed.")


if __name__ == "__main__":
    download_db_data()
