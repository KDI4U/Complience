import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd

from config import TARGET_DOWNLOAD_FOLDER, DB_OUTPUT


# =========================================================
# OUTPUT FILES
# =========================================================

# DB folder ke andar All_Merge.xlsx banega
OUTPUT_FOLDER = DB_OUTPUT.parent

EXCEL_OUTPUT = OUTPUT_FOLDER / "All_Merge.xlsx"
CSV_OUTPUT = OUTPUT_FOLDER / "All_Merge.csv"
REPORT_OUTPUT = OUTPUT_FOLDER / "Merge_Report.xlsx"


# =========================================================
# SUPPORTED FILE TYPES
# =========================================================

SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".csv"
}


# =========================================================
# READ EXCEL / CSV
# =========================================================

def read_file(file_path):

    file_path = Path(file_path)

    suffix = file_path.suffix.lower()

    # -----------------------------------------------------
    # CSV
    # -----------------------------------------------------

    if suffix == ".csv":

        try:
            return pd.read_csv(
                file_path,
                encoding="utf-8"
            )

        except UnicodeDecodeError:

            return pd.read_csv(
                file_path,
                encoding="cp1252"
            )

    # -----------------------------------------------------
    # EXCEL
    # -----------------------------------------------------

    if suffix in {".xlsx", ".xls"}:

        return pd.read_excel(
            file_path,
            sheet_name=0
        )

    raise ValueError(
        f"Unsupported file type: {file_path.name}"
    )


# =========================================================
# FIND DOWNLOAD FILES
# =========================================================

def find_input_files():

    folder = Path(TARGET_DOWNLOAD_FOLDER)

    if not folder.exists():

        raise FileNotFoundError(
            f"EPFO download folder nahi mila:\n{folder}"
        )

    files = []

    for file_path in sorted(folder.iterdir()):

        if not file_path.is_file():
            continue

        if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue

        # -------------------------------------------------
        # Apni output files ko dobara merge nahi karna
        # -------------------------------------------------

        if file_path.name.lower() in {
            "all_merge.xlsx",
            "all_merge.csv",
            "merge_report.xlsx"
        }:
            continue

        files.append(file_path)

    return files


# =========================================================
# MERGE FILES
# =========================================================

def merge_files():

    print()
    print("=" * 70)
    print("AUTOMATIC EXCEL MERGE")
    print("=" * 70)

    print()
    print(f"Input folder:")
    print(TARGET_DOWNLOAD_FOLDER)

    # =====================================================
    # FIND FILES
    # =====================================================

    files = find_input_files()

    if not files:

        raise FileNotFoundError(
            "EPFO download folder me koi Excel/CSV file nahi mili."
        )

    print()
    print(
        f"Total files found : {len(files)}"
    )

    # =====================================================
    # READ FILES
    # =====================================================

    data = []
    failed = []
    empty = []

    for i, file_path in enumerate(files, 1):

        print()
        print(
            f"[{i}/{len(files)}] Reading: "
            f"{file_path.name}"
        )

        try:

            df = read_file(file_path)

            # ------------------------------------------------
            # Empty file
            # ------------------------------------------------

            if df.empty:

                empty.append(
                    file_path.name
                )

                print(
                    "   STATUS : EMPTY"
                )

                continue

            # ------------------------------------------------
            # Valid file
            # ------------------------------------------------

            data.append(
                (
                    file_path.name,
                    df
                )
            )

            print(
                f"   STATUS : OK"
            )

            print(
                f"   Rows   : {len(df):,}"
            )

            print(
                f"   Columns: {len(df.columns)}"
            )

        except Exception as e:

            failed.append(
                {
                    "file": file_path.name,
                    "reason": str(e)
                }
            )

            print(
                f"   STATUS : FAILED"
            )

            print(
                f"   Reason : {e}"
            )

    # =====================================================
    # CHECK USABLE DATA
    # =====================================================

    if not data:

        raise RuntimeError(
            "Kisi bhi downloaded file se usable data nahi mila."
        )

    # =====================================================
    # MERGE
    # =====================================================

    print()
    print("=" * 70)
    print("FILES MERGE HO RAHI HAIN")
    print("=" * 70)

    merged = pd.concat(
        [
            df
            for name, df in data
        ],
        ignore_index=True,
        sort=False
    )

    # =====================================================
    # CREATE OUTPUT FOLDER
    # =====================================================

    OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    # =====================================================
    # SAVE EXCEL
    # =====================================================

    print()
    print("All_Merge.xlsx create ho rahi hai...")

    merged.to_excel(
        EXCEL_OUTPUT,
        sheet_name="All_Merge",
        index=False
    )

    # =====================================================
    # SAVE CSV
    # =====================================================

    print(
        "All_Merge.csv create ho rahi hai..."
    )

    merged.to_csv(
        CSV_OUTPUT,
        index=False,
        encoding="utf-8-sig"
    )

    # =====================================================
    # MERGE REPORT
    # =====================================================

    report_rows = []

    for item in failed:

        report_rows.append(
            {
                "File": item["file"],
                "Status": "FAILED",
                "Reason": item["reason"]
            }
        )

    for file_name in empty:

        report_rows.append(
            {
                "File": file_name,
                "Status": "EMPTY",
                "Reason": "File me data rows nahi mili."
            }
        )

    report_df = pd.DataFrame(
        report_rows,
        columns=[
            "File",
            "Status",
            "Reason"
        ]
    )

    summary_df = pd.DataFrame(
        [
            [
                "Input Folder",
                str(TARGET_DOWNLOAD_FOLDER)
            ],
            [
                "Total Files Found",
                len(files)
            ],
            [
                "Successfully Merged",
                len(data)
            ],
            [
                "Failed",
                len(failed)
            ],
            [
                "Empty",
                len(empty)
            ],
            [
                "Total Rows",
                len(merged)
            ],
            [
                "Total Columns",
                len(merged.columns)
            ]
        ],
        columns=[
            "Item",
            "Value"
        ]
    )

    with pd.ExcelWriter(
        REPORT_OUTPUT,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        report_df.to_excel(
            writer,
            sheet_name="Not Merged",
            index=False
        )

    # =====================================================
    # FINAL OUTPUT
    # =====================================================

    print()
    print("=" * 70)
    print("AUTOMATIC MERGE COMPLETED")
    print("=" * 70)

    print()
    print(
        f"Files found       : {len(files)}"
    )

    print(
        f"Successfully merged: {len(data)}"
    )

    print(
        f"Failed            : {len(failed)}"
    )

    print(
        f"Empty             : {len(empty)}"
    )

    print(
        f"Total rows        : {len(merged):,}"
    )

    print()
    print("OUTPUT FILES:")

    print(
        f"Excel  : {EXCEL_OUTPUT}"
    )

    print(
        f"CSV    : {CSV_OUTPUT}"
    )

    print(
        f"Report : {REPORT_OUTPUT}"
    )

    print()
    print("=" * 70)

    return True


# =========================================================
# DIRECT RUN
# =========================================================

if __name__ == "__main__":

    try:

        merge_files()

    except Exception as e:

        print()
        print("=" * 70)
        print("MERGE ERROR")
        print("=" * 70)

        print()
        print(str(e))

        raise