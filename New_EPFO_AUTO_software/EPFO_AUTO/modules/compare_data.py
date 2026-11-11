import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd

from config import (
    DB_OUTPUT,
    HRMS_UPLOAD_FOLDER,
    COMPARISON_OUTPUT,
    HRMS_UPLOAD_DIR,
    HRMS_UPLOAD_FILE,
)


# =========================================================
# EPFO MERGED FILE
# =========================================================

EPFO_MERGED_FILE = DB_OUTPUT.parent / "All_Merge.xlsx"


# =========================================================
# EXACT COLUMN MAPPING
# =========================================================
#
# All_Merge.xlsx              DB
# ---------------------------------------------------------
# UAN                         UAN
# PF NO                       PFNO
# Employee Name               EmployeeName
# Gender                      Gender
# DoB                         DoB
# DoJ                         DoJ
# Father's/Husband's Name     FatherHusbandName
# NameRelation                NameRelation
# Marital Status              MaritalStatus
# Mobile                      Mobile
# Email ID                    EmailID
# EPFO Employer Code          EPFOEmployerCode
# Nomination Filed            NominationFiled
# Is AADHAAR Verified         IsAadhaarVerified
# Face Auth Status            FaceAuthStatus
# =========================================================

COLUMN_MAPPING = {
    "UAN": "UAN",
    "PF NO": "PFNO",
    "Employee Name": "EmployeeName",
    "Gender": "Gender",
    "DoB": "DoB",
    "DoJ": "DoJ",
    "Father's/Husband's Name": "FatherHusbandName",
    "NameRelation": "NameRelation",
    "Marital Status": "MaritalStatus",
    "Mobile": "Mobile",
    "Email ID": "EmailID",
    "EPFO Employer Code": "EPFOEmployerCode",
    "Nomination Filed": "NominationFiled",
    "Is AADHAAR Verified": "IsAadhaarVerified",
    "Face Auth Status": "FaceAuthStatus",
}

COLUMN_ALIASES = {
    "UAN": ["UAN", "UAN NUMBER", "UAN NO"],
    "PF NO": ["PF NO", "PFNO", "PF_NO", "PF NUMBER", "PF NO."],
    "Employee Name": ["Employee Name", "EmployeeName", "MEMBER NAME", "EMPLOYEE NAME", "NAME"],
    "Gender": ["Gender", "GENDER"],
    "DoB": ["DoB", "DOB", "DATE OF BIRTH"],
    "DoJ": ["DoJ", "DOJ", "DATE OF JOINING"],
    "Father's/Husband's Name": [
        "Father's/Husband's Name",
        "Father's/Husband's",
        "FatherHusbandName",
        "FATHER'S/HUSBAND'S NAME",
        "FATHER/HUSBAND NAME",
        "Father / Husband Name",
        "FATHER NAME",
    ],
    "NameRelation": ["NameRelation", "NAME RELATION", "RELATION", "RELATIONSHIP"],
    "Marital Status": ["Marital Status", "MaritalStatus", "MARITAL_STATUS"],
    "Mobile": ["Mobile", "MOBILE", "MOBILE NO", "MOBILENO"],
    "Email ID": ["Email ID", "EmailID", "EMAIL", "EMAIL_ID"],
    "EPFO Employer Code": ["EPFO Employer Code", "EPFOEmployerCode", "ESTABLISHMENT ID", "EST ID"],
    "Nomination Filed": ["Nomination Filed", "NominationFiled", "NOMINATION FILED"],
    "Is AADHAAR Verified": ["Is AADHAAR Verified", "IsAadhaarVerified", "AADHAAR VERIFIED"],
    "Face Auth Status": ["Face Auth Status", "FaceAuthStatus", "FACE AUTH STATUS"],
}


def normalize_col_key(name):
    return "".join(ch for ch in str(name).upper() if ch.isalnum())


def resolve_column_mapping(epfo_columns):
    """
    All_Merge.xlsx ke actual columns ko DB columns ke sath map karta hai,
    chahe column me space ho ya na ho, ya alternate naming convention ho.
    """
    epfo_columns_list = list(epfo_columns)
    epfo_col_norm = {normalize_col_key(c): c for c in epfo_columns_list}

    resolved_mapping = {}
    missing_columns = []

    for primary_col, db_col in COLUMN_MAPPING.items():
        found_col = None
        if primary_col in epfo_columns_list:
            found_col = primary_col
        else:
            aliases = COLUMN_ALIASES.get(primary_col, [primary_col])
            for alias in aliases:
                if alias in epfo_columns_list:
                    found_col = alias
                    break
            if not found_col:
                norm_primary = normalize_col_key(primary_col)
                if norm_primary in epfo_col_norm:
                    found_col = epfo_col_norm[norm_primary]
                else:
                    for alias in aliases:
                        norm_alias = normalize_col_key(alias)
                        if norm_alias in epfo_col_norm:
                            found_col = epfo_col_norm[norm_alias]
                            break

        if found_col:
            resolved_mapping[found_col] = db_col
        else:
            missing_columns.append(primary_col)

    return resolved_mapping, missing_columns



# =========================================================
# BOOLEAN COLUMNS
# =========================================================
#
# All_Merge:
# YES / NO
#
# DB:
# TRUE / FALSE
#
# YES == TRUE
# NO  == FALSE
# =========================================================

BOOLEAN_COLUMNS = {
    "Nomination Filed",
    "NominationFiled",
    "Is AADHAAR Verified",
    "IsAadhaarVerified",
    "Face Auth Status",
    "FaceAuthStatus",
}


# =========================================================
# DATE COLUMNS
# =========================================================

DATE_COLUMNS = {
    "DoB",
    "DOB",
    "DoJ",
    "DOJ",
}


# =========================================================
# CLEAN EXCEL HEADERS
# =========================================================

def clean_headers(df):
    """
    Excel column names ke starting/ending spaces remove karta hai.
    """

    df = df.copy()

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    return df


# =========================================================
# UAN COLUMN FINDER
# =========================================================

def find_uan_column(df):

    for column in df.columns:

        normalized = (
            str(column)
            .strip()
            .upper()
            .replace("_", "")
            .replace("-", "")
            .replace(" ", "")
            .replace(".", "")
            .replace("'", "")
        )

        if normalized in {
            "UAN",
            "UANNUMBER",
            "UANNO",
        }:

            return column

    return None


# =========================================================
# NORMALIZE UAN
# =========================================================

def normalize_uan(value):

    if pd.isna(value):
        return ""

    value_string = str(value).strip()

    # Excel numeric UAN:
    # 1234567890.0
    # ->
    # 1234567890

    if value_string.endswith(".0"):

        try:

            numeric_value = float(value_string)

            if numeric_value.is_integer():

                value_string = str(
                    int(numeric_value)
                )

        except Exception:
            pass

    return value_string.strip()


# =========================================================
# NORMALIZE DATE
# =========================================================

def normalize_date(value):

    if pd.isna(value):
        return ""

    # -----------------------------------------------------
    # Pandas Timestamp
    # -----------------------------------------------------

    if isinstance(value, pd.Timestamp):

        return value.strftime("%Y-%m-%d")

    # -----------------------------------------------------
    # Python datetime/date
    # -----------------------------------------------------

    try:

        if hasattr(value, "strftime"):

            return value.strftime("%Y-%m-%d")

    except Exception:
        pass

    value_string = str(value).strip()

    if not value_string:
        return ""

    # =====================================================
    # DATE FORMATS
    # =====================================================

    formats = (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d/%b/%Y",
        "%d/%B/%Y",

        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y.%m.%d",

        "%d/%m/%y",
        "%d-%m-%y",
        "%d.%m.%y",
        "%d-%b-%y",
    )

    # -----------------------------------------------------
    # Explicit format check
    # -----------------------------------------------------

    for fmt in formats:

        try:

            parsed = pd.to_datetime(
                value_string,
                format=fmt,
                errors="coerce",
            )

            if not pd.isna(parsed):

                return parsed.strftime("%Y-%m-%d")

        except Exception:
            pass

    # -----------------------------------------------------
    # General fallback
    # -----------------------------------------------------

    try:

        parsed = pd.to_datetime(
            value_string,
            errors="coerce",
        )

        if not pd.isna(parsed):

            return parsed.strftime("%Y-%m-%d")

    except Exception:
        pass

    return value_string.upper()


# =========================================================
# NORMALIZE YES / NO
# =========================================================

def normalize_boolean(value):

    if pd.isna(value):
        return ""

    # -----------------------------------------------------
    # Python boolean
    # -----------------------------------------------------

    if isinstance(value, bool):

        if value:
            return "YES"

        return "NO"

    value_string = (
        str(value)
        .strip()
        .upper()
    )

    # -----------------------------------------------------
    # TRUE SIDE
    # -----------------------------------------------------

    if value_string in {
        "YES",
        "TRUE",
        "1",
        "Y",
        "T",
    }:

        return "YES"

    # -----------------------------------------------------
    # FALSE SIDE
    # -----------------------------------------------------

    if value_string in {
        "NO",
        "FALSE",
        "0",
        "N",
        "F",
    }:

        return "NO"

    return value_string


# =========================================================
# NORMALIZE GENERAL VALUE
# =========================================================

def normalize_value(
    value,
    all_merge_column,
):

    if pd.isna(value):
        return ""

    col_norm = "".join(ch for ch in str(all_merge_column).upper() if ch.isalnum())

    # =====================================================
    # BOOLEAN
    # =====================================================

    if all_merge_column in BOOLEAN_COLUMNS or col_norm in {"NOMINATIONFILED", "ISAADHAARVERIFIED", "FACEAUTHSTATUS"}:

        return normalize_boolean(value)

    # =====================================================
    # DATE
    # =====================================================

    if all_merge_column in DATE_COLUMNS or col_norm in {"DOB", "DOJ"}:

        return normalize_date(value)

    # =====================================================
    # PYTHON BOOLEAN
    # =====================================================

    if isinstance(value, bool):

        return (
            "TRUE"
            if value
            else "FALSE"
        )

    # =====================================================
    # FLOAT
    # =====================================================

    if isinstance(value, float):

        if value.is_integer():

            return str(
                int(value)
            )

    # =====================================================
    # INTEGER
    # =====================================================

    if isinstance(value, int):

        return str(value)

    # =====================================================
    # STRING
    # =====================================================

    value_string = str(value).strip()

    # -----------------------------------------------------
    # Remove .0 from Excel numbers
    # -----------------------------------------------------

    if value_string.endswith(".0"):

        try:

            numeric_value = float(
                value_string
            )

            if numeric_value.is_integer():

                value_string = str(
                    int(numeric_value)
                )

        except Exception:
            pass

    # -----------------------------------------------------
    # Remove extra spaces
    # -----------------------------------------------------

    value_string = " ".join(
        value_string.split()
    )

    # -----------------------------------------------------
    # Uppercase
    # -----------------------------------------------------

    return value_string.upper()


# =========================================================
# FIND DUPLICATE UAN ROWS
# =========================================================

def get_duplicate_rows(
    df,
    uan_column,
    source_name,
):

    if uan_column not in df.columns:

        return pd.DataFrame()

    temp = df.copy()

    temp["_UAN_KEY"] = temp[
        uan_column
    ].apply(
        normalize_uan
    )

    # Ignore empty UAN

    temp = temp[
        temp["_UAN_KEY"] != ""
    ].copy()

    duplicate_mask = temp[
        "_UAN_KEY"
    ].duplicated(
        keep=False
    )

    duplicates = temp[
        duplicate_mask
    ].copy()

    if duplicates.empty:

        return pd.DataFrame()

    duplicates["SOURCE"] = source_name

    columns = [
        "SOURCE",
        uan_column,
    ]

    remaining_columns = [
        col
        for col in duplicates.columns
        if col not in columns
        and col != "_UAN_KEY"
    ]

    return duplicates[
        columns + remaining_columns
    ]


# =========================================================
# MAIN COMPARISON
# =========================================================

def compare_data():

    print()
    print("=" * 80)
    print("EPFO DATABASE COMPARISON")
    print("=" * 80)

    # =====================================================
    # CHECK FILES
    # =====================================================

    if not DB_OUTPUT.exists():

        raise FileNotFoundError(
            f"\nDB file nahi mila:\n"
            f"{DB_OUTPUT}"
        )

    if not EPFO_MERGED_FILE.exists():

        raise FileNotFoundError(
            f"\nAll_Merge.xlsx nahi mila:\n"
            f"{EPFO_MERGED_FILE}"
        )

    print()
    print(f"DB FILE     : {DB_OUTPUT}")
    print(f"ALL_MERGE   : {EPFO_MERGED_FILE}")

    # =====================================================
    # READ FILES
    # =====================================================

    print()
    print("Excel files read ho rahe hain...")

    db_df = pd.read_excel(
        DB_OUTPUT,
        dtype=object,
    )

    epfo_df = pd.read_excel(
        EPFO_MERGED_FILE,
        dtype=object,
    )

    # =====================================================
    # CLEAN HEADERS
    # =====================================================

    db_df = clean_headers(db_df)
    epfo_df = clean_headers(epfo_df)

    print()
    print(f"DB records        : {len(db_df)}")
    print(f"All_Merge records : {len(epfo_df)}")

    print()
    print(f"DB columns        : {len(db_df.columns)}")
    print(f"All_Merge columns : {len(epfo_df.columns)}")

    # =====================================================
    # SHOW ACTUAL DB COLUMNS
    # =====================================================

    print()
    print("=" * 80)
    print("ACTUAL DB COLUMNS")
    print("=" * 80)

    for number, column in enumerate(
        db_df.columns,
        start=1,
    ):

        print(
            f"{number}. {repr(column)}"
        )

    # =====================================================
    # SHOW ACTUAL ALL_MERGE COLUMNS
    # =====================================================

    print()
    print("=" * 80)
    print("ACTUAL ALL_MERGE COLUMNS")
    print("=" * 80)

    for number, column in enumerate(
        epfo_df.columns,
        start=1,
    ):

        print(
            f"{number}. {repr(column)}"
        )

    # =====================================================
    # FIND UAN COLUMNS
    # =====================================================

    db_uan_column = find_uan_column(
        db_df
    )

    epfo_uan_column = find_uan_column(
        epfo_df
    )

    if db_uan_column is None:

        raise ValueError(
            "\nDB me UAN column nahi mila."
        )

    if epfo_uan_column is None:

        raise ValueError(
            "\nAll_Merge.xlsx me UAN column nahi mila."
        )

    print()
    print(
        f"DB UAN Column        : "
        f"{db_uan_column}"
    )

    print(
        f"All_Merge UAN Column : "
        f"{epfo_uan_column}"
    )

    # =====================================================
    # CHECK REQUIRED ALL_MERGE COLUMNS
    # =====================================================

    print()
    print(
        "All_Merge columns check ho rahe hain..."
    )

    active_mapping, missing_all_merge = resolve_column_mapping(
        epfo_df.columns
    )

    if missing_all_merge:

        raise ValueError(
            "\nAll_Merge.xlsx me required "
            "columns nahi mile:\n"
            + "\n".join(
                str(column)
                for column in missing_all_merge
            )
        )

    # =====================================================
    # CHECK REQUIRED DB COLUMNS
    # =====================================================

    print()
    print(
        "DB columns check ho rahe hain..."
    )

    missing_db = [
        db_column
        for db_column in active_mapping.values()
        if db_column not in db_df.columns
    ]

    if missing_db:

        raise ValueError(
            "\nDB me required columns nahi mile:\n"
            + "\n".join(
                str(column)
                for column in missing_db
            )
        )

    # =====================================================
    # PRINT EXACT MAPPING
    # =====================================================

    print()
    print("=" * 80)
    print("COLUMN MAPPING")
    print("=" * 80)

    comparison_columns = []

    for number, (
        all_merge_column,
        db_column,
    ) in enumerate(
        active_mapping.items(),
        start=1,
    ):

        print(
            f"{number}. "
            f"All_Merge [{all_merge_column}] "
            f"<--> "
            f"DB [{db_column}]"
        )

        # UAN comparison key hai
        if all_merge_column != epfo_uan_column:

            comparison_columns.append(
                all_merge_column
            )

    print()
    print(
        f"TOTAL ALL_MERGE COLUMNS : "
        f"{len(active_mapping)}"
    )

    print(
        "UAN MATCHING KEY        : 1"
    )

    print(
        f"ACTUAL COLUMNS COMPARED : "
        f"{len(comparison_columns)}"
    )

    # =====================================================
    # DUPLICATES
    # =====================================================

    print()
    print(
        "Duplicate UAN check ho raha hai..."
    )

    db_duplicates = get_duplicate_rows(
        db_df,
        db_uan_column,
        "DB",
    )

    epfo_duplicates = get_duplicate_rows(
        epfo_df,
        epfo_uan_column,
        "EPFO_DOWNLOAD",
    )

    duplicate_parts = []

    if not db_duplicates.empty:

        duplicate_parts.append(
            db_duplicates
        )

    if not epfo_duplicates.empty:

        duplicate_parts.append(
            epfo_duplicates
        )

    if duplicate_parts:

        duplicates_df = pd.concat(
            duplicate_parts,
            ignore_index=True,
        )

    else:

        duplicates_df = pd.DataFrame()

    # =====================================================
    # CREATE UAN KEY
    # =====================================================

    db_df = db_df.copy()
    epfo_df = epfo_df.copy()

    db_df["_UAN_KEY"] = db_df[
        db_uan_column
    ].apply(
        normalize_uan
    )

    epfo_df["_UAN_KEY"] = epfo_df[
        epfo_uan_column
    ].apply(
        normalize_uan
    )

    # =====================================================
    # REMOVE EMPTY UAN
    # =====================================================

    db_df = db_df[
        db_df["_UAN_KEY"] != ""
    ].copy()

    epfo_df = epfo_df[
        epfo_df["_UAN_KEY"] != ""
    ].copy()

    # =====================================================
    # FIRST RECORD FOR DUPLICATE UAN
    # =====================================================

    db_compare = db_df.drop_duplicates(
        subset=[
            "_UAN_KEY"
        ],
        keep="first",
    ).copy()

    epfo_compare = epfo_df.drop_duplicates(
        subset=[
            "_UAN_KEY"
        ],
        keep="first",
    ).copy()

    # =====================================================
    # DB INDEX
    # =====================================================

    db_index = db_compare.set_index(
        "_UAN_KEY",
        drop=False,
    )

    # =====================================================
    # COMPARISON
    # =====================================================

    main_records = []
    compare_records = []

    new_count = 0
    updated_count = 0
    same_count = 0

    print()
    print("=" * 80)
    print("DATA COMPARISON START")
    print("=" * 80)

    # =====================================================
    # EACH ALL_MERGE ROW
    # =====================================================

    for _, epfo_row in epfo_compare.iterrows():

        uan_key = epfo_row[
            "_UAN_KEY"
        ]

        # =================================================
        # NEW RECORD
        # =================================================

        if uan_key not in db_index.index:

            status = "NEW"
            changed_fields = ""

            new_count += 1

            main_row = epfo_row.drop(
                labels=[
                    "_UAN_KEY"
                ]
            ).to_dict()

            main_records.append(
                main_row
            )

        # =================================================
        # EXISTING RECORD
        # =================================================

        else:

            db_row = db_index.loc[
                uan_key
            ]

            if isinstance(db_row, pd.DataFrame):
                db_row = db_row.iloc[0]

            changed_fields_list = []

            # =============================================
            # COMPARE EACH ALL_MERGE COLUMN
            # =============================================

            for all_merge_column in comparison_columns:

                db_column = active_mapping[
                    all_merge_column
                ]

                # -----------------------------------------
                # All_Merge value
                # -----------------------------------------

                all_merge_value = normalize_value(
                    epfo_row[
                        all_merge_column
                    ],
                    all_merge_column,
                )

                # -----------------------------------------
                # DB value
                # -----------------------------------------

                db_value = normalize_value(
                    db_row[
                        db_column
                    ],
                    all_merge_column,
                )

                # -----------------------------------------
                # COMPARE
                # -----------------------------------------

                if (
                    all_merge_value
                    != db_value
                ):

                    changed_fields_list.append(
                        all_merge_column
                    )

            # =============================================
            # UPDATED
            # =============================================

            if changed_fields_list:

                status = "UPDATED"

                changed_fields = ", ".join(
                    changed_fields_list
                )

                updated_count += 1

                main_row = epfo_row.drop(
                    labels=[
                        "_UAN_KEY"
                    ]
                ).to_dict()

                main_records.append(
                    main_row
                )

            # =============================================
            # SAME
            # =============================================

            else:

                status = "SAME"
                changed_fields = ""

                same_count += 1

        # =================================================
        # COMPARE SHEET ROW
        # =================================================

        compare_row = epfo_row.drop(
            labels=[
                "_UAN_KEY"
            ]
        ).to_dict()

        compare_row[
            "STATUS"
        ] = status

        compare_row[
            "CHANGED_FIELDS"
        ] = changed_fields

        compare_records.append(
            compare_row
        )

    # =====================================================
    # CREATE DATAFRAMES
    # =====================================================

    main_df = pd.DataFrame(
        main_records
    )

    compare_df = pd.DataFrame(
        compare_records
    )

    # =====================================================
    # ORIGINAL ALL_MERGE COLUMN ORDER
    # =====================================================

    original_columns = [
        column
        for column in epfo_df.columns
        if column != "_UAN_KEY"
    ]

    # =====================================================
    # MAIN COLUMN ORDER
    # =====================================================

    if not main_df.empty:

        main_columns = [
            column
            for column in original_columns
            if column in main_df.columns
        ]

        main_df = main_df[
            main_columns
        ]

    # =====================================================
    # COMPARE COLUMN ORDER
    # =====================================================

    if not compare_df.empty:

        compare_columns = [
            column
            for column in original_columns
            if column in compare_df.columns
        ]

        compare_columns.extend(
            [
                "STATUS",
                "CHANGED_FIELDS",
            ]
        )

        compare_df = compare_df[
            compare_columns
        ]

    # =====================================================
    # CREATE FOLDERS
    # =====================================================

    HRMS_UPLOAD_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    HRMS_UPLOAD_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # =====================================================
    # CREATE EXCEL OUTPUT
    # =====================================================

    print()
    print(
        "EPFO_Comparison_Result.xlsx "
        "create ho rahi hai..."
    )

    with pd.ExcelWriter(
        COMPARISON_OUTPUT,
        engine="openpyxl",
    ) as writer:

        # -------------------------------------------------
        # MAIN
        # -------------------------------------------------

        main_df.to_excel(
            writer,
            sheet_name="MAIN",
            index=False,
        )

        # -------------------------------------------------
        # COMPARE
        # -------------------------------------------------

        compare_df.to_excel(
            writer,
            sheet_name="COMPARE",
            index=False,
        )

        # -------------------------------------------------
        # DUPLICATES
        # -------------------------------------------------

        if not duplicates_df.empty:

            duplicates_df.to_excel(
                writer,
                sheet_name="DUPLICATES",
                index=False,
            )

        else:

            empty_duplicates = pd.DataFrame(
                columns=[
                    "SOURCE",
                    "UAN",
                    "INFO",
                ]
            )

            empty_duplicates.to_excel(
                writer,
                sheet_name="DUPLICATES",
                index=False,
            )

    # =====================================================
    # CREATE HRMS CSV
    # =====================================================

    print()
    print(
        "HRMS upload CSV create ho rahi hai..."
    )

    # Only NEW + UPDATED

    main_df.to_csv(
        HRMS_UPLOAD_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    # =====================================================
    # FINAL RESULT
    # =====================================================

    print()
    print("=" * 80)
    print("COMPARISON COMPLETED")
    print("=" * 80)

    print()

    print(
        f"NEW records              : "
        f"{new_count}"
    )

    print(
        f"UPDATED records          : "
        f"{updated_count}"
    )

    print(
        f"SAME records             : "
        f"{same_count}"
    )

    print()

    print(
        f"DB duplicate rows        : "
        f"{len(db_duplicates)}"
    )

    print(
        f"All_Merge duplicate rows : "
        f"{len(epfo_duplicates)}"
    )

    print()

    print(
        f"Total columns            : "
        f"{len(active_mapping)}"
    )

    print(
        "UAN matching key         : 1"
    )

    print(
        f"Columns actually compared: "
        f"{len(comparison_columns)}"
    )

    print()

    print(
        f"MAIN records             : "
        f"{len(main_df)}"
    )

    print()

    print(
        "COMPARISON FILE:"
    )

    print(
        COMPARISON_OUTPUT
    )

    print()

    print(
        "HRMS UPLOAD CSV:"
    )

    print(
        HRMS_UPLOAD_FILE
    )

    print()

    print(
        "HRMS CSV me sirf NEW + UPDATED "
        "records hain."
    )

    print()
    print("=" * 80)

    return True


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    try:

        compare_data()

    except Exception as e:

        print()
        print("=" * 80)
        print("COMPARISON ERROR")
        print("=" * 80)

        print()

        print(
            str(e)
        )

        print()

        raise