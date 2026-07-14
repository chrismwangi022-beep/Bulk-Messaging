"""
=========================================================
DAILY DUES REPORT COMBINER
=========================================================

Combines all Excel/CSV files from Raw Data folder.

SAFEGUARDS:
- Preserves original mobile number formatting (NO loss of leading zero)
- Prevents pandas numeric conversion issues
- Does NOT modify any column values

FILTERS:
- Mobile column must NOT be empty
- Member Number must start with 100

Creates:
Daily_Dues_SMS.xlsx
=========================================================
"""

import os
import glob
import pandas as pd
import traceback


CONFIG = {

    "INPUT_FOLDER":
    r"C:\Users\ADMIN\Desktop\Christopher\Bulk Messaging\Dues\Raw Data",

    "OUTPUT_FOLDER":
    r"C:\Users\ADMIN\Desktop\Christopher\Bulk Messaging\Dues",

    "OUTPUT_FILE_NAME":
    "Daily_Dues_SMS.xlsx"

}


# ===============================
# READ FILE (FORCE TEXT MODE)
# ===============================

def read_file(path):

    ext = os.path.splitext(path)[1].lower()

    try:

        # 🔥 FORCE EVERYTHING AS STRING TO PRESERVE FORMATTING
        if ext in [".xlsx", ".xls", ".xlsm"]:
            df = pd.read_excel(path, dtype=str)

        elif ext == ".csv":
            df = pd.read_csv(path, dtype=str)

        else:
            return None

        return df

    except Exception as e:
        print(f"Failed reading {path}: {e}")
        return None


# ===============================
# MAIN COMBINER
# ===============================

def combine_files():

    folder = CONFIG["INPUT_FOLDER"]

    files = []

    for pattern in ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]:
        files.extend(glob.glob(os.path.join(folder, pattern)))

    if not files:
        print("No files found")
        return

    print(f"Found {len(files)} files")

    data = []
    removed_rows = 0

    for file in files:

        print("Reading:", os.path.basename(file))

        df = read_file(file)

        if df is None:
            continue

        before = len(df)

        # ===============================
        # FIND COLUMNS (SAFE MATCHING)
        # ===============================

        mobile_col = None
        member_col = None

        for col in df.columns:

            clean = str(col).lower().replace(" ", "")

            if clean in [
                "mobile#",
                "mobile",
                "mobileno",
                "mobilenumber",
                "phone",
                "phoneno"
            ]:
                mobile_col = col

            if clean in [
                "memberno",
                "membernumber",
                "member#",
                "member"
            ]:
                member_col = col

        # ===============================
        # FILTER 1: MOBILE (NO MODIFICATION)
        # ===============================

        if mobile_col is not None:

            mobile_mask = (
                df[mobile_col]
                .notna() &
                (df[mobile_col].astype(str).str.strip() != "")
            )

        else:
            mobile_mask = pd.Series([True] * len(df))

        # ===============================
        # FILTER 2: MEMBER NUMBER RULE
        # ===============================

        if member_col is not None:

            member_mask = (
                df[member_col]
                .astype(str)
                .str.strip()
                .str.startswith("100")
            )

        else:
            member_mask = pd.Series([True] * len(df))

        # ===============================
        # APPLY FILTERS
        # ===============================

        df = df[mobile_mask & member_mask]

        removed_rows += before - len(df)

        data.append(df)

    if not data:

        print("No usable files")
        return

    # ===============================
    # COMBINE
    # ===============================

    final = pd.concat(
        data,
        ignore_index=True,
        sort=False
    )

    os.makedirs(
        CONFIG["OUTPUT_FOLDER"],
        exist_ok=True
    )

    output = os.path.join(
        CONFIG["OUTPUT_FOLDER"],
        CONFIG["OUTPUT_FILE_NAME"]
    )

    # ===============================
    # EXPORT (SAFE EXCEL OUTPUT)
    # ===============================

    final.to_excel(
        output,
        index=False
    )

    print("==============================")
    print("COMPLETED")
    print("Final rows:", len(final))
    print("Removed invalid rows:", removed_rows)
    print("Saved:", output)
    print("==============================")



if __name__ == "__main__":

    try:
        combine_files()

    except Exception:
        traceback.print_exc()