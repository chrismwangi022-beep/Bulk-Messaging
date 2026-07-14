"""
================================================================================
ARREARS DATA PROCESSING SCRIPT
================================================================================
Reads all raw arrears Excel/CSV files from an input folder, combines them,
filters them down to valid, SMS-ready customer records, and writes a single
clean output file.

A row is kept only if:
  1. It has a MobileNo that is present, non-blank, and a valid-looking number.
  2. Its arrears "Days" value falls within MIN_DAYS_ARREARS..MAX_DAYS_ARREARS
     (inclusive), after all files have been merged.

Edit the CONFIG block below to point at your own folders/settings — nothing
else needs to change for day-to-day use.
================================================================================
"""

import os
import re
import glob
import traceback
import pandas as pd

# ==============================================================================
# CONFIG — EDIT THESE VALUES AS NEEDED
# ==============================================================================
CONFIG = {
    # Folder containing the raw branch arrears files (.xlsx, .xls, .xlsm, .csv)
    "INPUT_FOLDER": r"C:\Users\ADMIN\Desktop\Christopher\Arrears Reports\Arears Reports formating folder\Documents",

    # Folder where the cleaned, combined output will be written
    "OUTPUT_FOLDER": r"C:\Users\ADMIN\Desktop\Christopher\Bulk Messaging",

    # Final output file name. Extension decides the format (.xlsx or .csv)
    "OUTPUT_FILE_NAME": "Daily_Arrears_SMS.xlsx",

    # Arrears aging window (in days) — inclusive on both ends
    "MIN_DAYS_ARREARS": 1,
    "MAX_DAYS_ARREARS": 60,

    # Mobile number sanity-check bounds (digit count, after cleaning)
    "MOBILE_MIN_DIGITS": 9,
    "MOBILE_MAX_DIGITS": 13,

    # If True, normalizes common Kenyan mobile formats to local "07.../01..." form:
    #   2547XXXXXXXX / 2541XXXXXXXX  -> 07XXXXXXXX / 01XXXXXXXX
    #   7XXXXXXXX    / 1XXXXXXXX     -> 07XXXXXXXX / 01XXXXXXXX  (missing leading 0)
    "NORMALIZE_KENYAN_MOBILE": True,
}
# ==============================================================================


# ------------------------------------------------------------------------------
# Column standardization map.
# Canonical name -> list of raw header variants it should match (case- and
# punctuation-insensitive). Extend this if your branch files use other labels.
# ------------------------------------------------------------------------------
COLUMN_ALIASES = {
    "MemberNo": [
        "memberno", "member no", "memberid", "member id", "clientno", "client no",
        "accountno", "account no", "loanno", "loan no", "loan number", "id",
    ],
    "MemberName": [
        "membername", "member name", "clientname", "client name", "customername",
        "customer name", "name", "fullname", "full name",
    ],
    "MobileNo": [
        "mobileno", "mobile no", "mobilenumber", "mobile number", "phoneno",
        "phone no", "phonenumber", "phone number", "msisdn", "contact",
        "contactno", "contact no", "telno", "tel no", "cellphone", "cellno",
    ],
    "ProductName": [
        "productname", "product name", "product", "loanproduct", "loan product",
    ],
    "Date": [
        "date", "disbursementdate", "disbursement date", "loandate", "loan date",
        "loandisbursementdate",
    ],
    "Principle": ["principle", "principal"],
    "Interest": ["interest"],
    "TotalLoan": ["totalloan", "total loan", "loanamount", "loan amount"],
    "PBalance": ["pbalance", "principalbalance", "principal balance"],
    "IBalance": ["ibalance", "interestbalance", "interest balance"],
    "TotalBalance": [
        "totalbalance", "total balance", "balance", "outstandingbalance",
        "outstanding balance",
    ],
    "Days": [
        "days", "daysinarrears", "days in arrears", "arrearsdays", "arrears days",
        "dpd", "dayspastdue", "days past due", "agingdays", "aging days",
    ],
    "Total": ["total"],
    "DueDate": ["duedate", "due date"],
    "Installment": ["installment", "instalment"],
    "AmountDue": [
        "amountdue", "amount due", "arrearsamount", "arrears amount", "dueamount",
        "due amount", "amountinarrears", "amount in arrears",
    ],
}

# Final column order for the standardized SMS-ready output
OUTPUT_COLUMN_ORDER = [
    "MemberNo", "MemberName", "MobileNo", "ProductName", "Days", "AmountDue",
    "TotalBalance", "DueDate", "Date", "Principle", "Interest", "TotalLoan",
    "PBalance", "IBalance", "Installment", "Total", "SourceFile",
]

# Numeric columns (besides Days) that should be cleaned to numbers if present
EXTRA_NUMERIC_COLUMNS = [
    "AmountDue", "TotalBalance", "Principle", "Interest", "TotalLoan",
    "PBalance", "IBalance", "Total", "Installment",
]

# Date-like columns that should be parsed to datetime if present
DATE_COLUMNS = ["Date", "DueDate"]


def normalize_header(col):
    """Lowercase + strip everything except letters/digits, for fuzzy header matching."""
    return re.sub(r"[^a-z0-9]", "", str(col).lower())


def _build_reverse_lookup():
    lookup = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            lookup[normalize_header(alias)] = canonical
    return lookup


_REVERSE_LOOKUP = _build_reverse_lookup()


def map_columns(df):
    """
    Rename a raw dataframe's columns to canonical names wherever a recognized
    alias is found (case/punctuation-insensitive). Unrecognized columns are
    left untouched. If a file accidentally contains two columns that would
    map to the same canonical name, only the first is renamed (to avoid
    creating duplicate columns) and a warning is printed.
    """
    rename_map = {}
    seen_canonical = set()
    for col in df.columns:
        key = normalize_header(col)
        canonical = _REVERSE_LOOKUP.get(key)
        if canonical is None:
            continue
        if canonical in seen_canonical:
            print(f"    [WARN] Column '{col}' also maps to '{canonical}' — "
                  f"a column for that field was already found, ignoring this one.")
            continue
        rename_map[col] = canonical
        seen_canonical.add(canonical)
    return df.rename(columns=rename_map)


def read_one_file(path):
    """
    Read a single Excel/CSV file into a dataframe (all columns as text, so
    nothing like a mobile number's leading zero gets lost). Returns None
    (and prints a message) instead of raising, so one bad file never takes
    down the whole batch.
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xls", ".xlsm"):
            df = pd.read_excel(path, sheet_name=0, dtype=str)
        elif ext == ".csv":
            try:
                df = pd.read_csv(path, dtype=str, encoding="utf-8")
            except UnicodeDecodeError:
                df = pd.read_csv(path, dtype=str, encoding="latin1")
        else:
            print(f"  [SKIP] Unsupported file type: {os.path.basename(path)}")
            return None
    except Exception as exc:
        print(f"  [ERROR] Could not read {os.path.basename(path)}: {exc}")
        return None

    if df is None or df.empty:
        print(f"  [WARN] {os.path.basename(path)} has no rows — skipped.")
        return None

    df = map_columns(df)
    df["SourceFile"] = os.path.basename(path)
    return df


def load_all_files(input_folder):
    """Find and read every supported file in input_folder, returning one combined dataframe."""
    if not os.path.isdir(input_folder):
        print(f"[ERROR] Input folder not found: {input_folder}")
        return pd.DataFrame()

    patterns = ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]
    files = sorted(set(
        f for pattern in patterns for f in glob.glob(os.path.join(input_folder, pattern))
    ))

    if not files:
        print(f"[WARN] No Excel/CSV files found in: {input_folder}")
        return pd.DataFrame()

    print(f"Found {len(files)} file(s) in '{input_folder}':")
    frames = []
    for path in files:
        print(f"  - Reading {os.path.basename(path)} ...")
        df = read_one_file(path)
        if df is not None and not df.empty:
            print(f"    -> {len(df)} row(s) loaded")
            frames.append(df)

    if not frames:
        print("[WARN] No usable data found in any file.")
        return pd.DataFrame()

    # Merge ALL files together first; filtering happens afterwards on the combined set
    combined = pd.concat(frames, ignore_index=True, sort=False)
    return combined


def clean_numeric_series(series):
    """Vectorized cleanup of a text column into numbers (strips commas/spaces, coerces bad values to NaN)."""
    cleaned = series.fillna("").astype(str).str.strip()
    cleaned = cleaned.str.replace(",", "", regex=False)
    cleaned = cleaned.str.replace(r"\s+", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce")


def clean_mobile_series(series, cfg):
    """
    Vectorized cleanup + validation of the MobileNo column.

    Returns:
        cleaned (pd.Series[str]): normalized digit-only numbers
        is_valid (pd.Series[bool]): True where the row has a usable mobile number
    """
    s = series.fillna("").astype(str).str.strip()

    # Drop a stray trailing ".0" left behind when a number was read as a float
    s = s.str.replace(r"\.0+$", "", regex=True)

    # Keep digits and a leading "+" only
    s = s.str.replace(r"[^\d+]", "", regex=True)

    # Strip the "+" now that punctuation has served its purpose
    digits = s.str.replace(r"^\+", "", regex=True)

    if cfg.get("NORMALIZE_KENYAN_MOBILE", False):
        # 254 7XXXXXXXX / 254 1XXXXXXXX  ->  07XXXXXXXX / 01XXXXXXXX
        mask_254 = digits.str.match(r"^254[17]\d{8}$")
        digits = digits.where(~mask_254, "0" + digits.str[3:])

        # 7XXXXXXXX / 1XXXXXXXX (leading 0 dropped) -> 07XXXXXXXX / 01XXXXXXXX
        mask_missing_zero = digits.str.match(r"^[17]\d{8}$")
        digits = digits.where(~mask_missing_zero, "0" + digits)

    digit_count = digits.str.len()
    is_numeric = digits.str.match(r"^\d+$") & (digits != "")
    is_valid_length = digit_count.between(cfg["MOBILE_MIN_DIGITS"], cfg["MOBILE_MAX_DIGITS"])
    is_valid = is_numeric & is_valid_length

    return digits, is_valid


def process(cfg):
    print("=" * 70)
    print("ARREARS DATA PROCESSING")
    print("=" * 70)

    combined = load_all_files(cfg["INPUT_FOLDER"])
    total_loaded = len(combined)
    print(f"\nTotal rows combined from all files: {total_loaded}")

    if combined.empty:
        print("\nNothing to process — exiting.")
        return

    # Make sure every expected canonical column exists, even if a file lacked it
    for col in OUTPUT_COLUMN_ORDER:
        if col not in combined.columns:
            combined[col] = pd.NA

    # ---- Clean "Days" (arrears aging) ----
    combined["Days"] = clean_numeric_series(combined["Days"])

    # ---- Clean other numeric columns, where present ----
    for col in EXTRA_NUMERIC_COLUMNS:
        if col in combined.columns:
            combined[col] = clean_numeric_series(combined[col])

    # ---- Parse date-like columns, where present ----
    for col in DATE_COLUMNS:
        if col in combined.columns:
            combined[col] = pd.to_datetime(combined[col], errors="coerce")

    # ---- Clean & validate MobileNo ----
    cleaned_mobile, valid_mobile_mask = clean_mobile_series(combined["MobileNo"], cfg)
    combined["MobileNo"] = cleaned_mobile
    rejected_mobile_count = int((~valid_mobile_mask).sum())

    valid_df = combined[valid_mobile_mask].copy()

    # ---- Apply the arrears aging filter (after merging all files) ----
    min_days = cfg["MIN_DAYS_ARREARS"]
    max_days = cfg["MAX_DAYS_ARREARS"]
    days_in_range_mask = valid_df["Days"].notna() & valid_df["Days"].between(min_days, max_days)
    rejected_aging_count = int((~days_in_range_mask).sum())

    final_df = valid_df[days_in_range_mask].copy()

    # ---- Final column selection / order ----
    final_columns = [c for c in OUTPUT_COLUMN_ORDER if c in final_df.columns]
    final_df = final_df[final_columns].reset_index(drop=True)

    # ---- Save output ----
    os.makedirs(cfg["OUTPUT_FOLDER"], exist_ok=True)
    output_path = os.path.join(cfg["OUTPUT_FOLDER"], cfg["OUTPUT_FILE_NAME"])
    ext = os.path.splitext(output_path)[1].lower()

    if ext == ".csv":
        final_df.to_csv(output_path, index=False)
    else:
        final_df.to_excel(output_path, index=False)

    # ---- Summary ----
    print("\n" + "-" * 70)
    print("PROCESSING SUMMARY")
    print("-" * 70)
    print(f"Total rows combined from all input files            : {total_loaded}")
    print(f"Rejected — missing/invalid MobileNo                  : {rejected_mobile_count}")
    print(f"Rejected — outside {min_days}-{max_days} day arrears range          : {rejected_aging_count}")
    print(f"Valid SMS-ready records written to output            : {len(final_df)}")
    print(f"\nOutput file: {output_path}")
    print("=" * 70)


if __name__ == "__main__":
    try:
        process(CONFIG)
    except Exception:
        print("\n[FATAL] Processing failed with an unexpected error:")
        traceback.print_exc()
