# ╔════════════════════════════════════════════════════════════════════╗
# ║ HRIT UNIVERSITY - CORRECTED CIE + ESE AWARD SHEET GENERATOR       ║
# ║                                                                    ║
# ║ Fixes:                                                             ║
# ║   ✅ ESE course code from subjectcode / Column A                    ║
# ║   ✅ ESE subject name from subjectname / Column B                   ║
# ║   ✅ ESE marks from scaled_marks / Column U                         ║
# ║   ✅ Uni Roll No mapped using old DumP sheet logic                  ║
# ║   ✅ CIE MM = 40 always                                             ║
# ║   ✅ ESE MM = 60 always                                             ║
# ║   ✅ Semester from course code: 101 -> First, 202 -> Second         ║
# ║   ✅ Retry/backoff for Google Sheets quota 429 errors               ║
# ║   ✅ One tab per course code                                        ║
# ╚════════════════════════════════════════════════════════════════════╝


# ════════════════════════════════════════
# 1. INSTALL + AUTH
# ════════════════════════════════════════

!pip install gspread google-auth pandas --quiet

from google.colab import auth
auth.authenticate_user()

import gspread
from google.auth import default
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import pandas as pd
import numpy as np
import re
import time
import random
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

creds, _ = default()
gc = gspread.authorize(creds)
service = build("sheets", "v4", credentials=creds)

print("✅ Authenticated with Google Sheets")


# ════════════════════════════════════════
# 2. CONFIG
# ════════════════════════════════════════

SOURCE_SHEET_ID = "1e_GpPbBhkPvEl8DvOlYDUy8VmRtsk7EmT17adnGFwgY"

# This is from your old code — used for Admission/Enrollment ID -> Uni Roll No mapping
ATTEND_SHEET_ID = "1xIiMiQK0-0mStSfJbvuTS1FTTEjNyQahbHAZYjmyulk"
DUMP_TAB = "DumP"

# 0-based columns in DumP tab
DUMP_ADM_COL = 0     # Column A = Admission / Enrollment No
DUMP_ROLL_COL = 2    # Column C = Uni Roll No

CIE_DUMP_TAB_CANDIDATES = [
    "Dump CIE",
    "CIE Dump",
    "CIE dump",
    "cie dump",
    "CIE"
]

ESE_DUMP_TAB_CANDIDATES = [
    "Dump ESE",
    "ESE Dump",
    "ESE dump",
    "ese dump",
    "ESE"
]

COURSE_CODE_DUMP_TAB_CANDIDATES = [
    "Dump course code",
    "Dump Course Code",
    "Course Code Dump",
    "Course code dump",
    "course code dump",
    "Course Codes",
    "Course codes"
]

TEMPLATE_CIE_TAB_CANDIDATES = [
    "Template CIE",
    "template cie",
    "CIE Template"
]

TEMPLATE_ESE_TAB_CANDIDATES = [
    "Template ESE",
    "template ese",
    "ESE Template"
]

RUN_TIMESTAMP = datetime.now().strftime("%Y-%m-%d %H-%M-%S")

CIE_OUTPUT_TITLE = f"HRIT University - CIE Award Sheets - Corrected - Generated {RUN_TIMESTAMP}"
ESE_OUTPUT_TITLE = f"HRIT University - ESE Award Sheets - Corrected - Generated {RUN_TIMESTAMP}"

FACULTY_NAME_DEFAULT = "Faculty of Life Sciences"

CIE_MAX_MARKS = 40
ESE_MAX_MARKS = 60

CIE_TITLE_TEXT = "Award Sheet For CIE (Semester - Jan'26 - June'26 Session - 2025-2026. )"
ESE_TITLE_TEXT = "Award Sheet For ESE (Semester - Jan'26 - June'26 Session - 2025-2026. )"

LOGO_DRIVE_FILE_URL = "https://drive.google.com/file/d/1dLQpHFuzNH13sQTNSwy-_BK2GCHmjmve/view?usp=sharing"

DATA_START_ROW = 11
MIN_DISPLAY_ROWS = 13

# To avoid Google quota errors
SLEEP_BETWEEN_TABS_SECONDS = 2.5
LONG_SLEEP_AFTER_EVERY_N_TABS = 20
LONG_SLEEP_SECONDS = 35

print("✅ Configuration ready")


# ════════════════════════════════════════
# 3. SAFE GOOGLE API EXECUTION WITH RETRY
# ════════════════════════════════════════

def execute_with_retry(request, label="", max_retries=8):
    """
    Executes Google API request with retry for quota/rate-limit errors.
    """
    for attempt in range(max_retries):
        try:
            return request.execute()

        except HttpError as e:
            status = getattr(e.resp, "status", None)

            if status in [429, 500, 503]:
                wait = min(90, (2 ** attempt) + random.uniform(1, 4))
                print(f"⚠️ {label} hit {status}. Retry {attempt + 1}/{max_retries} after {wait:.1f}s...")
                time.sleep(wait)
                continue

            raise

    raise Exception(f"❌ Failed after retries: {label}")


def batch_update(spreadsheet_id, requests, label="batchUpdate"):
    if not requests:
        return

    req = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    )

    return execute_with_retry(req, label=label)


def values_update(spreadsheet_id, range_name, values, label="values.update"):
    req = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    )

    return execute_with_retry(req, label=label)


def values_batch_update(spreadsheet_id, data, label="values.batchUpdate"):
    req = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": data
        }
    )

    return execute_with_retry(req, label=label)


def values_clear(spreadsheet_id, range_name, label="values.clear"):
    req = service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        body={}
    )

    return execute_with_retry(req, label=label)


def copy_to(source_spreadsheet_id, source_sheet_id, destination_spreadsheet_id, label="copyTo"):
    req = service.spreadsheets().sheets().copyTo(
        spreadsheetId=source_spreadsheet_id,
        sheetId=source_sheet_id,
        body={"destinationSpreadsheetId": destination_spreadsheet_id}
    )

    return execute_with_retry(req, label=label)


print("✅ Retry helpers ready")


# ════════════════════════════════════════
# 4. BASIC HELPERS
# ════════════════════════════════════════

def clean(val):
    if pd.isna(val):
        return ""

    s = str(val).strip()

    if s.endswith(".0"):
        s = s[:-2]

    return s


def clean_number(val):
    """
    Cleans and rounds marks using normal academic rounding:
      32.4 -> 32
      32.5 -> 33
      32.6 -> 33

    Also preserves absent-like values.
    """
    if pd.isna(val):
        return ""

    s = str(val).strip()

    if not s:
        return ""

    lower = s.lower()

    if lower in ["absent", "ab", "a"]:
        return "Absent"

    try:
        rounded = Decimal(s).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        return str(int(rounded))

    except:
        return s


def normalize_header(h):
    """
    Very loose normalization:
      subject_code -> subjectcode
      subjectcode  -> subjectcode
      Student_Name -> studentname
    """
    return re.sub(r"[^a-z0-9]+", "", str(h).strip().lower())


def safe_sheet_title(title):
    title = clean(title)
    title = re.sub(r"[\[\]\*\/\\\?\:]", "-", title)
    title = title[:99]
    return title if title else "Sheet"


def quote_sheet_name(sheet_name):
    return "'" + sheet_name.replace("'", "''") + "'"


def extract_drive_file_id(url):
    if not url:
        return ""

    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)

    m = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)

    return ""


LOGO_FILE_ID = extract_drive_file_id(LOGO_DRIVE_FILE_URL)
LOGO_IMAGE_URL = f"https://drive.google.com/uc?export=view&id={LOGO_FILE_ID}" if LOGO_FILE_ID else ""

print("✅ Basic helpers ready")
print("✅ Logo file id:", LOGO_FILE_ID)


# ════════════════════════════════════════
# 5. SEMESTER LOGIC
# ════════════════════════════════════════

def extract_course_number(course_code):
    s = clean(course_code).upper()
    m = re.search(r"(\d{3})", s)
    return m.group(1) if m else ""


def semester_word_from_course_code(course_code):
    num = extract_course_number(course_code)

    if not num:
        return ""

    sem_digit = num[0]

    sem_map = {
        "1": "First",
        "2": "Second",
        "3": "Third",
        "4": "Fourth",
        "5": "Fifth",
        "6": "Sixth",
        "7": "Seventh",
        "8": "Eighth",
    }

    return sem_map.get(sem_digit, "")


def semester_number_from_course_code(course_code):
    num = extract_course_number(course_code)

    if not num:
        return 0

    try:
        return int(num[0])
    except:
        return 0


print("✅ Semester logic ready")
print("Example HRBOT 101 ->", semester_word_from_course_code("HRBOT 101"))
print("Example HRRMI 202 ->", semester_word_from_course_code("HRRMI 202"))


# ════════════════════════════════════════
# 6. MARKS TO WORDS
# ════════════════════════════════════════

ONES = [
    "Zero", "One", "Two", "Three", "Four", "Five",
    "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
    "Twelve", "Thirteen", "Fourteen", "Fifteen",
    "Sixteen", "Seventeen", "Eighteen", "Nineteen"
]

TENS = [
    "", "", "Twenty", "Thirty", "Forty", "Fifty",
    "Sixty", "Seventy", "Eighty", "Ninety"
]


def int_to_words(n):
    n = int(n)

    if 0 <= n < 20:
        return ONES[n]

    if 20 <= n < 100:
        tens = n // 10
        ones = n % 10
        return TENS[tens] if ones == 0 else f"{TENS[tens]} {ONES[ones]}"

    if n == 100:
        return "One Hundred"

    if 100 < n < 1000:
        hundreds = n // 100
        rem = n % 100

        if rem == 0:
            return f"{ONES[hundreds]} Hundred"

        return f"{ONES[hundreds]} Hundred {int_to_words(rem)}"

    return str(n)


def marks_to_words(mark):
    s = clean_number(mark)

    if s == "":
        return ""

    lower = s.lower()

    if lower in ["absent", "ab", "a"]:
        return "Absent"

    try:
        f = float(s)

        if f.is_integer():
            return int_to_words(int(f))

        whole = int(f)
        decimal_part = str(s).split(".")[1]

        decimal_words = " ".join(
            ONES[int(d)] for d in decimal_part if d.isdigit()
        )

        return f"{int_to_words(whole)} Point {decimal_words}"

    except:
        return s


print("✅ Marks-to-words helper ready")


# ════════════════════════════════════════
# 7. GOOGLE SHEET READ HELPERS
# ════════════════════════════════════════

def list_tabs(spreadsheet_id):
    wb = gc.open_by_key(spreadsheet_id)
    return [ws.title for ws in wb.worksheets()]


def find_tab_name(spreadsheet_id, candidates, keyword=None, required=True):
    tabs = list_tabs(spreadsheet_id)
    lower_map = {t.lower().strip(): t for t in tabs}

    for c in candidates:
        key = c.lower().strip()

        if key in lower_map:
            return lower_map[key]

    if keyword:
        keyword = keyword.lower()

        for t in tabs:
            if keyword in t.lower():
                return t

    if required:
        raise Exception(
            f"❌ Could not find tab. Tried candidates={candidates}, keyword={keyword}. "
            f"Available tabs={tabs}"
        )

    return None


def sheet_to_df(spreadsheet_id, tab_name, header_row=0):
    wb = gc.open_by_key(spreadsheet_id)
    ws = wb.worksheet(tab_name)
    raw = ws.get_all_values()

    if not raw:
        return pd.DataFrame()

    headers = [str(h).strip() for h in raw[header_row]]
    data = raw[header_row + 1:]

    n_cols = len(headers)

    fixed = []

    for r in data:
        if len(r) < n_cols:
            r = r + [""] * (n_cols - len(r))
        else:
            r = r[:n_cols]

        fixed.append(r)

    df = pd.DataFrame(fixed, columns=headers)
    df = df.replace("", pd.NA)

    print(f"✅ Loaded '{tab_name}' -> {len(df)} rows, {len(df.columns)} columns")
    print("   Columns:", list(df.columns))

    return df


def get_col_by_name_or_index(df, possible_names=None, fallback_index=None, required=False, label=""):
    possible_names = possible_names or []

    norm_to_actual = {
        normalize_header(c): c for c in df.columns
    }

    # Exact loose-normalized match
    for name in possible_names:
        key = normalize_header(name)

        if key in norm_to_actual:
            return norm_to_actual[key]

    # Contains match both ways
    for actual in df.columns:
        norm = normalize_header(actual)

        for name in possible_names:
            key = normalize_header(name)

            if key and (key in norm or norm in key):
                return actual

    # Fallback index
    if fallback_index is not None and fallback_index < len(df.columns):
        return df.columns[fallback_index]

    if required:
        raise Exception(
            f"❌ Required column not found for {label}. "
            f"Tried names={possible_names}, fallback_index={fallback_index}. "
            f"Available columns={list(df.columns)}"
        )

    return None


print("✅ Sheet helpers ready")


# ════════════════════════════════════════
# 8. LOAD SOURCE TABS
# ════════════════════════════════════════

print("\n📋 Available tabs in source workbook:")
for t in list_tabs(SOURCE_SHEET_ID):
    print("   -", t)

cie_dump_tab = find_tab_name(
    SOURCE_SHEET_ID,
    CIE_DUMP_TAB_CANDIDATES,
    keyword="cie",
    required=True
)

ese_dump_tab = find_tab_name(
    SOURCE_SHEET_ID,
    ESE_DUMP_TAB_CANDIDATES,
    keyword="ese",
    required=True
)

course_code_dump_tab = find_tab_name(
    SOURCE_SHEET_ID,
    COURSE_CODE_DUMP_TAB_CANDIDATES,
    keyword="course",
    required=False
)

template_cie_tab = find_tab_name(
    SOURCE_SHEET_ID,
    TEMPLATE_CIE_TAB_CANDIDATES,
    keyword="template cie",
    required=True
)

template_ese_tab = find_tab_name(
    SOURCE_SHEET_ID,
    TEMPLATE_ESE_TAB_CANDIDATES,
    keyword="template ese",
    required=True
)

print("\n✅ Detected tabs:")
print("   CIE Dump         :", cie_dump_tab)
print("   ESE Dump         :", ese_dump_tab)
print("   Course Code Dump :", course_code_dump_tab)
print("   Template CIE     :", template_cie_tab)
print("   Template ESE     :", template_ese_tab)

df_cie_raw = sheet_to_df(SOURCE_SHEET_ID, cie_dump_tab)
df_ese_raw = sheet_to_df(SOURCE_SHEET_ID, ese_dump_tab)

if course_code_dump_tab:
    df_course_code_raw = sheet_to_df(SOURCE_SHEET_ID, course_code_dump_tab)
else:
    df_course_code_raw = pd.DataFrame()


# ════════════════════════════════════════
# 9. LOAD UNI ROLL MAPPING FROM OLD DUMP SHEET
# ════════════════════════════════════════

print("\n📊 Loading Admission/Enrollment ID -> Uni Roll No mapping...")

attend_wb = gc.open_by_key(ATTEND_SHEET_ID)
dump_ws = attend_wb.worksheet(DUMP_TAB)
dump_values = dump_ws.get_all_values()

adm_to_roll = {}

for row in dump_values[1:]:
    if len(row) <= max(DUMP_ADM_COL, DUMP_ROLL_COL):
        continue

    adm = clean(row[DUMP_ADM_COL])
    roll = clean(row[DUMP_ROLL_COL])

    if adm and roll:
        adm_to_roll[adm] = roll

print(f"✅ Admission/Enrollment -> Uni Roll mappings loaded: {len(adm_to_roll)}")


def map_to_uni_roll(enrollment_id):
    """
    Converts enrollment/admission ID to Uni Roll No.
    If not found, keeps blank fallback marker.
    """
    eid = clean(enrollment_id)

    if not eid:
        return ""

    return adm_to_roll.get(eid, "")


# ════════════════════════════════════════
# 10. COURSE/SUBJECT MAPS
# ════════════════════════════════════════

def infer_course_name_from_program_or_code(program, code):
    p = clean(program).upper()
    c = clean(code).upper()

    if "HRBOT" in c or "OTT" in p or "OPERATION" in p:
        return "Bachelor in Anaesthesia and Operation Theatre Technology"

    if "HRRMI" in c or "RMIT" in p or "MRIT" in p or "RADIO" in p:
        return "Bachelor in Medical Radiology and Imaging Technology"

    if "HRBMLT" in c or "MLT" in p:
        return "Bachelor in Medical Laboratory Technology"

    if "HRBOP" in c or "OPT" in p or "OPTOMETRY" in p:
        return "Bachelor in Optometry"

    return clean(program)


# Use CIE as primary source for course code -> subject name and course name.
# Because user said ESE can use same course-code mapping as CIE.
def build_maps_from_cie(df):
    col_program = get_col_by_name_or_index(
        df,
        possible_names=["Adm_program", "Program", "Course Name", "course_name"],
        fallback_index=0,
        required=False,
        label="CIE Program"
    )

    col_sub_code = get_col_by_name_or_index(
        df,
        possible_names=["subject_code", "subjectcode", "Subject Code", "Course Code", "course_code"],
        fallback_index=5,
        required=True,
        label="CIE Subject Code"
    )

    col_sub_name = get_col_by_name_or_index(
        df,
        possible_names=["subject_name", "subjectname", "Subject Name"],
        fallback_index=6,
        required=True,
        label="CIE Subject Name"
    )

    subject_name_map = {}
    course_name_map = {}

    for _, row in df.iterrows():
        code = clean(row.get(col_sub_code, ""))

        if not code:
            continue

        sub_name = clean(row.get(col_sub_name, ""))
        program = clean(row.get(col_program, "")) if col_program else ""

        if sub_name and code not in subject_name_map:
            subject_name_map[code] = sub_name

        course_name = infer_course_name_from_program_or_code(program, code)

        if course_name and code not in course_name_map:
            course_name_map[code] = course_name

    return subject_name_map, course_name_map


cie_subject_name_map, cie_course_name_map = build_maps_from_cie(df_cie_raw)

print("\n✅ Built maps from CIE:")
print("   Subject names:", len(cie_subject_name_map))
print("   Course names :", len(cie_course_name_map))


# Also try course-code dump if useful.
# Your course code dump appeared as 2 columns both named HRRMI 101, likely no header.
# This handles both header and no-header style.
def build_maps_from_course_code_dump(df):
    """
    Builds subject map from Dump course code.

    Your dump currently has duplicate headers:
      ['HRRMI 101', 'HRRMI 101']

    So we avoid row.get(column_name) and use iloc instead.
    Expected no-header style:
      Column A = subject/course code
      Column B = subject name OR related mapped value
    """

    subject_map = {}
    course_map = {}

    if df is None or df.empty:
        return subject_map, course_map

    # Case 1: Try normal header detection first
    code_col = get_col_by_name_or_index(
        df,
        possible_names=[
            "subjectcode",
            "subject_code",
            "Subject Code",
            "coursecode",
            "course_code",
            "Course Code",
            "code"
        ],
        fallback_index=None,
        required=False,
        label="Course Dump Code"
    )

    name_col = get_col_by_name_or_index(
        df,
        possible_names=[
            "subjectname",
            "subject_name",
            "Subject Name",
            "name"
        ],
        fallback_index=None,
        required=False,
        label="Course Dump Subject Name"
    )

    # Only use normal mode if detected column names are not duplicated
    if code_col and name_col and code_col != name_col:
        try:
            for _, row in df.iterrows():
                code = clean(row[code_col])
                name = clean(row[name_col])

                if code and name:
                    subject_map[code] = name
        except Exception as e:
            print("⚠️ Normal course dump parsing failed, switching to iloc mode:", e)

    # Case 2: Duplicate headers / no real header mode
    # Use first two physical columns by position.
    if len(subject_map) == 0 and df.shape[1] >= 2:
        print("ℹ️ Course code dump seems to have duplicate/no headers. Using first 2 columns by position.")

        # Include column headers themselves as the first possible mapping row
        first_code = clean(df.columns[0])
        first_name = clean(df.columns[1])

        if first_code and first_name:
            subject_map[first_code] = first_name

        for i in range(len(df)):
            code = clean(df.iloc[i, 0])
            name = clean(df.iloc[i, 1])

            if code and name:
                subject_map[code] = name

    return subject_map, course_map


dump_subject_name_map, dump_course_name_map = build_maps_from_course_code_dump(df_course_code_raw)

print("✅ Built maps from course-code dump:")
print("   Subject names:", len(dump_subject_name_map))

# Final maps: CIE first, dump fills gaps
subject_name_map = dict(cie_subject_name_map)
subject_name_map.update({k: v for k, v in dump_subject_name_map.items() if k not in subject_name_map})

course_name_map = dict(cie_course_name_map)
course_name_map.update({k: v for k, v in dump_course_name_map.items() if k not in course_name_map})


# ════════════════════════════════════════
# 11. STANDARDIZE CIE
# ════════════════════════════════════════

def standardize_cie(df):
    col_program = get_col_by_name_or_index(
        df,
        possible_names=["Adm_program", "Program", "Course Name", "course_name"],
        fallback_index=0,
        required=False,
        label="CIE Program"
    )

    col_enrollment = get_col_by_name_or_index(
        df,
        possible_names=["enrollment_id", "enrollmentid", "Admission No", "admission_no", "student_id"],
        fallback_index=1,
        required=True,
        label="CIE Enrollment/Admission ID"
    )

    col_name = get_col_by_name_or_index(
        df,
        possible_names=["student_name", "studentname", "Student Name", "Candidate Name", "Name"],
        fallback_index=2,
        required=True,
        label="CIE Student Name"
    )

    col_sub_code = get_col_by_name_or_index(
        df,
        possible_names=["subject_code", "subjectcode", "Subject Code", "Course Code", "course_code"],
        fallback_index=5,
        required=True,
        label="CIE Subject Code"
    )

    col_sub_name = get_col_by_name_or_index(
        df,
        possible_names=["subject_name", "subjectname", "Subject Name"],
        fallback_index=6,
        required=True,
        label="CIE Subject Name"
    )

    col_marks = get_col_by_name_or_index(
        df,
        possible_names=["Final", "final", "Marks", "marks", "CIE Marks"],
        fallback_index=15,
        required=True,
        label="CIE Final Marks"
    )

    out = pd.DataFrame()

    out["course_code"] = df[col_sub_code].apply(clean)
    out["enrollment_id"] = df[col_enrollment].apply(clean)
    out["roll_no"] = out["enrollment_id"].apply(map_to_uni_roll)
    out["student_name"] = df[col_name].apply(clean)
    out["subject_name"] = df[col_sub_name].apply(clean)
    out["marks"] = df[col_marks].apply(clean_number)
    out["exam_type"] = "CIE"
    out["max_marks"] = CIE_MAX_MARKS

    if col_program:
        out["program_raw"] = df[col_program].apply(clean)
    else:
        out["program_raw"] = ""

    out["subject_name"] = out.apply(
        lambda r: subject_name_map.get(r["course_code"], r["subject_name"]),
        axis=1
    )

    out["course_name"] = out.apply(
        lambda r: course_name_map.get(
            r["course_code"],
            infer_course_name_from_program_or_code(r["program_raw"], r["course_code"])
        ),
        axis=1
    )

    out["faculty_name"] = FACULTY_NAME_DEFAULT
    out["semester_word"] = out["course_code"].apply(semester_word_from_course_code)

    out = out[out["course_code"].astype(str).str.strip() != ""].copy()
    out = out[out["enrollment_id"].astype(str).str.strip() != ""].copy()

    return out


df_cie = standardize_cie(df_cie_raw)

print("\n✅ CIE standardized")
print(df_cie.head())
print("CIE rows:", len(df_cie))
print("CIE unique course codes:", df_cie["course_code"].nunique())
print("CIE mapped uni rolls:", (df_cie["roll_no"] != "").sum(), "/", len(df_cie))


# ════════════════════════════════════════
# 12. STANDARDIZE ESE — CORRECTED
# ════════════════════════════════════════

def standardize_ese(df):
    """
    Correct ESE columns:
      A subjectcode
      B subjectname
      C semester
      I enrollment_id
      K Student_Name
      O program_name
      U scaled_marks
    """

    col_sub_code = get_col_by_name_or_index(
        df,
        possible_names=["subjectcode", "subject_code", "Subject Code", "coursecode", "course_code"],
        fallback_index=0,
        required=True,
        label="ESE Subject Code"
    )

    col_sub_name = get_col_by_name_or_index(
        df,
        possible_names=["subjectname", "subject_name", "Subject Name"],
        fallback_index=1,
        required=True,
        label="ESE Subject Name"
    )

    col_enrollment = get_col_by_name_or_index(
        df,
        possible_names=["enrollment_id", "enrollmentid", "Admission No", "admission_no"],
        fallback_index=8,
        required=True,
        label="ESE Enrollment/Admission ID"
    )

    col_name = get_col_by_name_or_index(
        df,
        possible_names=["Student_Name", "student_name", "studentname", "Student Name", "Candidate Name"],
        fallback_index=10,
        required=True,
        label="ESE Student Name"
    )

    col_program = get_col_by_name_or_index(
        df,
        possible_names=["program_name", "programname", "Program Name", "Course Name", "course_name"],
        fallback_index=14,
        required=False,
        label="ESE Program/Course Name"
    )

    col_marks = get_col_by_name_or_index(
        df,
        possible_names=["scaled_marks", "scaledmarks", "Final", "final", "Marks", "marks", "ESE Marks"],
        fallback_index=20,
        required=True,
        label="ESE Scaled Marks"
    )
    col_is_absent = get_col_by_name_or_index(
        df,
        possible_names=["is_absent", "isabsent", "Absent", "absent"],
        fallback_index=18,   # Column S
        required=True,
        label="ESE is_absent"
    )

    # Keep only rows where is_absent is false
    before_rows = len(df)

    df = df.copy()
    df["_is_absent_norm"] = df[col_is_absent].apply(lambda x: clean(x).lower())

    df = df[df["_is_absent_norm"].isin(["false", "0", "no", "n"])].copy()

    print(f"✅ ESE is_absent filter applied: kept {len(df)} / {before_rows} rows where is_absent = false")

    out = pd.DataFrame()

    out["course_code"] = df[col_sub_code].apply(clean)
    out["enrollment_id"] = df[col_enrollment].apply(clean)
    out["roll_no"] = out["enrollment_id"].apply(map_to_uni_roll)
    out["student_name"] = df[col_name].apply(clean)
    out["subject_name"] = df[col_sub_name].apply(clean)
    out["marks"] = df[col_marks].apply(clean_number)
    out["exam_type"] = "ESE"
    out["max_marks"] = ESE_MAX_MARKS

    if col_program:
        out["program_raw"] = df[col_program].apply(clean)
    else:
        out["program_raw"] = ""

    # Use CIE/dump mappings if same course code exists.
    out["subject_name"] = out.apply(
        lambda r: subject_name_map.get(r["course_code"], r["subject_name"]),
        axis=1
    )

    out["course_name"] = out.apply(
        lambda r: course_name_map.get(
            r["course_code"],
            infer_course_name_from_program_or_code(r["program_raw"], r["course_code"])
        ),
        axis=1
    )

    out["faculty_name"] = FACULTY_NAME_DEFAULT
    out["semester_word"] = out["course_code"].apply(semester_word_from_course_code)

    out = out[out["course_code"].astype(str).str.strip() != ""].copy()
    out = out[out["enrollment_id"].astype(str).str.strip() != ""].copy()

    return out


df_ese = standardize_ese(df_ese_raw)

print("\n✅ ESE standardized — corrected")
print(df_ese.head())
print("ESE rows:", len(df_ese))
print("ESE unique course codes:", df_ese["course_code"].nunique())
print("ESE mapped uni rolls:", (df_ese["roll_no"] != "").sum(), "/", len(df_ese))


# ════════════════════════════════════════
# 13. COURSE DATA
# ════════════════════════════════════════

def sort_roll_value(roll, enrollment_id=""):
    r = clean(roll)

    if not r:
        r = clean(enrollment_id)

    try:
        return (0, int(r))
    except:
        return (1, r)


def build_course_data(df, exam_type):
    course_data = {}

    for code in sorted(df["course_code"].dropna().unique().tolist()):
        df_c = df[df["course_code"] == code].copy()

        df_c["_sort"] = df_c.apply(
            lambda r: sort_roll_value(r.get("roll_no", ""), r.get("enrollment_id", "")),
            axis=1
        )

        df_c = df_c.sort_values("_sort").drop(columns=["_sort"])

        # Avoid duplicate enrollment under same course
        df_c = df_c.drop_duplicates(subset=["enrollment_id", "course_code"], keep="first")

        first = df_c.iloc[0]

        max_marks = CIE_MAX_MARKS if exam_type == "CIE" else ESE_MAX_MARKS

        students = []

        for _, row in df_c.iterrows():
            mark = clean_number(row.get("marks", ""))
            roll = clean(row.get("roll_no", ""))
            enrollment_id = clean(row.get("enrollment_id", ""))

            # If mapping missing, keep visible marker so you can audit.
            # If you want blank instead, replace fallback with "".
            if not roll:
                roll = f"(ADM: {enrollment_id})"

            students.append({
                "roll_no": roll,
                "enrollment_id": enrollment_id,
                "student_name": clean(row.get("student_name", "")),
                "marks": mark,
                "marks_word": marks_to_words(mark),
            })

        course_data[code] = {
            "exam_type": exam_type,
            "course_code": code,
            "course_name": clean(first.get("course_name", "")),
            "subject_name": clean(first.get("subject_name", "")),
            "faculty_name": clean(first.get("faculty_name", FACULTY_NAME_DEFAULT)) or FACULTY_NAME_DEFAULT,
            "semester_word": semester_word_from_course_code(code),
            "max_marks": max_marks,
            "students": students,
        }

    return course_data


cie_course_data = build_course_data(df_cie, "CIE")
ese_course_data = build_course_data(df_ese, "ESE")

print("\n✅ Course-wise data ready")

print("\n📘 CIE Courses:")
for code, info in cie_course_data.items():
    print(f"   {code:15s} | {info['semester_word']:8s} | {len(info['students']):3d} students | {info['subject_name']}")

print("\n📕 ESE Courses:")
for code, info in ese_course_data.items():
    print(f"   {code:15s} | {info['semester_word']:8s} | {len(info['students']):3d} students | {info['subject_name']}")


# ════════════════════════════════════════
# 14. SHEETS API FORMATTING HELPERS
# ════════════════════════════════════════

def get_sheet_id_by_title(spreadsheet_id, title):
    req = service.spreadsheets().get(spreadsheetId=spreadsheet_id)
    meta = execute_with_retry(req, label="spreadsheets.get")

    for s in meta["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]

    return None


def get_template_sheet_id(template_title):
    sid = get_sheet_id_by_title(SOURCE_SHEET_ID, template_title)

    if sid is None:
        raise Exception(f"❌ Template sheet not found: {template_title}")

    return sid


def copy_template_to_workbook(source_template_sheet_id, destination_spreadsheet_id, new_title):
    copied = copy_to(
        source_spreadsheet_id=SOURCE_SHEET_ID,
        source_sheet_id=source_template_sheet_id,
        destination_spreadsheet_id=destination_spreadsheet_id,
        label=f"copyTo {new_title}"
    )

    new_sheet_id = copied["sheetId"]

    batch_update(destination_spreadsheet_id, [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": new_sheet_id,
                    "title": new_title
                },
                "fields": "title"
            }
        }
    ], label=f"rename {new_title}")

    return new_sheet_id


def delete_default_sheet_if_exists(spreadsheet_id):
    wb = gc.open_by_key(spreadsheet_id)

    try:
        ws = wb.worksheet("Sheet1")

        if len(wb.worksheets()) > 1:
            wb.del_worksheet(ws)
            print("🗑️ Deleted default Sheet1")

    except Exception:
        pass


def solid_border():
    return {
        "style": "SOLID",
        "width": 1,
        "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1}
    }


def thin_border():
    return {
        "style": "SOLID",
        "width": 1,
        "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1}
    }


def fmt_range(sheet_id, r1, r2, c1, c2, user_format):
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2
            },
            "cell": {
                "userEnteredFormat": user_format
            },
            "fields": "userEnteredFormat(" + ",".join(user_format.keys()) + ")"
        }
    }


def merge_range(sheet_id, r1, r2, c1, c2):
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2
            },
            "mergeType": "MERGE_ALL"
        }
    }


def unmerge_range(sheet_id, r1, r2, c1, c2):
    return {
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2
            }
        }
    }


def row_height(sheet_id, start_row, end_row, px):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row,
                "endIndex": end_row
            },
            "properties": {
                "pixelSize": px
            },
            "fields": "pixelSize"
        }
    }


def freeze_rows(sheet_id, frozen_rows):
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {
                    "frozenRowCount": frozen_rows
                }
            },
            "fields": "gridProperties.frozenRowCount"
        }
    }


def ensure_grid_size(spreadsheet_id, sheet_id, min_rows=100, min_cols=7):
    batch_update(spreadsheet_id, [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {
                        "rowCount": min_rows,
                        "columnCount": min_cols
                    }
                },
                "fields": "gridProperties(rowCount,columnCount)"
            }
        }
    ], label="ensure grid size")


print("✅ API formatting helpers ready")


# ════════════════════════════════════════
# 15. FILL TEMPLATE TAB
# ════════════════════════════════════════

def fill_award_sheet(spreadsheet_id, sheet_title, sheet_id, info):
    exam_type = info["exam_type"]
    max_marks = info["max_marks"]
    students = info["students"]
    n = len(students)
    display_rows = max(MIN_DISPLAY_ROWS, n)

    q = quote_sheet_name(sheet_title)

    required_rows = DATA_START_ROW + display_rows + 5
    ensure_grid_size(spreadsheet_id, sheet_id, min_rows=max(100, required_rows), min_cols=7)

    # Clear old data area
    values_clear(
        spreadsheet_id,
        f"{q}!A{DATA_START_ROW}:G{required_rows}",
        label=f"clear data {sheet_title}"
    )

    title_text = CIE_TITLE_TEXT if exam_type == "CIE" else ESE_TITLE_TEXT
    marks_header = f"{exam_type} MARKS (MM: {max_marks})"

    header_data = [
        {
            "range": f"{q}!A4",
            "values": [[title_text]]
        },
        {
            "range": f"{q}!B5",
            "values": [[info["course_name"]]]
        },
        {
            "range": f"{q}!G5",
            "values": [[info["faculty_name"]]]
        },
        {
            "range": f"{q}!B6",
            "values": [[info["subject_name"]]]
        },
        {
            "range": f"{q}!E6",
            "values": [[info["course_code"]]]
        },
        {
            "range": f"{q}!G6",
            "values": [[info["semester_word"]]]
        },
        {
            "range": f"{q}!D9",
            "values": [[marks_header]]
        },
        {
            "range": f"{q}!D10",
            "values": [["Marks in Figure"]]
        },
        {
            "range": f"{q}!E10",
            "values": [["Marks in Word"]]
        },
    ]

    values_batch_update(
        spreadsheet_id,
        header_data,
        label=f"header update {sheet_title}"
    )

    data_rows = []

    for i in range(display_rows):
        if i < n:
            st = students[i]

            data_rows.append([
                i + 1,
                st["roll_no"],
                st["student_name"],
                st["marks"],
                st["marks_word"]
            ])

        else:
            data_rows.append([
                i + 1,
                "",
                "",
                "",
                ""
            ])

    values_update(
        spreadsheet_id,
        f"{q}!A{DATA_START_ROW}:E{DATA_START_ROW + display_rows - 1}",
        data_rows,
        label=f"student data {sheet_title}"
    )

    start_idx = DATA_START_ROW - 1
    end_idx = start_idx + display_rows

    requests = []

    # Unmerge and remerge marks word area E:G row-wise
    requests.append(unmerge_range(sheet_id, start_idx, end_idx, 4, 7))

    for r in range(start_idx, end_idx):
        requests.append(merge_range(sheet_id, r, r + 1, 4, 7))

    requests.append(
        fmt_range(sheet_id, start_idx, end_idx, 0, 7, {
            "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP",
            "textFormat": {
                "fontFamily": "Times New Roman",
                "fontSize": 11
            },
            "borders": {
                "top": thin_border(),
                "bottom": thin_border(),
                "left": thin_border(),
                "right": thin_border()
            }
        })
    )

    # Center S.No., Roll No., Marks figure
    requests.append(
        fmt_range(sheet_id, start_idx, end_idx, 0, 2, {
            "horizontalAlignment": "CENTER"
        })
    )

    requests.append(
        fmt_range(sheet_id, start_idx, end_idx, 3, 4, {
            "horizontalAlignment": "CENTER"
        })
    )

    # Student name left
    requests.append(
        fmt_range(sheet_id, start_idx, end_idx, 2, 3, {
            "horizontalAlignment": "LEFT"
        })
    )

    # Marks word left
    requests.append(
        fmt_range(sheet_id, start_idx, end_idx, 4, 7, {
            "horizontalAlignment": "LEFT"
        })
    )

    requests.append(row_height(sheet_id, start_idx, end_idx, 26))
    requests.append(freeze_rows(sheet_id, 10))

    batch_update(
        spreadsheet_id,
        requests,
        label=f"format data {sheet_title}"
    )

    print(f"   ✅ {exam_type} tab filled: {sheet_title:15s} | {n:3d} students")


# ════════════════════════════════════════
# 16. MASTER SUMMARY
# ════════════════════════════════════════

def course_sort_key(item):
    code, info = item

    return (
        semester_number_from_course_code(code),
        code
    )


def create_master_summary(spreadsheet_id, course_data, exam_type):
    wb = gc.open_by_key(spreadsheet_id)

    try:
        ws = wb.add_worksheet(
            title="Master Summary",
            rows=max(100, len(course_data) + 10),
            cols=9
        )
    except Exception:
        ws = wb.worksheet("Master Summary")

    rows = [
        ["HRIT UNIVERSITY", "", "", "", "", "", "", "", ""],
        [f"{exam_type} Award Sheets - Master Summary", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", ""],
        [
            "S. No.",
            "Course Code",
            "Course Name",
            "Subject Name",
            "Semester",
            "Faculty",
            "MM",
            "Students",
            "Missing Uni Roll Count"
        ]
    ]

    for i, (code, info) in enumerate(sorted(course_data.items(), key=course_sort_key), 1):
        missing_roll = sum(
            1 for s in info["students"]
            if clean(s["roll_no"]).startswith("(ADM:")
        )

        rows.append([
            i,
            code,
            info["course_name"],
            info["subject_name"],
            info["semester_word"],
            info["faculty_name"],
            info["max_marks"],
            len(info["students"]),
            missing_roll
        ])

    ws.update(values=rows, range_name="A1")

    print(f"✅ Master Summary created for {exam_type}")


# ════════════════════════════════════════
# 17. GENERATE WORKBOOK
# ════════════════════════════════════════

def generate_award_workbook(output_title, course_data, exam_type, template_tab_title):
    print("\n" + "═" * 90)
    print(f"🚀 Creating {exam_type} workbook")
    print("═" * 90)

    new_wb = gc.create(output_title)
    new_sid = new_wb.id

    print(f"✅ Created workbook: {output_title}")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{new_sid}/edit")

    template_sheet_id = get_template_sheet_id(template_tab_title)

    create_master_summary(new_sid, course_data, exam_type)

    print(f"\n📝 Creating {len(course_data)} tabs for {exam_type}...\n")

    count = 0

    for code, info in sorted(course_data.items(), key=course_sort_key):
        count += 1

        tab_title = safe_sheet_title(code)

        copied_sheet_id = copy_template_to_workbook(
            source_template_sheet_id=template_sheet_id,
            destination_spreadsheet_id=new_sid,
            new_title=tab_title
        )

        fill_award_sheet(
            spreadsheet_id=new_sid,
            sheet_title=tab_title,
            sheet_id=copied_sheet_id,
            info=info
        )

        # Gentle throttle
        time.sleep(SLEEP_BETWEEN_TABS_SECONDS)

        # Longer pause every N tabs to avoid 429
        if count % LONG_SLEEP_AFTER_EVERY_N_TABS == 0:
            print(f"⏳ Processed {count} tabs. Sleeping {LONG_SLEEP_SECONDS}s to avoid Google quota...")
            time.sleep(LONG_SLEEP_SECONDS)

    delete_default_sheet_if_exists(new_sid)

    print(f"\n✅ {exam_type} workbook completed")
    print(f"🔗 https://docs.google.com/spreadsheets/d/{new_sid}/edit")

    return new_sid


# ════════════════════════════════════════
# 18. RUN EVERYTHING
# ════════════════════════════════════════

print("\n🚀 Starting corrected full generation...\n")

cie_output_sid = generate_award_workbook(
    output_title=CIE_OUTPUT_TITLE,
    course_data=cie_course_data,
    exam_type="CIE",
    template_tab_title=template_cie_tab
)

ese_output_sid = generate_award_workbook(
    output_title=ESE_OUTPUT_TITLE,
    course_data=ese_course_data,
    exam_type="ESE",
    template_tab_title=template_ese_tab
)

print("\n" + "═" * 90)
print("✅ ALL DONE!")
print("═" * 90)

print("\n📘 CIE Award Sheets:")
print(f"https://docs.google.com/spreadsheets/d/{cie_output_sid}/edit")

print("\n📕 ESE Award Sheets:")
print(f"https://docs.google.com/spreadsheets/d/{ese_output_sid}/edit")


# ════════════════════════════════════════
# 19. FINAL SUMMARY
# ════════════════════════════════════════

def print_final_summary(course_data, exam_type):
    print("\n" + "─" * 130)
    print(f"📊 {exam_type} FINAL SUMMARY")
    print("─" * 130)
    print(f"{'Course Code':<15} {'Semester':<10} {'MM':<5} {'Students':>8} {'Missing Rolls':>14}  Subject Name")
    print("─" * 130)

    total_students = 0
    total_missing = 0

    for code, info in sorted(course_data.items(), key=course_sort_key):
        n = len(info["students"])
        missing_roll = sum(
            1 for s in info["students"]
            if clean(s["roll_no"]).startswith("(ADM:")
        )

        total_students += n
        total_missing += missing_roll

        print(
            f"{code:<15} "
            f"{info['semester_word']:<10} "
            f"{str(info['max_marks']):<5} "
            f"{n:>8} "
            f"{missing_roll:>14}  "
            f"{info['subject_name']}"
        )

    print("─" * 130)
    print(f"{'TOTAL':<32} {total_students:>8} {total_missing:>14}")


print_final_summary(cie_course_data, "CIE")
print_final_summary(ese_course_data, "ESE")
