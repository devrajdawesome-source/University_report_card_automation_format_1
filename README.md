A Google Colab-based automation script to generate corrected CIE and ESE award sheets for HRIT University using data from Google Sheets.

This project reads raw CIE, ESE, course-code, and student roll mapping sheets, standardizes the data, maps enrollment/admission IDs to university roll numbers, and generates separate formatted Google Sheets workbooks with one tab per course code.

✅ Key Features
✅ Generates separate workbooks for:
CIE Award Sheets
ESE Award Sheets
✅ Creates one tab per course code
✅ Copies formatting from predefined Google Sheets templates
✅ Maps Admission / Enrollment ID → University Roll No
✅ Uses corrected ESE column logic:
Course code from subjectcode / Column A
Subject name from subjectname / Column B
Marks from scaled_marks / Column U
Filters only students where is_absent = false
✅ Applies fixed maximum marks:
CIE MM = 40
ESE MM = 60
✅ Automatically detects semester from course code:
101 → First Semester
202 → Second Semester
301 → Third Semester
✅ Converts marks into words:
35 → Thirty Five
Absent → Absent
✅ Handles Google Sheets quota/rate-limit errors using retry + exponential backoff
✅ Creates a Master Summary tab for each generated workbook
🧾 Project Purpose
This script was built to automate the generation of university award sheets from raw evaluation dumps.

Instead of manually preparing individual course-wise sheets, this tool:

Reads source data from Google Sheets
Cleans and standardizes CIE/ESE records
Maps roll numbers
Copies official templates
Fills student-wise marks
Generates final formatted award sheet workbooks
