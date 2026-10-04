# Work Time Analyzer — Upgraded WOP/Date Edition

Built against the uploaded `sep2026.xlsx` structure.

Workbook inspected: ['WorkDurationReport']
First sheet shape: (269, 36)

Rules:
- P = Present
- A = Absent
- WOP = Saturday working day
- WO = Sunday weekly off
- WOP is counted as a working day.
- WOP/Saturday OT is excluded from regular Total OT.
- Date is included in daily records and month filtering.
- Export: Daily/Monthly × All/Selected Employees × PDF/Word/Image.

Deploy by replacing `app.py` and `requirements.txt` in the existing Streamlit GitHub repository.
