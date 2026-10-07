# Attendance Report Generator API

Turn a ZK attendance database (`ZK.db`) into a fully formatted Excel report in one request.

Built with **FastAPI**, **Pandas**, and **OpenPyXL**.

---

## Contents

- [Overview](#overview)
- [Quick Start](#quick-start)
- [Ways to Generate a Report](#ways-to-generate-a-report)
- [API Reference](#api-reference)
- [Excel Output](#excel-output)
- [Business Rules](#business-rules)
- [Error Handling](#error-handling)
- [Project Structure](#project-structure)
- [Production Checklist](#production-checklist)
- [Dependencies](#dependencies)

---

## Overview

| You provide              | You get                                                                                           |
| ------------------------ | ------------------------------------------------------------------------------------------------- |
| A `ZK.db` SQLite file    | An `.xlsx` workbook with two sheets: **Data** (raw punches) and **Attendance** (processed report) |
| A date range             | Late / early / overtime / penalty calculations per employee                                       |
| Optional public holidays | Automatic highlighting and OT redistribution                                                      |

**Highlights**

- Upload a database, pick dates, download the report
- Web interface, interactive API docs, or plain HTTP
- Output matches the original `saya.py` script, feature for feature
- Smart handling of accidental double punches (1-hour gap rule)
- Public holiday support with overtime redistribution

---

## Quick Start

**1. Install dependencies**

```bash
pip install -r requirements.txt
```

**2. Start the server**

```bash
python run_api.py
```

Or run uvicorn directly:

```bash
uvicorn attendance_api:app --host 0.0.0.0 --port 8000 --reload
```

**3. Open the app**

| URL                         | Purpose                        |
| --------------------------- | ------------------------------ |
| http://localhost:8000       | Web interface                  |
| http://localhost:8000/docs  | Interactive API docs (Swagger) |
| http://localhost:8000/redoc | Alternative API docs (ReDoc)   |

> **Which file starts the server?**
> `run_api.py` is the entry point and launches `attendance_api:app`. Do **not** use `python main.py`; it is a separate implementation that the API does not use.

---

## Ways to Generate a Report

### Option 1: Web Interface

1. Open http://localhost:8000
2. Select your `ZK.db` file
3. Choose a start date (and optionally an end date)
4. Optionally enter public holidays, comma-separated (e.g. `2025-06-05,2025-06-15`)
5. Click **Generate Report** and wait for the progress bar
6. Click **Download Report**

### Option 2: Interactive Docs

1. Open http://localhost:8000/docs
2. Expand **POST /generate-attendance-report** and click **Try it out**
3. Upload your `ZK.db` and fill in the dates
4. Click **Execute** and download the Excel file from the response

### Option 3: cURL

```bash
curl -X POST "http://localhost:8000/generate-attendance-report" \
  -F "db_file=@/path/to/your/ZK.db" \
  -F "start_date=2025-06-01" \
  -F "end_date=2025-06-30" \
  -F "public_holidays=2025-06-05,2025-06-15" \
  --output attendance_report.xlsx
```

### Option 4: Python

```python
import requests

url = "http://localhost:8000/generate-attendance-report"

with open("ZK.db", "rb") as db:
    response = requests.post(
        url,
        files={"db_file": db},
        data={
            "start_date": "2025-06-01",
            "end_date": "2025-06-30",
            "public_holidays": "2025-06-05,2025-06-15",
        },
    )

if response.status_code == 200:
    with open("attendance_report.xlsx", "wb") as f:
        f.write(response.content)
    print("Report generated successfully!")
else:
    print(f"Error: {response.json()}")
```

---

## API Reference

### `POST /generate-attendance-report`

Generates an Excel attendance report from an uploaded `ZK.db` file.

| Parameter         | Type | Required | Description                                       | Example                 |
| ----------------- | ---- | -------- | ------------------------------------------------- | ----------------------- |
| `db_file`         | File | Yes      | ZK SQLite database (`.db`)                        | `ZK.db`                 |
| `start_date`      | Form | Yes      | Start date, `YYYY-MM-DD`                          | `2025-06-01`            |
| `end_date`        | Form | No       | End date, `YYYY-MM-DD` (defaults to `start_date`) | `2025-06-30`            |
| `public_holidays` | Form | No       | Comma-separated dates, `YYYY-MM-DD`               | `2025-06-05,2025-06-15` |

**Responses**

| Outcome | Result                  |
| ------- | ----------------------- |
| Success | Excel file (`.xlsx`)    |
| Failure | JSON with error details |

---

## Excel Output

The workbook always contains two sheets, in this order:

| Order | Sheet          | Contents                                          |
| ----- | -------------- | ------------------------------------------------- |
| 1     | **Data**       | Raw punch times exactly as stored in the database |
| 2     | **Attendance** | Processed report with all calculations and totals |

Both sheets share the same company name, date range, and Tahoma styling.

### Sheet 1: Data

Raw punches, one row per employee per day.

| Date       | Employee ID | Employee Name | In    | Out   | In    | Out   | In  | Out |
| ---------- | ----------- | ------------- | ----- | ----- | ----- | ----- | --- | --- |
| 2025-01-15 | 101         | John Doe      | 08:00 | 12:00 | 13:00 | 17:00 |     |     |
| 2025-01-16 | 101         | John Doe      | 08:15 | 12:05 | 13:00 | 17:10 |     |     |
| 2025-01-17 | 101         | John Doe      | 07:55 | 11:58 | 13:02 | 17:05 |     |     |

**Features**

- Company name as title, "Raw Punch Data" subtitle, and date range
- Up to 3 In/Out pairs per day
- Sunday rows highlighted yellow
- Bordered cells and optimized column widths
- Frozen header row

### Sheet 2: Attendance

The full processed report.

**Calculations and logic**

- Late and early clock-in/out detection
- Overtime tiers (OT1, OT2, OT3), night shifts, penalties, and allowances
- Time format conversions and decimal calculations
- Employee grouping with totals
- Suspicious punch detection

**Formatting**

- Date range shown under the "Monthly Statement Report" subtitle (e.g. `2025-06-01 to 2025-06-30`)
- Outline grouping (collapsed by default) and frozen panes
- Borders and Tahoma font throughout

### Color Reference

| Color           | Hex       | Meaning                            |
| --------------- | --------- | ---------------------------------- |
| Yellow          | `#FFFF00` | Sunday row                         |
| Blue-gray       | `#B0C4DE` | Public holiday row                 |
| Light red       | `#FFCCCB` | Late / early clock-in or clock-out |
| Orange          | `#FFA500` | Suspicious punch pattern           |
| Bright red font | n/a       | Suspicious early clock-in          |

**Header colors:** light yellow `#FFF2CC`, orange `#FFCC99`, blue `#ADD8E6`, purple `#E6E6FA`, green `#E2EFDA`

---

## Business Rules

### Smart Punch Adjustment (1-Hour Gap Rule)

Accidental quick double-punches are handled automatically for both Clock-In/Clock-Out and In/Out pairs.

| Rule      | Behavior                                                                 |
| --------- | ------------------------------------------------------------------------ |
| Clock-In  | Always the first punch of the day                                        |
| Clock-Out | Must be at least 1 hour after Clock-In, otherwise skip to the next punch |
| In / Out  | Must be at least 1 hour apart, otherwise the Out skips to the next punch |
| Cascading | Adjustments carry forward to keep the sequence logical                   |

All raw punch data is preserved.

**Examples**

| Scenario             | Raw punches                                  | Result                                                       |
| -------------------- | -------------------------------------------- | ------------------------------------------------------------ |
| Normal, no change    | 8:00, 12:00, 1:00 PM, 5:00 PM                | Clock-In 8:00 · Clock-Out 12:00 · In 1:00 PM · Out 5:00 PM   |
| Clock-Out too close  | 8:00, 8:03, 12:00, 1:00 PM, 5:00 PM          | Clock-Out moves from 8:03 to 12:00; In 1:00 PM · Out 5:00 PM |
| In/Out too close     | 8:00, 12:00, 2:00 PM, 2:02 PM, 6:00 PM       | Out moves from 2:02 PM to 6:00 PM                            |
| Multiple adjustments | 8:00, 8:03, 12:00, 2:00 PM, 2:02 PM, 6:00 PM | Clock-In 8:00 · Clock-Out 12:00 · In 2:00 PM · Out 6:00 PM   |

### Public Holidays

Dates passed in `public_holidays` receive two treatments:

1. **Highlighting:** matching rows get the blue-gray background (`#B0C4DE`)
2. **OT redistribution:** OT1 and OT2 are moved into OT3, in both individual rows and TOTAL calculations

|                | OT1   | OT2   | OT3       |
| -------------- | ----- | ----- | --------- |
| Normal day     | 2.5 h | 0.0 h | 1.0 h     |
| Public holiday | 0.0 h | 0.0 h | **3.5 h** |

### Date Format

All dates use `YYYY-MM-DD` (e.g. `2025-06-01`).

---

## Error Handling

The API returns clear errors for:

- Invalid date formats
- Invalid file types (non-`.db` files)
- Database connection errors
- No data found for the specified date range
- Internal server errors

---

## Project Structure

```
├── attendance_api.py                         # Main FastAPI app: SQL logic, punch adjustment, Excel generation
├── run_api.py                                # Entry point: launches attendance_api:app
├── main.py                                   # Separate implementation (not used by the API)
├── requirements.txt                          # Python dependencies
├── README.md                                 # This documentation
├── saya.py                                   # Original script (reference)
└── multi_employee_attendance_converter.py    # Converts Excel to JSON (reference)
```

### How the Pieces Fit

| File                | Role                                                                                     | Used by API |
| ------------------- | ---------------------------------------------------------------------------------------- | ----------- |
| `attendance_api.py` | SQL queries (CTEs), punch adjustment, `generate_data_sheet()`, `generate_excel_report()` | Yes         |
| `run_api.py`        | Starts the server                                                                        | Yes         |
| `main.py`           | Alternate implementation                                                                 | No          |
| `saya.py`           | Original script, kept for reference                                                      | No          |

---

## Production Checklist

This API is designed for **internal use**. Before exposing it publicly, add:

- [ ] Authentication and authorization
- [ ] Rate limiting
- [ ] File size limits
- [ ] Input validation and sanitization
- [ ] HTTPS encryption

---

## Dependencies

| Package  | Purpose                                   |
| -------- | ----------------------------------------- |
| FastAPI  | Web framework                             |
| Uvicorn  | ASGI server                               |
| Pandas   | Data manipulation                         |
| OpenPyXL | Excel file generation                     |
| SQLite3  | Database connectivity (built into Python) |

---

## Verify Your Setup

1. Start the API: `python run_api.py`
2. Open http://localhost:8000/docs
3. Upload a `ZK.db` file with a date range
4. Download the Excel file
5. Confirm it has two sheets: **Data** (first) and **Attendance** (second)

---

<p align="center">
  <strong>Developer:</strong> <a href="https://github.com/Sayed-jibril">github.com/Sayed-jibril</a>
</p>
