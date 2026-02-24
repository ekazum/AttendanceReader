# AttendanceReader

A standalone Windows Desktop application that reads a **Malam Saar**
(מלם שכר) monthly attendance PDF and converts it into a structured
Excel (`.xlsx`) file – **no Python or runtime installation required** by the
end-user.

---

## Features

- Simple GUI: browse for the PDF, click **Export to Excel**, done.
- Uses `pdfplumber` to extract text directly from the PDF — no internet or
  cloud credentials required.
- Parses the horizontal date/time grid produced by Malam Saar.
- Handles Hebrew RTL text (day letters, column headers, activity names).
- Translates Hebrew activity labels to English (e.g. חופשה → Vacation,
  מחלה → Sick, מילואים → Military Reserve).
- Writes a two-sheet Excel workbook:

  **Sheet 1 – Daily Attendance** (one row per calendar day):

  | Column | Header | Description |
  |--------|--------|-------------|
  | A | Employee ID | Employee identity number |
  | B | Employee Name | Employee full name |
  | C | Tag Number | Badge / tag number |
  | D | Salary Month | MM/YYYY of the pay period |
  | E | Date | Full date (DD/MM/YYYY) |
  | F | Day of Week | Hebrew day letter |
  | G | Day Type | Activity in English (Work, Vacation, Sick, …) |
  | H | Shift Premium | `True` if a shift-premium marker (`*`) is present |
  | I | Entry (Actual) | Actual clock-in time |
  | J | Exit (Actual) | Actual clock-out time |
  | K | Total Present Hours | Total time present (`[h]:mm`) |
  | L | Entry (For Pay) | Payable entry time |
  | M | Exit (For Pay) | Payable exit time |
  | N | Total For Pay Hours | Total payable hours (`[h]:mm`) |
  | O | Standard Hours | Standard contracted hours (`[h]:mm`) |
  | P | OT 100% | Overtime at 100% rate (`[h]:mm`) |
  | Q | OT 125% | Overtime at 125% rate (`[h]:mm`) |
  | R | OT 150% | Overtime at 150% rate (`[h]:mm`) |
  | S | OT 200% | Overtime at 200% rate (`[h]:mm`) |
  | T | Shift Bonus 87 | Shift bonus at 87% (`[h]:mm`) |
  | U | Shift Bonus 50 | Shift bonus at 50% (`[h]:mm`) |
  | V | Shift Bonus 20 | Shift bonus at 20% (`[h]:mm`) |
  | W | Deduction | Deduction hours (`[h]:mm`) |

  **Sheet 2 – Monthly Summary** (one row per salary month, SUMIF formulas
  referencing Sheet 1):

  | Column | Header |
  |--------|--------|
  | A | Salary Month |
  | B | Work Days |
  | C | Vacation Days |
  | D | Sick Days |
  | E | No Report Days |
  | F | Total Present Hours |
  | G | Total For Pay Hours |
  | H | Total OT 100% |
  | I | Total OT 125% |
  | J | Total OT 150% |
  | K | Total OT 200% |
  | L | Military Reserve Days |
  | M | On-Call Days |
  | N | Work Accident Days |
  | O | Recess Days |

---

## Configuration – `columns_config.json`

The PDF parser locates each data column by its horizontal x-coordinate range.
These ranges are stored in `columns_config.json` (included in the repo and
bundled into the `.exe`).

If a copy of `columns_config.json` is placed **next to the executable**, it
takes precedence over the bundled defaults — allowing you to fine-tune ranges
for different PDF layouts without rebuilding.

If neither file is found at runtime, the application writes the default config
to the executable's directory automatically.

---

## Running from source

```bash
pip install -r requirements.txt
python malam_saar_attendance.py
```

## Building a standalone `.exe`

**Windows (Command Prompt):**

```bat
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name MalamSaarAttendance ^
    --add-data "columns_config.json;." ^
    malam_saar_attendance.py
```

**Linux / macOS (Bash):**

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name MalamSaarAttendance \
    --add-data "columns_config.json:." \
    malam_saar_attendance.py
```

The resulting executable is `dist/MalamSaarAttendance` (or `dist\MalamSaarAttendance.exe` on Windows).

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pdfplumber | 0.11.4 | Extract text and word positions from PDF |
| openpyxl | 3.1.5 | Write `.xlsx` files |
| tkinter | stdlib | Native Windows GUI (bundled with Python) |

