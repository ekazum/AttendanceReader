# AttendanceReader

A standalone Windows Desktop application that reads a **Malam Sachar Plus**
(מלם שכר Plus) monthly attendance PDF and converts it into a structured
Excel (`.xlsx`) file – **no Python or runtime installation required** by the
end-user.

---

## Features

- Simple GUI: browse for the PDF, pick an output path, one click to convert.
- Parses the horizontal date/time grid produced by Malam Sachar Plus.
- Handles Hebrew RTL text (day letters א–ש, column headers).
- Output `.xlsx` columns:
  | Column | Hebrew header | Meaning |
  |--------|--------------|---------|
  | A | תאריך | Date (DD/MM) |
  | B | יום בשבוע | Day of the week |
  | C | שעת כניסה | Entry time |
  | D | שעת יציאה | Exit time |
  | E | סה"כ שעות ליום | Total hours per day |

---

## Running from source

### Prerequisites

- Python 3.10+
- `pip install -r requirements.txt`

### Run

```bash
python attendance_reader.py
```

---

## Building a standalone `.exe` (Windows)

### Prerequisites

```bash
pip install -r requirements.txt
pip install pyinstaller
```

### Build command

```bash
pyinstaller --onefile --windowed --name AttendanceReader attendance_reader.py
```

The resulting executable is written to `dist\AttendanceReader.exe`.  
Distribute that single file – no Python or additional libraries are needed on
the target machine.

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pdfplumber | 0.11.4 | Extract text/word positions from PDF |
| openpyxl | 3.1.5 | Write `.xlsx` files |
| tkinter | stdlib | Native Windows GUI (bundled with Python) |
