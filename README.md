# AttendanceReader

A standalone Windows Desktop application that reads a **Malam Sachar Plus**
(מלם שכר Plus) monthly attendance PDF and converts it into a structured
Excel (`.xlsx`) file – **no Python or runtime installation required** by the
end-user.

Two application modes are provided:

| Application | File | When to use |
|-------------|------|-------------|
| Local PDF parser | `attendance_reader.py` | PDFs up to ~100 pages; no cloud required |
| Document AI (cloud) | `attendance_reader_docai.py` | PDFs > 100 pages; uses Google Cloud Document AI Batch Processing |

---

## Features

- Simple GUI: browse for the PDF, pick an output path, one click to convert.
- Parses the horizontal date/time grid produced by Malam Sachar Plus.
- Handles Hebrew RTL text (day letters א–ש, column headers).
- Output `.xlsx` columns:
  | Column | Hebrew header | Meaning |
  |--------|--------------|---------|
  | A | תאריך | Date (DD/MM/YY or DD/MM/YYYY) |
  | B | יום בשבוע | Day of the week |
  | C | שעת כניסה | Entry time |
  | D | שעת יציאה | Exit time |
  | E | סה"כ שעות ליום | Total hours per day |
  | F | סוג יום | Day type (e.g. חופשה, מחלה) |

---

## Application 1 – Local PDF Parser (`attendance_reader.py`)

Uses `pdfplumber` to extract text directly from the PDF — no internet or
cloud credentials required.

### Running from source

```bash
pip install -r requirements.txt
python attendance_reader.py
```

### Building a standalone `.exe`

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name AttendanceReader attendance_reader.py
```

---

## Application 2 – Document AI Batch Processor (`attendance_reader_docai.py`)

For large PDFs (> 100 pages) that exceed synchronous Document AI limits.
Uploads the PDF to Google Cloud Storage, runs a Document AI Batch Processing
job, retrieves the sharded JSON results via the Document AI Toolbox, and
converts them to Excel.

### Google Cloud prerequisites

1. A GCP project with the **Document AI API** and **Cloud Storage API** enabled.
2. A Document AI processor (e.g. Document OCR or Form Parser) in a supported region.
3. A GCS bucket in the same region as the processor.
4. A service-account JSON key with the following roles:
   - `roles/documentai.apiUser`
   - `roles/storage.objectAdmin` (on the GCS bucket)

### Configuration

Copy `.env.example` to `.env` in the same directory as the executable and
fill in your values.  **Never commit `.env` or the service-account key.**

```ini
PROJECT_ID=your-gcp-project-id
LOCATION=us
PROCESSOR_ID=your-processor-id
GCS_BUCKET_NAME=your-gcs-bucket-name
GOOGLE_APPLICATION_CREDENTIALS=C:\path\to\your-service-account-key.json
```

### Running from source

```bash
pip install -r requirements.txt
python attendance_reader_docai.py
```

### Building a standalone `.exe`

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name AttendanceReaderDocAI \
    --exclude-module pdfplumber \
    attendance_reader_docai.py
```

The resulting executable is `dist\AttendanceReaderDocAI.exe`.
Place the `.env` file and the service-account JSON key in the **same folder**
as the executable — they are loaded at runtime and must **not** be bundled
inside the `.exe`.

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pdfplumber | 0.11.4 | Extract text from PDF (local mode) |
| openpyxl | 3.1.5 | Write `.xlsx` files |
| google-cloud-documentai | 3.10.0 | Document AI Batch Processing API client |
| google-cloud-storage | 3.9.0 | GCS upload/download |
| google-cloud-documentai-toolbox | 0.15.0a0 | Parse sharded batch results |
| python-dotenv | 1.2.1 | Load `.env` configuration |
| tkinter | stdlib | Native Windows GUI (bundled with Python) |

