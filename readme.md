# DOCX / PDF vs YAML Comparator

A web-based tool that extracts environment variables from a Confluence-exported **DOCX or PDF** file, compares them against Helm **YAML values files** (`prd.values.yaml` / `uat.values.yaml`), and generates a color-coded **Excel report** of all mismatches.

---

## What It Does

For each application folder (e.g. `realtimeapi`, `predictiontimeapp`), the tool:

1. Reads environment variable **Name** and **Value** from the uploaded DOCX/PDF
2. Looks up the same **Name** in the corresponding `prd.values.yaml` and `uat.values.yaml`
3. Applies comparison rules (see below)
4. Outputs an Excel file with only the **mismatched rows**

---

## Comparison Rules

| PDF Value Type | Example | What is Checked |
|---|---|---|
| **Placeholder** `<...>` | `<TIMEZONE>` | Only checks if the **Name exists** in YAML |
| **Actual value** | `LRD` | Checks if **Name exists AND value matches** (case-insensitive) |
| **Secret ref** | `BIDB` | Checks if PDF value is **contained in** the `secretKeyRef.key` field |

### Status Values in Excel

| Color | Status | Meaning |
|---|---|---|
| 🔴 Red | `NOT_FOUND` | Key is missing from YAML entirely |
| 🟠 Orange | `VALUE_MISMATCH` | Key exists but value does not match |

> Rows with `MATCHED` or `FOUND` status are **not included** in the Excel — only problems are shown.

---

## Project Structure

```
project/
│
├── backend/
│   ├── app.py              # Flask API server
│   ├── parser_utils.py     # DOCX / PDF extraction logic
│   ├── comparator.py       # YAML reading and comparison logic
│   └── requirements.txt    # Python dependencies
│
└── frontend/
    ├── index.html          # UI
    ├── script.js           # Upload and download logic
    └── style.css           # Styling
```

---

## Prerequisites

- **Python 3.9+**
- **pip**
- A modern browser (Chrome / Edge recommended for folder upload support)

---

## Installation

### 1. Clone or download the project

```bash
git clone <your-repo-url>
cd project/backend
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

`requirements.txt` includes:
```
Flask==3.0.3
flask-cors==4.0.1
pdfplumber==0.11.4
python-docx==1.1.2
PyYAML==6.0.2
pandas==2.2.3
openpyxl==3.1.5
Werkzeug==3.0.4
```

### 3. Start the Flask server

```bash
cd backend
python app.py
```

Server starts at: `http://127.0.0.1:5000`

### 4. Open the frontend

Open `frontend/index.html` directly in your browser — **no extra server needed**.

---

## Required File Structure

### DOCX / PDF (Confluence Export)

The document must contain sections with the **application folder name** as a heading, followed by a table with `Name` and `Value` columns:

```
realtimeapi
┌──────────────────────────────┬──────────────────────────┐
│ Name                         │ Value                    │
├──────────────────────────────┼──────────────────────────┤
│ TZ                           │ <TIMEZONE>               │
│ DB_SID                       │ LRD                      │
│ OTEL_EXPORTER_OTLP_ENDPOINT  │ <Open telemetry endpoint>│
└──────────────────────────────┴──────────────────────────┘
```

### App Folder (YAML Files)

The uploaded folder **must be named `app`** and contain one subfolder per application. Each subfolder must have exactly:
- `prd.values.yaml` — production values
- `uat.values.yaml` — UAT values

```
app/                          ← Upload this folder
├── realtimeapi/
│   ├── prd.values.yaml
│   └── uat.values.yaml
├── predictiontimeapp/
│   ├── prd.values.yaml
│   └── uat.values.yaml
├── StopPrediction/
│   ├── prd.values.yaml
│   └── uat.values.yaml
└── can-dataprocessor/
    ├── prd.values.yaml
    └── uat.values.yaml
```



## How to Use

1. **Start the backend** (`python app.py`)
2. **Open** `index.html` in your browser
3. **Upload PDF or DOCX** — the Confluence export file
4. **Upload App Folder** — select the `app` folder containing all YAML subfolders
5. Click **Compare & Download Excel**
6. The Excel file downloads automatically with two sheets:
   - `PROD_MISMATCH` — mismatches found in `prd.values.yaml`
   - `UAT_MISMATCH` — mismatches found in `uat.values.yaml`

---

## Excel Output Format

Each sheet contains these columns:

| Column | Description |
|---|---|
| `Folder` | Application name (e.g. `realtimeapi`) |
| `Name` | Environment variable name |
| `PDF Value` | Value from the Confluence document |
| `YAML Value` | Value found in the YAML file (or `NOT FOUND`) |
| `File` | `PRD` or `UAT` |
| `YAML File Found` | `YES` if the YAML file exists for this folder |
| `Status` | `NOT_FOUND` or `VALUE_MISMATCH` |

---

## Troubleshooting

### ❌ "No service folders found inside the uploaded 'app' folder"
- Make sure the uploaded folder is named **exactly `app`**
- Each subfolder must contain `prd.values.yaml` or `uat.values.yaml`

### ❌ "No data could be extracted from the document"
- Check that the folder names in the document **match** the subfolder names inside `app/`
- Folder name matching is **case-insensitive** (e.g. `StopPrediction` matches `stopprediction`)

### ❌ YAML parse error
- The tool handles UTF-8 BOM and Windows CRLF line endings automatically
- If a YAML file is corrupted, it will be **skipped with a warning** in the terminal and all other folders will still be compared

### ❌ Values showing as VALUE_MISMATCH when they look the same
- Comparison is **case-insensitive** — `RealTime.Alerts.Alert` == `REALTIME.Alerts.Alert` ✅
- For `secretKeyRef` values — the PDF value only needs to be **contained in** the `key` field (e.g. PDF `BIDB` matches YAML key `BIDB_USERNAME`) ✅
- Check for invisible characters or trailing backticks in the DOCX — these are stripped automatically

---

## Notes

- The `uploads/` folder inside `backend/` is **cleared on every new comparison** — do not store anything important there
- The tool runs entirely **locally** — no data is sent to any external server
- Both **PDF** (pdfplumber) and **DOCX** (python-docx) formats are supported