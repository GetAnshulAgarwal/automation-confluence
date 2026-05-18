import yaml
import os
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# =========================
# NORMALIZE FOLDER NAME
# =========================
def normalize_folder_name(text):
    """
    BUG FIX: Original code was missing .replace(" ", "")
    This caused folder matching to fail when YAML folder names had spaces.
    """
    return (
        str(text)
        .strip()
        .replace(":", "")
        .replace(".", "")
        .replace(" ", "")   # <-- THIS WAS MISSING IN ORIGINAL
        .lower()
    )


# =========================
# EXTRACT ONLY ENV VARIABLES
# =========================

# All possible keys in Helm values YAML that hold env var lists
ENV_LIST_KEYS = {"env", "extraEnvs", "extraEnv", "envVars", "environment"}


def process_env_list(env_list, result):
    """
    Process a list of {name, value/valueFrom} dicts into result dict.

    For valueFrom.secretKeyRef entries, we store a dict with both
    the sentinel and the actual key name, e.g.:
      {"type": "secretKeyRef", "key": "BIDB_USERNAME"}
    This allows compare_data to check if the PDF value is contained
    in the secretKeyRef key name (case-insensitive).
    """
    for env_item in env_list:
        if not isinstance(env_item, dict):
            continue
        if "name" not in env_item:
            continue

        key = str(env_item.get("name")).strip()
        if not key:
            continue

        if "value" in env_item:
            result[key] = str(env_item.get("value")).strip()
        elif "valueFrom" in env_item:
            value_from = env_item.get("valueFrom", {})
            secret_key = ""
            if isinstance(value_from, dict):
                # Try secretKeyRef first, then configMapKeyRef
                ref = value_from.get("secretKeyRef") or value_from.get("configMapKeyRef") or {}
                if isinstance(ref, dict):
                    secret_key = str(ref.get("key", "")).strip()
                # fieldRef (e.g. metadata.name) — store fieldPath
                elif "fieldRef" in value_from:
                    secret_key = str(value_from["fieldRef"].get("fieldPath", "")).strip()
            result[key] = {
                "type": "secretKeyRef",
                "key": secret_key
            }


def extract_envs(data, result):
    """
    Recursively walk the YAML structure and collect all env vars.
    Handles both 'env' and 'extraEnvs' (and other common variants).
    """
    if isinstance(data, dict):
        for key, value in data.items():
            if key in ENV_LIST_KEYS and isinstance(value, list):
                process_env_list(value, result)
            else:
                extract_envs(value, result)

    elif isinstance(data, list):
        for item in data:
            extract_envs(item, result)


# =========================
# READ YAML FILE
# =========================
def read_yaml_file(file_path):
    """
    Robustly read a YAML file handling:
    - UTF-8 with BOM (common in Windows-saved files)
    - UTF-16 encoding
    - Trailing whitespace / Windows CRLF line endings
    - Multiple YAML documents in one file
    """
    # Try encodings in order
    encodings = ["utf-8-sig", "utf-8", "utf-16", "latin-1"]
    raw_text = None

    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                raw_text = f.read()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if raw_text is None:
        print(f"WARNING: Could not read file with any encoding: {file_path}")
        return {}

    # Normalize line endings (Windows CRLF -> LF)
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    # YAML does not allow TAB characters anywhere except inside quoted scalars.
    # Strategy: process line by line —
    #   - Lines that are purely indentation + key: keep leading spaces, strip trailing tabs
    #   - Lines that contain a quoted value: replace ALL tabs with spaces safely
    #   - Trailing tabs/spaces at end of any line: strip them
    import re as _re
    def _sanitize_tabs(text):
        clean_lines = []
        for line in text.split("\n"):
            # Strip trailing whitespace/tabs from every line
            line = line.rstrip()
            # Replace any remaining tab that appears mid-line (inside values)
            # We replace tab with a single space everywhere — YAML indentation
            # uses spaces anyway, so leading tabs would be a structural error
            # in the source file; replacing them with spaces is the safest option.
            line = line.replace("\t", " ")
            clean_lines.append(line)
        return "\n".join(clean_lines)
    raw_text = _sanitize_tabs(raw_text)

    # Try loading — first as single doc, then as multi-doc
    result = {}
    try:
        data = yaml.safe_load(raw_text)
        if data is None:
            return {}
        extract_envs(data, result)

    except yaml.YAMLError:
        # Fallback: try as multiple YAML documents
        try:
            for doc in yaml.safe_load_all(raw_text):
                if doc:
                    extract_envs(doc, result)
        except yaml.YAMLError as e:
            print(f"WARNING: Could not parse YAML file: {file_path} — {e}")
            return {}

    return result


# =========================
# FIND PROD / UAT YAML FILES
# =========================
def find_yaml_files(base_path, folder_name):
    prod_file = None
    uat_file = None

    folder_name_norm = normalize_folder_name(folder_name)

    if not os.path.exists(base_path):
        return prod_file, uat_file

    for item in os.listdir(base_path):
        item_path = os.path.join(base_path, item)

        if not os.path.isdir(item_path):
            continue

        if normalize_folder_name(item) == folder_name_norm:
            for file in os.listdir(item_path):
                full_path = os.path.join(item_path, file)

                file_lower = file.lower()
                # Accept all known prod file name variants
                if file_lower in ("prd.values.yaml", "prod-ireland.values.yaml", "prd-ireland.values.yaml"):
                    prod_file = full_path
                # Accept all known uat file name variants
                elif file_lower in ("uat.values.yaml", "uat-ireland.values.yaml"):
                    uat_file = full_path

            break

    return prod_file, uat_file


# =========================
# PLACEHOLDER CHECK
# =========================
def is_placeholder(value):
    """
    If value is like <something> it means the PDF just marks
    that this field should exist — only check name presence in YAML.
    """
    value = str(value).strip()
    return value.startswith("<") and value.endswith(">")


# =========================
# COMPARE DATA
# =========================
def compare_data(doc_data, yaml_base_path):
    """
    Compares extracted doc data against YAML files.

    Two comparison modes based on PDF value:
    - Placeholder (<value>): Only check if Name EXISTS in YAML
    - Actual value: Check if Name exists AND Value matches

    Returns:
    - prod_rows: all comparison rows for PROD env
    - uat_rows:  all comparison rows for UAT env
    (Both include ALL rows; filtering to mismatch-only happens in generate_excel)
    """
    prod_rows = []
    uat_rows = []

    for folder, rows in doc_data.items():

        prod_file, uat_file = find_yaml_files(yaml_base_path, folder)

        # Check if the subfolder itself exists on disk
        folder_norm = normalize_folder_name(folder)
        folder_exists_on_disk = False
        if os.path.exists(yaml_base_path):
            for item in os.listdir(yaml_base_path):
                if normalize_folder_name(item) == folder_norm:
                    folder_exists_on_disk = True
                    break

        yaml_files = {
            "prod": prod_file,
            "uat": uat_file
        }

        for env_name, yaml_file in yaml_files.items():

            current_list = prod_rows if env_name == "prod" else uat_rows

            # --- Case: subfolder not present in uploaded app folder ---
            if not folder_exists_on_disk:
                for row in rows:
                    current_list.append({
                        "Folder": folder,
                        "Name": str(row["name"]).strip(),
                        "PDF Value": str(row["value"]).strip(),
                        "YAML Value": "Folder not present",
                        "File": "Folder not present",
                        "YAML File Found": "Folder not present",
                        "Status": "FOLDER_NOT_PRESENT"
                    })
                continue

            yaml_data = {}
            if yaml_file and os.path.exists(yaml_file):
                yaml_data = read_yaml_file(yaml_file)

            for row in rows:
                pdf_name = str(row["name"]).strip()
                pdf_value = str(row["value"]).strip()

                yaml_value = yaml_data.get(pdf_name, "NOT FOUND")

                # Resolve yaml_value display string and secretKeyRef key
                secret_ref_key = ""
                if isinstance(yaml_value, dict) and yaml_value.get("type") == "secretKeyRef":
                    secret_ref_key = yaml_value.get("key", "")
                    yaml_value_display = f"valueFrom.secretKeyRef (key: {secret_ref_key})" if secret_ref_key else "valueFrom.secretKeyRef"
                else:
                    yaml_value_display = str(yaml_value)

                # --- Condition 1: Placeholder value like <Open telemetry endpoint> ---
                # Only check if the NAME exists in YAML
                if is_placeholder(pdf_value):
                    if pdf_name in yaml_data:
                        status = "FOUND"
                    else:
                        status = "NOT_FOUND"

                # --- Condition 2: Actual value ---
                # Check both NAME and VALUE match (case-insensitive)
                else:
                    if pdf_name not in yaml_data:
                        status = "NOT_FOUND"

                    elif isinstance(yaml_data[pdf_name], dict) and yaml_data[pdf_name].get("type") == "secretKeyRef":
                        # valueFrom.secretKeyRef case:
                        # Check if pdf_value is contained in the secretKeyRef key (case-insensitive)
                        # e.g. PDF value "BIDB" matches secretKeyRef key "BIDB_USERNAME" -> MATCHED
                        if secret_ref_key and pdf_value.upper() in secret_ref_key.upper():
                            status = "MATCHED"
                        else:
                            status = "VALUE_MISMATCH"

                    elif pdf_value.strip().lower() == str(yaml_data[pdf_name]).strip().lower():
                        # Case-insensitive match
                        # e.g. "RealTime.Alerts.Alert" == "REALTIME.Alerts.Alert" -> MATCHED
                        status = "MATCHED"

                    else:
                        status = "VALUE_MISMATCH"

                current_list.append({
                    "Folder": folder,
                    "Name": pdf_name,
                    "PDF Value": pdf_value,
                    "YAML Value": yaml_value_display,
                    "File": env_name.upper(),
                    "YAML File Found": "YES" if yaml_file else "NO",
                    "YAML File Path": yaml_file if yaml_file else "NOT FOUND",
                    "Status": status
                })

    return prod_rows, uat_rows


# =========================
# COLOR MAP FOR STATUS
# =========================
STATUS_COLORS = {
    "NOT_FOUND":          "FFFF0000",  # Red
    "VALUE_MISMATCH":     "FFFFA500",  # Orange
    "FOUND":              "FF90EE90",  # Light green (placeholder found)
    "MATCHED":            "FF00CC00",  # Green
    "FOLDER_NOT_PRESENT": "FFFF00FF",  # Magenta/Purple
}


def apply_excel_styling(ws, df):
    """
    Apply header styling and row color coding based on Status column.
    """
    # Header row styling
    header_fill = PatternFill(start_color="FF2F5496", end_color="FF2F5496", fill_type="solid")
    header_font = Font(color="FFFFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Get Status column index (1-based)
    status_col_idx = None
    for idx, col in enumerate(df.columns, start=1):
        if col == "Status":
            status_col_idx = idx
            break

    # Color each data row based on Status
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        status_cell = ws.cell(row=row_idx, column=status_col_idx)
        status_value = str(status_cell.value) if status_cell.value else ""

        color = STATUS_COLORS.get(status_value)
        if color:
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for cell in row:
                cell.fill = fill

    # Auto-fit column widths
    for col_idx, col in enumerate(df.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col))
        for row in ws.iter_rows(min_row=2):
            cell_val = str(row[col_idx - 1].value or "")
            if len(cell_val) > max_len:
                max_len = len(cell_val)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)


# =========================
# GENERATE EXCEL
# =========================
def generate_excel(prod_rows, uat_rows, output_file):
    """
    Generate Excel with two sheets: PROD_MISMATCH and UAT_MISMATCH.

    Only rows with status NOT_FOUND or VALUE_MISMATCH are written.
    (For placeholders, NOT_FOUND means name missing in YAML = problem)
    Rows with MATCHED or FOUND are considered passing — excluded.

    Color coding:
    - Red    = NOT_FOUND (key missing in YAML entirely)
    - Orange = VALUE_MISMATCH (key exists but value differs)
    """
    columns = [
        "Folder",
        "Name",
        "PDF Value",
        "YAML Value",
        "File",
        "YAML File Found",
        "Status"
    ]

    # Filter: only keep mismatch rows
    mismatch_statuses = {"NOT_FOUND", "VALUE_MISMATCH", "FOLDER_NOT_PRESENT"}

    prod_mismatch = [r for r in prod_rows if r.get("Status") in mismatch_statuses]
    uat_mismatch  = [r for r in uat_rows  if r.get("Status") in mismatch_statuses]

    prod_df = pd.DataFrame(prod_mismatch, columns=columns) if prod_mismatch else pd.DataFrame(columns=columns)
    uat_df  = pd.DataFrame(uat_mismatch,  columns=columns) if uat_mismatch  else pd.DataFrame(columns=columns)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        prod_df.to_excel(writer, sheet_name="PROD_MISMATCH", index=False)
        uat_df.to_excel(writer,  sheet_name="UAT_MISMATCH",  index=False)

        # Apply styling
        apply_excel_styling(writer.sheets["PROD_MISMATCH"], prod_df)
        apply_excel_styling(writer.sheets["UAT_MISMATCH"],  uat_df)