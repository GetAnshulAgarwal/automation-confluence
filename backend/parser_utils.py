import pdfplumber
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from docx.document import Document as DocxDocument
import re


# =========================
# BASIC HELPERS
# =========================
def clean_text(text):
    if text is None:
        return ""
    return str(text).strip().strip("`").strip()


def normalize_folder_name(text):
    text = clean_text(text)
    text = text.replace(":", "")
    text = text.replace(".", "")
    text = text.replace(" ", "")
    return text.lower()


def collapse_repeated_folder(text):
    """
    Handles cases where DOCX repeats folder name in merged cells.
    e.g. "realtimeapirealtimeapi" -> "realtimeapi"
    """
    text = normalize_folder_name(text)

    if not text:
        return ""

    for count in range(2, 10):
        if len(text) % count == 0:
            size = len(text) // count
            part = text[:size]
            if part * count == text:
                return part

    return text


def normalize_header(text):
    text = clean_text(text).lower()
    text = text.replace(" ", "")
    text = collapse_repeated_folder(text)

    if "name" in text:
        return "name"
    if "key" in text:
        return "name"
    if "value" in text:
        return "value"

    return text


def is_valid_env_name(text):
    """
    Valid env variable name: letters (upper or lower), digits, underscores,
    dots, hyphens. Must start with a letter or underscore or digit.
    Mixed-case names like "DatabaseMachineIP" and "Port" are allowed
    since some apps use camelCase env vars.
    """
    text = clean_text(text)
    return bool(re.match(r"^[A-Za-z0-9_][A-Za-z0-9_.\-]*$", text))


def get_known_folder_match(text, known_folders):
    """
    Check if this text IS a folder name (or contains one).
    Returns the matched folder name (normalized) or None.
    """
    collapsed = collapse_repeated_folder(text)

    if not collapsed:
        return None

    known_normalized = [normalize_folder_name(f) for f in known_folders]

    # Exact match first
    for folder in known_normalized:
        if collapsed == folder:
            return folder

    # Substring match (folder name inside cell text)
    for folder in known_normalized:
        if folder and folder in collapsed:
            return folder

    return None


def add_row(extracted_data, folder, name, value):
    folder = collapse_repeated_folder(folder)
    name = clean_text(name)
    value = clean_text(value)

    if not folder or not name:
        return

    if not is_valid_env_name(name):
        return

    if folder not in extracted_data:
        extracted_data[folder] = []

    extracted_data[folder].append({
        "name": name,
        "value": value
    })


def remove_duplicate_rows(extracted_data):
    final_data = {}

    for folder, rows in extracted_data.items():
        folder = collapse_repeated_folder(folder)
        seen = set()
        final_rows = []

        for row in rows:
            key = (row["name"], row["value"])
            if key not in seen:
                seen.add(key)
                final_rows.append(row)

        if final_rows:
            final_data[folder] = final_rows

    return final_data


# =========================
# DOCX BLOCK ITERATOR
# =========================
def iter_block_items(parent):
    """
    Yields Paragraph and Table blocks in document order.
    Works for both Document root and nested Cell parents.
    """
    if isinstance(parent, DocxDocument):
        parent_element = parent.element.body
    elif isinstance(parent, _Cell):
        parent_element = parent._tc
    else:
        return

    for child in parent_element.iterchildren():
        tag = child.tag
        if tag.endswith("}p") or tag.endswith("p"):
            yield Paragraph(child, parent)
        elif tag.endswith("}tbl") or tag.endswith("tbl"):
            yield Table(child, parent)


def get_row_text(row):
    """Get combined unique text from all cells in a row."""
    texts = []
    for cell in row.cells:
        text = clean_text(cell.text)
        if text and text not in texts:
            texts.append(text)
    return " ".join(texts)


def find_header_indexes(row):
    """
    Detect if this row is a header row with Name and Value columns.
    Returns (name_index, value_index) or (None, None).
    """
    headers = [normalize_header(cell.text) for cell in row.cells]

    if "name" in headers and "value" in headers:
        return headers.index("name"), headers.index("value")

    return None, None


# =========================
# TABLE PARSER
# =========================
def parse_table(table, extracted_data, known_folders, current_folder):
    """
    Parse a DOCX table to extract env name/value pairs.

    STRICT RULE: name_index/value_index are LOCAL to this table only.
    Data rows are extracted ONLY if a Name|Value header was found
    in THIS exact table. No state is carried in from outside.

    This prevents false extraction from description tables, heading rows,
    or any table that doesn't explicitly have a Name/Value header.
    """
    # These are LOCAL to this table — never inherited from outside
    name_index = None
    value_index = None
    header_found_in_this_table = False

    # --- PASS 1: Scan rows for folder headings, headers, and data ---
    for row in table.rows:

        # --- Folder heading check ---
        # CRITICAL FIX: Only check the FIRST cell for folder name.
        # Checking all cells (get_row_text) causes false matches when the
        # VALUE column contains folder-like substrings, e.g.:
        #   "KAFKA_TOPIC_ACTUALDRIVINGDWELL | Driving_DwellTime" -> falsely matches "dwelltime"
        #   "REALTIME_MODEL_FOLDER_DWELL | /san/.../DwellTime"   -> falsely matches "dwelltime"
        #   "OTEL_SERVICE_NAME | dwelltime_basemodel"            -> falsely matches "dwelltime"
        # By only checking the first cell, we ensure only true heading rows are matched.
        first_cell_text = clean_text(row.cells[0].text) if row.cells else ""

        # A row is a folder heading only if first cell matches AND it's not
        # a valid env var name (env var names are all-caps+underscores,
        # folder names like "realtimeapi" are lowercase alphanumeric)
        folder_match = get_known_folder_match(first_cell_text, known_folders)
        is_header_row = find_header_indexes(row)[0] is not None

        if folder_match and not is_header_row and not is_valid_env_name(first_cell_text):
            current_folder = folder_match
            name_index = None
            value_index = None
            header_found_in_this_table = False
            continue

        # --- Header row check (Name | Value columns) ---
        possible_name_idx, possible_value_idx = find_header_indexes(row)
        if possible_name_idx is not None and possible_value_idx is not None:
            name_index = possible_name_idx
            value_index = possible_value_idx
            header_found_in_this_table = True
            continue

        # --- Data row extraction ---
        # STRICT: Only extract if a Name|Value header was found in THIS table
        if (
            header_found_in_this_table
            and current_folder
            and name_index is not None
            and value_index is not None
        ):
            cells = row.cells
            if len(cells) > max(name_index, value_index):
                name = clean_text(cells[name_index].text)
                value = clean_text(cells[value_index].text)
                # Skip if name is empty or contains spaces (not an env var)
                if name and " " not in name:
                    add_row(extracted_data, current_folder, name, value)

    # --- PASS 2: Recurse into nested tables inside cells ---
    for row in table.rows:
        for cell in row.cells:
            for block in iter_block_items(cell):
                if isinstance(block, Table):
                    current_folder = parse_table(
                        block,
                        extracted_data,
                        known_folders,
                        current_folder
                    )

    return current_folder


# =========================
# DOCX MAIN PARSER
# =========================
def parse_docx(docx_path, known_folders=None):
    doc = Document(docx_path)

    if known_folders is None:
        known_folders = []

    extracted_data = {}
    current_folder = None

    for block in iter_block_items(doc):

        if isinstance(block, Paragraph):
            folder_match = get_known_folder_match(
                block.text, known_folders
            )
            if folder_match:
                current_folder = folder_match

        elif isinstance(block, Table):
            current_folder = parse_table(
                block, extracted_data, known_folders, current_folder
            )

    return remove_duplicate_rows(extracted_data)


# =========================
# PDF PARSER
# =========================
def parse_pdf(pdf_path, known_folders=None):
    """
    PDF parser using pdfplumber.
    First tries structured table extraction per page,
    falls back to plain text line scanning.
    """
    extracted_data = {}
    current_folder = None

    if known_folders is None:
        known_folders = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            # Try structured table extraction first
            tables = page.extract_tables()

            if tables:
                inside_table = False
                name_idx = 0
                value_idx = 1

                for table in tables:
                    for row in table:
                        if not row:
                            continue

                        row_clean = [clean_text(cell) for cell in row]
                        row_text = " ".join(row_clean)

                        # Check folder heading
                        folder_match = get_known_folder_match(
                            row_text, known_folders
                        )
                        if folder_match:
                            current_folder = folder_match
                            inside_table = False
                            continue

                        # Check header row
                        normalized_cells = [normalize_header(c) for c in row_clean]
                        if "name" in normalized_cells and "value" in normalized_cells:
                            inside_table = True
                            name_idx = normalized_cells.index("name")
                            value_idx = normalized_cells.index("value")
                            continue

                        # Data row
                        if inside_table and current_folder:
                            if len(row_clean) > max(name_idx, value_idx):
                                name = row_clean[name_idx]
                                value = row_clean[value_idx]
                                add_row(extracted_data, current_folder, name, value)

            else:
                # Fallback: plain text extraction
                text = page.extract_text()
                if not text:
                    continue

                inside_table = False
                lines = text.split("\n")

                for line in lines:
                    line = clean_text(line)

                    folder_match = get_known_folder_match(line, known_folders)
                    if folder_match:
                        current_folder = folder_match
                        inside_table = False
                        continue

                    normalized_line = normalize_header(line)
                    if "name" in normalized_line and "value" in normalized_line:
                        inside_table = True
                        continue

                    if inside_table and current_folder:
                        parts = re.split(r"\s{2,}", line)
                        if len(parts) >= 2:
                            name = clean_text(parts[0])
                            value = clean_text(" ".join(parts[1:]))
                            add_row(extracted_data, current_folder, name, value)

    return remove_duplicate_rows(extracted_data)