from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import shutil

from parser_utils import parse_pdf, parse_docx
from comparator import compare_data, generate_excel

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================
# GET SUBFOLDER NAMES
# =========================
def get_app_folder_names(yaml_base_path):
    folders = []
    if not os.path.exists(yaml_base_path):
        return folders
    for item in os.listdir(yaml_base_path):
        item_path = os.path.join(yaml_base_path, item)
        if os.path.isdir(item_path):
            folders.append(item)
    return folders


# =========================
# GET FOLDERS API
# Returns list of subfolders inside the app folder
# =========================
@app.route("/get-folders", methods=["POST"])
def get_folders():
    try:
        data = request.get_json()
        app_folder_path = (data or {}).get("appFolderPath", "").strip().strip('"').strip("'").strip()

        if not app_folder_path:
            return jsonify({"error": "appFolderPath is required"}), 400

        if not os.path.isdir(app_folder_path):
            return jsonify({"error": f"Folder not found: {app_folder_path}"}), 400

        folders = sorted([
            item for item in os.listdir(app_folder_path)
            if os.path.isdir(os.path.join(app_folder_path, item))
        ])

        return jsonify({"folders": folders})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# =========================
# COMPARE API
# Uses local file paths — no upload size limit
# =========================
@app.route("/compare", methods=["POST"])
def compare():
    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "JSON body required"}), 400

        # Strip whitespace and surrounding quotes
        # (users sometimes copy paths with quotes e.g. "C:\path\file.docx")
        document_path   = data.get("documentPath", "").strip().strip('"').strip("'").strip()
        app_folder_path = data.get("appFolderPath", "").strip().strip('"').strip("'").strip()
        # selectedFolders: list of folder names chosen by user in UI
        # If empty/not provided -> use ALL folders
        selected_folders = data.get("selectedFolders", [])

        # -------------------------
        # Validate inputs
        # -------------------------
        if not document_path:
            return jsonify({"error": "documentPath is required"}), 400

        if not app_folder_path:
            return jsonify({"error": "appFolderPath is required"}), 400

        if not os.path.isfile(document_path):
            return jsonify({"error": f"Document file not found: {document_path}"}), 400

        if not os.path.isdir(app_folder_path):
            return jsonify({"error": f"App folder not found: {app_folder_path}"}), 400

        # -------------------------
        # Resolve yaml_base_path and known_folders
        # -------------------------
        yaml_base_path = app_folder_path

        all_folders = get_app_folder_names(yaml_base_path)

        # If user selected specific folders, use only those
        # Otherwise use all folders
        if selected_folders:
            known_folders = [f for f in all_folders if f in selected_folders]
        else:
            known_folders = all_folders

        print(f"\nDOCUMENT       : {document_path}")
        print(f"APP FOLDER     : {yaml_base_path}")
        print(f"KNOWN FOLDERS  : {known_folders}\n")

        if not known_folders:
            return jsonify({
                "error": (
                    "No service folders found inside the app folder. "
                    "Expected: app/realtimeapi/prd.values.yaml, etc."
                )
            }), 400

        # -------------------------
        # Parse DOCX or PDF
        # -------------------------
        doc_lower = document_path.lower()

        if doc_lower.endswith(".pdf"):
            extracted_data = parse_pdf(document_path, known_folders)

        elif doc_lower.endswith(".docx"):
            extracted_data = parse_docx(document_path, known_folders)

        else:
            return jsonify({"error": "Only PDF or DOCX files are allowed"}), 400

        print("EXTRACTED DATA:")
        for folder, rows in extracted_data.items():
            print(f"  {folder}: {len(rows)} rows")

        if not extracted_data:
            return jsonify({
                "error": (
                    "No data could be extracted from the document. "
                    "Check that folder names in the document match the app folder."
                )
            }), 400

        # -------------------------
        # Compare against YAML
        # -------------------------
        prod_rows, uat_rows = compare_data(extracted_data, yaml_base_path)

        prod_mis = sum(1 for r in prod_rows if r["Status"] in ("NOT_FOUND", "VALUE_MISMATCH", "FOLDER_NOT_PRESENT"))
        uat_mis  = sum(1 for r in uat_rows  if r["Status"] in ("NOT_FOUND", "VALUE_MISMATCH", "FOLDER_NOT_PRESENT"))
        print(f"\nPROD mismatches: {prod_mis}")
        print(f"UAT  mismatches: {uat_mis}")

        # -------------------------
        # Generate Excel
        # -------------------------
        if os.path.exists(UPLOAD_FOLDER):
            shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        output_excel = os.path.join(UPLOAD_FOLDER, "comparison_result.xlsx")
        generate_excel(prod_rows, uat_rows, output_excel)

        return send_file(
            output_excel,
            as_attachment=True,
            download_name="comparison_result.xlsx"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, threaded=True)