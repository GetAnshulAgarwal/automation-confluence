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
# SAVE UPLOADED FOLDER FILES
# =========================
def save_uploaded_yaml_files(files, save_base_path):
    """
    Save browser-uploaded folder files preserving relative paths.

    Browser sends files with webkitRelativePath as filename, e.g.:
      app/realtimeapi/prd.values.yaml
      app/realtimeapi/uat.values.yaml
      app/predictiontimeapp/prd.values.yaml
      ...

    After save, disk structure becomes:
      yaml_folder/app/realtimeapi/prd.values.yaml
      yaml_folder/app/realtimeapi/uat.values.yaml
      yaml_folder/app/predictiontimeapp/prd.values.yaml
      ...
    """
    os.makedirs(save_base_path, exist_ok=True)

    for file in files:
        # file.filename contains the full relative path sent by browser
        relative_path = file.filename.replace("\\", "/")
        safe_parts = [
            part for part in relative_path.split("/")
            if part not in ["", ".", ".."]
        ]

        if not safe_parts:
            continue

        final_path = os.path.join(save_base_path, *safe_parts)
        os.makedirs(os.path.dirname(final_path), exist_ok=True)
        file.save(final_path)


# =========================
# GET SUBFOLDER NAMES
# =========================
def get_app_folder_names(yaml_base_path):
    """
    Returns list of subdirectory names directly inside yaml_base_path.
    These are the service folder names: realtimeapi, predictiontimeapp, etc.
    """
    folders = []

    if not os.path.exists(yaml_base_path):
        return folders

    for item in os.listdir(yaml_base_path):
        item_path = os.path.join(yaml_base_path, item)
        if os.path.isdir(item_path):
            folders.append(item)

    return folders


# =========================
# COMPARE API
# =========================
@app.route("/compare", methods=["POST"])
def compare():
    try:
        # Clean old uploads
        if os.path.exists(UPLOAD_FOLDER):
            shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        # Read uploaded files
        document_file = request.files.get("document")
        yaml_files = request.files.getlist("yamlFiles")

        if not document_file:
            return jsonify({"error": "PDF or DOCX file is required"}), 400

        if not yaml_files:
            return jsonify({"error": "App folder is required"}), 400

        # -------------------------
        # Save the document file
        # -------------------------
        document_path = os.path.join(UPLOAD_FOLDER, document_file.filename)
        document_file.save(document_path)

        # -------------------------
        # Save all yaml folder files
        # Disk result:
        #   uploads/yaml_folder/app/realtimeapi/prd.values.yaml
        #   uploads/yaml_folder/app/realtimeapi/uat.values.yaml
        #   uploads/yaml_folder/app/predictiontimeapp/prd.values.yaml
        #   ...
        # -------------------------
        yaml_upload_path = os.path.join(UPLOAD_FOLDER, "yaml_folder")
        save_uploaded_yaml_files(yaml_files, yaml_upload_path)

        # -------------------------
        # yaml_base_path = uploads/yaml_folder/app/
        # This is the folder that directly contains service subfolders.
        # Structure is always: app/ --> realtimeapi/, predictiontimeapp/, ...
        # -------------------------
        yaml_base_path = os.path.join(yaml_upload_path, "app")

        if not os.path.exists(yaml_base_path):
            # Safety fallback: if "app" subfolder not found, log and use root
            print(f"WARNING: 'app' folder not found inside upload. Using root: {yaml_upload_path}")
            print(f"Files saved under: {yaml_upload_path}")
            print(f"Contents: {os.listdir(yaml_upload_path)}")
            yaml_base_path = yaml_upload_path

        # -------------------------
        # Get service folder names
        # e.g. ["realtimeapi", "predictiontimeapp", "StopPrediction", ...]
        # -------------------------
        known_folders = get_app_folder_names(yaml_base_path)

        print(f"\nYAML BASE PATH : {yaml_base_path}")
        print(f"KNOWN FOLDERS  : {known_folders}\n")

        if not known_folders:
            return jsonify({
                "error": (
                    "No service folders found inside the uploaded 'app' folder. "
                    "Expected structure: app/realtimeapi/prd.values.yaml, etc."
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
                    "Check that folder names in the document match the uploaded app folder."
                )
            }), 400

        # -------------------------
        # Compare against YAML files
        # -------------------------
        prd_rows, uat_rows = compare_data(extracted_data, yaml_base_path)

        prd_mismatch = sum(1 for r in prd_rows if r["Status"] in ("NOT_FOUND", "VALUE_MISMATCH"))
        uat_mismatch  = sum(1 for r in uat_rows  if r["Status"] in ("NOT_FOUND", "VALUE_MISMATCH"))
        print(f"\nprd mismatches: {prd_mismatch}")
        print(f"UAT  mismatches: {uat_mismatch}")

        # -------------------------
        # Generate and return Excel
        # -------------------------
        output_excel = os.path.join(UPLOAD_FOLDER, "comparison_result.xlsx")
        generate_excel(prd_rows, uat_rows, output_excel)

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
    app.run(debug=True)