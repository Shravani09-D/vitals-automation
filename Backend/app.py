import os
import traceback
import tempfile
from uuid import uuid4
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

from vitals_automation import process_file

app = Flask(__name__)
CORS(app)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"docx"}
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def home():
    return jsonify({"message": "Backend running"}), 200


import tempfile

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["file"]

        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400

        if not allowed_file(file.filename):
            return jsonify({"error": "Only DOCX files are allowed"}), 400

        original_filename = secure_filename(file.filename)
        base_name = os.path.splitext(original_filename)[0]

        # create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_input:
            file.save(temp_input.name)
            input_path = temp_input.name

        # output file (same as before)
        output_filename = f"{base_name}_output.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        # remove old output if exists
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except PermissionError:
                return jsonify({
                    "error": f"Close the file before processing: {output_filename}"
                }), 400

        # process file
        process_file(input_path, output_path)

        # delete temp file AFTER processing
        try:
            os.remove(input_path)
        except:
            pass

        if not os.path.exists(output_path):
            return jsonify({"error": "Output file was not created"}), 500

        return jsonify({
            "message": "File processed successfully",
            "output_file": output_filename,
            "download_url": f"http://127.0.0.1:5000/download/{output_filename}"
        }), 200

    except Exception as e:
        print("=== BACKEND ERROR ===")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/download/<path:filename>")
def download_file(filename):
    try:
        return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 404


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)