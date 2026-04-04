import os
import traceback
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
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def home():
    return jsonify({"message": "Backend running"}), 200


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

        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)

        file.save(input_path)

        output_filename = filename.rsplit(".", 1)[0] + "_output.docx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        print("=== FILE RECEIVED ===", filename)
        print("=== INPUT PATH ===", input_path)
        print("=== OUTPUT PATH ===", output_path)

        process_file(input_path, output_path)

        download_url = request.host_url.rstrip("/") + "/download/" + output_filename

        return jsonify({
            "message": "File processed successfully",
            "output_file": output_filename,
            "download_url": download_url
        }), 200

    except Exception as e:
        print("=== BACKEND ERROR ===")
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/download/<path:filename>")
def download_file(filename):
    try:
        return send_from_directory(
            OUTPUT_FOLDER,
            filename,
            as_attachment=True
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 404


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)