from flask import Flask, render_template, request
import os
import subprocess
import sys
from flask import send_from_directory

app = Flask(__name__)


UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/download/<filename>')
def download_file(filename):

    full_path = os.path.join("output", filename)

    print("DOWNLOAD REQUEST:", filename)
    print("FILE EXISTS:", os.path.exists(full_path))
    print("FULL PATH:", os.path.abspath(full_path))

    return send_from_directory(
        "output",
        filename,
        as_attachment=True
    )

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/run", methods=["POST"])
def run_all():

    if "excel_file" not in request.files:
        return "No file uploaded"

    file = request.files["excel_file"]

    if file.filename == "":
        return "No file selected"

    file_path = os.path.join(
        UPLOAD_FOLDER,
        file.filename
    )

    file.save(file_path)

    try:

        # Run PPT1
        result1 = subprocess.run(
            [sys.executable, "ppt1.py", file_path],
            capture_output=True,
            text=True
        )

        if result1.returncode != 0:
            return f"<pre>PPT1 Error\n\n{result1.stderr}</pre>"

        # Run PPT2
        result2 = subprocess.run(
            [sys.executable, "ppt2.py", file_path],
            capture_output=True,
            text=True
        )

        if result2.returncode != 0:
            return f"<pre>PPT2 Error\n\n{result2.stderr}</pre>"

        # Run PPT3
        result3 = subprocess.run(
            [sys.executable, "ppt3.py", file_path],
            capture_output=True,
            text=True
        )

        if result3.returncode != 0:
            return f"<pre>PPT3 Error\n\n{result3.stderr}</pre>"

        
        print("OUTPUT FOLDER CONTENTS:")
        print(os.listdir("output"))

        return """
        <h2>Reports Generated Successfully</h2>

        <br>

        <a href="/download/ERC_Defect_Analysis_RITES_FINAL.pptx">
        Download PPT1
        </a>

        <br><br>

        <a href="/download/Vendor_PO_Quality_Analysis.pptx">
        Download PPT2
        </a>

        <br><br>

        <a href="/download/Defect_Percentage_Pie_RITES_Final.pptx">
        Download PPT3
        </a>
        """

    except Exception as e:
        return str(e)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
    