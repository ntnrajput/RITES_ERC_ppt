from flask import Flask, render_template, request
import os
import subprocess

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
            ["python", "ppt1.py", file_path],
            capture_output=True,
            text=True
        )

        if result1.returncode != 0:
            return f"<pre>PPT1 Error\n\n{result1.stderr}</pre>"

        # Run PPT2
        result2 = subprocess.run(
            ["python", "ppt2.py", file_path],
            capture_output=True,
            text=True
        )

        if result2.returncode != 0:
            return f"<pre>PPT2 Error\n\n{result2.stderr}</pre>"

        # Run PPT3
        result3 = subprocess.run(
            ["python", "ppt3.py", file_path],
            capture_output=True,
            text=True
        )

        if result3.returncode != 0:
            return f"<pre>PPT3 Error\n\n{result3.stderr}</pre>"

        return """
        <h2>Success</h2>
        <p>PPT1 Generated</p>
        <p>PPT2 Generated</p>
        <p>PPT3 Generated</p>
        <a href="/">Back</a>
        """

    except Exception as e:
        return str(e)

if __name__ == "__main__":
    app.run(debug=True)
    