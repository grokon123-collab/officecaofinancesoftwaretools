from flask import Flask, render_template, request, jsonify, send_file, after_this_request
import subprocess, threading, os, json, shutil

app = Flask(__name__)

LOG_FILE = "script_output.log"
DOWNLOAD_DIR = "downloads"
RESULT_FILE = os.path.join(DOWNLOAD_DIR, "UofT_Staff_Report.xlsx")


def run_script_background(selected_departments):
    """Run crawler_script.py with selected departments asynchronously."""
    # Ensure download folder exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Clear old logs before each run
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)

    # Run the crawler in a subprocess
    with open(LOG_FILE, "w") as log:
        process = subprocess.Popen(
            ["python", "crawler_script.py", json.dumps(selected_departments)],
            stdout=log,
            stderr=subprocess.STDOUT,
            text=True
        )
        process.wait()


@app.route("/")
def index():
    departments = [
    'Department of Radiation Oncology',
    'Department of Speech - Language Pathology',
    'Department of Medical Biophysics',
    'Department of Obstetrics And Gynaecology',
    'Laboratory Medicine And Pathobiology',
    'Department of Pharmacology',
    'Discovery Commons',
    'Graduate Department of Rehabilitation Science',
    'Department of Occupational Science & Therapy',
    'Department of Biochemistry',
    'Division of Comparative Medicine',
    'Med: Office of The Dean',
    'Department of Nutritional Sciences',
    'Banting & Best Diabetes Centre',
    'Division of Teaching Laboratories',
    'Division of (Department of Surgery) Anatomy',
    'Department of Otolaryngology - Head & Neck Surgery',
    'Postgraduate Medical Education',
    'Standardized Patient Program',
    'History of Medicine Program',
    'Department of Anesthesiology & Pain Medicine',
    'Faculty of Medicine',
    'Med Store',
    'Department of Physiology',
    'Rehabilitation Sciences Sector',
    'Terrence Donnelly Centre for Cellular and Biomolecular Research',
    'Molecular Genetics',
    'Level 3 Facility',
    'Department of Ophthalmology & Vision Sciences',
    'Department of Immunology',
    'Department of Medical Imaging',
    'Department of Physical Therapy',
    'Structural Genomics Consortium',
    ]
    return render_template("index.html", departments=departments)


@app.route("/run-script", methods=["POST"])
def run_script():
    data = request.get_json()
    selected = data.get("departments", [])

    thread = threading.Thread(target=run_script_background, args=(selected,), daemon=True)
    thread.start()
    return jsonify({"status": "Crawler started."})


@app.route("/is-done")
def is_done():
    """Check if crawler has finished and file exists."""
    return jsonify({"done": os.path.exists(RESULT_FILE)})


@app.route("/download-result")
def download_result():
    """Download the Excel result and clean up temporary files afterward."""
    if not os.path.exists(RESULT_FILE):
        return "No report available."

    @after_this_request
    def cleanup(response):
        """Remove downloads folder and log after sending the file."""
        try:
            if os.path.exists(DOWNLOAD_DIR):
                shutil.rmtree(DOWNLOAD_DIR)
                os.makedirs(DOWNLOAD_DIR, exist_ok=True)
            if os.path.exists(LOG_FILE):
                os.remove(LOG_FILE)
            print("✅ Cleaned up downloads and logs.")
        except Exception as e:
            print(f"⚠️ Cleanup error: {e}")
        return response

    return send_file(RESULT_FILE, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)

