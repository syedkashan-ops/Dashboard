from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

from flask import Flask, Response, render_template_string, request, send_file
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
EXPECTED_INPUT_FILES = {
    "Lubricant_Product_Liters.xlsx",
    "Sale data FY 25-26.xlsx",
    "TM Name.xlsx",
    "Source_Cost_Center_Data.xlsx",
    "Outlet_Master.xlsx",
    "Open_Orders.xlsx",
    "Area Objectives.xlsx",
    "Sale data FY 26-27.xlsx",
    "Transit.xlsx",
    "Outlet Objectives.xlsx",
}

INPUT_FIELDS = [
    ("input_lubricant", "Lubricant_Product_Liters.xlsx", "Lubricant Product Liters"),
    ("input_sales_2526", "Sale data FY 25-26.xlsx", "Last Year Sales"),
    ("input_tm_name", "TM Name.xlsx", "TM Name"),
    ("input_cost_center", "Source_Cost_Center_Data.xlsx", "Source Cost Center Data"),
    ("input_outlet_master", "Outlet_Master.xlsx", "Outlet Master"),
    ("input_open_orders", "Open_Orders.xlsx", "Open Orders"),
    ("input_area_objectives", "Area Objectives.xlsx", "Area Objectives"),
    ("input_sales_2627", "Sale data FY 26-27.xlsx", "This Year Sales"),
    ("input_transit", "Transit.xlsx", "Transit"),
    ("input_outlet_objectives", "Outlet Objectives.xlsx", "Monthly Objectives"),
]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024  # 250 MB total request limit

PAGE = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Sales Dashboard Generator</title>
  <style>
    * { box-sizing: border-box; }
    :root { font-family: Inter, "Segoe UI", Arial, sans-serif; color:#18233a; background:#eef4fb; }
    body { margin:0; min-height:100vh; background:linear-gradient(135deg,#eaf4ff 0%,#f8f2ff 52%,#effcf6 100%); }
    .topbar { background:linear-gradient(90deg,#173b70,#3b5fc0,#6b4fc8); color:#fff; padding:25px 20px; box-shadow:0 6px 20px rgba(32,54,110,.18); }
    .topbar-inner { max-width:1120px; margin:auto; display:flex; align-items:center; gap:15px; }
    .logo { width:52px; height:52px; border-radius:15px; display:grid; place-items:center; background:rgba(255,255,255,.18); font-size:27px; }
    .topbar h1 { margin:0; font-size:27px; }
    .topbar p { margin:5px 0 0; opacity:.86; }
    main { max-width:1120px; margin:26px auto 55px; padding:0 18px; }
    .intro { display:flex; justify-content:space-between; gap:14px; align-items:center; background:#fff; border:1px solid #dfe8f4; border-radius:18px; padding:18px 20px; box-shadow:0 8px 26px rgba(31,58,100,.07); }
    .intro strong { display:block; font-size:17px; margin-bottom:4px; }
    .badge { white-space:nowrap; border-radius:999px; background:#e9f8ef; color:#19734a; padding:8px 12px; font-weight:700; font-size:13px; }
    .section { background:#fff; border:1px solid #dfe8f4; border-radius:20px; padding:22px; margin-top:20px; box-shadow:0 8px 26px rgba(31,58,100,.07); }
    .section-head { display:flex; align-items:center; gap:11px; margin-bottom:7px; }
    .num { width:34px; height:34px; border-radius:11px; display:grid; place-items:center; color:#fff; font-weight:800; background:linear-gradient(135deg,#315dcc,#7354cf); }
    h2 { margin:0; font-size:20px; }
    .hint { color:#68758b; margin:0 0 18px 45px; line-height:1.45; font-size:14px; }
    .daily { border:2px dashed #9eb8e8; border-radius:16px; padding:18px; background:#f7faff; }
    .daily input { width:100%; }
    .file-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:13px; }
    .file-card { border:1px solid #dfe7f3; border-radius:15px; padding:14px; background:linear-gradient(180deg,#fff,#f9fbff); transition:.18s ease; }
    .file-card:hover { transform:translateY(-1px); border-color:#a9bde4; box-shadow:0 7px 18px rgba(40,67,120,.08); }
    .file-title { font-weight:750; font-size:14px; margin-bottom:5px; word-break:break-word; }
    .file-sub { color:#76839a; font-size:12px; margin-bottom:10px; }
    input[type=file] { width:100%; color:#526078; font-size:13px; }
    input[type=file]::file-selector-button { border:0; padding:8px 11px; border-radius:8px; background:#e9efff; color:#294f9f; font-weight:700; cursor:pointer; margin-right:8px; }
    .action { margin-top:21px; display:flex; gap:12px; align-items:center; }
    button { flex:1; border:0; border-radius:14px; padding:15px 20px; font-size:16px; font-weight:800; cursor:pointer; color:#fff; background:linear-gradient(90deg,#245bc4,#624fc7); box-shadow:0 9px 20px rgba(50,77,177,.2); }
    button:hover { filter:brightness(1.04); }
    .privacy { color:#657389; font-size:13px; }
    .footer-note { margin-top:14px; background:#edf8f2; color:#276247; border:1px solid #cdebdc; border-radius:13px; padding:12px 14px; font-size:13px; }
    #loading { display:none; position:fixed; inset:0; background:rgba(20,31,54,.68); z-index:50; place-items:center; backdrop-filter:blur(3px); }
    #loading.show { display:grid; }
    .loading-card { width:min(430px,90vw); background:#fff; border-radius:20px; padding:28px; text-align:center; box-shadow:0 18px 60px rgba(0,0,0,.25); }
    .spinner { width:48px; height:48px; margin:0 auto 18px; border:5px solid #e1e8f5; border-top-color:#4d5fd0; border-radius:50%; animation:spin .9s linear infinite; }
    @keyframes spin { to { transform:rotate(360deg); } }
    .loading-card h3 { margin:0 0 7px; }
    .loading-card p { margin:0; color:#6f7b90; line-height:1.5; }
    @media(max-width:760px){ .file-grid{grid-template-columns:1fr;} .intro{align-items:flex-start;flex-direction:column;} .hint{margin-left:0;} }
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-inner">
      <div class="logo">📊</div>
      <div><h1>Sales Dashboard Generator</h1><p>Upload input workbooks → generate dashboard → download Result.xlsx</p></div>
    </div>
  </header>

  <main>
    <div class="intro">
      <div><strong>Private processing for each request</strong><span style="color:#6c7890;font-size:14px">Your uploaded files are processed in a separate temporary workspace.</span></div>
      <div class="badge">● Server Ready</div>
    </div>

    <form id="dashboardForm" action="/generate" method="post" enctype="multipart/form-data">
      <section class="section">
        <div class="section-head"><div class="num">1</div><h2>Daily Input</h2></div>
        <p class="hint">Upload one or more daily SAP sales exports. This is separate from the 10 master/input workbooks below.</p>
        <div class="daily"><input type="file" name="daily_files" accept=".xlsx,.xlsm" multiple required></div>
      </section>

      <section class="section">
        <div class="section-head"><div class="num">2</div><h2>10 Input Workbooks</h2></div>
        <p class="hint">Each input has its own named upload option. If you leave one blank, the packaged/default version on the server will be used.</p>
        <div class="file-grid">
          {% for field, name, label in input_fields %}
          <div class="file-card">
            <div class="file-title">{{ loop.index }}. {{ label }}</div>
            <div class="file-sub">File used by system: {{ name }} · Upload replacement (optional)</div>
            <input type="file" name="{{ field }}" accept=".xlsx,.xlsm">
          </div>
          {% endfor %}
        </div>
      </section>

      <section class="section">
        <div class="section-head"><div class="num">3</div><h2>Generate Result</h2></div>
        <p class="hint">Click once. Keep this page open while the server reads, merges, builds and saves the workbook.</p>
        <div class="action"><button type="submit">⚡ Generate & Download Result.xlsx</button></div>
        <div class="footer-note">✓ No Python is required on the user's laptop after this application is deployed on the company server.</div>
      </section>
    </form>
  </main>

  <div id="loading">
    <div class="loading-card"><div class="spinner"></div><h3>Generating your dashboard…</h3><p>Uploading files and creating Result.xlsx. Large Excel files can take several minutes. Please keep this page open.</p></div>
  </div>
  <script>
    document.getElementById('dashboardForm').addEventListener('submit', function(){ document.getElementById('loading').classList.add('show'); });
    window.addEventListener('pageshow', function(){ document.getElementById('loading').classList.remove('show'); });
  </script>
</body>
</html>
"""


def _valid_excel(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _copy_project_for_job(job_dir: Path) -> None:
    # Copy only the Python application and default reference files needed by the job.
    for filename in ("config.py", "processor.py", "excel_writer.py", "incremental_merger.py", "dashboard.py"):
        shutil.copy2(BASE_DIR / filename, job_dir / filename)

    shutil.copytree(BASE_DIR / "Input", job_dir / "Input")
    (job_dir / "Daily_Input").mkdir(parents=True, exist_ok=True)
    (job_dir / "Output").mkdir(parents=True, exist_ok=True)
    (job_dir / "Backup").mkdir(parents=True, exist_ok=True)


@app.get("/")
def index() -> str:
    return render_template_string(PAGE, input_fields=INPUT_FIELDS)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/generate")
def generate():
    daily_files = [f for f in request.files.getlist("daily_files") if f and f.filename]
    reference_files = []
    for field, expected_name, label in INPUT_FIELDS:
        upload = request.files.get(field)
        if upload and upload.filename:
            reference_files.append((upload, expected_name))

    if not daily_files:
        return Response("Please upload at least one daily SAP Excel file.", status=400, mimetype="text/plain")

    bad_daily = [f.filename for f in daily_files if not _valid_excel(f.filename)]
    if bad_daily:
        return Response("Only .xlsx or .xlsm files are allowed.", status=400, mimetype="text/plain")

    job_root = Path(tempfile.gettempdir()) / "sales_dashboard_jobs"
    job_root.mkdir(parents=True, exist_ok=True)
    job_dir = job_root / uuid.uuid4().hex
    job_dir.mkdir()

    try:
        _copy_project_for_job(job_dir)

        # Save private daily sales uploads.
        for idx, upload in enumerate(daily_files, start=1):
            safe = secure_filename(upload.filename) or f"daily_{idx}.xlsx"
            if not _valid_excel(safe):
                raise ValueError("Daily upload must be .xlsx or .xlsm.")
            upload.save(job_dir / "Daily_Input" / safe)

        # Replace only recognized reference filenames. This prevents an upload from
        # overwriting Python/source files or escaping the Input directory.
        for upload, expected_name in reference_files:
            if not _valid_excel(upload.filename):
                raise ValueError(f"{expected_name} must be an .xlsx or .xlsm file.")
            # Save to the exact filename expected by the existing dashboard code,
            # so users do not have to manually rename their uploaded workbook.
            upload.save(job_dir / "Input" / expected_name)

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        completed = subprocess.run(
            [sys.executable, "dashboard.py"],
            cwd=job_dir,
            env=env,
            capture_output=True,
            text=True,
            timeout=900,
        )

        result = job_dir / "Output" / "Result.xlsx"
        if completed.returncode != 0 or not result.exists():
            details = (completed.stdout + "\n" + completed.stderr).strip()
            raise RuntimeError(details[-8000:] or "Dashboard generation failed.")

        # Read the result into memory before deleting the private workspace.
        data = result.read_bytes()
    except subprocess.TimeoutExpired:
        shutil.rmtree(job_dir, ignore_errors=True)
        return Response("Dashboard generation timed out. Please contact IT.", status=500, mimetype="text/plain")
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        return Response(f"Dashboard generation failed:\n\n{exc}", status=500, mimetype="text/plain")

    shutil.rmtree(job_dir, ignore_errors=True)
    return Response(
        data,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Result.xlsx"'},
    )


if __name__ == "__main__":
    # Development/test use only. IT should run run_web_server.py (Waitress) in production.
    app.run(host="0.0.0.0", port=8080, debug=False)
