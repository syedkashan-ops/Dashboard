import streamlit as st
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

st.set_page_config(
    page_title="Sales Dashboard Generator",
    page_icon="📊",
    layout="wide"
)

INPUT_FIELDS = [
    ("Lubricant_Product_Liters.xlsx", "Lubricant Product Liters"),
    ("Sale data FY 25-26.xlsx", "Last Year Sales"),
    ("TM Name.xlsx", "TM Name"),
    ("Source_Cost_Center_Data.xlsx", "Source Cost Center Data"),
    ("Outlet_Master.xlsx", "Outlet Master"),
    ("Open_Orders.xlsx", "Open Orders"),
    ("Area Objectives.xlsx", "Area Objectives"),
    ("Sale data FY 26-27.xlsx", "This Year Sales"),
    ("Transit.xlsx", "Transit"),
    ("Outlet Objectives.xlsx", "Monthly Objectives"),
]

st.title("📊 Sales Dashboard Generator")

st.info(
    "Upload daily SAP sales files, optionally replace any input workbook, "
    "then generate and download Result.xlsx."
)

st.subheader("Step 1 — Daily SAP Input")

daily_files = st.file_uploader(
    "Upload one or more daily SAP sales files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

st.subheader("Step 2 — Optional Replacement Input Files")

uploaded_references = {}

for filename, label in INPUT_FIELDS:

    uploaded_references[filename] = st.file_uploader(
        f"{label} (Optional)",
        type=["xlsx", "xlsm"],
        key=filename
    )

st.subheader("Step 3 — Generate Dashboard")

if st.button("⚡ Generate Result.xlsx", use_container_width=True):

    if not daily_files:
        st.error("Please upload at least one daily SAP Excel file.")
        st.stop()

    with st.spinner("Generating dashboard... Please wait."):

        job_root = Path(tempfile.gettempdir()) / "sales_dashboard_jobs"
        job_root.mkdir(parents=True, exist_ok=True)

        job_dir = job_root / uuid.uuid4().hex
        job_dir.mkdir()

        try:

            # Copy Python files
            python_files = [
                "config.py",
                "processor.py",
                "excel_writer.py",
                "incremental_merger.py",
                "dashboard.py"
            ]

            for filename in python_files:
                shutil.copy2(
                    BASE_DIR / filename,
                    job_dir / filename
                )

            # Copy default Input folder
            shutil.copytree(
                BASE_DIR / "Input",
                job_dir / "Input"
            )

            # Create required folders
            (job_dir / "Daily_Input").mkdir(
                parents=True,
                exist_ok=True
            )

            (job_dir / "Output").mkdir(
                parents=True,
                exist_ok=True
            )

            (job_dir / "Backup").mkdir(
                parents=True,
                exist_ok=True
            )

            # Save daily SAP files
            for uploaded_file in daily_files:

                file_path = (
                    job_dir
                    / "Daily_Input"
                    / uploaded_file.name
                )

                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

            # Replace optional input files
            for expected_name, uploaded_file in uploaded_references.items():

                if uploaded_file is not None:

                    file_path = (
                        job_dir
                        / "Input"
                        / expected_name
                    )

                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

            # Run dashboard.py
            completed = subprocess.run(
                [sys.executable, "dashboard.py"],
                cwd=job_dir,
                capture_output=True,
                text=True,
                timeout=900
            )

            result_file = (
                job_dir
                / "Output"
                / "Result.xlsx"
            )

            if (
                completed.returncode != 0
                or not result_file.exists()
            ):

                error_details = (
                    completed.stdout
                    + "\n"
                    + completed.stderr
                )

                st.error("Dashboard generation failed.")

                st.code(error_details[-8000:])

                st.stop()

            result_data = result_file.read_bytes()

            st.success(
                "Dashboard generated successfully!"
            )

            st.download_button(
                label="⬇️ Download Result.xlsx",
                data=result_data,
                file_name="Result.xlsx",
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except subprocess.TimeoutExpired:

            st.error(
                "Dashboard generation timed out."
            )

        except Exception as e:

            st.error(
                f"Dashboard generation failed: {e}"
            )

        finally:

            shutil.rmtree(
                job_dir,
                ignore_errors=True
            )
