import streamlit as st
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

st.set_page_config(
    page_title="Sales Dashboard Generator",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# INPUT FILES
# ============================================================

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


# ============================================================
# PAGE HEADER
# ============================================================

st.title("📊 Sales Dashboard Generator")

st.write(
    "Upload your daily SAP sales files and generate the updated "
    "Excel dashboard."
)


# ============================================================
# DAILY INPUT FILES
# ============================================================

st.subheader("Step 1 — Daily SAP Input Files")

daily_files = st.file_uploader(
    "Upload one or more daily SAP sales files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)


# ============================================================
# OPTIONAL INPUT FILES
# ============================================================

st.subheader("Step 2 — Replace Input Files (Optional)")

st.info(
    "If you do not upload a file below, the default file available "
    "in the Input folder will be used."
)

uploaded_references = {}

for filename, label in INPUT_FIELDS:

    uploaded_references[filename] = st.file_uploader(
        f"{label} (Optional)",
        type=["xlsx", "xlsm"],
        key=f"upload_{filename}"
    )


# ============================================================
# GENERATE BUTTON
# ============================================================

st.subheader("Step 3 — Generate Dashboard")

generate_button = st.button(
    "⚡ Generate Result.xlsx",
    use_container_width=True
)


# ============================================================
# DASHBOARD PROCESSING
# ============================================================

if generate_button:

    if not daily_files:

        st.error(
            "Please upload at least one Daily SAP Excel file."
        )

        st.stop()


    # Status area
    status = st.empty()

    # Progress bar
    progress_bar = st.progress(0)


    job_dir = None


    try:

        # ----------------------------------------------------
        # STEP 1 - CREATE TEMPORARY WORKSPACE
        # ----------------------------------------------------

        status.info(
            "Step 1/6: Preparing temporary workspace..."
        )

        progress_bar.progress(5)

        job_root = (
            Path(tempfile.gettempdir())
            / "sales_dashboard_jobs"
        )

        job_root.mkdir(
            parents=True,
            exist_ok=True
        )

        job_dir = (
            job_root
            / uuid.uuid4().hex
        )

        job_dir.mkdir(
            parents=True,
            exist_ok=True
        )


        # ----------------------------------------------------
        # STEP 2 - COPY APPLICATION FILES
        # ----------------------------------------------------

        status.info(
            "Step 2/6: Copying dashboard application files..."
        )

        progress_bar.progress(15)


        python_files = [
            "config.py",
            "processor.py",
            "excel_writer.py",
            "incremental_merger.py",
            "dashboard.py"
        ]


        for filename in python_files:

            source_file = (
                BASE_DIR / filename
            )

            destination_file = (
                job_dir / filename
            )

            if not source_file.exists():

                raise FileNotFoundError(
                    f"Required file not found: {filename}"
                )

            shutil.copy2(
                source_file,
                destination_file
            )


        # ----------------------------------------------------
        # STEP 3 - COPY DEFAULT INPUT FOLDER
        # ----------------------------------------------------

        status.info(
            "Step 3/6: Preparing default Excel input files..."
        )

        progress_bar.progress(25)


        source_input_dir = (
            BASE_DIR / "Input"
        )

        destination_input_dir = (
            job_dir / "Input"
        )


        if not source_input_dir.exists():

            raise FileNotFoundError(
                "Input folder was not found in the GitHub repository."
            )


        shutil.copytree(
            source_input_dir,
            destination_input_dir
        )


        # ----------------------------------------------------
        # CREATE REQUIRED FOLDERS
        # ----------------------------------------------------

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


        # ----------------------------------------------------
        # STEP 4 - SAVE UPLOADED DAILY FILES
        # ----------------------------------------------------

        status.info(
            "Step 4/6: Saving uploaded SAP sales files..."
        )

        progress_bar.progress(40)


        for uploaded_file in daily_files:

            file_path = (
                job_dir
                / "Daily_Input"
                / uploaded_file.name
            )


            with open(
                file_path,
                "wb"
            ) as f:

                f.write(
                    uploaded_file.getbuffer()
                )


        # ----------------------------------------------------
        # REPLACE OPTIONAL INPUT FILES
        # ----------------------------------------------------

        for expected_name, uploaded_file in (
            uploaded_references.items()
        ):

            if uploaded_file is not None:

                file_path = (
                    job_dir
                    / "Input"
                    / expected_name
                )


                with open(
                    file_path,
                    "wb"
                ) as f:

                    f.write(
                        uploaded_file.getbuffer()
                    )


        # ----------------------------------------------------
        # STEP 5 - RUN DASHBOARD PROCESS
        # ----------------------------------------------------

        status.warning(
            "Step 5/6: Processing Excel files and generating "
            "dashboard... This may take several minutes."
        )

        progress_bar.progress(60)


        completed = subprocess.run(

            [
                sys.executable,
                "dashboard.py"
            ],

            cwd=job_dir,

            capture_output=True,

            text=True,

            timeout=900

        )


        # ----------------------------------------------------
        # STEP 6 - CHECK RESULT FILE
        # ----------------------------------------------------

        status.info(
            "Step 6/6: Checking generated Result.xlsx..."
        )

        progress_bar.progress(90)


        result_file = (
            job_dir
            / "Output"
            / "Result.xlsx"
        )


        # ----------------------------------------------------
        # CHECK FOR PYTHON ERROR
        # ----------------------------------------------------

        if completed.returncode != 0:

            progress_bar.empty()

            status.empty()


            st.error(
                "❌ Dashboard generation failed."
            )


            st.subheader(
                "Error Details"
            )


            error_output = (
                "STANDARD OUTPUT:\n"
                + completed.stdout
                + "\n\n"
                + "ERROR OUTPUT:\n"
                + completed.stderr
            )


            st.code(
                error_output[-12000:]
            )


            st.stop()


        # ----------------------------------------------------
        # CHECK IF RESULT FILE EXISTS
        # ----------------------------------------------------

        if not result_file.exists():

            progress_bar.empty()

            status.empty()


            st.error(
                "❌ Processing completed, but Result.xlsx "
                "was not created."
            )


            st.subheader(
                "Application Output"
            )


            output_text = (
                "STANDARD OUTPUT:\n"
                + completed.stdout
                + "\n\n"
                + "ERROR OUTPUT:\n"
                + completed.stderr
            )


            st.code(
                output_text[-12000:]
            )


            st.stop()


        # ----------------------------------------------------
        # READ GENERATED EXCEL FILE
        # ----------------------------------------------------

        status.info(
            "Reading generated Result.xlsx..."
        )

        progress_bar.progress(95)


        result_data = (
            result_file.read_bytes()
        )


        # ----------------------------------------------------
        # SUCCESS
        # ----------------------------------------------------

        progress_bar.progress(100)


        status.success(
            "✅ Dashboard generated successfully!"
        )


        st.balloons()


        # ----------------------------------------------------
        # DOWNLOAD BUTTON
        # ----------------------------------------------------

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


    # ========================================================
    # TIMEOUT ERROR
    # ========================================================

    except subprocess.TimeoutExpired:

        progress_bar.empty()

        status.empty()


        st.error(
            "⏰ Dashboard generation exceeded the 15-minute "
            "limit and was stopped."
        )


        st.warning(
            "The processing is taking too long for the Streamlit "
            "Cloud server. The next step will be to optimize the "
            "Excel processing code."
        )


    # ========================================================
    # OTHER ERRORS
    # ========================================================

    except Exception as e:

        progress_bar.empty()

        status.empty()


        st.error(
            f"❌ Error occurred: {str(e)}"
        )


        st.exception(e)


    # ========================================================
    # CLEAN TEMPORARY FILES
    # ========================================================

    finally:

        if job_dir is not None:

            shutil.rmtree(
                job_dir,
                ignore_errors=True
            )
