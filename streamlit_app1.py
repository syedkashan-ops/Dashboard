import base64
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

import requests
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent

# ---------------------------------------------------------
# GitHub repository settings
# ---------------------------------------------------------

GITHUB_OWNER = "syedkashan-ops"
GITHUB_REPO = "Dashboard"
GITHUB_BRANCH = "main"

MASTER_FILE_NAME = "Sale data FY 26-27.xlsx"
MASTER_FILE_PATH = f"Input/{MASTER_FILE_NAME}"


# ---------------------------------------------------------
# Streamlit page
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# GitHub helper functions
# ---------------------------------------------------------

def get_github_token():
    """
    Read GitHub token securely from Streamlit Secrets.
    """

    try:
        return st.secrets["github"]["token"]
    except Exception:
        return None


def upload_master_to_github(master_file: Path):
    """
    Update Input/Sale data FY 26-27.xlsx
    in the GitHub repository.
    """

    token = get_github_token()

    if not token:
        raise RuntimeError(
            "GitHub token was not found in Streamlit Secrets."
        )

    api_url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/contents/"
        f"{MASTER_FILE_PATH}"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    # -----------------------------------------------------
    # Get current SHA of master file
    # -----------------------------------------------------

    response = requests.get(
        api_url,
        headers=headers,
        params={"ref": GITHUB_BRANCH},
        timeout=60,
    )

    if response.status_code != 200:
        raise RuntimeError(
            "Could not read the current master file "
            f"from GitHub.\n"
            f"Status: {response.status_code}\n"
            f"Details: {response.text}"
        )

    current_file = response.json()

    current_sha = current_file["sha"]

    # -----------------------------------------------------
    # Read updated master file
    # -----------------------------------------------------

    file_content = master_file.read_bytes()

    encoded_content = base64.b64encode(
        file_content
    ).decode("utf-8")

    # -----------------------------------------------------
    # Update file in GitHub
    # -----------------------------------------------------

    payload = {
        "message": (
            "Automatic update: "
            "merge new daily SAP sales data"
        ),
        "content": encoded_content,
        "sha": current_sha,
        "branch": GITHUB_BRANCH,
    }

    response = requests.put(
        api_url,
        headers=headers,
        json=payload,
        timeout=300,
    )

    if response.status_code not in (200, 201):

        raise RuntimeError(
            "Could not update the master sales file "
            f"in GitHub.\n"
            f"Status: {response.status_code}\n"
            f"Details: {response.text}"
        )

    return response.json()


# ---------------------------------------------------------
# User interface
# ---------------------------------------------------------

st.title("📊 Sales Dashboard Generator")

st.info(
    "Upload only new SAP sales data. "
    "The system will merge new records into the master "
    "sales file, skip duplicates, generate the dashboard, "
    "and save the updated master for the next run."
)


# ---------------------------------------------------------
# Step 1
# ---------------------------------------------------------

st.subheader("Step 1 — Daily SAP Input")

daily_files = st.file_uploader(
    "Upload new daily SAP sales file",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)


# ---------------------------------------------------------
# Step 2
# ---------------------------------------------------------

st.subheader(
    "Step 2 — Optional Replacement Input Files"
)

uploaded_references = {}

for filename, label in INPUT_FIELDS:

    # Do not allow replacing the master through optional input.
    # It is updated automatically through the daily merge.
    if filename == MASTER_FILE_NAME:
        continue

    uploaded_references[filename] = st.file_uploader(
        f"{label} (Optional)",
        type=["xlsx", "xlsm"],
        key=filename,
    )


# ---------------------------------------------------------
# Step 3
# ---------------------------------------------------------

st.subheader("Step 3 — Generate Dashboard")


if st.button(
    "⚡ Generate Result.xlsx",
    use_container_width=True,
):

    if not daily_files:

        st.error(
            "Please upload at least one daily SAP Excel file."
        )

        st.stop()


    # -----------------------------------------------------
    # Check GitHub token before starting long processing
    # -----------------------------------------------------

    github_token = get_github_token()

    if not github_token:

        st.error(
            "GitHub token is missing. "
            "Please add it in Streamlit Secrets."
        )

        st.stop()


    with st.spinner(
        "Generating dashboard... Please wait. "
        "This may take several minutes."
    ):

        job_root = (
            Path(tempfile.gettempdir())
            / "sales_dashboard_jobs"
        )

        job_root.mkdir(
            parents=True,
            exist_ok=True,
        )

        job_dir = (
            job_root
            / uuid.uuid4().hex
        )

        job_dir.mkdir()


        try:

            # -------------------------------------------------
            # Copy Python files
            # -------------------------------------------------

            python_files = [
                "config.py",
                "processor.py",
                "excel_writer.py",
                "incremental_merger.py",
                "dashboard.py",
            ]

            for filename in python_files:

                shutil.copy2(
                    BASE_DIR / filename,
                    job_dir / filename,
                )


            # -------------------------------------------------
            # Copy Input folder
            # -------------------------------------------------

            shutil.copytree(
                BASE_DIR / "Input",
                job_dir / "Input",
            )


            # -------------------------------------------------
            # Create required folders
            # -------------------------------------------------

            (
                job_dir / "Daily_Input"
            ).mkdir(
                parents=True,
                exist_ok=True,
            )

            (
                job_dir / "Output"
            ).mkdir(
                parents=True,
                exist_ok=True,
            )

            (
                job_dir / "Backup"
            ).mkdir(
                parents=True,
                exist_ok=True,
            )


            # -------------------------------------------------
            # Save uploaded daily SAP files
            # -------------------------------------------------

            for uploaded_file in daily_files:

                file_path = (
                    job_dir
                    / "Daily_Input"
                    / uploaded_file.name
                )

                with open(
                    file_path,
                    "wb",
                ) as f:

                    f.write(
                        uploaded_file.getbuffer()
                    )


            # -------------------------------------------------
            # Replace optional input files
            # -------------------------------------------------

            for (
                expected_name,
                uploaded_file
            ) in uploaded_references.items():

                if uploaded_file is not None:

                    file_path = (
                        job_dir
                        / "Input"
                        / expected_name
                    )

                    with open(
                        file_path,
                        "wb",
                    ) as f:

                        f.write(
                            uploaded_file.getbuffer()
                        )


            # -------------------------------------------------
            # Run dashboard
            # -------------------------------------------------

            completed = subprocess.run(
                [
                    sys.executable,
                    "dashboard.py",
                ],
                cwd=job_dir,
                capture_output=True,
                text=True,
                timeout=900,
            )


            result_file = (
                job_dir
                / "Output"
                / "Result.xlsx"
            )


            # -------------------------------------------------
            # Check dashboard result
            # -------------------------------------------------

            if (
                completed.returncode != 0
                or not result_file.exists()
            ):

                error_details = (
                    completed.stdout
                    + "\n"
                    + completed.stderr
                )

                st.error(
                    "Dashboard generation failed."
                )

                st.code(
                    error_details[-8000:]
                )

                st.stop()


            # -------------------------------------------------
            # IMPORTANT:
            # Upload updated master to GitHub
            # -------------------------------------------------

            updated_master_file = (
                job_dir
                / "Input"
                / MASTER_FILE_NAME
            )


            if not updated_master_file.exists():

                raise RuntimeError(
                    "Updated master sales file was not found."
                )


            with st.spinner(
                "Saving updated sales master to GitHub..."
            ):

                upload_master_to_github(
                    updated_master_file
                )


            # -------------------------------------------------
            # Read Result.xlsx
            # -------------------------------------------------

            result_data = (
                result_file.read_bytes()
            )


            st.success(
                "Dashboard generated successfully! "
                "The updated sales master has also been "
                "saved to GitHub."
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
                use_container_width=True,
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
                ignore_errors=True,
            )
