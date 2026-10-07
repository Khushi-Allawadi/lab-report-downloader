import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
from pathlib import Path
from datetime import datetime
import tempfile
import zipfile
import shutil
import os
import re
import time


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Lab Report Downloader",
    page_icon="📄",
    layout="centered"
)


# ============================================================
# CONFIGURATION
# ============================================================

LABCONNECT_URL = "https://lms.mylabconnect.co.uk/lab-connect"

SEARCH_SEL = "input[placeholder*='Search by patient']"


# ============================================================
# PAGE TITLE
# ============================================================

st.title("📄 Lab Report Downloader")

st.write(
    "Upload an Excel file containing case numbers and download "
    "the Patient & Invoice Reports from LabConnect."
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Settings")

    st.info(
        "The application uses Chromium on the server so it can "
        "run from Streamlit Cloud."
    )

    browser_type = "Chromium"


# ============================================================
# LOGIN DETAILS
# ============================================================

st.subheader("LabConnect Login")

username = st.text_input(
    "LabConnect Username",
    type="text"
)

password = st.text_input(
    "LabConnect Password",
    type="password"
)


# ============================================================
# EXCEL UPLOAD
# ============================================================

st.subheader("Upload Case List")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"]
)


# ============================================================
# LOAD CASE NUMBERS
# ============================================================

def load_cases(file_path):
    """
    Read case numbers from the uploaded Excel file.
    Attempts to identify the Case Number column automatically.
    """

    try:
        df = pd.read_excel(file_path)

    except Exception as e:
        raise Exception(f"Could not read Excel file: {e}")

    if df.empty:
        raise Exception("The Excel file is empty.")

    # Possible column names
    possible_columns = [
        "Case Number",
        "Case No",
        "Case",
        "CaseNumber",
        "Case_Number",
        "Case number",
        "CASE NUMBER",
        "CASE NO",
        "CASE"
    ]

    case_column = None

    for col in possible_columns:
        if col in df.columns:
            case_column = col
            break

    # If exact match wasn't found, look for columns
    # containing the word "case"
    if case_column is None:

        for col in df.columns:
            if "case" in str(col).lower():
                case_column = col
                break

    if case_column is None:
        raise Exception(
            "Could not find a Case Number column in the Excel file. "
            f"Available columns: {list(df.columns)}"
        )

    cases = (
        df[case_column]
        .dropna()
        .astype(str)
        .str.strip()
    )

    # Remove Excel-style .0 from numeric case numbers
    cases = cases.str.replace(
        r"\.0$",
        "",
        regex=True
    )

    # Remove blank values
    cases = cases[cases != ""]

    cases = cases.tolist()

    if not cases:
        raise Exception(
            "No case numbers were found in the Excel file."
        )

    return cases


# ============================================================
# LOGIN
# ============================================================

def do_login(page, username, password):

    page.goto(
        LABCONNECT_URL,
        wait_until="domcontentloaded",
        timeout=120000
    )

    page.wait_for_timeout(3000)

    # Username
    username_selectors = [
        "input[type='email']",
        "input[name='username']",
        "input[name='email']",
        "input[placeholder*='Username']",
        "input[placeholder*='username']",
        "input[placeholder*='Email']",
        "input[placeholder*='email']"
    ]

    username_box = None

    for selector in username_selectors:

        try:
            locator = page.locator(selector)

            if locator.count() > 0:
                username_box = locator.first
                break

        except Exception:
            pass

    if username_box is None:
        raise Exception(
            "Could not find the LabConnect username field."
        )

    username_box.fill(username)

    # Password
    password_selectors = [
        "input[type='password']",
        "input[name='password']",
        "input[placeholder*='Password']",
        "input[placeholder*='password']"
    ]

    password_box = None

    for selector in password_selectors:

        try:
            locator = page.locator(selector)

            if locator.count() > 0:
                password_box = locator.first
                break

        except Exception:
            pass

    if password_box is None:
        raise Exception(
            "Could not find the LabConnect password field."
        )

    password_box.fill(password)

    # Login button
    login_selectors = [
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Login')",
        "button:has-text('Log in')",
        "button:has-text('Sign in')",
        "text=Login",
        "text=Log in",
        "text=Sign in"
    ]

    login_button = None

    for selector in login_selectors:

        try:
            locator = page.locator(selector)

            if locator.count() > 0:
                login_button = locator.first
                break

        except Exception:
            pass

    if login_button is None:
        raise Exception(
            "Could not find the LabConnect login button."
        )

    login_button.click()

    page.wait_for_timeout(5000)


# ============================================================
# OPEN HOME PAGE
# ============================================================

def open_home(page):

    try:
        page.wait_for_load_state(
            "domcontentloaded",
            timeout=60000
        )
    except Exception:
        pass

    page.wait_for_timeout(3000)


# ============================================================
# PROCESS ONE CASE
# ============================================================

def process_case(page, case_number, output_folder):

    try:

        # ----------------------------------------------------
        # Search box
        # ----------------------------------------------------

        search_box = page.locator(SEARCH_SEL)

        search_box.wait_for(
            state="visible",
            timeout=30000
        )

        search_box.fill(str(case_number))

        page.wait_for_timeout(2000)

        # ----------------------------------------------------
        # Search
        # ----------------------------------------------------

        # Try pressing Enter first
        search_box.press("Enter")

        page.wait_for_timeout(3000)

        # ----------------------------------------------------
        # Look for case/result
        # ----------------------------------------------------

        possible_case_selectors = [
            f"text={case_number}",
            f"td:has-text('{case_number}')",
            f"div:has-text('{case_number}')",
            f"a:has-text('{case_number}')"
        ]

        case_found = False

        for selector in possible_case_selectors:

            try:

                locator = page.locator(selector)

                if locator.count() > 0:

                    locator.first.click()

                    case_found = True

                    break

            except Exception:
                pass

        if not case_found:

            # Sometimes search automatically opens
            # the case, so don't immediately fail.
            page.wait_for_timeout(2000)

        # ----------------------------------------------------
        # Patient & Invoice Report
        # ----------------------------------------------------

        report_selectors = [
            "text=Patient & Invoice Report",
            "text=Patient and Invoice Report",
            "text=Patient & Invoice",
            "button:has-text('Patient & Invoice Report')",
            "a:has-text('Patient & Invoice Report')"
        ]

        report_button = None

        for selector in report_selectors:

            try:

                locator = page.locator(selector)

                if locator.count() > 0:

                    report_button = locator.first

                    break

            except Exception:
                pass

        if report_button is None:

            raise Exception(
                "Patient & Invoice Report button was not found."
            )

        # ----------------------------------------------------
        # Download PDF
        # ----------------------------------------------------

        with page.expect_download(
            timeout=60000
        ) as download_info:

            report_button.click()

        download = download_info.value

        # ----------------------------------------------------
        # Save PDF
        # ----------------------------------------------------

        safe_case_number = re.sub(
            r"[^A-Za-z0-9_\-]",
            "_",
            str(case_number)
        )

        filename = (
            f"{safe_case_number}_Patient_Invoice_Report.pdf"
        )

        save_path = Path(output_folder) / filename

        download.save_as(str(save_path))

        return True, f"Downloaded {filename}"

    except Exception as e:

        return False, str(e)


# ============================================================
# CREATE ZIP
# ============================================================

def create_zip(folder):

    folder = Path(folder)

    zip_path = folder.parent / "Lab_Reports.zip"

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for file in folder.iterdir():

            if file.is_file():

                zip_file.write(
                    file,
                    arcname=file.name
                )

    return zip_path


# ============================================================
# MAIN BUTTON
# ============================================================

if st.button(
    "🚀 Start Download",
    type="primary",
    use_container_width=True
):

    # --------------------------------------------------------
    # Validate username
    # --------------------------------------------------------

    if not username:

        st.error(
            "Please enter your LabConnect username."
        )

        st.stop()

    # --------------------------------------------------------
    # Validate password
    # --------------------------------------------------------

    if not password:

        st.error(
            "Please enter your LabConnect password."
        )

        st.stop()

    # --------------------------------------------------------
    # Validate Excel
    # --------------------------------------------------------

    if uploaded_file is None:

        st.error(
            "Please upload an Excel file."
        )

        st.stop()

    # --------------------------------------------------------
    # Create temporary directory
    # --------------------------------------------------------

    temp_dir = Path(
        tempfile.mkdtemp(
            prefix="lab_reports_"
        )
    )

    excel_path = temp_dir / uploaded_file.name

    output_folder = temp_dir / "Reports"

    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    try:

        # ----------------------------------------------------
        # Save uploaded Excel
        # ----------------------------------------------------

        with open(
            excel_path,
            "wb"
        ) as f:

            f.write(
                uploaded_file.getbuffer()
            )

        # ----------------------------------------------------
        # Read cases
        # ----------------------------------------------------

        st.info(
            "Reading case numbers from Excel..."
        )

        cases = load_cases(
            excel_path
        )

        st.success(
            f"Found {len(cases)} case(s)."
        )

        # ----------------------------------------------------
        # Progress UI
        # ----------------------------------------------------

        progress_bar = st.progress(0)

        status_text = st.empty()

        results_container = st.empty()

        successful = []
        failed = []

        # ----------------------------------------------------
        # Start Playwright
        # ----------------------------------------------------

        status_text.info(
            "Starting Chromium browser..."
        )

        with sync_playwright() as p:

            # =================================================
            # IMPORTANT:
            # Streamlit Cloud uses the Linux Chromium package
            # installed through packages.txt.
            # =================================================

            chromium_path = shutil.which(
                "chromium"
            )

            if not chromium_path:

                # Try chromium-browser as fallback
                chromium_path = shutil.which(
                    "chromium-browser"
                )

            if not chromium_path:

                raise RuntimeError(
                    "Chromium was not found on the Streamlit "
                    "server. Make sure packages.txt contains "
                    "the line: chromium"
                )

            status_text.info(
                f"Using Chromium: {chromium_path}"
            )

            browser = p.chromium.launch(
                headless=True,
                executable_path=chromium_path,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--no-first-run",
                    "--no-zygote",
                    "--disable-setuid-sandbox"
                ]
            )

            # ------------------------------------------------
            # Browser context
            # ------------------------------------------------

            context = browser.new_context(
                accept_downloads=True,
                viewport={
                    "width": 1440,
                    "height": 900
                }
            )

            page = context.new_page()

            # ------------------------------------------------
            # Login
            # ------------------------------------------------

            status_text.info(
                "Opening LabConnect and logging in..."
            )

            do_login(
                page,
                username,
                password
            )

            open_home(page)

            status_text.success(
                "Logged into LabConnect."
            )

            # ------------------------------------------------
            # Process cases
            # ------------------------------------------------

            for index, case_number in enumerate(cases):

                current_number = index + 1

                status_text.info(
                    f"Processing case "
                    f"{current_number}/{len(cases)}: "
                    f"{case_number}"
                )

                success, message = process_case(
                    page,
                    case_number,
                    output_folder
                )

                if success:

                    successful.append(
                        case_number
                    )

                else:

                    failed.append(
                        (
                            case_number,
                            message
                        )
                    )

                # Progress
                progress = (
                    current_number /
                    len(cases)
                )

                progress_bar.progress(
                    progress
                )

            # ------------------------------------------------
            # Close browser
            # ------------------------------------------------

            context.close()
            browser.close()

        # ====================================================
        # CREATE ZIP
        # ====================================================

        status_text.info(
            "Creating ZIP file..."
        )

        zip_path = create_zip(
            output_folder
        )

        # Read ZIP into memory
        with open(
            zip_path,
            "rb"
        ) as f:

            zip_bytes = f.read()

        # ====================================================
        # RESULTS
        # ====================================================

        st.divider()

        st.subheader(
            "Download Complete"
        )

        col1, col2 = st.columns(2)

        with col1:

            st.metric(
                "Successful",
                len(successful)
            )

        with col2:

            st.metric(
                "Failed",
                len(failed)
            )

        # ----------------------------------------------------
        # Download button
        # ----------------------------------------------------

        st.download_button(
            label="⬇️ Download All Reports (ZIP)",
            data=zip_bytes,
            file_name=(
                "Lab_Reports_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                + ".zip"
            ),
            mime="application/zip",
            use_container_width=True
        )

        # ====================================================
        # FAILED CASES
        # ====================================================

        if failed:

            st.warning(
                f"{len(failed)} case(s) could not be downloaded."
            )

            with st.expander(
                "View Failed Cases"
            ):

                for case_number, error in failed:

                    st.write(
                        f"**{case_number}** — {error}"
                    )

        # ====================================================
        # SUCCESSFUL CASES
        # ====================================================

        if successful:

            with st.expander(
                "View Successful Cases"
            ):

                for case_number in successful:

                    st.write(
                        f"✅ {case_number}"
                    )

    except Exception as e:

        st.error(
            "Automation error:"
        )

        st.code(
            str(e)
        )

    finally:

        # ----------------------------------------------------
        # Cleanup temporary files
        # ----------------------------------------------------

        try:

            shutil.rmtree(
                temp_dir,
                ignore_errors=True
            )

        except Exception:
            pass
