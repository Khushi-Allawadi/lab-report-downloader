import streamlit as st
import pandas as pd
from playwright.sync_api import sync_playwright
from pathlib import Path
from datetime import datetime
import tempfile
import zipfile
import shutil
import re


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
# HEADER
# ============================================================

st.title("📄 Lab Report Downloader")

st.write(
    "Upload an Excel file containing case numbers and download "
    "the Patient & Invoice Reports from LabConnect."
)

st.divider()


# ============================================================
# LOGIN
# ============================================================

st.subheader("LabConnect Login")

username = st.text_input(
    "LabConnect Username"
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
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):

    if pd.isna(value):
        return ""

    text = str(value).strip().lower()

    text = re.sub(
        r"[^a-z0-9\s]",
        "",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip()

    return text


# ============================================================
# LOAD CASE NUMBERS
# ============================================================

def load_cases(file_path):

    try:

        # Read without assuming first row is header
        df = pd.read_excel(
            file_path,
            header=None
        )

    except Exception as e:

        raise Exception(
            f"Could not read Excel file: {e}"
        )

    if df.empty:

        raise Exception(
            "The Excel file is empty."
        )

    # --------------------------------------------------------
    # Find Case No. anywhere in sheet
    # --------------------------------------------------------

    case_row = None
    case_col = None

    for row_index in range(
        len(df)
    ):

        for col_index in range(
            len(df.columns)
        ):

            value = df.iloc[
                row_index,
                col_index
            ]

            normalized = normalize_text(
                value
            )

            if normalized in [
                "case no",
                "case number",
                "caseno",
                "case id",
                "caseid"
            ]:

                case_row = row_index
                case_col = col_index

                break

        if case_col is not None:
            break

    # --------------------------------------------------------
    # Not found
    # --------------------------------------------------------

    if case_col is None:

        preview = df.head(
            20
        ).to_string(
            index=False,
            header=False
        )

        raise Exception(
            "Could not find 'Case No.' anywhere in "
            "the Excel sheet.\n\n"
            "First 20 rows detected:\n\n"
            + preview
        )

    # --------------------------------------------------------
    # Values underneath Case No.
    # --------------------------------------------------------

    case_values = df.iloc[
        case_row + 1:,
        case_col
    ]

    case_values = case_values.dropna()

    case_values = case_values.astype(
        str
    )

    case_values = case_values.str.strip()

    # Remove Excel .0
    case_values = case_values.str.replace(
        r"\.0$",
        "",
        regex=True
    )

    case_values = case_values[
        case_values != ""
    ]

    # Remove repeated header
    case_values = case_values[
        ~case_values.apply(
            normalize_text
        ).isin([
            "case no",
            "case number",
            "caseno",
            "case id",
            "caseid"
        ])
    ]

    cases = case_values.tolist()

    # Remove duplicates while maintaining order
    cases = list(
        dict.fromkeys(
            cases
        )
    )

    if not cases:

        raise Exception(
            "The Case No. column was found, "
            "but no case numbers were found underneath it."
        )

    return cases


# ============================================================
# FIND USERNAME
# ============================================================

def find_username_field(page):

    selectors = [

        "input[type='email']",
        "input[name='username']",
        "input[name='email']",
        "input[placeholder*='Username']",
        "input[placeholder*='username']",
        "input[placeholder*='Email']",
        "input[placeholder*='email']"
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                return locator.first

        except Exception:

            pass

    return None


# ============================================================
# FIND PASSWORD
# ============================================================

def find_password_field(page):

    selectors = [

        "input[type='password']",
        "input[name='password']",
        "input[placeholder*='Password']",
        "input[placeholder*='password']"
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                return locator.first

        except Exception:

            pass

    return None


# ============================================================
# LOGIN
# ============================================================

def do_login(
    page,
    username,
    password
):

    page.goto(
        LABCONNECT_URL,
        wait_until="domcontentloaded",
        timeout=120000
    )

    page.wait_for_timeout(
        5000
    )

    # --------------------------------------------------------
    # Username
    # --------------------------------------------------------

    username_box = find_username_field(
        page
    )

    if username_box is None:

        raise Exception(
            "Could not find the LabConnect username field."
        )

    username_box.fill(
        username
    )

    # --------------------------------------------------------
    # Password
    # --------------------------------------------------------

    password_box = find_password_field(
        page
    )

    if password_box is None:

        raise Exception(
            "Could not find the LabConnect password field."
        )

    password_box.fill(
        password
    )

    # --------------------------------------------------------
    # Login button
    # --------------------------------------------------------

    login_selectors = [

        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Login')",
        "button:has-text('Log in')",
        "button:has-text('Sign in')",
        "input[value='Login']",
        "input[value='Log in']",
        "input[value='Sign in']"
    ]

    login_button = None

    for selector in login_selectors:

        try:

            locator = page.locator(
                selector
            )

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

    page.wait_for_timeout(
        7000
    )


# ============================================================
# SEARCH CASE
# ============================================================

def search_case(
    page,
    case_number
):

    search_box = page.locator(
        SEARCH_SEL
    )

    search_box.wait_for(
        state="visible",
        timeout=30000
    )

    search_box.fill(
        ""
    )

    search_box.fill(
        str(case_number)
    )

    page.wait_for_timeout(
        1500
    )

    search_box.press(
        "Enter"
    )

    page.wait_for_timeout(
        4000
    )


# ============================================================
# OPEN CASE
# ============================================================

def open_case(
    page,
    case_number
):

    selectors = [

        f"text={case_number}",
        f"td:has-text('{case_number}')",
        f"a:has-text('{case_number}')",
        f"div:has-text('{case_number}')"
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                locator.first.click()

                page.wait_for_timeout(
                    3000
                )

                return True

        except Exception:

            pass

    return False


# ============================================================
# DOWNLOAD REPORT
# ============================================================

def download_report(
    page,
    case_number,
    output_folder
):

    report_selectors = [

        "text=Patient & Invoice Report",

        "text=Patient and Invoice Report",

        "text=Patient & Invoice",

        "button:has-text('Patient & Invoice Report')",

        "a:has-text('Patient & Invoice Report')"
    ]

    report_button = None

    # --------------------------------------------------------
    # Find button
    # --------------------------------------------------------

    for selector in report_selectors:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                report_button = locator.first

                break

        except Exception:

            pass

    if report_button is None:

        raise Exception(
            "Patient & Invoice Report button was not found."
        )

    # --------------------------------------------------------
    # Download
    # --------------------------------------------------------

    try:

        with page.expect_download(
            timeout=120000
        ) as download_info:

            report_button.click()

        download = download_info.value

    except Exception as e:

        raise Exception(
            f"Report button was found, but the PDF "
            f"download did not start: {e}"
        )

    # --------------------------------------------------------
    # Filename
    # --------------------------------------------------------

    safe_case_number = re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        str(case_number)
    )

    filename = (
        f"{safe_case_number}"
    )

    save_path = (
        Path(output_folder)
        / filename
    )

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    download.save_as(
        str(save_path)
    )

    # --------------------------------------------------------
    # VERIFY FILE EXISTS
    # --------------------------------------------------------

    if not save_path.exists():

        raise Exception(
            "Download was reported by Playwright, "
            "but the PDF file was not created."
        )

    # --------------------------------------------------------
    # VERIFY FILE SIZE
    # --------------------------------------------------------

    file_size = save_path.stat().st_size

    if file_size == 0:

        # Delete empty file
        try:
            save_path.unlink()
        except Exception:
            pass

        raise Exception(
            "The PDF file was created but is 0 bytes."
        )

    # --------------------------------------------------------
    # VERIFY PDF SIGNATURE
    # --------------------------------------------------------

    try:

        with open(
            save_path,
            "rb"
        ) as pdf_file:

            first_bytes = pdf_file.read(
                5
            )

        if first_bytes != b"%PDF-":

            raise Exception(
                "The downloaded file is not a valid PDF."
            )

    except Exception as e:

        try:
            save_path.unlink()
        except Exception:
            pass

        raise Exception(
            str(e)
        )

    return save_path


# ============================================================
# PROCESS CASE
# ============================================================

def process_case(
    page,
    case_number,
    output_folder
):

    try:

        # ----------------------------------------------------
        # Search
        # ----------------------------------------------------

        search_case(
            page,
            case_number
        )

        # ----------------------------------------------------
        # Open case
        # ----------------------------------------------------

        case_opened = open_case(
            page,
            case_number
        )

        if not case_opened:

            raise Exception(
                f"Could not open case {case_number}."
            )

        # ----------------------------------------------------
        # Download
        # ----------------------------------------------------

        report_path = download_report(
            page,
            case_number,
            output_folder
        )

        # ----------------------------------------------------
        # Final verification
        # ----------------------------------------------------

        file_size = report_path.stat().st_size

        return True, (
            report_path.name,
            file_size
        )

    except Exception as e:

        return False, str(e)


# ============================================================
# CREATE VERIFIED ZIP
# ============================================================

def create_zip(
    output_folder
):

    output_folder = Path(
        output_folder
    )

    # --------------------------------------------------------
    # Only include actual PDF files
    # --------------------------------------------------------

    pdf_files = []

    for file in output_folder.iterdir():

        if not file.is_file():
            continue

        if file.suffix.lower() != ".pdf":
            continue

        if file.stat().st_size <= 0:
            continue

        pdf_files.append(
            file
        )

    # --------------------------------------------------------
    # No PDFs
    # --------------------------------------------------------

    if not pdf_files:

        raise Exception(
            "NO PDF FILES WERE FOUND AFTER THE "
            "AUTOMATION FINISHED.\n\n"
            "The ZIP was NOT created because there "
            "are no valid PDF files to put inside it."
        )

    # --------------------------------------------------------
    # Create ZIP
    # --------------------------------------------------------

    zip_path = (
        output_folder.parent
        / "Lab_Reports.zip"
    )

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for pdf_file in pdf_files:

            zip_file.write(
                pdf_file,
                arcname=pdf_file.name
            )

    # --------------------------------------------------------
    # VERIFY ZIP
    # --------------------------------------------------------

    if not zip_path.exists():

        raise Exception(
            "ZIP file was not created."
        )

    zip_size = zip_path.stat().st_size

    if zip_size <= 22:

        raise Exception(
            "The ZIP file appears to be empty."
        )

    # --------------------------------------------------------
    # Verify ZIP contents
    # --------------------------------------------------------

    with zipfile.ZipFile(
        zip_path,
        "r"
    ) as zip_file:

        zip_contents = zip_file.namelist()

    if not zip_contents:

        raise Exception(
            "ZIP was created but contains no files."
        )

    return zip_path, pdf_files, zip_contents


# ============================================================
# START DOWNLOAD
# ============================================================

if st.button(
    "🚀 Start Download",
    type="primary",
    use_container_width=True
):

    # ========================================================
    # VALIDATION
    # ========================================================

    if not username:

        st.error(
            "Please enter your LabConnect username."
        )

        st.stop()

    if not password:

        st.error(
            "Please enter your LabConnect password."
        )

        st.stop()

    if uploaded_file is None:

        st.error(
            "Please upload an Excel file."
        )

        st.stop()

    # ========================================================
    # TEMP FOLDER
    # ========================================================

    temp_dir = Path(
        tempfile.mkdtemp(
            prefix="lab_reports_"
        )
    )

    excel_path = (
        temp_dir
        / uploaded_file.name
    )

    output_folder = (
        temp_dir
        / "Reports"
    )

    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    try:

        # ====================================================
        # SAVE EXCEL
        # ====================================================

        with open(
            excel_path,
            "wb"
        ) as file:

            file.write(
                uploaded_file.getbuffer()
            )

        # ====================================================
        # READ CASES
        # ====================================================

        st.info(
            "Reading Case No. from Excel..."
        )

        cases = load_cases(
            excel_path
        )

        st.success(
            f"Found {len(cases)} case(s)."
        )

        with st.expander(
            "📋 View Case Numbers"
        ):

            st.write(
                cases
            )

        # ====================================================
        # PROGRESS
        # ====================================================

        progress_bar = st.progress(
            0
        )

        status_text = st.empty()

        successful = []

        failed = []

        # ====================================================
        # START PLAYWRIGHT
        # ====================================================

        status_text.info(
            "Starting Chromium..."
        )

        with sync_playwright() as p:

            # ------------------------------------------------
            # Find Chromium
            # ------------------------------------------------

            chromium_path = shutil.which(
                "chromium"
            )

            if not chromium_path:

                chromium_path = shutil.which(
                    "chromium-browser"
                )

            if not chromium_path:

                raise RuntimeError(
                    "Chromium was not found on the "
                    "Streamlit server."
                )

            status_text.info(
                "Chromium found. Opening LabConnect..."
            )

            # ------------------------------------------------
            # Browser
            # ------------------------------------------------

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

            context = browser.new_context(
                accept_downloads=True,
                viewport={
                    "width": 1440,
                    "height": 900
                }
            )

            page = context.new_page()

            # =================================================
            # LOGIN
            # =================================================

            status_text.info(
                "Logging into LabConnect..."
            )

            do_login(
                page,
                username,
                password
            )

            status_text.success(
                "Logged into LabConnect."
            )

            # =================================================
            # PROCESS CASES
            # =================================================

            total_cases = len(
                cases
            )

            for index, case_number in enumerate(
                cases
            ):

                position = index + 1

                status_text.info(
                    f"Processing "
                    f"{position}/{total_cases}: "
                    f"{case_number}"
                )

                success, result = process_case(
                    page,
                    case_number,
                    output_folder
                )

                if success:

                    filename, file_size = result

                    successful.append(
                        (
                            case_number,
                            filename,
                            file_size
                        )
                    )

                else:

                    failed.append(
                        (
                            case_number,
                            result
                        )
                    )

                progress_bar.progress(
                    position / total_cases
                )

            # =================================================
            # CLOSE BROWSER
            # =================================================

            context.close()

            browser.close()

        # ====================================================
        # IMPORTANT:
        # VERIFY FILES BEFORE ZIP
        # ====================================================

        actual_pdfs = [

            file for file in output_folder.iterdir()

            if (
                file.is_file()
                and file.suffix.lower() == ".pdf"
                and file.stat().st_size > 0
            )
        ]

        # ====================================================
        # SHOW ACTUAL FILE COUNT
        # ====================================================

        st.divider()

        st.subheader(
            "📊 Download Verification"
        )

        st.write(
            f"Cases processed: **{len(cases)}**"
        )

        st.write(
            f"Valid PDFs actually created: "
            f"**{len(actual_pdfs)}**"
        )

        st.write(
            f"Failed cases: **{len(failed)}**"
        )

        # ====================================================
        # SHOW ACTUAL FILES
        # ====================================================

        if actual_pdfs:

            with st.expander(
                "📁 View Actual PDF Files"
            ):

                for pdf in actual_pdfs:

                    size = pdf.stat().st_size

                    st.write(
                        f"✅ {pdf.name} "
                        f"— {size:,} bytes"
                    )

        # ====================================================
        # NO FILES = STOP
        # ====================================================

        if not actual_pdfs:

            st.error(
                "❌ ZERO PDF FILES WERE ACTUALLY CREATED."
            )

            st.warning(
                "The automation processed the cases, "
                "but no valid PDF files were found. "
                "I have NOT created an empty ZIP."
            )

            if failed:

                with st.expander(
                    "❌ View Failed Cases"
                ):

                    for case_number, error in failed:

                        st.write(
                            f"**{case_number}**"
                        )

                        st.code(
                            error
                        )

            st.stop()

        # ====================================================
        # CREATE VERIFIED ZIP
        # ====================================================

        status_text.info(
            "Creating and verifying ZIP..."
        )

        zip_path, pdf_files, zip_contents = create_zip(
            output_folder
        )

        # ====================================================
        # READ ZIP
        # ====================================================

        with open(
            zip_path,
            "rb"
        ) as file:

            zip_bytes = file.read()

        # ====================================================
        # FINAL VERIFICATION
        # ====================================================

        if len(zip_bytes) == 0:

            raise Exception(
                "The ZIP file contains zero bytes."
            )

        if not zip_contents:

            raise Exception(
                "The ZIP file contains no files."
            )

        # ====================================================
        # SUCCESS
        # ====================================================

        status_text.success(
            "✅ Reports downloaded and ZIP verified."
        )

        st.success(
            f"✅ {len(pdf_files)} PDF file(s) "
            f"successfully added to the ZIP."
        )

        # ====================================================
        # DOWNLOAD ZIP
        # ====================================================

        st.download_button(
            label="⬇️ DOWNLOAD ALL REPORTS (ZIP)",
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
        # ZIP CONTENTS
        # ====================================================

        with st.expander(
            "📦 View ZIP Contents"
        ):

            for filename in zip_contents:

                st.write(
                    f"✅ {filename}"
                )

        # ====================================================
        # FAILED CASES
        # ====================================================

        if failed:

            st.warning(
                f"{len(failed)} case(s) failed."
            )

            with st.expander(
                "❌ View Failed Cases"
            ):

                for case_number, error in failed:

                    st.write(
                        f"**{case_number}**"
                    )

                    st.code(
                        error
                    )

    # ========================================================
    # ERROR
    # ========================================================

    except Exception as e:

        st.error(
            "❌ Automation error"
        )

        st.code(
            str(e)
        )

    finally:

        # Cleanup only happens after the ZIP bytes
        # have already been loaded into memory.

        try:

            shutil.rmtree(
                temp_dir,
                ignore_errors=True
            )

        except Exception:

            pass
