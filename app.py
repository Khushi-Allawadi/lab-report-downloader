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
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("Settings")

    st.info(
        "The application uses Chromium on the server."
    )


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

    # Remove punctuation
    text = re.sub(
        r"[^a-z0-9\s]",
        "",
        text
    )

    # Remove extra spaces
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

    """
    Reads the Excel file without assuming that the first row
    contains the headers.

    It searches the entire sheet for:

        Case No.
        Case No
        Case Number
        CaseNumber
        Case ID

    Then it reads the values underneath that column.
    """

    try:

        # IMPORTANT:
        # header=None means we DON'T assume row 1 is the header.
        df = pd.read_excel(
            file_path,
            header=None
        )

    except Exception as e:

        raise Exception(
            f"Could not read Excel file: {e}"
        )

    # --------------------------------------------------------
    # Empty file
    # --------------------------------------------------------

    if df.empty:

        raise Exception(
            "The Excel file is empty."
        )

    # --------------------------------------------------------
    # Find Case No. anywhere in the sheet
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
    # Case No. wasn't found
    # --------------------------------------------------------

    if case_col is None:

        # Create a useful preview for debugging
        preview = df.head(
            20
        ).to_string(
            index=False,
            header=False
        )

        raise Exception(
            "Could not find 'Case No.' anywhere in the "
            "Excel sheet.\n\n"
            "First 20 rows detected:\n\n"
            + preview
        )

    # --------------------------------------------------------
    # Get everything BELOW Case No.
    # --------------------------------------------------------

    case_values = df.iloc[
        case_row + 1:,
        case_col
    ]

    # --------------------------------------------------------
    # Remove empty cells
    # --------------------------------------------------------

    case_values = case_values.dropna()

    # --------------------------------------------------------
    # Convert to string
    # --------------------------------------------------------

    case_values = case_values.astype(
        str
    )

    # --------------------------------------------------------
    # Remove spaces
    # --------------------------------------------------------

    case_values = case_values.str.strip()

    # --------------------------------------------------------
    # Remove Excel .0
    #
    # Example:
    # 123456.0 -> 123456
    # --------------------------------------------------------

    case_values = case_values.str.replace(
        r"\.0$",
        "",
        regex=True
    )

    # --------------------------------------------------------
    # Remove blank values
    # --------------------------------------------------------

    case_values = case_values[
        case_values != ""
    ]

    # --------------------------------------------------------
    # Remove repeated header if present
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Convert to list
    # --------------------------------------------------------

    cases = case_values.tolist()

    # --------------------------------------------------------
    # Remove duplicate case numbers
    # --------------------------------------------------------

    cases = list(
        dict.fromkeys(
            cases
        )
    )

    # --------------------------------------------------------
    # Validate
    # --------------------------------------------------------

    if not cases:

        raise Exception(
            "The 'Case No.' column was found, "
            "but there are no case numbers underneath it."
        )

    return cases


# ============================================================
# FIND USERNAME FIELD
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

            continue

    return None


# ============================================================
# FIND PASSWORD FIELD
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

            continue

    return None


# ============================================================
# LOGIN
# ============================================================

def do_login(
    page,
    username,
    password
):

    # --------------------------------------------------------
    # Open LabConnect
    # --------------------------------------------------------

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

            continue

    if login_button is None:

        raise Exception(
            "Could not find the LabConnect login button."
        )

    # --------------------------------------------------------
    # Login
    # --------------------------------------------------------

    login_button.click()

    page.wait_for_timeout(
        7000
    )


# ============================================================
# SEARCH FOR CASE
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

    # Clear old search
    search_box.fill(
        ""
    )

    # Enter case number
    search_box.fill(
        str(case_number)
    )

    page.wait_for_timeout(
        1500
    )

    # Search
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

            continue

    return False


# ============================================================
# DOWNLOAD PATIENT & INVOICE REPORT
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

    for selector in report_selectors:

        try:

            locator = page.locator(
                selector
            )

            if locator.count() > 0:

                report_button = locator.first

                break

        except Exception:

            continue

    if report_button is None:

        raise Exception(
            "Patient & Invoice Report button was not found."
        )

    # --------------------------------------------------------
    # Download PDF
    # --------------------------------------------------------

    with page.expect_download(
        timeout=60000
    ) as download_info:

        report_button.click()

    download = download_info.value

    # --------------------------------------------------------
    # Safe filename
    # --------------------------------------------------------

    safe_case_number = re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        str(case_number)
    )

    filename = (
        f"{safe_case_number}"
        f"_Patient_Invoice_Report.pdf"
    )

    save_path = (
        Path(output_folder)
        / filename
    )

    download.save_as(
        str(save_path)
    )

    return save_path


# ============================================================
# PROCESS ONE CASE
# ============================================================

def process_case(
    page,
    case_number,
    output_folder
):

    try:

        # Search
        search_case(
            page,
            case_number
        )

        # Open case
        open_case(
            page,
            case_number
        )

        # Download
        report_path = download_report(
            page,
            case_number,
            output_folder
        )

        return True, report_path.name

    except Exception as e:

        return False, str(e)


# ============================================================
# CREATE ZIP
# ============================================================

def create_zip(
    output_folder
):

    output_folder = Path(
        output_folder
    )

    zip_path = (
        output_folder.parent
        / "Lab_Reports.zip"
    )

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for file in output_folder.iterdir():

            if file.is_file():

                zip_file.write(
                    file,
                    arcname=file.name
                )

    return zip_path


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
    # TEMPORARY FOLDER
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

        # ----------------------------------------------------
        # Show cases
        # ----------------------------------------------------

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
        # PLAYWRIGHT
        # ====================================================

        status_text.info(
            "Starting Chromium..."
        )

        with sync_playwright() as p:

            # ------------------------------------------------
            # Find system Chromium
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
                    "Streamlit server.\n\n"
                    "Make sure packages.txt contains:\n"
                    "chromium"
                )

            status_text.info(
                "Chromium found. Opening LabConnect..."
            )

            # ------------------------------------------------
            # Launch browser
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

                position = (
                    index + 1
                )

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

                    successful.append(
                        case_number
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
        # CREATE ZIP
        # ====================================================

        status_text.info(
            "Creating ZIP file..."
        )

        zip_path = create_zip(
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
        # RESULTS
        # ====================================================

        st.divider()

        st.subheader(
            "✅ Download Complete"
        )

        col1, col2 = st.columns(
            2
        )

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

        # ====================================================
        # DOWNLOAD BUTTON
        # ====================================================

        st.download_button(
            label="⬇️ Download All Reports",
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

        # ====================================================
        # SUCCESSFUL CASES
        # ====================================================

        if successful:

            with st.expander(
                "✅ View Successful Cases"
            ):

                for case_number in successful:

                    st.write(
                        f"✓ {case_number}"
                    )

        status_text.success(
            "All processing is complete."
        )

    # ========================================================
    # ERROR
    # ========================================================

    except Exception as e:

        st.error(
            "Automation error"
        )

        st.code(
            str(e)
        )

    # ========================================================
    # CLEANUP
    # ========================================================

    finally:

        try:

            shutil.rmtree(
                temp_dir,
                ignore_errors=True
            )

        except Exception:

            pass
