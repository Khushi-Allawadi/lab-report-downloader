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


# ============================================================
# CONFIG
# ============================================================

LABCONNECT_URL = "https://lms.mylabconnect.co.uk/lab-connect"

SEARCH_SEL = "input[placeholder*='Search by patient']"


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Lab Report Downloader",
    page_icon="📄",
    layout="centered"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #666;
        margin-bottom: 25px;
    }

    .success-box {
        padding: 15px;
        border-radius: 10px;
        background-color: #e9f7ef;
        border: 1px solid #b7e4c7;
    }

    .error-box {
        padding: 15px;
        border-radius: 10px;
        background-color: #fdecec;
        border: 1px solid #f5b5b5;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">📄 Lab Report Downloader</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Upload your Case No. Excel file and download the Patient & Invoice Reports.'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# FUNCTIONS
# ============================================================

def load_cases(excel_path):

    raw = pd.read_excel(
        excel_path,
        header=None,
        dtype=str
    )

    for r in range(min(15, len(raw))):

        for c in range(raw.shape[1]):

            cell = (
                str(raw.iat[r, c])
                .strip()
                .lower()
                .replace(" ", "")
            )

            if cell in ("caseno.", "caseno", "caseno:"):

                vals = (
                    raw.iloc[r + 1:, c]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                return list(
                    dict.fromkeys(
                        v for v in vals
                        if v and v.lower() != "nan"
                    )
                )

    raise ValueError(
        "Could not find a 'Case No.' column in the first rows of the Excel file."
    )


def safe_filename(value):

    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        str(value)
    )


def do_login(page, username, password):

    page.get_by_placeholder(
        "Enter your username"
    ).fill(username)

    page.get_by_placeholder(
        "Enter your password"
    ).fill(password)

    page.locator(
        "button.login-submit-btn"
    ).click()

    try:

        page.locator(
            SEARCH_SEL
        ).wait_for(timeout=45000)

    except Exception:

        raise RuntimeError(
            "Login failed. Please check your LabConnect username and password."
        )


def open_home(page, username, password):

    page.goto(LABCONNECT_URL)

    page.wait_for_load_state(
        "domcontentloaded"
    )

    try:

        page.locator(
            SEARCH_SEL
        ).wait_for(timeout=8000)

    except Exception:

        if page.get_by_placeholder(
            "Enter your username"
        ).count() > 0:

            do_login(
                page,
                username,
                password
            )

        else:

            page.locator(
                SEARCH_SEL
            ).wait_for(timeout=20000)


def process_case(page, case, output_dir):

    step = "search box"

    try:

        # ------------------------------------------------
        # Search Case
        # ------------------------------------------------

        box = page.locator(
            SEARCH_SEL
        )

        box.click()

        box.fill("")

        box.fill(case)


        # ------------------------------------------------
        # Select Case
        # ------------------------------------------------

        step = "search result"

        result = page.get_by_text(
            case,
            exact=True
        ).first

        result.wait_for(
            timeout=15000
        )

        result.click(
            force=True
        )


        # ------------------------------------------------
        # Patient & Invoice Report
        # ------------------------------------------------

        step = "report button"

        btn = page.locator(
            "button:has-text('Patient & Invoice Report')"
        )

        btn.wait_for(
            timeout=20000
        )

        btn.scroll_into_view_if_needed()


        # ------------------------------------------------
        # Download
        # ------------------------------------------------

        step = "download"

        with page.expect_download(
            timeout=60000
        ) as dl:

            btn.evaluate(
                "el => el.click()"
            )

        download = dl.value

        filename = safe_filename(
            case
        ) + ".pdf"

        download.save_as(
            str(
                output_dir / filename
            )
        )

        return True, ""

    except Exception as e:

        return False, (
            f"step={step}: "
            f"{str(e).splitlines()[0][:200]}"
        )


def create_zip(folder):

    zip_path = folder.parent / "Lab_Reports.zip"

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zip_file:

        for file in folder.glob("*.pdf"):

            zip_file.write(
                file,
                arcname=file.name
            )

    return zip_path


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("Settings")

    st.info(
        "Your LabConnect credentials are used only for this session "
        "and are not saved by this app."
    )

    browser_type = st.selectbox(
        "Browser",
        [
            "Chromium",
            "Chrome",
            "Edge"
        ]
    )


# ============================================================
# LOGIN
# ============================================================

st.subheader("1. LabConnect Login")

username = st.text_input(
    "LabConnect Username"
)

password = st.text_input(
    "LabConnect Password",
    type="password"
)


# ============================================================
# EXCEL
# ============================================================

st.subheader("2. Upload Case Excel")

uploaded_file = st.file_uploader(
    "Upload Excel file containing Case No.",
    type=["xlsx", "xls"]
)


# ============================================================
# START
# ============================================================

start = st.button(
    "🚀 Start Download",
    type="primary",
    use_container_width=True
)


# ============================================================
# MAIN PROCESS
# ============================================================

if start:

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
            "Please upload the Excel file."
        )

        st.stop()


    # --------------------------------------------------------
    # Temporary working directory
    # --------------------------------------------------------

    work_dir = Path(
        tempfile.mkdtemp(
            prefix="lab_reports_"
        )
    )

    excel_path = (
        work_dir /
        uploaded_file.name
    )

    output_dir = (
        work_dir /
        "Reports"
    )

    output_dir.mkdir(
        exist_ok=True
    )


    # --------------------------------------------------------
    # Save uploaded Excel
    # --------------------------------------------------------

    with open(
        excel_path,
        "wb"
    ) as f:

        f.write(
            uploaded_file.getbuffer()
        )


    # --------------------------------------------------------
    # Read cases
    # --------------------------------------------------------

    try:

        cases = load_cases(
            excel_path
        )

    except Exception as e:

        st.error(
            f"Could not read Excel: {e}"
        )

        shutil.rmtree(
            work_dir,
            ignore_errors=True
        )

        st.stop()


    if not cases:

        st.error(
            "No Case No. values were found."
        )

        shutil.rmtree(
            work_dir,
            ignore_errors=True
        )

        st.stop()


    st.success(
        f"Found {len(cases)} cases."
    )


    # --------------------------------------------------------
    # Progress
    # --------------------------------------------------------

    progress = st.progress(
        0
    )

    status_text = st.empty()

    log_area = st.empty()

    successful = []

    failed = []


    # --------------------------------------------------------
    # Playwright
    # --------------------------------------------------------

    try:

        with sync_playwright() as p:

            # Browser selection

            if browser_type == "Chrome":

                browser = p.chromium.launch(
                    headless=True,
                    channel="chrome"
                )

            elif browser_type == "Edge":

                browser = p.chromium.launch(
                    headless=True,
                    channel="msedge"
                )

            else:

                browser = p.chromium.launch(
                    headless=True
                )


            context = browser.new_context(
                accept_downloads=True
            )

            page = context.new_page()

            page.set_default_timeout(
                20000
            )


            # ------------------------------------------------
            # Login
            # ------------------------------------------------

            status_text.info(
                "Logging into LabConnect..."
            )

            open_home(
                page,
                username,
                password
            )


            status_text.success(
                "Logged into LabConnect."
            )


            # ------------------------------------------------
            # Process cases
            # ------------------------------------------------

            for i, case in enumerate(
                cases,
                start=1
            ):

                status_text.info(
                    f"Processing {i}/{len(cases)}: {case}"
                )

                success = False
                detail = ""


                # Retry once

                for attempt in range(2):

                    try:

                        success, detail = process_case(
                            page,
                            case,
                            output_dir
                        )

                        if success:

                            break

                        open_home(
                            page,
                            username,
                            password
                        )

                    except Exception as e:

                        detail = str(e)

                        try:

                            open_home(
                                page,
                                username,
                                password
                            )

                        except Exception:

                            pass


                # --------------------------------------------
                # Record result
                # --------------------------------------------

                if success:

                    successful.append(
                        case
                    )

                    log_message = (
                        f"✅ {case} - Downloaded"
                    )

                else:

                    failed.append(
                        {
                            "case": case,
                            "error": detail
                        }
                    )

                    log_message = (
                        f"❌ {case} - {detail}"
                    )


                # --------------------------------------------
                # Update UI
                # --------------------------------------------

                current_progress = (
                    i / len(cases)
                )

                progress.progress(
                    current_progress
                )

                log_area.write(
                    log_message
                )


            browser.close()


    except Exception as e:

        st.error(
            f"Automation error: {e}"
        )

        shutil.rmtree(
            work_dir,
            ignore_errors=True
        )

        st.stop()


    # ========================================================
    # RESULTS
    # ========================================================

    st.divider()

    st.subheader(
        "Results"
    )


    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "Downloaded",
            len(successful)
        )

    with col2:

        st.metric(
            "Failed",
            len(failed)
        )


    # --------------------------------------------------------
    # Failed cases
    # --------------------------------------------------------

    if failed:

        st.warning(
            "Some cases could not be downloaded."
        )

        failed_df = pd.DataFrame(
            failed
        )

        st.dataframe(
            failed_df,
            use_container_width=True
        )


    # --------------------------------------------------------
    # ZIP
    # --------------------------------------------------------

    pdf_files = list(
        output_dir.glob("*.pdf")
    )


    if pdf_files:

        zip_path = create_zip(
            output_dir
        )

        st.success(
            f"{len(pdf_files)} PDF reports are ready."
        )


        with open(
            zip_path,
            "rb"
        ) as f:

            st.download_button(
                label="📦 Download All Reports (ZIP)",
                data=f.read(),
                file_name="Lab_Reports.zip",
                mime="application/zip",
                use_container_width=True
            )


        # Individual PDFs

        with st.expander(
            "Download individual reports"
        ):

            for pdf in pdf_files:

                with open(
                    pdf,
                    "rb"
                ) as f:

                    st.download_button(
                        label=f"📄 {pdf.name}",
                        data=f.read(),
                        file_name=pdf.name,
                        mime="application/pdf"
                    )


    else:

        st.error(
            "No reports were downloaded."
        )


    # --------------------------------------------------------
    # Cleanup
    # --------------------------------------------------------

    try:

        shutil.rmtree(
            work_dir,
            ignore_errors=True
        )

    except Exception:

        pass