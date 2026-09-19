"""
Data Quality & Analytics Platform - Combined App
------------------------------------------------------------------
Three features, switchable from a styled sidebar:

  1. Data Upload      - upload one or more CSV/Excel/PDF/TXT/images
  2. Data Quality Report - full quality report for EVERY uploaded file
  3. Image Sorting    - upload multiple images, auto-sort by category

Run this with: streamlit run main_app.py

Required packages:
  pip install streamlit pandas openpyxl Pillow pypdf tensorflow numpy
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from PIL import Image
from pypdf import PdfReader
from collections import defaultdict


# ==================================================================
# PAGE SETUP
# ==================================================================
st.set_page_config(
    page_title="Clarifile - Upload, Check, Understand Any File",
    page_icon="📊",
    layout="wide",
)

# ------------------------------------------------------------------
# LANDING PAGE
# Shown once per session, before the actual app. A pure black
# background with a centered title and one clear entry button.
# "entered_app" in session_state gates everything below it - the
# rest of the file (sidebar, features) only renders after the
# button is clicked.
# ------------------------------------------------------------------
if "entered_app" not in st.session_state:
    st.session_state["entered_app"] = False

if not st.session_state["entered_app"]:
    # Force a black page background (Streamlit's default theme
    # background is overridden here specifically for the landing
    # page) and hide the sidebar entirely while on this screen.
    st.markdown(
        """
        <style>
        [data-testid="stAppViewContainer"], [data-testid="stHeader"], body {
            background-color: #000000 !important;
        }
        section[data-testid="stSidebar"] {
            display: none;
        }
        .landing-title {
            text-align: center;
            font-size: 3.2rem;
            font-weight: 800;
            color: #FFFFFF;
            margin-top: 14vh;
            margin-bottom: 0.5rem;
            letter-spacing: -0.02em;
        }
        .landing-subtitle {
            text-align: center;
            font-size: 1.15rem;
            color: #9CA3AF;
            margin-bottom: 3rem;
        }
        div[data-testid="stButton"] {
            display: flex;
            justify-content: center;
        }
        div[data-testid="stButton"] button {
            background: linear-gradient(135deg, #2DD4BF, #818CF8);
            color: #000000;
            font-weight: 700;
            font-size: 1.1rem;
            padding: 0.9rem 2.8rem;
            border-radius: 999px;
            border: none;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        div[data-testid="stButton"] button:hover {
            transform: scale(1.04);
            box-shadow: 0 0 24px rgba(45, 212, 191, 0.45);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="landing-title">Welcome to Clarifile</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="landing-subtitle">Upload any file. Get clarity on what\'s inside.</div>',
        unsafe_allow_html=True,
    )

    _, center_col, _ = st.columns([1, 1, 1])
    with center_col:
        if st.button("Enter Platform →", key="enter_platform_button", use_container_width=True):
            st.session_state["entered_app"] = True
            st.rerun()

    # Stops execution here - nothing below runs until the button
    # above is clicked and the page reruns.
    st.stop()


# Each feature gets its own accent color, used consistently in the
# sidebar nav button, the page's icon, and its section headers -
# this is what makes the three features feel visually distinct
# rather than uniform.
FEATURE_ACCENTS = {
    "Data Upload": {"color": "#2DD4BF", "icon": "📤"},
    "Data Quality Report": {"color": "#F5A524", "icon": "🩺"},
    "Image Sorting by Category": {"color": "#818CF8", "icon": "🖼️"},
}

# ------------------------------------------------------------------
# Sidebar navigation styling
# Streamlit doesn't let us fully restyle st.radio, so instead we use
# real st.button widgets (one per feature) and track which one is
# "active" in session_state. Custom CSS below gives each button an
# icon-friendly layout, a colored left border on the active one, and
# a muted look on the inactive ones.
# ------------------------------------------------------------------
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] button {
        width: 100%;
        text-align: left;
        border-radius: 8px;
        border: 1px solid rgba(255,255,255,0.08);
        padding: 0.65rem 0.9rem;
        margin-bottom: 0.4rem;
        font-size: 0.95rem;
        transition: all 0.15s ease;
    }
    section[data-testid="stSidebar"] button:hover {
        border-color: rgba(255,255,255,0.25);
        transform: translateX(2px);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "current_page" not in st.session_state:
    st.session_state["current_page"] = "Data Upload"

st.sidebar.markdown("### Choose a feature")

for feature_name, style in FEATURE_ACCENTS.items():
    is_active = st.session_state["current_page"] == feature_name
    button_label = f"{style['icon']}  {feature_name}"

    # The active button gets a colored border via type="primary"
    # (Streamlit tints primary buttons with the theme accent); the
    # rest stay as plain secondary buttons.
    if st.sidebar.button(
        button_label,
        key=f"nav_{feature_name}",
        type="primary" if is_active else "secondary",
        use_container_width=True,
    ):
        st.session_state["current_page"] = feature_name
        st.rerun()

page = st.session_state["current_page"]
accent = FEATURE_ACCENTS[page]["color"]


# ==================================================================
# SHARED CONSTANTS
# ==================================================================
SUPPORTED_EXTENSIONS = ["csv", "xlsx", "xls", "pdf", "txt", "jpg", "jpeg"]


# ==================================================================
# FEATURE 1: DATA UPLOAD
# ==================================================================

def validate_file(uploaded_file):
    """
    Checks whether an uploaded file is a supported type and non-empty.
    Returns a tuple: (is_valid: bool, extension: str, error_message: str)
    """
    filename = uploaded_file.name
    extension = filename.split(".")[-1].lower() if "." in filename else ""

    if extension not in SUPPORTED_EXTENSIONS:
        return False, extension, (
            f"'.{extension}' is not a supported file type. "
            f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS).upper()}."
        )

    if uploaded_file.size == 0:
        return False, extension, "The uploaded file is empty (0 bytes)."

    return True, extension, ""


def load_csv(uploaded_file):
    return pd.read_csv(uploaded_file)


def load_excel(uploaded_file, extension):
    return pd.read_excel(uploaded_file)


def load_txt(uploaded_file):
    """
    Tries to detect delimited data in a .txt file; falls back to
    loading it as one row per line if no clear delimiter is found.
    """
    raw_bytes = uploaded_file.read()
    text = raw_bytes.decode("utf-8", errors="replace")
    sample_lines = [line for line in text.splitlines()[:10] if line.strip()]

    candidate_delimiters = [",", "\t", ";", "|"]
    detected_delimiter = None

    for delimiter in candidate_delimiters:
        if not sample_lines:
            break
        counts_per_line = [line.count(delimiter) for line in sample_lines]
        if counts_per_line[0] > 0 and len(set(counts_per_line)) == 1:
            detected_delimiter = delimiter
            break

    if detected_delimiter is not None:
        try:
            df = pd.read_csv(BytesIO(raw_bytes), sep=detected_delimiter)
            if df.shape[1] > 1:
                return df
        except Exception:
            pass

    lines = text.splitlines()
    return pd.DataFrame({"line_text": lines})


def load_pdf(uploaded_file):
    """Extracts text per page from a PDF into a DataFrame."""
    reader = PdfReader(BytesIO(uploaded_file.read()))

    if len(reader.pages) == 0:
        raise ValueError("The PDF file contains no pages.")

    page_numbers = []
    page_texts = []

    for page_index, page in enumerate(reader.pages, start=1):
        extracted = page.extract_text() or ""
        page_numbers.append(page_index)
        page_texts.append(extracted.strip())

    return pd.DataFrame({"page_number": page_numbers, "extracted_text": page_texts})


def load_image_metadata(uploaded_file):
    """Reads image metadata (not pixel data) into a DataFrame."""
    image_bytes = uploaded_file.read()
    image = Image.open(BytesIO(image_bytes))

    metadata = {
        "filename": [uploaded_file.name],
        "format": [image.format],
        "width_px": [image.width],
        "height_px": [image.height],
        "color_mode": [image.mode],
        "file_size_kb": [round(len(image_bytes) / 1024, 2)],
    }
    return pd.DataFrame(metadata), image


def load_dataset(uploaded_file, extension):
    """Routes the uploaded file to the correct loader based on extension."""
    if extension == "csv":
        return load_csv(uploaded_file), None
    elif extension in ("xlsx", "xls"):
        return load_excel(uploaded_file, extension), None
    elif extension == "txt":
        return load_txt(uploaded_file), None
    elif extension == "pdf":
        return load_pdf(uploaded_file), None
    elif extension in ("jpg", "jpeg"):
        return load_image_metadata(uploaded_file)
    else:
        raise ValueError(f"No loader available for '.{extension}' files.")


def display_dataset_information(df, filename):
    col1, col2, col3 = st.columns(3)
    col1.metric("Dataset", filename)
    col2.metric("Rows", f"{df.shape[0]:,}")
    col3.metric("Columns", f"{df.shape[1]:,}")


def display_image_preview(image):
    st.subheader("Image Preview")
    st.image(image, use_container_width=True)


def display_dataset_preview(df):
    st.subheader("Dataset Preview")
    st.dataframe(df.head(50), use_container_width=True)
    if df.shape[0] > 50:
        st.caption(f"Showing the first 50 of {df.shape[0]:,} rows.")


def display_column_information(df):
    st.subheader("Column Information")
    column_info = pd.DataFrame({
        "Column Name": df.columns,
        "Data Type": [str(dtype) for dtype in df.dtypes],
        "Non-Null Count": [df[col].notna().sum() for col in df.columns],
    })
    st.dataframe(column_info, use_container_width=True, hide_index=True)


def display_basic_summary(df):
    st.subheader("Summary Statistics")
    summary = df.describe(include="all").transpose()
    st.dataframe(summary, use_container_width=True)

    missing_counts = df.isna().sum()
    total_missing = int(missing_counts.sum())

    if total_missing > 0:
        st.warning(
            f"This dataset contains {total_missing:,} missing value(s) "
            f"across {int((missing_counts > 0).sum())} column(s)."
        )
    else:
        st.info("No missing values were detected in this dataset.")


def process_one_upload(uploaded_file):
    """
    Runs one uploaded file through validation and loading. Returns
    a dict describing the outcome, so the caller can display it (or
    an error) without needing to know the loading details.
    """
    filename = uploaded_file.name

    is_valid, extension, error_message = validate_file(uploaded_file)
    if not is_valid:
        return {"filename": filename, "error": error_message}

    try:
        df, image = load_dataset(uploaded_file, extension)
    except (pd.errors.EmptyDataError, ValueError):
        return {
            "filename": filename,
            "error": (
                "The file was read, but it doesn't appear to contain any "
                "usable data. Please check the file and try again."
            ),
        }
    except (pd.errors.ParserError, UnicodeDecodeError):
        return {
            "filename": filename,
            "error": (
                "This file appears to be corrupted or not formatted "
                f"correctly for a .{extension} file."
            ),
        }
    except Exception:
        return {
            "filename": filename,
            "error": (
                "Something went wrong while reading this file. It may be "
                "corrupted or in an unexpected format."
            ),
        }

    if df.shape[1] == 0:
        return {"filename": filename, "error": "This file doesn't contain any columns that could be read."}

    if df.shape[0] == 0:
        return {"filename": filename, "error": "This dataset has no rows - only column headers were found."}

    return {"filename": filename, "df": df, "image": image, "error": None}


def run_data_upload_page():
    st.markdown(
        f"<h1 style='color:{accent}'>{FEATURE_ACCENTS[page]['icon']} Data Upload</h1>",
        unsafe_allow_html=True,
    )
    st.caption(
        "Upload one or more files. Each one gets its own overview below, "
        "and all of them become available on the Data Quality Report page."
    )

    # accept_multiple_files=True lets someone upload several datasets
    # at once (e.g., a whole batch of CSVs), instead of one at a time.
    uploaded_files = st.file_uploader(
        "Choose one or more files",
        type=SUPPORTED_EXTENSIONS,
        accept_multiple_files=True,
    )
    st.caption("Supported formats: CSV, XLSX, XLS, PDF, TXT, JPG, JPEG")

    if not uploaded_files:
        st.info("👆 Upload one or more files to get started.")
        return

    # This dict holds EVERY successfully loaded dataset, keyed by
    # filename, so the Data Quality Report page can show a full
    # report for each one - not just the most recently uploaded file.
    if "uploaded_datasets" not in st.session_state:
        st.session_state["uploaded_datasets"] = {}

    for uploaded_file in uploaded_files:
        result = process_one_upload(uploaded_file)
        filename = result["filename"]

        with st.container(border=True):
            if result["error"]:
                st.error(f"**{filename}**: {result['error']}")
                continue

            df = result["df"]
            image = result["image"]

            st.success(f"'{filename}' was uploaded and read successfully.")

            # Keep this dataset available for the Data Quality Report page.
            st.session_state["uploaded_datasets"][filename] = df

            display_dataset_information(df, filename)

            if image is not None:
                display_image_preview(image)

            with st.expander("View dataset details", expanded=(len(uploaded_files) == 1)):
                display_dataset_preview(df)
                display_column_information(df)
                display_basic_summary(df)


# ==================================================================
# FEATURE 2: DATA QUALITY REPORT
# ==================================================================
#
# Checks every uploaded dataset for common data quality problems and
# produces an overall quality score (0-100) for EACH one. Reuses
# whatever was uploaded on the Data Upload page (stored in
# st.session_state["uploaded_datasets"]).

def check_missing_values(df):
    """Per-column breakdown of missing values, plus totals."""
    missing_per_column = df.isna().sum()
    total_cells = df.shape[0] * df.shape[1]
    total_missing = int(missing_per_column.sum())
    percent_missing = (total_missing / total_cells * 100) if total_cells > 0 else 0

    breakdown = pd.DataFrame({
        "Column": df.columns,
        "Missing Count": missing_per_column.values,
        "Missing %": (missing_per_column.values / len(df) * 100).round(2),
    })
    breakdown = breakdown[breakdown["Missing Count"] > 0].sort_values(
        "Missing Count", ascending=False
    )

    return breakdown, total_missing, percent_missing


def check_duplicate_rows(df):
    """Fully duplicated rows: count, percentage, and a preview."""
    duplicate_mask = df.duplicated(keep=False)
    duplicate_count = int(df.duplicated(keep="first").sum())
    percent_duplicate = (duplicate_count / len(df) * 100) if len(df) > 0 else 0
    duplicate_rows_preview = df[duplicate_mask].head(20)

    return duplicate_count, percent_duplicate, duplicate_rows_preview


def check_inconsistent_types(df):
    """
    Flags "object" (text) columns that contain a mix of types under
    the hood - e.g., a column that's mostly numbers but has a few
    text entries like "N/A" mixed in.
    """
    inconsistent_columns = []

    for column in df.columns:
        if df[column].dtype == "object":
            non_null_values = df[column].dropna()
            if len(non_null_values) == 0:
                continue

            numeric_convertible = pd.to_numeric(non_null_values, errors="coerce").notna()
            percent_numeric = numeric_convertible.mean() * 100

            if 0 < percent_numeric < 100:
                inconsistent_columns.append({
                    "Column": column,
                    "Issue": f"Mix of numeric and text values ({percent_numeric:.0f}% look numeric)",
                })

    return pd.DataFrame(inconsistent_columns)


def check_outliers(df):
    """
    Uses the IQR (Interquartile Range) method to flag unusually
    extreme values in each numeric column.
    """
    outlier_summary = []
    numeric_columns = df.select_dtypes(include=[np.number]).columns

    for column in numeric_columns:
        values = df[column].dropna()
        if len(values) < 4:
            continue

        q1 = values.quantile(0.25)
        q3 = values.quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr

        outliers = values[(values < lower_bound) | (values > upper_bound)]

        if len(outliers) > 0:
            outlier_summary.append({
                "Column": column,
                "Outlier Count": len(outliers),
                "Outlier %": round(len(outliers) / len(values) * 100, 2),
                "Normal Range": f"{lower_bound:.2f} to {upper_bound:.2f}",
            })

    return pd.DataFrame(outlier_summary)


# Column names that strongly suggest the values should never be
# negative. This is a heuristic (based on common naming patterns),
# not a guarantee - it's meant to catch likely data-entry errors.
NON_NEGATIVE_NAME_HINTS = [
    "age", "price", "cost", "amount", "quantity", "qty", "count",
    "total", "weight", "height", "length", "distance", "salary",
    "income", "duration", "score", "rating", "fare",
]


def check_invalid_values(df):
    """
    Flags likely-invalid values based on common-sense rules:
      - Negative numbers in columns whose name suggests they should
        never be negative (e.g., "age", "price", "quantity").
      - Values of exactly 0 in columns like "age" that are almost
        certainly placeholders for missing data, not real zeros.
    This is heuristic, not exhaustive - it's meant to catch obvious,
    common data-entry mistakes.
    """
    invalid_summary = []
    numeric_columns = df.select_dtypes(include=[np.number]).columns

    for column in numeric_columns:
        column_lower = column.lower()
        matches_hint = any(hint in column_lower for hint in NON_NEGATIVE_NAME_HINTS)

        if not matches_hint:
            continue

        values = df[column].dropna()
        if len(values) == 0:
            continue

        negative_count = int((values < 0).sum())
        if negative_count > 0:
            invalid_summary.append({
                "Column": column,
                "Issue": "Negative values found",
                "Count": negative_count,
                "Example": values[values < 0].iloc[0],
            })

    return pd.DataFrame(invalid_summary)


def check_formatting_issues(df):
    """
    Flags text columns with formatting inconsistencies that are easy
    to miss visually but cause real problems - e.g., "USA" and " USA "
    or "usa" being treated as different categories.
    """
    formatting_issues = []
    text_columns = df.select_dtypes(include=["object"]).columns

    for column in text_columns:
        values = df[column].dropna().astype(str)
        if len(values) == 0:
            continue

        has_leading_trailing_space = (values != values.str.strip()).sum()
        # Compare how many unique values exist normally vs. after
        # lowercasing + stripping whitespace - a big drop means many
        # "different" values are actually the same thing, just
        # formatted inconsistently (e.g., "USA" vs "usa" vs " USA").
        unique_raw = values.nunique()
        unique_normalized = values.str.strip().str.lower().nunique()
        collapsed_count = unique_raw - unique_normalized

        if has_leading_trailing_space > 0 or collapsed_count > 0:
            formatting_issues.append({
                "Column": column,
                "Extra Whitespace": int(has_leading_trailing_space),
                "Inconsistent Capitalization/Spacing": int(collapsed_count),
            })

    return pd.DataFrame(formatting_issues)


def check_constant_columns(df):
    """
    Flags columns where every non-missing value is identical (zero
    variance). These columns carry no information and are usually
    either a data export mistake or a leftover placeholder column.
    """
    constant_columns = []

    for column in df.columns:
        non_null_values = df[column].dropna()
        if len(non_null_values) == 0:
            continue
        if non_null_values.nunique() == 1:
            constant_columns.append({
                "Column": column,
                "Constant Value": non_null_values.iloc[0],
            })

    return pd.DataFrame(constant_columns)


def check_high_cardinality(df):
    """
    Flags text columns where almost every value is unique (>95%).
    This usually means the column is actually an ID or free-text
    field, not a category - useful to know since it changes how the
    column should be analyzed later (e.g., it shouldn't be charted
    as a category).
    """
    high_cardinality_columns = []

    for column in df.select_dtypes(include=["object"]).columns:
        non_null_values = df[column].dropna()
        if len(non_null_values) < 10:
            continue

        percent_unique = non_null_values.nunique() / len(non_null_values) * 100
        if percent_unique > 95:
            high_cardinality_columns.append({
                "Column": column,
                "Unique %": round(percent_unique, 1),
                "Likely Type": "ID or free text (not a category)",
            })

    return pd.DataFrame(high_cardinality_columns)


def calculate_quality_score(percent_missing, percent_duplicate, inconsistent_df,
                             outlier_df, invalid_df, formatting_df, constant_df):
    """
    Combines every check into a single 0-100 quality score. Starts
    at 100 and subtracts penalty points for each issue found,
    weighted by severity and capped so no single issue can dominate
    the score by itself.
    """
    score = 100.0

    score -= min(percent_missing * 0.5, 25)
    score -= min(percent_duplicate * 0.5, 15)
    score -= min(len(inconsistent_df) * 4, 15)
    score -= min(len(outlier_df) * 2, 10)
    score -= min(len(invalid_df) * 5, 15)
    score -= min(len(formatting_df) * 3, 10)
    score -= min(len(constant_df) * 3, 10)

    return max(round(score), 0)


def display_quality_score(score, accent_color):
    if score >= 90:
        label, color = "Excellent", "#22C55E"
    elif score >= 75:
        label, color = "Good", "#3B82F6"
    elif score >= 50:
        label, color = "Needs Attention", "#F59E0B"
    else:
        label, color = "Poor", "#EF4444"

    st.markdown(
        f"""
        <div style="display:flex; align-items:baseline; gap:0.75rem; margin-bottom:0.25rem;">
            <span style="font-size:2rem; font-weight:700;">{score}/100</span>
            <span style="font-size:1.1rem; font-weight:600; color:{color};">{label}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.progress(score / 100)


def render_quality_report_for_one_dataset(filename, df, accent_color):
    """Runs every check on one dataset and renders its full report."""
    missing_breakdown, total_missing, percent_missing = check_missing_values(df)
    duplicate_count, percent_duplicate, duplicate_preview = check_duplicate_rows(df)
    inconsistent_df = check_inconsistent_types(df)
    outlier_df = check_outliers(df)
    invalid_df = check_invalid_values(df)
    formatting_df = check_formatting_issues(df)
    constant_df = check_constant_columns(df)
    high_cardinality_df = check_high_cardinality(df)

    score = calculate_quality_score(
        percent_missing, percent_duplicate, inconsistent_df,
        outlier_df, invalid_df, formatting_df, constant_df
    )

    display_quality_score(score, accent_color)
    st.caption(f"{df.shape[0]:,} rows x {df.shape[1]:,} columns")

    st.divider()

    checks = [
        ("🔍 Missing Values", total_missing > 0,
         f"Found **{total_missing:,}** missing value(s) ({percent_missing:.1f}% of all cells).",
         missing_breakdown),
        ("🔁 Duplicate Rows", duplicate_count > 0,
         f"Found **{duplicate_count:,}** duplicate row(s) ({percent_duplicate:.1f}% of the dataset).",
         duplicate_preview),
        ("⚠️ Inconsistent Data Types", not inconsistent_df.empty,
         f"Found **{len(inconsistent_df)}** column(s) with mixed data types.",
         inconsistent_df),
        ("📈 Outliers", not outlier_df.empty,
         f"Found potential outliers in **{len(outlier_df)}** numeric column(s) (IQR method).",
         outlier_df),
        ("🚫 Invalid Values", not invalid_df.empty,
         f"Found **{len(invalid_df)}** column(s) with values that look invalid (e.g., negative ages).",
         invalid_df),
        ("✂️ Formatting Issues", not formatting_df.empty,
         f"Found **{len(formatting_df)}** text column(s) with inconsistent spacing/capitalization.",
         formatting_df),
        ("🟰 Constant Columns", not constant_df.empty,
         f"Found **{len(constant_df)}** column(s) where every value is identical.",
         constant_df),
        ("🆔 High-Cardinality Columns", not high_cardinality_df.empty,
         f"Found **{len(high_cardinality_df)}** column(s) that look like IDs or free text, not categories.",
         high_cardinality_df),
    ]

    for title, has_issue, message, detail_df in checks:
        st.subheader(title)
        if has_issue:
            st.write(message)
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
        else:
            st.success("No issues detected.")
        st.divider()


def run_data_quality_report_page():
    st.markdown(
        f"<h1 style='color:{accent}'>{FEATURE_ACCENTS[page]['icon']} Data Quality Report</h1>",
        unsafe_allow_html=True,
    )
    st.caption(
        "A full quality report for every file you've uploaded - not just "
        "the most recent one."
    )

    datasets = st.session_state.get("uploaded_datasets", {})

    if not datasets:
        st.warning(
            "No datasets found. Please upload one or more files on the "
            "**Data Upload** page first, then come back here."
        )
        return

    st.info(f"Analyzing **{len(datasets)}** dataset(s).")

    # Each uploaded file gets its own titled, collapsible section, so
    # uploading several files shows several full reports at once,
    # rather than only ever showing the most recent one.
    for filename, df in datasets.items():
        with st.expander(f"📄 {filename}", expanded=(len(datasets) == 1)):
            render_quality_report_for_one_dataset(filename, df, accent)


# ==================================================================
# FEATURE 3: IMAGE SORTING BY CATEGORY
# ==================================================================

@st.cache_resource
def load_classification_model():
    """
    Loads the pretrained MobileNetV2 model once and caches it.
    Returns None if TensorFlow fails to load (e.g., a Windows DLL
    error), so the rest of the app can keep working instead of
    crashing entirely.
    """
    try:
        from tensorflow.keras.applications.mobilenet_v2 import MobileNetV2
        return MobileNetV2(weights="imagenet")
    except Exception as e:
        st.session_state["tf_load_error"] = str(e)
        return None


def classify_image(model, image: Image.Image) -> str:
    """Returns the single best-guess category label for one image."""
    from tensorflow.keras.applications.mobilenet_v2 import (
        preprocess_input,
        decode_predictions,
    )

    resized = image.convert("RGB").resize((224, 224))
    array = np.array(resized)
    array = np.expand_dims(array, axis=0)
    array = preprocess_input(array)

    predictions = model.predict(array, verbose=0)
    decoded = decode_predictions(predictions, top=1)[0][0]

    label = decoded[1].replace("_", " ").title()
    return label


def run_image_sorting_page():
    st.markdown(
        f"<h1 style='color:{accent}'>{FEATURE_ACCENTS[page]['icon']} Image Sorting by Category</h1>",
        unsafe_allow_html=True,
    )
    st.caption("Upload multiple images and they'll be automatically grouped by what they show.")

    uploaded_files = st.file_uploader(
        "Upload images",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("👆 Upload one or more images to sort them by category.")
        return

    st.info(f"Processing {len(uploaded_files)} image(s)... this may take a moment.")

    model = load_classification_model()

    if model is None:
        st.error(
            "Image sorting isn't available right now because TensorFlow "
            "couldn't load on this machine. This is usually a missing "
            "Microsoft Visual C++ Redistributable on Windows, not a bug "
            "in the app. Install it from https://aka.ms/vs/17/release/vc_redist.x64.exe, "
            "restart your computer, and try again. The Data Upload feature "
            "will still work in the meantime."
        )
        return

    categorized_images = defaultdict(list)
    progress_bar = st.progress(0)

    for index, uploaded_file in enumerate(uploaded_files):
        image = Image.open(uploaded_file)
        category = classify_image(model, image)
        categorized_images[category].append((uploaded_file.name, image))
        progress_bar.progress((index + 1) / len(uploaded_files))

    st.success(f"Done! Found {len(categorized_images)} categories.")

    for category, images_in_category in categorized_images.items():
        with st.expander(f"📁 {category} ({len(images_in_category)} image(s))", expanded=True):
            columns = st.columns(4)
            for i, (filename, image) in enumerate(images_in_category):
                with columns[i % 4]:
                    st.image(image, caption=filename, use_container_width=True)


# ==================================================================
# MAIN ROUTER
# ==================================================================
if page == "Data Upload":
    run_data_upload_page()
elif page == "Data Quality Report":
    run_data_quality_report_page()
elif page == "Image Sorting by Category":
    run_image_sorting_page()
