"""
Data Quality & Analytics Platform - Combined App
------------------------------------------------------------------
This single file combines two features into one app with a sidebar
to switch between them:

  1. Data Upload      - upload CSV/Excel/PDF/TXT/images, see an overview
  2. Image Sorting    - upload multiple images, auto-sort by category

Run this with: streamlit run main_app.py

Required packages:
  pip install streamlit pandas openpyxl Pillow pypdf tensorflow
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
st.set_page_config(page_title="Data Quality & Analytics Platform", layout="wide")

# st.sidebar puts this selection box on the left-hand side, so the
# user can switch between features without leaving the app.
page = st.sidebar.radio(
    "Choose a feature:",
    ["Data Upload", "Data Quality Report", "Image Sorting by Category"]
)


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


def run_data_upload_page():
    st.title("Data Quality & Analytics Platform")
    st.caption("Part 1: Data Upload")
    st.header("Upload Your Dataset")

    uploaded_file = st.file_uploader("Choose a file", type=SUPPORTED_EXTENSIONS)
    st.caption("Supported formats: CSV, XLSX, XLS, PDF, TXT, JPG, JPEG")

    if uploaded_file is None:
        st.info("👆 Upload a file to get started.")
        return

    filename = uploaded_file.name

    is_valid, extension, error_message = validate_file(uploaded_file)
    if not is_valid:
        st.error(error_message)
        return

    try:
        df, image = load_dataset(uploaded_file, extension)
    except (pd.errors.EmptyDataError, ValueError):
        st.error(
            "The file was read, but it doesn't appear to contain any "
            "usable data. Please check the file and try again."
        )
        return
    except (pd.errors.ParserError, UnicodeDecodeError):
        st.error(
            "This file appears to be corrupted or not formatted correctly "
            f"for a .{extension} file. Please verify the file and try "
            "uploading it again."
        )
        return
    except Exception:
        st.error(
            "Something went wrong while reading this file. It may be "
            "corrupted or in an unexpected format. Please try a "
            "different file."
        )
        return

    if df.shape[1] == 0:
        st.error("This file doesn't contain any columns that could be read.")
        return

    if df.shape[0] == 0:
        st.warning("This dataset has no rows - only column headers were found.")
        return

    st.success(f"'{filename}' was uploaded and read successfully.")

    # Save the dataset in session_state so the Data Quality Report page
    # can use it without asking the user to upload it a second time.
    st.session_state["uploaded_df"] = df
    st.session_state["uploaded_filename"] = filename

    display_dataset_information(df, filename)

    if image is not None:
        display_image_preview(image)

    display_dataset_preview(df)
    display_column_information(df)
    display_basic_summary(df)


# ==================================================================
# FEATURE 2: DATA QUALITY REPORT
# ==================================================================
#
# This feature checks an uploaded dataset for common data quality
# problems and produces an overall quality score (0-100). It reuses
# whatever dataset was uploaded on the Data Upload page (stored in
# st.session_state), so the user doesn't need to upload it twice.

def check_missing_values(df):
    """
    Returns a per-column breakdown of missing values, plus the total
    count and percentage across the whole dataset.
    """
    missing_per_column = df.isna().sum()
    total_cells = df.shape[0] * df.shape[1]
    total_missing = int(missing_per_column.sum())
    percent_missing = (total_missing / total_cells * 100) if total_cells > 0 else 0

    breakdown = pd.DataFrame({
        "Column": df.columns,
        "Missing Count": missing_per_column.values,
        "Missing %": (missing_per_column.values / len(df) * 100).round(2),
    })
    # Only show columns that actually have missing values, sorted worst-first.
    breakdown = breakdown[breakdown["Missing Count"] > 0].sort_values(
        "Missing Count", ascending=False
    )

    return breakdown, total_missing, percent_missing


def check_duplicate_rows(df):
    """
    Returns the number of fully duplicated rows and their percentage
    of the dataset, plus a preview of the duplicate rows themselves.
    """
    duplicate_mask = df.duplicated(keep=False)
    duplicate_count = int(df.duplicated(keep="first").sum())
    percent_duplicate = (duplicate_count / len(df) * 100) if len(df) > 0 else 0
    duplicate_rows_preview = df[duplicate_mask].head(20)

    return duplicate_count, percent_duplicate, duplicate_rows_preview


def check_inconsistent_types(df):
    """
    Flags "object" (text) columns that actually contain a mix of
    types under the hood - e.g., a column that's mostly numbers but
    has a few text entries like "N/A" or "unknown" mixed in. This is
    a common real-world data quality issue that a plain dtype check
    misses, since pandas just labels the whole column as "object."
    """
    inconsistent_columns = []

    for column in df.columns:
        # Only check "object" columns - numeric/datetime columns are
        # already internally consistent by definition.
        if df[column].dtype == "object":
            non_null_values = df[column].dropna()
            if len(non_null_values) == 0:
                continue

            # Try to see how many values in this "text" column could
            # actually be parsed as numbers.
            numeric_convertible = pd.to_numeric(non_null_values, errors="coerce").notna()
            percent_numeric = numeric_convertible.mean() * 100

            # If SOME but not ALL values are numeric, that's a mixed-type
            # column - e.g., a "quantity" column with "5", "10", "unknown".
            if 0 < percent_numeric < 100:
                inconsistent_columns.append({
                    "Column": column,
                    "Issue": f"Mix of numeric and text values ({percent_numeric:.0f}% look numeric)",
                })

    return pd.DataFrame(inconsistent_columns)


def check_outliers(df):
    """
    Uses the IQR (Interquartile Range) method - a standard statistical
    approach - to flag unusually extreme values in each numeric column.
    Any value more than 1.5x the IQR beyond the 25th/75th percentile is
    flagged as a potential outlier.
    """
    outlier_summary = []
    numeric_columns = df.select_dtypes(include=[np.number]).columns

    for column in numeric_columns:
        values = df[column].dropna()
        if len(values) < 4:
            # Not enough data points to meaningfully calculate quartiles.
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


def calculate_quality_score(df, total_missing, percent_missing, duplicate_count,
                             percent_duplicate, inconsistent_df, outlier_df):
    """
    Combines every check above into a single 0-100 quality score.
    Starts at 100 and subtracts penalty points for each issue found,
    weighted by how severe/widespread that issue is. This is a
    reasonable, transparent scoring approach - not the only valid one,
    but it's easy to explain and adjust later.
    """
    score = 100.0

    # Missing values: penalize proportionally, capped so a single
    # issue can't zero out the whole score by itself.
    score -= min(percent_missing * 0.5, 30)

    # Duplicate rows: similarly capped.
    score -= min(percent_duplicate * 0.5, 20)

    # Inconsistent types: fixed penalty per affected column.
    score -= min(len(inconsistent_df) * 5, 20)

    # Outliers: small penalty per affected column, since some
    # outliers are expected/legitimate in real data.
    score -= min(len(outlier_df) * 3, 15)

    return max(round(score), 0)


def display_quality_score(score):
    """Shows the overall score with a color-coded label."""
    if score >= 90:
        label, color = "Excellent", "green"
    elif score >= 75:
        label, color = "Good", "blue"
    elif score >= 50:
        label, color = "Needs Attention", "orange"
    else:
        label, color = "Poor", "red"

    st.markdown(f"## Overall Data Quality Score: {score}/100")
    st.markdown(f":{color}[**{label}**]")
    st.progress(score / 100)


def run_data_quality_report_page():
    st.title("Data Quality Report")
    st.caption("Part 2: Automated data quality checks")

    # Reuse the dataset uploaded on the Data Upload page, if any.
    df = st.session_state.get("uploaded_df")
    filename = st.session_state.get("uploaded_filename", "your dataset")

    if df is None:
        st.warning(
            "No dataset found. Please upload a file on the **Data Upload** "
            "page first, then come back here."
        )
        return

    st.info(f"Analyzing: **{filename}** ({df.shape[0]:,} rows, {df.shape[1]:,} columns)")

    # Run every check.
    missing_breakdown, total_missing, percent_missing = check_missing_values(df)
    duplicate_count, percent_duplicate, duplicate_preview = check_duplicate_rows(df)
    inconsistent_df = check_inconsistent_types(df)
    outlier_df = check_outliers(df)

    # Combine into one overall score.
    score = calculate_quality_score(
        df, total_missing, percent_missing, duplicate_count,
        percent_duplicate, inconsistent_df, outlier_df
    )

    display_quality_score(score)

    st.divider()

    # ------------------------------------------------------------------
    # Missing values section
    # ------------------------------------------------------------------
    st.subheader("🔍 Missing Values")
    if total_missing > 0:
        st.write(
            f"Found **{total_missing:,}** missing value(s) "
            f"({percent_missing:.1f}% of all cells)."
        )
        st.dataframe(missing_breakdown, use_container_width=True, hide_index=True)
    else:
        st.success("No missing values detected.")

    st.divider()

    # ------------------------------------------------------------------
    # Duplicate rows section
    # ------------------------------------------------------------------
    st.subheader("🔁 Duplicate Rows")
    if duplicate_count > 0:
        st.write(
            f"Found **{duplicate_count:,}** duplicate row(s) "
            f"({percent_duplicate:.1f}% of the dataset)."
        )
        st.caption("Preview of duplicated rows (showing up to 20):")
        st.dataframe(duplicate_preview, use_container_width=True)
    else:
        st.success("No duplicate rows detected.")

    st.divider()

    # ------------------------------------------------------------------
    # Inconsistent data types section
    # ------------------------------------------------------------------
    st.subheader("⚠️ Inconsistent Data Types")
    if not inconsistent_df.empty:
        st.write(
            f"Found **{len(inconsistent_df)}** column(s) with mixed "
            "data types (e.g., numbers and text mixed together)."
        )
        st.dataframe(inconsistent_df, use_container_width=True, hide_index=True)
    else:
        st.success("No inconsistent data types detected.")

    st.divider()

    # ------------------------------------------------------------------
    # Outliers section
    # ------------------------------------------------------------------
    st.subheader("📈 Outliers")
    if not outlier_df.empty:
        st.write(
            f"Found potential outliers in **{len(outlier_df)}** numeric "
            "column(s), based on the IQR method (values far outside the "
            "typical range for that column)."
        )
        st.dataframe(outlier_df, use_container_width=True, hide_index=True)
    else:
        st.success("No significant outliers detected in numeric columns.")


# ==================================================================
# FEATURE 2: IMAGE SORTING BY CATEGORY
# ==================================================================

@st.cache_resource
def load_classification_model():
    """
    Loads the pretrained MobileNetV2 model once and caches it, so it
    doesn't reload every time a new image is uploaded. Imported
    inside the function so this heavy library only loads when the
    Image Sorting page is actually used.

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
    st.title("🖼️ Image Sorting by Category")
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
# Runs whichever page the user picked from the sidebar above.
# ==================================================================
if page == "Data Upload":
    run_data_upload_page()
elif page == "Data Quality Report":
    run_data_quality_report_page()
elif page == "Image Sorting by Category":
    run_image_sorting_page()
