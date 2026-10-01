"""Data Quality & Analytics Platform
Unified Python application: Part 1 Data Upload, Part 2 Data Quality Report,
and Part 3 Analytics & Charts.

Run with:
    streamlit run main_app.py

The application is pure Python and can be developed/edited in VS Code or
RStudio. RStudio should run this file with a configured Python environment
(e.g. through its Terminal or reticulate environment).
"""

# ================= PART 1: DATA UPLOAD =================
"""
Data Quality & Analytics Platform
Part 1: Data Upload (merged implementation)
This file contains ONLY the Data Upload module. The Data Quality Report (#2),
Charts / Analytics (#3) and later modules are intentionally not built here.
They will read the uploaded dataset through get_uploaded_dataset(), defined
in the DATA HAND-OFF section below.
"""
# ==========================================
# IMPORTS
# ==========================================
# Import re so uploaded file names can be cleaned of unsafe characters.
import re
# Import os so only the base file name (never a folder path) is trusted.
import os
# Import zipfile so corrupted .xlsx workbooks (which are zip archives) can be detected.
import zipfile
# Import dataclass to define a small, typed container for an uploaded dataset.
from dataclasses import dataclass
# Import datetime to record when a dataset was uploaded.
from datetime import datetime
# Import BytesIO / StringIO so uploaded bytes/text can be read like files, entirely in memory.
from io import BytesIO, StringIO
# Import Streamlit to build the web interface (uploader, tables, messages).
import streamlit as st
# Import pandas, the DataFrame library every uploaded dataset is loaded into.
import pandas as pd
# Import Pillow's Image to open uploaded JPG/JPEG files and read their metadata.
from PIL import Image, UnidentifiedImageError
# Import pypdf to extract text from uploaded PDF files, plus its read error type.
from pypdf import PdfReader
from pypdf.errors import PdfReadError
# ==========================================
# CONFIGURATION
# ==========================================
# Visible name of the platform; shown as the page title and browser-tab title.
APP_TITLE = "Data Quality & Analytics Platform"
# Single source of truth for accepted extensions. Both the uploader widget and
# the validation function use this list so they can never disagree.
SUPPORTED_EXTENSIONS = ["csv", "xlsx", "xls", "pdf", "txt", "jpg", "jpeg"]
# Maximum accepted upload size (MB) so a huge file cannot exhaust local memory.
MAX_FILE_SIZE_MB = 200
# Number of rows shown in the preview table so large datasets do not slow the browser.
PREVIEW_ROW_LIMIT = 50
# Key under which the loaded dataset is kept in Streamlit's session state,
# so later modules (Quality Report, Charts) can retrieve it.
SESSION_KEY = "uploaded_dataset"
# Text encodings tried in order when decoding CSV/TXT files. latin-1 is last
# because it accepts any byte, guaranteeing a result for text-like files.
TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
# Delimiters checked when deciding whether a .txt file is really tabular data.
TXT_CANDIDATE_DELIMITERS = [",", "\t", ";", "|"]
# ==========================================
# ERROR HANDLING
# ==========================================
class UploadError(Exception):
    """
    Raised whenever an upload cannot be processed. The message is written for
    the end user (no technical details), so the interface can show it directly
    with st.error() instead of a raw traceback.
    """
# Generic message for unexpected failures; deliberately hides internal details.
UNEXPECTED_ERROR_MESSAGE = (
    "Something went wrong while reading this file. It may be corrupted or in "
    "an unexpected format. Please try a different file."
)
# ==========================================
# DATA HAND-OFF TO LATER MODULES
# ==========================================
# Container describing one successfully uploaded dataset. Later modules receive
# this object, so they get the DataFrame plus context (name, type, size, time).
@dataclass
class UploadedDataset:
    # Cleaned file name shown to the user (e.g. "customers.csv").
    name: str
    # Lower-case extension without the dot (e.g. "csv").
    extension: str
    # The dataset itself, loaded into memory as a pandas DataFrame.
    dataframe: pd.DataFrame
    # PIL image for JPG/JPEG uploads (None for every other file type).
    image: object
    # Size of the original upload in bytes.
    file_size_bytes: int
    # Moment the file was processed.
    uploaded_at: datetime
# Save the dataset in session state so it survives Streamlit reruns and can be
# read by Part 2 (Data Quality Report) and Part 3 (Charts).
def store_uploaded_dataset(dataset):
    st.session_state[SESSION_KEY] = dataset
# Public accessor for later modules: returns the current UploadedDataset,
# or None when nothing valid has been uploaded yet.
def get_uploaded_dataset():
    return st.session_state.get(SESSION_KEY)
# Remove any stored dataset, used when the file is removed or fails validation,
# so later modules never analyse stale or invalid data.
def clear_uploaded_dataset():
    st.session_state.pop(SESSION_KEY, None)
# ==========================================
# FILE VALIDATION
# ==========================================
# Clean an uploaded file name before it is displayed or stored.
# Browsers can send odd names, so only the base name is kept, control
# characters are removed, and the length is limited. The file is never
# written to disk or executed, so the name is used for display/type only.
def sanitize_filename(filename):
    # Drop any folder components (handles both / and \ separators).
    base_name = os.path.basename(str(filename).replace("\\", "/"))
    # Remove control characters that could break the display.
    base_name = re.sub(r"[\x00-\x1f\x7f]", "", base_name).strip()
    # Fall back to a neutral name if nothing usable remains.
    if not base_name:
        base_name = "uploaded_file"
    # Cap the length so very long names do not break the layout.
    return base_name[:150]
# Extract the lower-case extension so ".CSV" and ".csv" are treated the same.
def get_extension(filename):
    return filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
# Convert a byte count into a readable string (e.g. 2048 -> "2.0 KB").
def format_file_size(size_bytes):
    size = float(size_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
# Check that the file's actual bytes plausibly match its extension, so a
# renamed executable or random binary is rejected instead of parsed.
# Returns True when the content is consistent with the extension.
def content_matches_extension(extension, raw_bytes):
    # Only the first bytes ("magic number") are needed to identify each format.
    header = raw_bytes[:1024]
    if extension == "pdf":
        # Real PDFs contain the "%PDF" marker at the start.
        return b"%PDF" in header
    if extension in ("jpg", "jpeg"):
        # JPEG images always start with these three bytes.
        return header.startswith(b"\xff\xd8\xff")
    if extension == "xlsx":
        # .xlsx files are zip archives, which begin with "PK".
        return header.startswith(b"PK")
    if extension == "xls":
        # Legacy .xls is an OLE container; some tools save zip-based files as .xls.
        return header.startswith(b"\xd0\xcf\x11\xe0") or header.startswith(b"PK")
    # CSV/TXT are text: a NUL byte in the first 8 KB means the file is binary.
    return b"\x00" not in raw_bytes[:8192]
# Validate name, size and content BEFORE any parsing so obviously bad uploads
# are rejected with a clear message. Returns (is_valid, extension, error_message).
def validate_file(filename, raw_bytes):
    extension = get_extension(filename)
    # Prevent unsupported file types from being processed.
    if extension not in SUPPORTED_EXTENSIONS:
        shown = extension if extension else "(none)"
        return False, extension, (
            f"'.{shown}' is not a supported file type. "
            f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS).upper()}."
        )
    # Reject zero-byte files.
    if len(raw_bytes) == 0:
        return False, extension, "The uploaded file is empty (0 bytes)."
    # Reject files larger than the configured limit.
    if len(raw_bytes) > MAX_FILE_SIZE_MB * 1024 * 1024:
        return False, extension, (
            f"The file is too large ({format_file_size(len(raw_bytes))}). "
            f"The maximum allowed size is {MAX_FILE_SIZE_MB} MB."
        )
    # Reject files whose content does not match the claimed type.
    if not content_matches_extension(extension, raw_bytes):
        return False, extension, (
            f"This file does not look like a real .{extension} file. "
            "It may be corrupted or have the wrong extension."
        )
    # All basic checks passed.
    return True, extension, ""
# ==========================================
# DATA UPLOAD (LOADERS)
# ==========================================
# Each loader turns raw uploaded bytes into a DataFrame or raises UploadError
# with a user-friendly message. Working from bytes (not the Streamlit widget
# object) avoids file-pointer problems on reruns and keeps loaders testable.
# Decode text bytes using the first encoding that works. Handles the common
# case of CSV/TXT files saved by Excel (cp1252) rather than UTF-8.
def decode_text(raw_bytes):
    for encoding in TEXT_ENCODINGS:
        try:
            return raw_bytes.decode(encoding)
        except UnicodeDecodeError:
            # Try the next encoding in the list.
            continue
    # Unreachable in practice (latin-1 accepts every byte); kept as a safe fallback.
    raise UploadError("The text in this file could not be decoded. Please save it as UTF-8 and try again.")
# Load a CSV file into a DataFrame.
def load_csv(raw_bytes):
    text = decode_text(raw_bytes)
    try:
        return pd.read_csv(StringIO(text))
    except pd.errors.EmptyDataError:
        # No columns/rows at all (e.g. a file with only blank lines).
        raise UploadError("The CSV file doesn't contain any usable data.")
    except pd.errors.ParserError:
        # Ragged rows, broken quoting, etc.
        raise UploadError(
            "This CSV file appears to be corrupted or not formatted correctly. "
            "Please verify the file and try again."
        )
# Load the first sheet of an Excel workbook (.xlsx via openpyxl, .xls via xlrd).
def load_excel(raw_bytes, extension):
    # Zip-based workbooks (start with "PK") use openpyxl; old OLE files use xlrd.
    engine = "openpyxl" if raw_bytes[:2] == b"PK" else "xlrd"
    try:
        return pd.read_excel(BytesIO(raw_bytes), engine=engine)
    except ImportError:
        # The engine package is not installed.
        package = "openpyxl" if engine == "openpyxl" else "xlrd"
        raise UploadError(
            f"Reading .{extension} files requires the '{package}' package. "
            "Install it with: pip install -r requirements.txt"
        )
    except (zipfile.BadZipFile, ValueError, KeyError, OSError):
        # Damaged workbook, or a file that is not really an Excel file.
        raise UploadError(
            f"This file appears to be corrupted or is not a valid .{extension} "
            "workbook. Please verify the file and try again."
        )
    except Exception as error:
        # xlrd raises its own error type for damaged .xls files; treat any
        # remaining Excel parsing failure as a corrupted workbook.
        if error.__class__.__name__ in ("XLRDError", "InvalidFileException"):
            raise UploadError(
                f"This file appears to be corrupted or is not a valid .{extension} "
                "workbook. Please verify the file and try again."
            )
        raise
# Load a .txt file: delimited data becomes columns, otherwise one row per line.
def load_txt(raw_bytes):
    text = decode_text(raw_bytes)
    # Look at up to 10 non-blank lines to judge whether the text is tabular.
    sample_lines = [line for line in text.splitlines()[:10] if line.strip()]
    # Trust a delimiter only if it appears the same number of times on every
    # sampled line (a strong sign of real structure, unlike prose with commas).
    detected_delimiter = None
    for delimiter in TXT_CANDIDATE_DELIMITERS:
        if not sample_lines:
            break
        counts = [line.count(delimiter) for line in sample_lines]
        if counts[0] > 0 and len(set(counts)) == 1:
            detected_delimiter = delimiter
            break
    # Try structured parsing first.
    if detected_delimiter is not None:
        try:
            df = pd.read_csv(StringIO(text), sep=detected_delimiter)
            if df.shape[1] > 1:
                return df
        except (pd.errors.ParserError, pd.errors.EmptyDataError):
            # Fall through to the line-by-line fallback below.
            pass
    # Fallback: one row per line so .txt files always give something to inspect.
    return pd.DataFrame({"line_text": text.splitlines()})
# Load a PDF as a DataFrame with one row per page (page_number, extracted_text).
# Table reconstruction is a much bigger feature and is out of scope for Part 1.
def load_pdf(raw_bytes):
    try:
        reader = PdfReader(BytesIO(raw_bytes))
        # Encrypted PDFs are attempted with an empty password (common "owner-only" protection).
        if reader.is_encrypted and not reader.decrypt(""):
            raise UploadError("This PDF is password-protected and cannot be read.")
        if len(reader.pages) == 0:
            raise UploadError("The PDF file contains no pages.")
        page_numbers = []
        page_texts = []
        # Extract text per page; scanned pages without OCR simply yield "".
        for page_index, page in enumerate(reader.pages, start=1):
            page_numbers.append(page_index)
            page_texts.append((page.extract_text() or "").strip())
        return pd.DataFrame({"page_number": page_numbers, "extracted_text": page_texts})
    except UploadError:
        # Already user-friendly; pass it through unchanged.
        raise
    except (PdfReadError, ValueError, KeyError, OSError, NotImplementedError):
        raise UploadError(
            "This PDF appears to be corrupted or cannot be read. "
            "Please verify the file and try again."
        )
# Load a JPG/JPEG as a one-row DataFrame of metadata. Images aren't tabular,
# so the DataFrame describes the file; the picture itself is returned too so
# the UI can display it. Returns (metadata_dataframe, PIL_image).
def load_image(raw_bytes, filename):
    try:
        image = Image.open(BytesIO(raw_bytes))
        # load() fully decodes the pixels, catching truncated/corrupted images now.
        image.load()
    except (UnidentifiedImageError, OSError, ValueError, SyntaxError):
        raise UploadError(
            "This image appears to be corrupted or is not a valid JPG/JPEG file."
        )
    metadata = pd.DataFrame({
        "filename": [filename],
        "format": [image.format],
        "width_px": [image.width],
        "height_px": [image.height],
        "color_mode": [image.mode],
        "file_size_kb": [round(len(raw_bytes) / 1024, 2)],
    })
    return metadata, image
# Route validated bytes to the correct loader by extension.
# Returns (dataframe, image_or_None).
def load_dataset(raw_bytes, extension, filename):
    if extension == "csv":
        return load_csv(raw_bytes), None
    if extension in ("xlsx", "xls"):
        return load_excel(raw_bytes, extension), None
    if extension == "txt":
        return load_txt(raw_bytes), None
    if extension == "pdf":
        return load_pdf(raw_bytes), None
    if extension in ("jpg", "jpeg"):
        return load_image(raw_bytes, filename)
    # Safety net: validate_file() should already have rejected this.
    raise UploadError(f"No loader is available for '.{extension}' files.")
# Full upload pipeline: clean name -> validate -> load -> check structure.
# Contains no Streamlit calls, so it can be tested or reused by other modules.
# Returns an UploadedDataset or raises UploadError.
def process_upload(filename, raw_bytes):
    safe_name = sanitize_filename(filename)
    is_valid, extension, error_message = validate_file(safe_name, raw_bytes)
    if not is_valid:
        raise UploadError(error_message)
    df, image = load_dataset(raw_bytes, extension, safe_name)
    # A dataset with no columns has nothing to analyse.
    if df.shape[1] == 0:
        raise UploadError("This file doesn't contain any columns that could be read.")
    return UploadedDataset(
        name=safe_name,
        extension=extension,
        dataframe=df,
        image=image,
        file_size_bytes=len(raw_bytes),
        uploaded_at=datetime.now(),
    )
# ==========================================
# DATA PREVIEW
# ==========================================
# Render a table full-width. Newer Streamlit uses width="stretch"; older
# versions use use_container_width, so both are supported.
def show_table(data, **kwargs):
    try:
        st.dataframe(data, width="stretch", **kwargs)
    except Exception:
        st.dataframe(data, use_container_width=True, **kwargs)
# Render an image full-width, with the same version compatibility as show_table.
def show_image(image):
    try:
        st.image(image, width="stretch")
    except Exception:
        st.image(image, use_container_width=True)
# Show the dataset name, row count and column count as three metrics.
def display_dataset_information(dataset):
    df = dataset.dataframe
    # The heading text "Dataset Overview" labels this section.
    st.subheader("Dataset Overview")
    col1, col2, col3 = st.columns(3)
    # Each label ("Dataset", "Rows", "Columns") appears above its value.
    col1.metric("Dataset", dataset.name)
    col2.metric("Rows", f"{df.shape[0]:,}")
    col3.metric("Columns", f"{df.shape[1]:,}")
    # Extra file facts so the user can confirm what was read.
    st.caption(
        f"File type: {dataset.extension.upper()} | "
        f"Size: {format_file_size(dataset.file_size_bytes)}"
    )
# Show the actual uploaded picture (JPG/JPEG uploads only).
def display_image_preview(image):
    st.subheader("Image Preview")
    show_image(image)
# Show the first rows so the user can confirm the data was read correctly.
def display_dataset_preview(df):
    st.subheader("Preview")
    show_table(df.head(PREVIEW_ROW_LIMIT))
    # Tell the user when only part of the data is shown.
    if df.shape[0] > PREVIEW_ROW_LIMIT:
        st.caption(f"Showing the first {PREVIEW_ROW_LIMIT} of {df.shape[0]:,} rows.")
# Show each column's name, detected pandas data type and non-null count.
def display_column_information(df):
    st.subheader("Column Information")
    column_info = pd.DataFrame({
        "Column Name": [str(name) for name in df.columns],
        "Data Type": [str(dtype) for dtype in df.dtypes],
        # df.notna().sum() works even if column names are duplicated.
        "Non-Null Count": df.notna().sum().tolist(),
    })
    show_table(column_info, hide_index=True)
# Show pandas' describe() summary plus a quick missing-value notice.
def display_basic_summary(df):
    st.subheader("Basic Dataset Summary")
    try:
        # include="all" summarises numeric and text columns together.
        summary = df.describe(include="all").transpose()
        # Convert to text so mixed-type summary cells display reliably.
        show_table(summary.astype(object).fillna("").astype(str))
    except Exception:
        # A summary problem should not block the rest of the results.
        st.info("A statistical summary could not be generated for this dataset.")
    missing_counts = df.isna().sum()
    total_missing = int(missing_counts.sum())
    if total_missing > 0:
        # Amber warning: worth attention, not a failure.
        st.warning(
            f"This dataset contains {total_missing:,} missing value(s) "
            f"across {int((missing_counts > 0).sum())} column(s)."
        )
    else:
        # Neutral blue notice: nothing to fix.
        st.info("No missing values were detected in this dataset.")
# ==========================================
# APPLICATION ENTRY POINT
# ==========================================
# Build the page: uploader -> processing -> success/errors -> results.
# Run the app when executed by Streamlit (streamlit run app.py). Importing this
# file from another module will not launch the interface.

# ============== PART 2: DATA QUALITY REPORT ==============
"""
data_quality_report.py
Part 2 of the Data Quality & Analytics Platform: the Data Quality Report.
This module is READ-ONLY. It never deletes rows, deletes columns, corrects
values, or overwrites the DataFrame produced by Part 1 (Data Upload). Every
check works on a private, cleaned-up view of the data that is created inside
prepare_dataframe() and thrown away after the report is built.
The module has two layers:
1. ANALYSIS LAYER (pure Python + pandas + numpy, no Streamlit calls).
   Functions such as calculate_missing_values(), detect_outliers() and
   calculate_quality_score(). Because they do not touch Streamlit they are
   easy to test.
2. DISPLAY LAYER (Streamlit).
   Functions named display_*() that turn analysis results into headings,
   metrics, tables and messages. display_quality_report() is the single
   entry point that app.py calls.
Expected workflow:
    Data Upload (Part 1) -> uploaded DataFrame -> run_quality_analysis()
                         -> display_quality_report()
"""
import datetime as dt
import logging
import re
import warnings
import numpy as np
import pandas as pd
import streamlit as st
# A module-level logger. Technical details of unexpected errors are written
# here (to the terminal running Streamlit) instead of being shown to the user.
logger = logging.getLogger(__name__)
# =============================================================================
# CONFIGURATION
# Every threshold used by the report lives here so it is easy to find, read,
# and change. These are data-quality INDICATORS chosen by this application;
# they are not universal rules.
# =============================================================================
# --- Part 1 integration: where the uploaded DataFrame is expected --------------
# Part 1 (Data Upload) is expected to store the DataFrame in st.session_state.
# get_uploaded_dataframe() checks these keys in order and uses the first one
# that holds a DataFrame. If your Part 1 code uses a different key, add it here
# (or, better, pass the DataFrame directly to display_quality_report()).
DATAFRAME_SESSION_KEYS = ("df", "dataframe", "uploaded_df", "uploaded_dataframe", "data")
DATASET_NAME_SESSION_KEYS = ("dataset_name", "file_name", "filename", "uploaded_file_name")
# --- Missing-value severity thresholds (percent of a column that is missing) --
MISSING_LOW_MAX = 5.0        # 0% < missing < 5%    -> "Low concern"
MISSING_MODERATE_MAX = 20.0  # 5% <= missing < 20%  -> "Moderate concern"; 20%+ -> "High concern"
# --- Generic severity thresholds for "how many values are affected" issues ----
SEVERITY_HIGH_PCT = 10.0     # 10% or more of the values affected -> High
SEVERITY_MEDIUM_PCT = 1.0    # 1% or more of the values affected  -> Medium (below that -> Low)
# --- Outlier detection (IQR method) -------------------------------------------
IQR_MULTIPLIER = 1.5         # Standard "Tukey fence" multiplier
MIN_VALUES_FOR_OUTLIERS = 4  # Fewer numeric values than this: quartiles are not meaningful
OUTLIER_MEDIUM_PCT = 5.0     # 5% or more potential outliers -> Medium issue (below -> Low)
# --- Type detection ------------------------------------------------------------
NUMERIC_TEXT_MIN_SHARE = 0.90  # A text column is "numeric-looking" if 90%+ of its values parse as numbers
DATE_TEXT_MIN_SHARE = 0.50     # A text column is "date-looking" if 50%+ of its values have a date shape
DATE_NAME_HINT_MIN_SHARE = 0.10  # ...or 10%+ if the column NAME also suggests dates (e.g. "order_date")
DATE_NAME_HINTS = {"date", "dob", "datetime", "timestamp", "time"}
# --- Categorical consistency ----------------------------------------------------
MAX_UNIQUE_FOR_CONSISTENCY = 5000  # Skip columns with more distinct values than this (keeps the check fast)
# --- Validity rules that depend on column NAMES (heuristics, clearly labelled) --
# If a column name contains one of these words, negative values are reported as
# "Invalid Values" because such quantities are normally not negative.
NON_NEGATIVE_NAME_TOKENS = {
    "age", "price", "income", "salary", "quantity", "qty", "amount", "cost",
    "revenue", "sales", "count", "population", "weight", "height", "distance", "duration",
}
MAX_REASONABLE_AGE = 120  # Values above this in a column named like "age" are reported as invalid
# Words in a column name that suggest the column is meant to hold unique identifiers.
ID_NAME_TOKENS = {"id", "uid", "uuid", "guid", "key", "pk"}
# Text values that often stand in for "no value" (compared case-insensitively).
PLACEHOLDER_TOKENS = {"n/a", "na", "#n/a", "null", "nan", "nil", "-", "--", "?", "missing", "undefined"}
# --- Issue names (used in the issue table and to drive the recommendations) ----
ISSUE_MISSING = "Missing Values"
ISSUE_EMPTY = "Empty Column"
ISSUE_DUP_ROWS = "Duplicate Rows"
ISSUE_DUP_ID = "Duplicate Values in ID-like Column"
ISSUE_CONSTANT = "Constant Column"
ISSUE_OUTLIERS = "Potential Outliers"
ISSUE_INCONSISTENT = "Inconsistent Categorical Values"
ISSUE_NUMERIC_TEXT = "Numeric Values Stored as Text"
ISSUE_DATE_TEXT = "Dates Stored as Text"
ISSUE_MIXED_TYPES = "Mixed Data Types"
ISSUE_INVALID = "Invalid Values"
ISSUE_PLACEHOLDER = "Placeholder Text Values"
ISSUE_INVALID_DATES = "Invalid Dates"
SEVERITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}
ISSUE_COLUMNS = ["Severity", "Column", "Issue", "Details"]
# --- Date-shape patterns (used only to DECIDE whether text looks like a date) --
# Matches shapes such as 2023-05-01, 2023/5/1, 05/01/2023, 5.1.23, and optional times.
_NUMERIC_DATE_SHAPE = (
    r"(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})"
    r"(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s?[AaPp][Mm]|Z|[+-]\d{2}:?\d{2})?)?"
)
# Matches shapes such as "March 5, 2020", "5 Mar 2020", "Mar 5 2020".
_MONTH = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
_NAMED_DATE_SHAPE = (
    rf"(?:{_MONTH}\s+\d{{1,2}}(?:st|nd|rd|th)?,?\s+\d{{2,4}}"
    rf"|\d{{1,2}}(?:st|nd|rd|th)?\s+{_MONTH},?\s+\d{{2,4}})"
)
DATE_SHAPE_PATTERN = rf"^(?:{_NUMERIC_DATE_SHAPE}|{_NAMED_DATE_SHAPE})$"
# =============================================================================
# SECTION 1: SMALL HELPER FUNCTIONS
# =============================================================================
def describe_dtype(dtype):
    """Return a friendlier text label for a pandas dtype (used in the Data Types table)."""
    name = str(dtype)
    if name == "object":
        return "object (text or mixed)"
    if name in ("str", "string"):
        return "text (string)"
    return name
def make_unique_labels(columns):
    """
    Turn column labels into unique strings.
    Real-world files sometimes contain duplicate or non-text column names. The
    analysis needs every column to be addressable by a unique string, so the
    analysis copy uses labels such as "Name", "Name.1". The user's original
    DataFrame is not renamed.
    """
    labels, seen = [], {}
    for column in columns:
        base = str(column)
        if base in seen:
            seen[base] += 1
            labels.append(f"{base}.{seen[base]}")
        else:
            seen[base] = 0
            labels.append(base)
    return labels
def is_text_column(series):
    """True for object/string columns (the columns that can contain free text)."""
    if isinstance(series.dtype, pd.CategoricalDtype):
        return False
    return pd.api.types.is_object_dtype(series.dtype) or pd.api.types.is_string_dtype(series.dtype)
def clean_blank_strings(series):
    """
    Return a copy of a column in which blank/whitespace-only text becomes a
    real missing value (NaN). Non-text columns are returned unchanged.
    Why: a cell containing "   " is not meaningful data. Treating blanks as
    missing makes the completeness numbers reflect what a person would call
    "empty". This only affects the private analysis copy.
    """
    if isinstance(series.dtype, pd.CategoricalDtype):
        series = series.astype(object)
    if not is_text_column(series):
        return series
    try:
        stripped = series.astype("string").str.strip()
        is_blank = stripped.eq("").fillna(False).astype(bool)
        return series.mask(is_blank)
    except Exception:  # unusual object contents: leave the column as it is
        return series
def prepare_dataframe(df):
    """
    Build the private analysis copy of the uploaded DataFrame.
    Returns (analysis_df, original_dtypes):
      * analysis_df: same data, unique string column labels, a clean 0..n-1
        index, and blank strings turned into NaN. The uploaded DataFrame itself
        is never modified.
      * original_dtypes: {column label: friendly dtype label} taken from the
        ORIGINAL data so the Data Types table reports what the user really has.
    """
    # reset_index(drop=True) returns a NEW DataFrame, so renaming its columns
    # below cannot affect the user's DataFrame.
    work = df.reset_index(drop=True)
    labels = make_unique_labels(df.columns)
    work.columns = labels
    original_dtypes = {label: describe_dtype(dtype) for label, dtype in zip(labels, df.dtypes)}
    # Build the analysis DataFrame column by column (blank text -> NaN).
    cleaned_columns = {label: clean_blank_strings(work[label]) for label in labels}
    analysis_df = pd.DataFrame(cleaned_columns, index=work.index)
    return analysis_df, original_dtypes
def check_dataframe(df):
    """
    Validate the input before analysis.
    Returns (level, message). level is None when the DataFrame is usable,
    otherwise "error" or "warning", and message is the user-friendly text
    that display_quality_report() shows.
    """
    if df is None:
        return "warning", ("No dataset is available yet. Please upload a dataset in the "
                           "Data Upload section first, then return to this page.")
    if not isinstance(df, pd.DataFrame):
        return "error", ("The uploaded data is not in a supported table format, so a Data "
                         "Quality Report cannot be created.")
    if df.shape[1] == 0:
        return "error", ("The dataset has no columns, so a Data Quality Report cannot be created.")
    if df.shape[0] == 0:
        return "warning", ("The dataset has columns but no rows, so there is nothing to analyze. "
                           "Please upload a file that contains data.")
    return None, None
def column_name_tokens(name):
    """
    Split a column name into lowercase words. Handles snake_case, spaces and
    camelCase: "CustomerID" -> {"customer", "id"}, "annual_income" -> {"annual", "income"}.
    """
    spaced = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", str(name))
    return {token for token in re.split(r"[^A-Za-z0-9]+", spaced.lower()) if token}
def is_id_like_column(name):
    """True if the column NAME suggests an identifier (e.g. "id", "customer_id", "OrderID")."""
    return bool(column_name_tokens(name) & ID_NAME_TOKENS)
def safe_nunique(series):
    """Count distinct non-null values; falls back to text comparison for unhashable values (lists, dicts)."""
    try:
        return int(series.nunique(dropna=True))
    except TypeError:
        return int(series.dropna().astype(str).nunique())
def first_examples(values, limit=3, width=24):
    """Return up to `limit` example values as a short quoted string, for issue details."""
    examples = []
    for value in pd.Series(values).dropna().astype(str).unique()[:limit]:
        text = value if len(value) <= width else value[: width - 1] + "…"
        examples.append(f"'{text}'")
    return ", ".join(examples)
def placeholder_mask(series):
    """Boolean Series: True where a text value looks like a placeholder such as 'N/A' or 'null'."""
    lowered = series.astype("string").str.strip().str.lower()
    return lowered.isin(PLACEHOLDER_TOKENS).fillna(False).astype(bool)
def numeric_text_profile(series):
    """
    Measure how numeric a TEXT column is.
    Returns a dict with:
      considered   - non-missing, non-placeholder values examined
      numeric      - how many of those can be read as numbers
      share        - numeric / considered (0..1)
      non_numeric_mask - Boolean Series (full length) marking the entries that are not numbers
    """
    present = series.notna() & ~placeholder_mask(series)
    texts = series[present].astype("string").str.strip().str.replace(",", "", regex=False)
    converted = pd.to_numeric(texts, errors="coerce")
    is_number = converted.notna()
    non_numeric_mask = pd.Series(False, index=series.index)
    non_numeric_mask.loc[texts.index[~is_number.to_numpy()]] = True
    considered = int(len(texts))
    numeric = int(is_number.sum())
    return {
        "considered": considered,
        "numeric": numeric,
        "share": (numeric / considered) if considered else 0.0,
        "non_numeric_mask": non_numeric_mask,
    }
def python_type_label(value):
    """Group Python value types into broad labels used to detect mixed data types."""
    if isinstance(value, (bool, np.bool_)):
        return "boolean"
    if isinstance(value, (int, float, np.integer, np.floating)):
        return "number"
    if isinstance(value, str):
        return "text"
    if isinstance(value, (pd.Timestamp, dt.date, dt.datetime)):
        return "date/time"
    return type(value).__name__
def get_numeric_values(series):
    """Return the finite numeric values of a column as floats (missing and infinite values removed)."""
    try:
        values = pd.to_numeric(series, errors="coerce").astype("float64")
    except (TypeError, ValueError):
        return pd.Series(dtype="float64")
    return values.replace([np.inf, -np.inf], np.nan).dropna()
def parse_dates(texts):
    """
    Parse text into datetimes WITHOUT raising errors: anything that cannot be
    read as a real date becomes NaT (missing). Several strategies are tried so
    that mixed formats and time zones are handled.
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        for kwargs in ({"format": "mixed"}, {"format": "mixed", "utc": True}, {}):
            try:
                return pd.to_datetime(texts, errors="coerce", **kwargs)
            except (TypeError, ValueError):
                continue
    return pd.Series(pd.NaT, index=texts.index)
def looks_like_date_column(series, column_name):
    """
    Decide whether a TEXT column appears to contain dates.
    It counts the values whose SHAPE looks like a date (e.g. 2023-05-01 or
    "March 5, 2020"). This does not check that the dates are valid; invalid
    ones are reported later by analyze_datetime_quality().
    """
    present = series.notna() & ~placeholder_mask(series)
    texts = series[present].astype(str).str.strip()
    if texts.empty:
        return False
    share = texts.str.match(DATE_SHAPE_PATTERN, flags=re.IGNORECASE).mean()
    if share >= DATE_TEXT_MIN_SHARE:
        return True
    return bool(share >= DATE_NAME_HINT_MIN_SHARE and (column_name_tokens(column_name) & DATE_NAME_HINTS))
def severity_from_percentage(percentage):
    """Map 'percent of values affected' to High / Medium / Low."""
    if percentage >= SEVERITY_HIGH_PCT:
        return "High"
    if percentage >= SEVERITY_MEDIUM_PCT:
        return "Medium"
    return "Low"
def format_timestamp(timestamp):
    """Format a Timestamp as a date, adding the time only when it is not midnight."""
    if pd.isna(timestamp):
        return "n/a"
    try:
        if timestamp.hour == 0 and timestamp.minute == 0 and timestamp.second == 0:
            return timestamp.strftime("%Y-%m-%d")
        return timestamp.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(timestamp)
# =============================================================================
# SECTION 2: COLUMN CLASSIFICATION AND DATASET OVERVIEW
# =============================================================================
def classify_columns(df):
    """
    Sort every column into one group:
      numeric     - integer / float columns (booleans are NOT counted as numeric)
      datetime    - columns that pandas already stores as dates
      date_text   - text columns whose values look like dates
      categorical - everything else (text, categories, booleans, ...)
    """
    groups = {"numeric": [], "datetime": [], "date_text": [], "categorical": []}
    for column in df.columns:
        series = df[column]
        if pd.api.types.is_bool_dtype(series.dtype):
            groups["categorical"].append(column)
        elif pd.api.types.is_numeric_dtype(series.dtype):
            groups["numeric"].append(column)
        elif pd.api.types.is_datetime64_any_dtype(series.dtype):
            groups["datetime"].append(column)
        elif is_text_column(series) and looks_like_date_column(series, column):
            groups["date_text"].append(column)
        else:
            groups["categorical"].append(column)
    return groups
def build_overview(df, groups, dataset_name):
    """Collect the high-level facts shown in the Dataset Overview section."""
    rows, columns = df.shape
    return {
        "dataset_name": dataset_name,
        "rows": int(rows),
        "columns": int(columns),
        "total_cells": int(rows * columns),
        "numeric_columns": len(groups["numeric"]),
        "categorical_columns": len(groups["categorical"]),
        # Date/time columns include real datetime columns AND text columns that look like dates.
        "datetime_columns": len(groups["datetime"]) + len(groups["date_text"]),
    }
# =============================================================================
# SECTION 3: MISSING VALUES, DUPLICATES, EMPTY AND CONSTANT COLUMNS
# =============================================================================
def classify_missing_severity(percentage):
    """Translate a missing-value percentage into this application's label (an indicator, not a universal rule)."""
    if percentage == 0:
        return "Complete"
    if percentage < MISSING_LOW_MAX:
        return "Low concern"
    if percentage < MISSING_MODERATE_MAX:
        return "Moderate concern"
    return "High concern"
def calculate_missing_values(df):
    """Return one row per column: missing count, missing percentage and severity label."""
    # Calculate the total number of missing values in each column.
    # This information is used by the Data Quality Report to measure
    # dataset completeness.
    missing_counts = df.isnull().sum()
    # Calculate the percentage of missing values in each column.
    # This percentage allows the report to identify columns that
    # may require additional data-quality attention.
    total_rows = len(df)
    missing_percentages = (missing_counts / total_rows * 100) if total_rows else missing_counts * 0.0
    table = pd.DataFrame({
        "Column": missing_counts.index,
        "Missing Count": missing_counts.to_numpy(dtype=int),
        "Missing %": missing_percentages.to_numpy(dtype=float).round(2),
    })
    # Attach the severity label (Complete / Low / Moderate / High concern).
    table["Severity"] = table["Missing %"].apply(classify_missing_severity)
    return table.sort_values("Missing %", ascending=False, kind="stable").reset_index(drop=True)
def calculate_duplicates(df):
    """Count rows that are exact repeats of an earlier row (the first occurrence is not counted)."""
    try:
        duplicate_flags = df.duplicated()
    except TypeError:
        # Unhashable cell contents (lists, dicts): compare the text form instead.
        duplicate_flags = df.astype(str).duplicated()
    # Number of duplicate rows, and that number as a percentage of all rows.
    duplicate_rows = int(duplicate_flags.sum())
    percentage = (duplicate_rows / len(df) * 100) if len(df) else 0.0
    return {"duplicate_rows": duplicate_rows, "percentage": round(percentage, 2), "total_rows": int(len(df))}
def calculate_duplicate_values(df):
    """
    Per column: how many values repeat an earlier value.
    Duplicate Values = non-null values - distinct values, i.e. the number of
    entries that repeat something already seen. A repeat is NOT automatically a
    problem: repeated values are normal in columns like Country or Gender and
    are only unexpected in identifier columns.
    """
    records = []
    for column in df.columns:
        non_null = int(df[column].notna().sum())
        unique = safe_nunique(df[column])
        duplicates = non_null - unique
        records.append({
            "Column": column,
            "Non-Null Values": non_null,
            "Unique Values": unique,
            "Duplicate Values": duplicates,
            "Duplicate %": round(duplicates / non_null * 100, 2) if non_null else 0.0,
            "Looks Like an ID Column": is_id_like_column(column),
        })
    return pd.DataFrame(records)
def detect_empty_columns(df):
    """Columns where every value is missing or blank."""
    records = []
    total_rows = len(df)
    for column in df.columns:
        missing = int(df[column].isna().sum())
        if total_rows and missing == total_rows:
            records.append({
                "Column": column,
                "Number of Values": total_rows,
                "Missing Values": missing,
                "Missing %": round(missing / total_rows * 100, 2),
            })
    return pd.DataFrame(records, columns=["Column", "Number of Values", "Missing Values", "Missing %"])
def detect_constant_columns(df):
    """Columns whose non-null values are all identical (needs at least 2 rows to be meaningful)."""
    columns = ["Column", "Unique Values", "Value", "Non-Null Values"]
    if len(df) < 2:
        return pd.DataFrame(columns=columns)
    records = []
    for column in df.columns:
        non_null = df[column].dropna()
        if len(non_null) > 0 and safe_nunique(non_null) == 1:
            value = str(non_null.iloc[0])
            records.append({
                "Column": column,
                "Unique Values": 1,
                "Value": value if len(value) <= 60 else value[:59] + "…",
                "Non-Null Values": int(len(non_null)),
            })
    return pd.DataFrame(records, columns=columns)
# =============================================================================
# SECTION 4: DATA TYPES, INVALID VALUES, DATES
# =============================================================================
def analyze_data_types(df, original_dtypes, groups):
    """
    Build the Data Types table and detect type-related concerns.
    Returns (table, findings, mixed_type_cells):
      table  - Column, Detected Data Type, Non-Null Values, Unique Values, Potential Concern
      findings - issue dicts (Severity, Column, Issue, Details)
      mixed_type_cells - number of cells in the minority type of mixed columns (feeds the Consistency score)
    """
    records, findings, mixed_type_cells = [], [], 0
    for column in df.columns:
        series = df[column]
        non_null = int(series.notna().sum())
        concerns = []
        if is_text_column(series) and non_null > 0:
            # (a) Numeric-looking values stored as text.
            profile = numeric_text_profile(series)
            if profile["considered"] > 0 and profile["share"] >= NUMERIC_TEXT_MIN_SHARE and profile["numeric"] > 0:
                concerns.append("Numeric-looking values stored as text")
                findings.append({
                    "Severity": "Low", "Column": column, "Issue": ISSUE_NUMERIC_TEXT,
                    "Details": (f"{profile['numeric']:,} of {profile['considered']:,} values look like numbers "
                                f"but are stored as text (may be intentional, e.g. ZIP codes or IDs)"),
                })
            # (b) Date-looking values stored as strings.
            if column in groups["date_text"]:
                concerns.append("Date-looking values stored as text")
                findings.append({
                    "Severity": "Low", "Column": column, "Issue": ISSUE_DATE_TEXT,
                    "Details": "Values look like dates but the column is stored as text, not as a date/time type",
                })
            # (c) Mixed Python types inside one object column (e.g. numbers and text together).
            if pd.api.types.is_object_dtype(series.dtype):
                type_counts = series.dropna().map(python_type_label).value_counts()
                if len(type_counts) > 1:
                    minority = int(type_counts.sum() - type_counts.iloc[0])
                    mixed_type_cells += minority
                    summary = ", ".join(f"{label}: {count:,}" for label, count in type_counts.items())
                    concerns.append("Mixed data types")
                    findings.append({
                        "Severity": "Medium", "Column": column, "Issue": ISSUE_MIXED_TYPES,
                        "Details": f"Column holds more than one kind of value ({summary})",
                    })
        records.append({
            "Column": column,
            "Detected Data Type": original_dtypes.get(column, str(series.dtype)),
            "Non-Null Values": non_null,
            "Unique Values": safe_nunique(series),
            "Potential Concern": "; ".join(concerns) if concerns else "None detected",
        })
    return pd.DataFrame(records), findings, mixed_type_cells
def detect_invalid_values(df, groups):
    """
    Detect values that are unlikely to be valid. Rules used:
      * Numeric columns: infinite values; negative values in columns whose NAME
        suggests they should not be negative (age, price, income, ...); ages above 120.
      * Text columns: placeholder text such as 'N/A' or 'null'; non-numeric
        entries inside an otherwise numeric-looking column.
    (Invalid dates are handled in analyze_datetime_quality().)
    Returns (findings, invalid_counts) where invalid_counts maps
    column -> number of distinct invalid cells (each cell counted once).
    """
    findings, invalid_counts = [], {}
    # ---- numeric columns ----------------------------------------------------
    for column in groups["numeric"]:
        try:
            raw = pd.to_numeric(df[column], errors="coerce").astype("float64")
        except (TypeError, ValueError):
            continue
        infinite = np.isinf(raw)
        finite = raw.where(~infinite)
        tokens = column_name_tokens(column)
        invalid_mask = infinite.copy()
        if infinite.sum() > 0:
            findings.append({
                "Severity": "High", "Column": column, "Issue": ISSUE_INVALID,
                "Details": f"{int(infinite.sum()):,} infinite values",
            })
        if tokens & NON_NEGATIVE_NAME_TOKENS:
            negative = finite < 0
            if negative.sum() > 0:
                invalid_mask |= negative
                findings.append({
                    "Severity": "High", "Column": column, "Issue": ISSUE_INVALID,
                    "Details": (f"{int(negative.sum()):,} negative values (the column name suggests "
                                f"values should not be negative; please verify)"),
                })
        if "age" in tokens:
            too_old = finite > MAX_REASONABLE_AGE
            if too_old.sum() > 0:
                invalid_mask |= too_old
                findings.append({
                    "Severity": "Medium", "Column": column, "Issue": ISSUE_INVALID,
                    "Details": f"{int(too_old.sum()):,} values above {MAX_REASONABLE_AGE} in an age-like column",
                })
        if invalid_mask.sum() > 0:
            invalid_counts[column] = invalid_counts.get(column, 0) + int(invalid_mask.sum())
    # ---- text columns (date-looking columns are checked separately) -----------
    for column in groups["categorical"]:
        series = df[column]
        if not is_text_column(series) and not pd.api.types.is_object_dtype(series.dtype):
            continue
        non_null = int(series.notna().sum())
        if non_null == 0:
            continue
        placeholders = placeholder_mask(series)
        invalid_mask = placeholders.copy()
        if placeholders.sum() > 0:
            pct = placeholders.sum() / non_null * 100
            findings.append({
                "Severity": severity_from_percentage(pct), "Column": column, "Issue": ISSUE_PLACEHOLDER,
                "Details": (f"{int(placeholders.sum()):,} values such as "
                            f"{first_examples(series[placeholders])} may stand in for missing data"),
            })
        profile = numeric_text_profile(series)
        # Only treat leftovers as invalid if the column is overwhelmingly numeric.
        if profile["considered"] > 0 and NUMERIC_TEXT_MIN_SHARE <= profile["share"] < 1.0:
            bad = profile["non_numeric_mask"]
            invalid_mask |= bad
            pct = bad.sum() / non_null * 100
            findings.append({
                "Severity": severity_from_percentage(pct), "Column": column, "Issue": ISSUE_INVALID,
                "Details": (f"{int(bad.sum()):,} entries are not numbers in a mostly numeric column "
                            f"(e.g. {first_examples(series[bad])})"),
            })
        if invalid_mask.sum() > 0:
            invalid_counts[column] = invalid_counts.get(column, 0) + int(invalid_mask.sum())
    return findings, invalid_counts
def analyze_datetime_quality(df, groups):
    """
    Date/time quality for real datetime columns and for text columns that look like dates.
    Returns (table, findings, invalid_counts). Nothing is converted in the
    user's data; text dates are parsed only inside this function to measure quality.
    """
    columns = ["Column", "Column Type", "Non-Null Values", "Missing Dates", "Invalid Dates",
               "Minimum Date", "Maximum Date", "Unique Dates"]
    records, findings, invalid_counts = [], [], {}
    for column in groups["datetime"] + groups["date_text"]:
        series = df[column]
        missing = int(series.isna().sum())
        non_null = int(series.notna().sum())
        is_text = column in groups["date_text"]
        invalid = 0
        if is_text:
            # Parse text values (excluding placeholders such as 'N/A', which are reported elsewhere).
            present = series.notna() & ~placeholder_mask(series)
            texts = series[present].astype(str).str.strip()
            parsed = parse_dates(texts)
            bad = parsed.isna()
            invalid = int(bad.sum())
            valid = parsed.dropna()
            if invalid > 0:
                pct = invalid / max(len(texts), 1) * 100
                invalid_counts[column] = invalid
                findings.append({
                    "Severity": severity_from_percentage(pct), "Column": column, "Issue": ISSUE_INVALID_DATES,
                    "Details": (f"{invalid:,} of {len(texts):,} values could not be read as valid dates "
                                f"(e.g. {first_examples(texts[bad])})"),
                })
        else:
            valid = series.dropna()
        try:
            minimum, maximum = format_timestamp(valid.min()), format_timestamp(valid.max())
        except Exception:
            minimum = maximum = "n/a"
        try:
            unique_dates = int(valid.dt.date.nunique())
        except Exception:
            unique_dates = int(len({v.date() for v in valid if hasattr(v, "date")}))
        records.append({
            "Column": column,
            "Column Type": "Text that looks like dates" if is_text else "Date/time",
            "Non-Null Values": non_null,
            "Missing Dates": missing,
            "Invalid Dates": invalid,
            "Minimum Date": minimum if len(valid) else "n/a",
            "Maximum Date": maximum if len(valid) else "n/a",
            "Unique Dates": unique_dates if len(valid) else 0,
        })
    return pd.DataFrame(records, columns=columns), findings, invalid_counts
# =============================================================================
# SECTION 5: NUMERIC QUALITY AND OUTLIERS
# =============================================================================
def detect_outliers(df, numeric_columns):
    """
    Potential outliers with the IQR method.
      Q1 = 25th percentile, Q3 = 75th percentile, IQR = Q3 - Q1
      Lower bound = Q1 - 1.5 * IQR,  Upper bound = Q3 + 1.5 * IQR
      A value outside [lower, upper] is a POTENTIAL outlier.
    A potential outlier is not necessarily an error; it may be a legitimate observation.
    """
    columns = ["Column", "Q1", "Q3", "IQR", "Lower Bound", "Upper Bound",
               "Potential Outliers", "Outlier %", "Note"]
    records = []
    for column in numeric_columns:
        values = get_numeric_values(df[column])
        record = {"Column": column, "Q1": np.nan, "Q3": np.nan, "IQR": np.nan, "Lower Bound": np.nan,
                  "Upper Bound": np.nan, "Potential Outliers": 0, "Outlier %": 0.0, "Note": ""}
        if len(values) < MIN_VALUES_FOR_OUTLIERS:
            record["Note"] = "Too few numeric values for the IQR method"
            records.append(record)
            continue
        # Q1, Q3 and the interquartile range (IQR).
        q1, q3 = float(values.quantile(0.25)), float(values.quantile(0.75))
        iqr = q3 - q1
        record.update({"Q1": q1, "Q3": q3, "IQR": iqr})
        if iqr == 0:
            # If the middle 50% of values are identical, every different value would be
            # flagged, which is not informative. Report this instead of guessing.
            record["Lower Bound"], record["Upper Bound"] = q1, q3
            record["Note"] = "IQR is 0, so the IQR method is not informative for this column"
        else:
            lower, upper = q1 - IQR_MULTIPLIER * iqr, q3 + IQR_MULTIPLIER * iqr
            outliers = int(((values < lower) | (values > upper)).sum())
            record.update({
                "Lower Bound": lower, "Upper Bound": upper, "Potential Outliers": outliers,
                "Outlier %": outliers / len(values) * 100,
            })
        records.append(record)
    table = pd.DataFrame(records, columns=columns)
    float_columns = ["Q1", "Q3", "IQR", "Lower Bound", "Upper Bound", "Outlier %"]
    table[float_columns] = table[float_columns].astype(float).round(4)
    return table
def analyze_numeric_quality(df, numeric_columns, outliers_table):
    """Summary statistics per numerical column (min, max, mean, median, std, missing, unique, outliers)."""
    columns = ["Column", "Minimum", "Maximum", "Mean", "Median", "Std Deviation",
               "Missing Values", "Unique Values", "Potential Outliers"]
    outlier_lookup = {}
    if not outliers_table.empty:
        outlier_lookup = dict(zip(outliers_table["Column"], outliers_table["Potential Outliers"]))
    records = []
    for column in numeric_columns:
        values = get_numeric_values(df[column])
        has_values = len(values) > 0
        records.append({
            "Column": column,
            "Minimum": float(values.min()) if has_values else np.nan,
            "Maximum": float(values.max()) if has_values else np.nan,
            "Mean": float(values.mean()) if has_values else np.nan,
            "Median": float(values.median()) if has_values else np.nan,
            "Std Deviation": float(values.std()) if len(values) > 1 else np.nan,
            "Missing Values": int(df[column].isna().sum()),
            "Unique Values": safe_nunique(df[column]),
            "Potential Outliers": int(outlier_lookup.get(column, 0)),
        })
    table = pd.DataFrame(records, columns=columns)
    stat_columns = ["Minimum", "Maximum", "Mean", "Median", "Std Deviation"]
    table[stat_columns] = table[stat_columns].astype(float).round(4)
    return table
# =============================================================================
# SECTION 6: CATEGORICAL CONSISTENCY
# =============================================================================
def normalize_category(value):
    """
    Reduce a category to a comparison key: lowercase, with spaces, punctuation
    and underscores removed. "New York", "new york", "NEW YORK" and "NewYork"
    all become "newyork".
    """
    return re.sub(r"[\W_]+", "", str(value).lower())
def analyze_categorical_consistency(df, categorical_columns):
    """
    Look for categories that probably mean the same thing but are written differently.
    Returns (table, findings, inconsistent_cells):
      inconsistent_cells counts rows written differently from the most common
      spelling in their group. This feeds the Consistency score.
    """
    columns = ["Column", "Comparison Key", "Variants Found", "Difference Type", "Rows Not Matching Most Common Spelling"]
    records, findings, inconsistent_cells = [], [], 0
    for column in categorical_columns:
        series = df[column]
        non_null = int(series.notna().sum())
        if non_null == 0:
            continue
        try:
            counts = series.dropna().astype(str).value_counts()
        except Exception:
            continue
        if len(counts) > MAX_UNIQUE_FOR_CONSISTENCY:
            continue  # too many distinct values: likely free text or identifiers
        # Group spellings that share the same comparison key.
        grouped = {}
        for variant, count in counts.items():
            key = normalize_category(variant)
            if not any(character.isalpha() for character in key):
                continue  # purely numeric keys are ignored (e.g. "1.0" vs "10" are different numbers)
            grouped.setdefault(key, []).append((variant, int(count)))
        column_groups, column_rows = 0, 0
        example_group = ""
        for key, variants in grouped.items():
            if len(variants) < 2:
                continue
            variants.sort(key=lambda item: -item[1])  # most common spelling first
            rows_differing = sum(count for _, count in variants[1:])
            collapsed = {re.sub(r"\s+", " ", v.strip()).lower() for v, _ in variants}
            difference = "Capitalization or extra whitespace" if len(collapsed) == 1 else "Spacing or punctuation"
            records.append({
                "Column": column,
                "Comparison Key": key,
                "Variants Found": ", ".join(f"'{v}' ({c:,})" for v, c in variants),
                "Difference Type": difference,
                "Rows Not Matching Most Common Spelling": rows_differing,
            })
            column_groups += 1
            column_rows += rows_differing
            if not example_group:
                example_group = " / ".join(f"'{v}'" for v, _ in variants[:3])
        if column_groups:
            inconsistent_cells += column_rows
            pct = column_rows / non_null * 100
            severity = "Medium" if (column_groups >= 3 or pct >= 5) else "Low"
            findings.append({
                "Severity": severity, "Column": column, "Issue": ISSUE_INCONSISTENT,
                "Details": (f"{column_groups} group(s) of values may be the same category written differently "
                            f"(e.g. {example_group}); {column_rows:,} rows differ from the most common spelling"),
            })
    return pd.DataFrame(records, columns=columns), findings, inconsistent_cells
# =============================================================================
# SECTION 7: OVERALL SCORE, ISSUE LIST, RECOMMENDATIONS
# =============================================================================
def calculate_quality_score(total_rows, total_columns, missing_cells, duplicate_rows,
                            id_uniqueness_percentages, invalid_cells, inconsistent_cells):
    """
    Overall Data Quality Score (0-100). Four dimensions, equally weighted.
    Every input is a plain count taken from the dataset, so the score is
    reproducible: the same dataset always gives the same score.
    """
    # Total number of cells, and how many of them contain something (are not missing).
    total_cells = total_rows * total_columns
    non_missing_cells = max(total_cells - missing_cells, 0)
    # COMPLETENESS = share of cells that are NOT missing.
    #   completeness = (1 - missing cells / total cells) * 100
    completeness = (1 - missing_cells / total_cells) * 100 if total_cells else 0.0
    # UNIQUENESS = how free the dataset is of unexpected repeats. Two parts:
    #   (a) row uniqueness = (1 - duplicate rows / total rows) * 100
    #   (b) for columns NAMED like identifiers (id, uuid, key ...) the average
    #       share of values that are unique.
    # If there are ID-like columns, uniqueness is the average of (a) and (b);
    # otherwise it is just (a). Other columns are NOT expected to be unique,
    # so repeated values in them do not reduce this score.
    row_uniqueness = (1 - duplicate_rows / total_rows) * 100 if total_rows else 0.0
    if id_uniqueness_percentages:
        id_uniqueness = float(np.mean(id_uniqueness_percentages))
        uniqueness = (row_uniqueness + id_uniqueness) / 2
    else:
        uniqueness = row_uniqueness
    # VALIDITY = share of non-missing cells that are not flagged invalid
    #   (infinite/negative-where-unexpected/impossible ages, placeholder text,
    #    non-numeric entries in numeric columns, unparseable dates).
    #   validity = (1 - invalid cells / non-missing cells) * 100
    validity = (1 - min(invalid_cells, non_missing_cells) / non_missing_cells) * 100 if non_missing_cells else 0.0
    # CONSISTENCY = share of non-missing cells that are not written inconsistently
    #   (category spellings that differ from the most common spelling of the
    #    same category, and minority types inside mixed-type columns).
    #   consistency = (1 - inconsistent cells / non-missing cells) * 100
    consistency = (1 - min(inconsistent_cells, non_missing_cells) / non_missing_cells) * 100 if non_missing_cells else 0.0
    # Keep every dimension between 0 and 100.
    dimensions = {
        "completeness": float(np.clip(completeness, 0, 100)),
        "uniqueness": float(np.clip(uniqueness, 0, 100)),
        "validity": float(np.clip(validity, 0, 100)),
        "consistency": float(np.clip(consistency, 0, 100)),
    }
    # OVERALL SCORE = simple average of the four dimensions (equal weights).
    overall = float(np.mean(list(dimensions.values())))
    # Score bands: reading aid used by this application only.
    if overall >= 90:
        rating = "Excellent"
    elif overall >= 75:
        rating = "Good"
    elif overall >= 50:
        rating = "Fair"
    else:
        rating = "Needs attention"
    return {**dimensions, "overall": overall, "rating": rating,
            "missing_cells": int(missing_cells), "duplicate_rows": int(duplicate_rows),
            "invalid_cells": int(invalid_cells), "inconsistent_cells": int(inconsistent_cells)}
def build_issue_list(missing, duplicates, duplicate_values, empty_columns, constant_columns,
                     outliers, extra_findings):
    """
    Combine every detected problem into one table: Severity, Column, Issue, Details.
    `extra_findings` holds findings already produced by the type, validity,
    date and consistency checks. Wording is deliberately cautious ("potential",
    "may") because the report detects patterns; it does not know the user's intent.
    """
    issues = list(extra_findings)
    # Empty columns (High).
    for _, row in empty_columns.iterrows():
        issues.append({"Severity": "High", "Column": row["Column"], "Issue": ISSUE_EMPTY,
                       "Details": f"All {int(row['Number of Values']):,} values are missing or blank"})
    empty_names = set(empty_columns["Column"]) if not empty_columns.empty else set()
    # Missing values (severity from the missing-value thresholds; fully empty columns already reported).
    severity_map = {"High concern": "High", "Moderate concern": "Medium", "Low concern": "Low"}
    for _, row in missing.iterrows():
        if row["Missing Count"] > 0 and row["Column"] not in empty_names:
            issues.append({"Severity": severity_map[row["Severity"]], "Column": row["Column"], "Issue": ISSUE_MISSING,
                           "Details": f"{int(row['Missing Count']):,} missing values ({row['Missing %']:.1f}% of the column)"})
    # Duplicate rows.
    if duplicates["duplicate_rows"] > 0:
        issues.append({"Severity": severity_from_percentage(duplicates["percentage"]), "Column": "(all columns)",
                       "Issue": ISSUE_DUP_ROWS,
                       "Details": f"{duplicates['duplicate_rows']:,} rows ({duplicates['percentage']:.1f}%) exactly repeat an earlier row"})
    # Duplicate values in ID-like columns.
    if not duplicate_values.empty:
        for _, row in duplicate_values[duplicate_values["Looks Like an ID Column"]].iterrows():
            if row["Duplicate Values"] > 0:
                issues.append({"Severity": "High", "Column": row["Column"], "Issue": ISSUE_DUP_ID,
                               "Details": (f"{int(row['Duplicate Values']):,} repeated values ({row['Duplicate %']:.1f}%); "
                                           f"columns named like identifiers are usually expected to be unique")})
    # Constant columns.
    for _, row in constant_columns.iterrows():
        issues.append({"Severity": "Low", "Column": row["Column"], "Issue": ISSUE_CONSTANT,
                       "Details": f"Every non-null value is '{row['Value']}'; may add limited analytical value"})
    # Potential outliers.
    if not outliers.empty:
        for _, row in outliers[outliers["Potential Outliers"] > 0].iterrows():
            severity = "Medium" if row["Outlier %"] >= OUTLIER_MEDIUM_PCT else "Low"
            issues.append({"Severity": severity, "Column": row["Column"], "Issue": ISSUE_OUTLIERS,
                           "Details": (f"{int(row['Potential Outliers']):,} potential outliers ({row['Outlier %']:.1f}%) "
                                       f"outside {row['Lower Bound']:.6g} to {row['Upper Bound']:.6g} (IQR method); may be legitimate")})
    table = pd.DataFrame(issues, columns=ISSUE_COLUMNS)
    if table.empty:
        return table
    table["_order"] = table["Severity"].map(SEVERITY_ORDER)
    table = table.sort_values(["_order", "Column"], kind="stable").drop(columns="_order")
    return table.reset_index(drop=True)
def generate_recommendations(issues):
    """Turn the detected issues into practical, numbered actions. Only issues that exist produce a recommendation."""
    if issues.empty:
        return ["No issues were detected by the checks in this report. Continue to review the "
                "sections above and confirm the data matches what you expect."]
    def columns_for(*issue_names, severity=None):
        subset = issues[issues["Issue"].isin(issue_names)]
        if severity is not None:
            subset = subset[subset["Severity"].isin(severity)]
        names = list(dict.fromkeys(subset["Column"]))
        shown = ", ".join(f"`{name}`" for name in names[:5])
        return names, shown + (f" and {len(names) - 5} more" if len(names) > 5 else "")
    recommendations = []
    names, text = columns_for(ISSUE_EMPTY)
    if names:
        recommendations.append(f"Review empty columns ({text}). Confirm whether they were exported by mistake or can be left out of your analysis.")
    names, text = columns_for(ISSUE_MISSING, severity=("High", "Medium"))
    if names:
        recommendations.append(f"Review columns with high or moderate missing-value percentages ({text}) and decide how to handle them "
                               f"(for example collect the data, fill it in, or exclude the column from analysis).")
    if (issues["Issue"] == ISSUE_DUP_ROWS).any():
        recommendations.append("Investigate potential duplicate records. Confirm whether repeated rows are genuine repeats "
                               "or accidental copies before removing anything.")
    names, text = columns_for(ISSUE_DUP_ID)
    if names:
        recommendations.append(f"Check ID-like columns ({text}) for repeated values and confirm whether repeats are expected.")
    names, text = columns_for(ISSUE_INVALID)
    if names:
        recommendations.append(f"Review invalid values in {text} (for example negative, infinite, impossible or non-numeric entries) and correct them at the source.")
    names, text = columns_for(ISSUE_INVALID_DATES)
    if names:
        recommendations.append(f"Check the date values in {text} that could not be read as valid dates and confirm the expected date format.")
    names, text = columns_for(ISSUE_PLACEHOLDER)
    if names:
        recommendations.append(f"Decide how placeholder text such as N/A or null in {text} should be treated (for example as true missing values).")
    names, text = columns_for(ISSUE_OUTLIERS)
    if names:
        recommendations.append(f"Review potential outliers in {text} before removing them; they may be legitimate observations.")
    names, text = columns_for(ISSUE_INCONSISTENT)
    if names:
        recommendations.append(f"Standardize inconsistent categorical values in {text} so that one category has one spelling.")
    names, text = columns_for(ISSUE_MIXED_TYPES)
    if names:
        recommendations.append(f"Review columns containing mixed data types ({text}) and decide on one consistent type.")
    names, text = columns_for(ISSUE_NUMERIC_TEXT)
    if names:
        recommendations.append(f"Confirm whether numeric-looking text columns ({text}) should be stored as numbers. "
                               f"Identifiers such as ZIP codes are often better kept as text.")
    names, text = columns_for(ISSUE_DATE_TEXT)
    if names:
        recommendations.append(f"Consider storing date-looking text columns ({text}) as a date/time type once you have confirmed their format.")
    names, text = columns_for(ISSUE_CONSTANT)
    if names:
        recommendations.append(f"Review constant columns ({text}). They add little analytical information but should not be deleted automatically.")
    return recommendations
# =============================================================================
# SECTION 8: RUN THE FULL ANALYSIS
# =============================================================================
def _safe_run(results, label, function, default, *args):
    """
    Run one check. If it fails for any unexpected reason, record a friendly
    warning, log the technical details privately, and continue with `default`
    so one broken check cannot stop the whole report.
    """
    try:
        return function(*args)
    except Exception:
        logger.exception("Data quality check failed: %s", label)
        results["warnings"].append(f"The '{label}' check could not be completed and was skipped.")
        return default
def run_quality_analysis(df, dataset_name="Uploaded dataset"):
    """
    Run every data-quality check and return one results dictionary.
    The uploaded DataFrame is only READ. Raises ValueError if the DataFrame is
    unusable (display_quality_report() checks this first and shows a message).
    """
    level, message = check_dataframe(df)
    if level is not None:
        raise ValueError(message)
    results = {"warnings": []}
    empty = pd.DataFrame()
    analysis_df, original_dtypes = prepare_dataframe(df)
    groups = classify_columns(analysis_df)
    results["groups"] = groups
    results["overview"] = build_overview(analysis_df, groups, dataset_name)
    # Individual checks (each protected by _safe_run).
    missing = _safe_run(results, "Missing Values", calculate_missing_values, empty, analysis_df)
    duplicates = _safe_run(results, "Duplicate Rows", calculate_duplicates,
                           {"duplicate_rows": 0, "percentage": 0.0, "total_rows": len(analysis_df)}, analysis_df)
    duplicate_values = _safe_run(results, "Duplicate Values", calculate_duplicate_values, empty, analysis_df)
    data_types, type_findings, mixed_cells = _safe_run(
        results, "Data Types", analyze_data_types, (empty, [], 0), analysis_df, original_dtypes, groups)
    empty_columns = _safe_run(results, "Empty Columns", detect_empty_columns, empty, analysis_df)
    constant_columns = _safe_run(results, "Constant Columns", detect_constant_columns, empty, analysis_df)
    outliers = _safe_run(results, "Potential Outliers", detect_outliers, empty, analysis_df, groups["numeric"])
    numeric_quality = _safe_run(results, "Numeric Data Quality", analyze_numeric_quality, empty,
                                analysis_df, groups["numeric"], outliers)
    categorical, category_findings, inconsistent_cells = _safe_run(
        results, "Categorical Consistency", analyze_categorical_consistency, (empty, [], 0),
        analysis_df, groups["categorical"])
    invalid_findings, invalid_counts = _safe_run(
        results, "Invalid Values", detect_invalid_values, ([], {}), analysis_df, groups)
    datetime_quality, date_findings, invalid_date_counts = _safe_run(
        results, "Date/Time Quality", analyze_datetime_quality, (empty, [], {}), analysis_df, groups)
    # Gather counts for the score. Invalid cells from the value checks and the date check are disjoint
    # (placeholders are excluded from the date check), so they can be added.
    missing_cells = int(missing["Missing Count"].sum()) if not missing.empty else 0
    invalid_cells = sum(invalid_counts.values()) + sum(invalid_date_counts.values())
    consistency_cells = inconsistent_cells + mixed_cells
    id_uniqueness = []
    if not duplicate_values.empty:
        for _, row in duplicate_values[duplicate_values["Looks Like an ID Column"]].iterrows():
            if row["Non-Null Values"] > 0:
                id_uniqueness.append(100 - row["Duplicate Values"] / row["Non-Null Values"] * 100)
    score = calculate_quality_score(
        total_rows=len(analysis_df), total_columns=analysis_df.shape[1], missing_cells=missing_cells,
        duplicate_rows=duplicates["duplicate_rows"], id_uniqueness_percentages=id_uniqueness,
        invalid_cells=invalid_cells, inconsistent_cells=consistency_cells)
    extra = type_findings + invalid_findings + date_findings + category_findings
    issues = _safe_run(results, "Issue List", build_issue_list, pd.DataFrame(columns=ISSUE_COLUMNS),
                       missing, duplicates, duplicate_values, empty_columns, constant_columns, outliers, extra)
    recommendations = _safe_run(results, "Recommendations", generate_recommendations, [], issues)
    results.update({
        "missing": missing, "duplicates": duplicates, "duplicate_values": duplicate_values,
        "data_types": data_types, "empty_columns": empty_columns, "constant_columns": constant_columns,
        "outliers": outliers, "numeric_quality": numeric_quality, "categorical": categorical,
        "datetime_quality": datetime_quality, "score": score, "issues": issues,
        "recommendations": recommendations,
    })
    return results
# =============================================================================
# SECTION 9: CONNECTING TO PART 1 (DATA UPLOAD)
# =============================================================================
def get_uploaded_dataframe():
    """
    Fetch the DataFrame created by Part 1 (Data Upload) from st.session_state.
    Returns (DataFrame or None, dataset_name). None means Part 1 has not stored
    a dataset yet. The DataFrame is returned as-is: this module never modifies it.
    """
    # Unified app integration: Part 1 stores an UploadedDataset object.
    uploaded = st.session_state.get(SESSION_KEY)
    if isinstance(uploaded, UploadedDataset):
        return uploaded.dataframe, uploaded.name

    dataframe = None
    for key in DATAFRAME_SESSION_KEYS:
        candidate = st.session_state.get(key)
        if isinstance(candidate, pd.DataFrame):
            dataframe = candidate
            break
    dataset_name = "Uploaded dataset"
    for key in DATASET_NAME_SESSION_KEYS:
        candidate = st.session_state.get(key)
        if candidate:
            dataset_name = str(candidate)
            break
    return dataframe, dataset_name
def dataframe_fingerprint(df):
    """
    A cheap identity for a DataFrame, used to avoid recomputing the report when
    Streamlit reruns the script without the data having changed.
    """
    try:
        content_hash = int(pd.util.hash_pandas_object(df, index=True).sum())
    except Exception:  # unhashable cell contents: fall back to object identity
        content_hash = id(df)
    return (df.shape, tuple(map(str, df.columns)), tuple(map(str, df.dtypes)), content_hash)
def get_or_compute_results(df, dataset_name):
    """Return cached results for this exact dataset, or run the analysis and cache it in the session."""
    fingerprint = (dataset_name, dataframe_fingerprint(df))
    cached = st.session_state.get("dq_report_cache")
    if cached and cached["fingerprint"] == fingerprint:
        return cached["results"]
    results = run_quality_analysis(df, dataset_name)
    st.session_state["dq_report_cache"] = {"fingerprint": fingerprint, "results": results}
    return results
# =============================================================================
# SECTION 10: STREAMLIT DISPLAY FUNCTIONS
# Each function below draws ONE section of the report. The comments explain the
# link between the Python code, the user-facing text, and what appears on screen.
# =============================================================================
def display_overview(results):
    """Dataset Overview: name, size, and column-type counts."""
    overview = results["overview"]
    # The heading "Dataset Overview" appears as a section title at the top of the
    # report and tells the user that the metrics below are high-level facts
    # (size and column types) about the uploaded dataset.
    st.subheader("Dataset Overview")
    # The line "Dataset name: <file name>" appears under the heading and shows
    # which uploaded dataset this report describes (the name Part 1 stored).
    st.markdown(f"**Dataset name:** `{overview['dataset_name']}`")
    # st.columns(4) makes four side-by-side slots. Each st.metric below shows a
    # label (for example "Rows") with its number, computed from the DataFrame shape.
    first_row = st.columns(4)
    first_row[0].metric("Rows", f"{overview['rows']:,}")                # visible label "Rows": number of records
    first_row[1].metric("Columns", f"{overview['columns']:,}")           # visible label "Columns": number of fields
    first_row[2].metric("Total Cells", f"{overview['total_cells']:,}")   # rows x columns
    first_row[3].metric("Numerical Columns", f"{overview['numeric_columns']:,}")
    second_row = st.columns(4)
    second_row[0].metric("Categorical Columns", f"{overview['categorical_columns']:,}")  # text/category/boolean columns
    second_row[1].metric("Date/Time Columns", f"{overview['datetime_columns']:,}")       # real dates + date-looking text
def display_score(results):
    """Overall Data Quality Score and its four dimensions."""
    score = results["score"]
    # The heading "Overall Data Quality Score" introduces the single 0-100 number
    # that summarises the dataset, plus the four dimensions it is built from.
    st.subheader("Overall Data Quality Score")
    # The big metric shows the score as, for example, "87/100". It is the average
    # of Completeness, Uniqueness, Validity and Consistency computed by
    # calculate_quality_score().
    st.metric("Overall Data Quality Score", f"{score['overall']:.0f}/100")
    # A colored message shows the score band. Green ("success") for 90+, blue
    # ("info") for 75-89, and yellow ("warning") below 75. The text tells the user
    # the band name and reminds them these bands are this application's own scale.
    message = f"Rating: {score['rating']} (score bands used by this application: 90+ Excellent, 75-89 Good, 50-74 Fair, below 50 Needs attention)."
    if score["overall"] >= 90:
        st.success(message)
    elif score["overall"] >= 75:
        st.info(message)
    else:
        st.warning(message)
    # Four metrics show each dimension as a percentage under the overall score.
    columns = st.columns(4)
    columns[0].metric("Completeness", f"{score['completeness']:.1f}%")  # share of cells that are not missing
    columns[1].metric("Uniqueness", f"{score['uniqueness']:.1f}%")      # freedom from duplicate rows / repeated IDs
    columns[2].metric("Validity", f"{score['validity']:.1f}%")          # share of values not flagged invalid
    columns[3].metric("Consistency", f"{score['consistency']:.1f}%")    # share of values written consistently
    # The expander titled "How is this score calculated?" is collapsed by default;
    # clicking it reveals the plain-language formulas so the score is transparent.
    with st.expander("How is this score calculated?"):
        # This block of text appears inside the expander and explains each formula
        # and shows the actual counts used for THIS dataset.
        st.markdown(
            "The score is the simple average of four dimensions (each 0-100%).\n\n"
            f"- **Completeness** = (1 - missing cells / total cells) x 100. Missing cells here: {score['missing_cells']:,}.\n"
            f"- **Uniqueness** = (1 - duplicate rows / total rows) x 100. If some columns are named like identifiers "
            f"(id, key, uuid...), it is the average of that value and the share of unique values in those columns. "
            f"Duplicate rows here: {score['duplicate_rows']:,}.\n"
            f"- **Validity** = (1 - invalid cells / non-missing cells) x 100. Invalid cells here: {score['invalid_cells']:,} "
            f"(infinite values, negatives or impossible ages in columns whose names suggest they should not have them, "
            f"placeholder text like N/A, non-numeric entries in numeric columns, unparseable dates).\n"
            f"- **Consistency** = (1 - inconsistent cells / non-missing cells) x 100. Inconsistent cells here: "
            f"{score['inconsistent_cells']:,} (category spellings that differ from the most common spelling, and "
            f"minority types in mixed-type columns).\n\n"
            "The same dataset always produces the same score. The rules and thresholds are indicators chosen by this "
            "application, not universal standards."
        )
def display_missing_values(results):
    """Completeness section: missing values per column."""
    missing = results["missing"]
    # The heading "Missing Values" (Completeness) marks the section that shows how
    # much data is absent in each column.
    st.subheader("Missing Values")
    if missing.empty:
        # This warning appears if the missing-value check failed, so the user knows the section is unavailable.
        st.warning("Missing-value information is not available for this dataset.")
        return
    with_missing = missing[missing["Missing Count"] > 0]
    # A caption under the heading explains that the labels are this application's own indicators.
    st.caption("Severity labels are data-quality indicators used by this application "
               f"(Complete = 0%, Low concern < {MISSING_LOW_MAX:g}%, Moderate concern < {MISSING_MODERATE_MAX:g}%, "
               "High concern otherwise). They are not universal rules.")
    if with_missing.empty:
        # This green message appears when every cell has a value.
        st.success("No missing values were detected. Every column is complete.")
    else:
        # A summary sentence tells the user how many columns contain missing data.
        st.markdown(f"**{len(with_missing)} of {len(missing)} columns contain missing values.** "
                    f"Columns with the highest percentage of missing data:")
        # This table lists (up to) the five columns with the highest missing percentage.
        st.dataframe(with_missing.head(5), hide_index=True)
    # The expander "All columns: missing-value details" holds the full table (every column,
    # including complete ones) so nothing is hidden from the user.
    with st.expander("All columns: missing-value details"):
        st.dataframe(missing, hide_index=True)
def display_duplicate_rows(results):
    """Duplicate rows count and percentage."""
    duplicates = results["duplicates"]
    # The heading "Duplicate Rows" starts the section about rows that repeat exactly.
    st.subheader("Duplicate Rows")
    # The metric shows the label "Duplicate Rows" with the count (for example 342)
    # and a small note such as "2.4% of dataset" that gives the percentage.
    st.metric("Duplicate Rows", f"{duplicates['duplicate_rows']:,}",
              delta=f"{duplicates['percentage']:.1f}% of dataset", delta_color="off")
    # A caption explains how duplicates are counted.
    st.caption("A row is counted as a duplicate when every column matches an earlier row. The first occurrence is not counted.")
def display_duplicate_values(results):
    """Duplicate values per column."""
    table = results["duplicate_values"]
    # The heading "Duplicate Values" starts the section about repeated values within each column.
    st.subheader("Duplicate Values")
    # This caption explains that repeated values are not automatically a problem.
    st.caption("Uniqueness expectations depend on the purpose of a column. Repeated values are normal in a column such as "
               "Country or Gender, but usually unexpected in an identifier column. Columns whose names look like "
               "identifiers are marked in the last column.")
    if table.empty:
        st.warning("Duplicate-value information is not available for this dataset.")
        return
    # The table shows, for each column, the unique values, duplicate values and duplicate percentage.
    st.dataframe(table, hide_index=True)
def display_data_types(results):
    """Data types and potential type concerns."""
    table = results["data_types"]
    # The heading "Data Types" starts the section describing how each column is stored.
    st.subheader("Data Types")
    # The caption reminds the user that no data has been converted.
    st.caption("The report only detects potential type problems. Your original data has not been changed.")
    if table.empty:
        st.warning("Data-type information is not available for this dataset.")
        return
    # The table lists each column's detected type, non-null count, unique count and any potential concern.
    st.dataframe(table, hide_index=True)
def display_empty_columns(results):
    """Empty columns."""
    table = results["empty_columns"]
    # The heading "Empty Columns" starts the section about columns with no data.
    st.subheader("Empty Columns")
    if table.empty:
        # This green message appears when every column contains at least one value.
        st.success("No empty columns were detected.")
    else:
        # This warning appears when at least one column has no data; the table below names them.
        st.warning(f"{len(table)} empty column(s) detected: every value is missing or blank.")
        st.dataframe(table, hide_index=True)
def display_constant_columns(results):
    """Constant columns."""
    table = results["constant_columns"]
    # The heading "Constant Columns" starts the section about columns with only one distinct value.
    st.subheader("Constant Columns")
    if table.empty:
        # This green message appears when no column has just one repeated value.
        st.success("No constant columns were detected.")
    else:
        # This information message explains what a constant column means and that it is not automatically bad.
        st.info(f"{len(table)} constant column(s) detected. A column where every value is the same may provide "
                "limited analytical value, but it should not be deleted automatically: it may still be needed as context.")
        st.dataframe(table, hide_index=True)
def display_numeric_quality(results):
    """Numeric data quality summary."""
    table = results["numeric_quality"]
    # The heading "Numeric Data Quality" starts the section with summary statistics for numerical columns.
    st.subheader("Numeric Data Quality")
    if table.empty:
        # This message appears when the dataset has no numerical columns.
        st.info("No numerical columns were found in this dataset.")
        return
    # The table shows min, max, mean, median, standard deviation, missing, unique and potential outlier counts.
    st.dataframe(table, hide_index=True)
def display_outliers(results):
    """Potential outliers by IQR."""
    table = results["outliers"]
    # The heading "Potential Outliers" starts the section that lists unusually high/low numeric values.
    st.subheader("Potential Outliers")
    if table.empty:
        # This message appears when there are no numerical columns to check.
        st.info("No numerical columns were found, so outlier detection was skipped.")
        return
    # The caption explains the method and that outliers are not automatically errors.
    st.caption("Method: IQR. Values below Q1 - 1.5 x IQR or above Q3 + 1.5 x IQR are flagged as POTENTIAL outliers. "
               "A potential outlier is not necessarily an error; it may be a legitimate observation.")
    # The table shows Q1, Q3, IQR, bounds, the number and percentage of potential outliers for each numerical column.
    st.dataframe(table, hide_index=True)
def display_categorical_consistency(results):
    """Categorical consistency."""
    table = results["categorical"]
    # The heading "Categorical Consistency" starts the section that looks for the same category written in different ways.
    st.subheader("Categorical Consistency")
    if table.empty:
        # This green message appears when no suspicious spelling variations were found.
        st.success("No potentially inconsistent categorical values were detected.")
        return
    # This warning explains what the table shows and that nothing was changed.
    st.warning("Some values may represent the same category written differently (for example 'New York', "
               "'new york', 'NEW YORK'). These are possibilities, not certainties. Nothing has been changed.")
    st.dataframe(table, hide_index=True)
def display_datetime_quality(results):
    """Date/time quality."""
    table = results["datetime_quality"]
    # The heading "Date/Time Quality" starts the section about date columns.
    st.subheader("Date/Time Quality")
    if table.empty:
        # This message appears when no date or date-looking columns were detected.
        st.info("No date/time columns were detected in this dataset.")
        return
    # The caption clarifies that text dates are only parsed temporarily for measurement.
    st.caption("Text columns that look like dates are parsed only to measure quality. The original values are not modified.")
    st.dataframe(table, hide_index=True)
def display_issues(results):
    """Centralized data-quality issue table with a severity filter."""
    issues = results["issues"]
    # The heading "Data Quality Issues" starts the central list of everything the checks detected.
    st.subheader("Data Quality Issues")
    if issues.empty:
        # This green message appears when the checks found nothing to report.
        st.success("No data quality issues were detected by the checks in this report.")
        return
    # Three metrics show how many High / Medium / Low severity issues were found.
    counts = issues["Severity"].value_counts()
    columns = st.columns(3)
    columns[0].metric("High Severity", int(counts.get("High", 0)))
    columns[1].metric("Medium Severity", int(counts.get("Medium", 0)))
    columns[2].metric("Low Severity", int(counts.get("Low", 0)))
    # The caption reminds the user that severity is an indicator.
    st.caption("Severity is an indicator based on thresholds used by this application. Findings describe patterns detected "
               "in the data; please review them in the context of your dataset.")
    # The multiselect labelled "Show severities" lets the user choose which severities appear in the table below.
    selected = st.multiselect("Show severities", ["High", "Medium", "Low"], default=["High", "Medium", "Low"],
                              key="dq_severity_filter")
    filtered = issues[issues["Severity"].isin(selected)]
    # The table lists Severity, Column, Issue and Details for every issue that matches the filter.
    st.dataframe(filtered, hide_index=True)
def generate_pdf_report(results):
    """
    Builds a downloadable PDF summarizing one dataset's quality report:
    overview, score, every detected issue, and the recommendations.
    Returns raw PDF bytes, ready to hand to st.download_button.

    Uses explicit new_x/new_y cursor positioning on every cell/multi_cell
    call - fpdf2's newer versions do not reset the cursor to the left
    margin by default, which silently corrupts layout on any multi_cell
    call after the first one if left unspecified.
    """
    from fpdf import FPDF
    from fpdf.enums import XPos, YPos

    class ReportPDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 16)
            self.cell(0, 10, "Data Quality Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
            self.set_font("Helvetica", "", 10)
            self.set_text_color(120, 120, 120)
            self.cell(0, 6, "Generated by Clarifile", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
            self.set_text_color(0, 0, 0)
            self.ln(4)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "", 8)
            self.set_text_color(150, 150, 150)
            self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def clean(text):
        # PDF's default font only supports latin-1; strip anything
        # outside that range (e.g., curly quotes, backticks used for
        # markdown-style emphasis) rather than letting it crash.
        return str(text).replace("`", "'").encode("latin-1", errors="replace").decode("latin-1")

    pdf = ReportPDF()
    pdf.add_page()

    overview = results["overview"]
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, clean(f"Dataset: {overview['dataset_name']}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 8, clean(f"Rows: {overview['rows']:,}   Columns: {overview['columns']:,}"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)

    score = results["score"]
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean(f"Overall Quality Score: {score['overall']:.0f}/100 ({score['rating']})"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, clean(f"Completeness: {score['completeness']:.1f}%   Uniqueness: {score['uniqueness']:.1f}%"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 7, clean(f"Validity: {score['validity']:.1f}%   Consistency: {score['consistency']:.1f}%"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, "Issues Found", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 10)
    issues = results["issues"]
    if issues.empty:
        pdf.multi_cell(0, 6, "No issues were detected.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        for _, row in issues.iterrows():
            text = f"[{row['Severity']}] {row['Column']} - {row['Issue']}: {row['Details']}"
            pdf.multi_cell(0, 6, clean(text), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, "Recommendations", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 10)
    for i, rec in enumerate(results["recommendations"], start=1):
        pdf.multi_cell(0, 6, clean(f"{i}. {rec}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # dest="S" returns the PDF as a string/bytearray instead of writing
    # to disk - needed so it can be handed directly to st.download_button.
    return bytes(pdf.output())


def display_export_section(results):
    """
    Part 4: Exportable Report. Lets the user download the same
    findings shown in the Data Quality Report as a PDF, or the raw
    issue list as a CSV for further analysis in a spreadsheet.
    """
    st.subheader("Exportable Report")
    st.caption(
        "Download this dataset's quality findings to share or keep, "
        "without needing to come back to this app."
    )

    dataset_name = results["overview"]["dataset_name"]
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in dataset_name)

    col1, col2 = st.columns(2)

    with col1:
        pdf_bytes = generate_pdf_report(results)
        st.download_button(
            "Download PDF Report",
            data=pdf_bytes,
            file_name=f"quality_report_{safe_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col2:
        csv_bytes = results["issues"].to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Issues as CSV",
            data=csv_bytes,
            file_name=f"issues_{safe_name}.csv",
            mime="text/csv",
            use_container_width=True,
        )


def display_recommendations(results):
    """Recommended actions."""
    # The heading "Recommendations" starts the list of suggested next steps.
    st.subheader("Recommendations")
    # The heading "Recommended Actions" (bold text) introduces the numbered list below.
    st.markdown("**Recommended Actions**")
    # Each recommendation appears as a numbered line. They are generated from the issues actually
    # detected in this dataset (see generate_recommendations()).
    st.markdown("\n".join(f"{number}. {text}" for number, text in enumerate(results["recommendations"], start=1)))
    # This caption confirms the report is read-only: it never edits the user's data.
    st.caption("This report is read-only. Your uploaded dataset has not been modified. Apply any fixes to a copy of the data.")
def display_quality_report(df, dataset_name="Uploaded dataset"):
    """
    Entry point for Part 2. app.py calls this with the DataFrame from Part 1.
    Order of the page: overview -> score -> completeness -> duplicates -> types ->
    empty/constant -> numeric/outliers -> categories -> dates -> issues -> recommendations.
    """
    # The heading "Data Quality Report" is the main title of this page, so the
    # user immediately knows which module they are viewing.
    st.header("Data Quality Report")
    # Validate the input first. Depending on the problem, the message below is shown as an
    # error (red) or warning (yellow) box, and the report stops there.
    level, message = check_dataframe(df)
    if level == "error":
        st.error(message)          # red box: the data cannot be used at all (e.g. no columns)
        return
    if level == "warning":
        st.warning(message)        # yellow box: e.g. no dataset uploaded yet, or no rows
        if isinstance(df, pd.DataFrame) and df.shape[1] > 0:
            # A short line lists the columns so the user can see what was uploaded.
            st.caption("Columns found: " + ", ".join(map(str, df.columns)))
        return
    # A spinner with the text "Analyzing dataset quality..." is visible while the checks run.
    try:
        with st.spinner("Analyzing dataset quality..."):
            results = get_or_compute_results(df, dataset_name)
    except Exception:
        # Full technical details go to the log, not to the user.
        logger.exception("Data quality analysis failed")
        # This red box appears if something unexpected stops the entire analysis.
        st.error("The data quality analysis could not be completed for this dataset. "
                 "Please check that the file was read correctly and try again.")
        return
    # This green message confirms that the analysis finished and the report below is ready.
    st.success("Data quality analysis completed.")
    # Any individual check that failed is listed here as a warning, so the user knows a section may be incomplete.
    for warning_text in results["warnings"]:
        st.warning(warning_text)
    display_overview(results)
    st.divider()
    display_score(results)
    st.divider()
    display_missing_values(results)
    st.divider()
    display_duplicate_rows(results)
    display_duplicate_values(results)
    st.divider()
    display_data_types(results)
    st.divider()
    display_empty_columns(results)
    display_constant_columns(results)
    st.divider()
    display_numeric_quality(results)
    display_outliers(results)
    st.divider()
    display_categorical_consistency(results)
    st.divider()
    display_datetime_quality(results)
    st.divider()
    display_issues(results)
    st.divider()
    display_recommendations(results)
    st.divider()
    display_export_section(results)

# ================ PART 3: ANALYTICS & CHARTS ================
"""
analytics.py
Data Quality & Analytics Platform
Part 3: Analytics & Charts Module
This module is intentionally self-contained. It exposes a single public
entry point, render_analytics(df), which the main app.py router calls
after Part 1 (Data Upload) and Part 2 (Data Quality Report) have already
produced and validated a Pandas DataFrame.
Workflow this module expects:
    Data Upload (Part 1)
          |
    Pandas DataFrame
          |
    Data Quality Report (Part 2)
          |
    Analytics & Charts (Part 3)   <-- this file
This file does NOT create sample/fake data. It only analyzes whatever
real DataFrame is handed to it by render_analytics(df).
"""
# ---------------------------------------------------------
# IMPORT REQUIRED LIBRARIES
# ---------------------------------------------------------
# Import Streamlit to build the Analytics & Charts user interface:
# headers, dropdowns, buttons, tabs, and messages.
import streamlit as st
# Import Pandas to inspect and manipulate the uploaded DataFrame
# (column detection, grouping, aggregation, describe(), etc.).
import pandas as pd
# Import NumPy for numeric helper operations (e.g. safe handling of
# infinities/NaNs before they reach Plotly).
import numpy as np
# Import Plotly Express to build interactive charts (histograms, box
# plots, bar charts, scatter plots, heatmaps, line charts) that the
# user can zoom, pan, and hover over inside the Streamlit app.
import plotly.express as px
# ---------------------------------------------------------
# COLUMN DETECTION
# ---------------------------------------------------------
# These three functions classify the DataFrame's columns into the three
# groups the rest of the module needs: numerical, categorical, and
# datetime. Every chart/statistic function below relies on one of these
# instead of re-detecting types itself, so the classification logic
# lives in exactly one place.
def detect_numeric_columns(df):
    """Return the list of column names holding numeric data (int/float)."""
    # select_dtypes(include=["number"]) captures both integer and float
    # columns, which is what histograms, box plots, correlations, and
    # scatter plots all require.
    return df.select_dtypes(include=["number"]).columns.tolist()
def detect_categorical_columns(df):
    """Return categorical/text columns while excluding date-like text columns."""
    # Object/category columns are candidates for categorical analysis.
    candidates = df.select_dtypes(include=["object", "category"]).columns.tolist()
    # Date-like text should be handled by the time-series/date logic instead
    # of being offered as an ordinary category.
    datetime_columns = set(detect_datetime_columns(df))
    return [column for column in candidates if column not in datetime_columns]
def detect_datetime_columns(df):
    """
    Return the list of column names holding datetime data.
    In addition to columns Pandas already parsed as datetime64, this
    also checks object/string columns that LOOK like dates (e.g. a
    "date" column read from CSV as plain text), since Part 1's loaders
    do not always parse dates automatically. A column is only included
    if it can be converted to datetime without producing mostly-invalid
    values, so we don't misclassify ordinary text columns as dates.
    """
    datetime_columns = df.select_dtypes(include=["datetime", "datetimetz"]).columns.tolist()
    # Look for object columns that are actually dates stored as text.
    candidate_columns = df.select_dtypes(include=["object"]).columns.tolist()
    for column in candidate_columns:
        # Skip columns that are already confirmed numeric-like text or
        # empty, since attempting a date conversion on them is wasted
        # work and can raise noisy warnings.
        non_null_sample = df[column].dropna()
        if non_null_sample.empty:
            continue
        # Try converting a small sample first; this keeps detection fast
        # on very large datasets instead of parsing every row up front.
        sample = non_null_sample.head(25)
        try:
            # format="mixed" avoids ambiguous-format warnings on modern pandas
            # while still allowing a column containing more than one date shape.
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                try:
                    parsed_sample = pd.to_datetime(sample, errors="coerce", format="mixed")
                except TypeError:
                    # Compatibility fallback for older pandas releases.
                    parsed_sample = pd.to_datetime(sample, errors="coerce")
        except Exception:
            continue
        # Only treat the column as a real date column if the large
        # majority of the sampled values successfully parsed. This
        # avoids false positives on text columns that merely contain a
        # few number-like tokens.
        if parsed_sample.notna().mean() >= 0.8:
            datetime_columns.append(column)
    return datetime_columns
# ---------------------------------------------------------
# DATASET OVERVIEW
# ---------------------------------------------------------
def render_dataset_overview(df, dataset_name, numeric_columns, categorical_columns, datetime_columns):
    """
    Display the high-level shape of the dataset: its name, row count,
    column count, and how many columns fall into each detected type.
    This orients the user before they dive into specific charts below.
    """
    # Create the "Analytics & Charts" heading. This string produces the
    # visible section title that tells the user they are now viewing
    # the analytics portion of the platform (as opposed to Upload or
    # the Quality Report).
    st.header("Analytics & Charts")
    st.caption("Part 3: Analytics & Charts")
    # "Dataset Overview" subheader introduces the metrics row below it.
    st.subheader("Dataset Overview")
    # st.columns(5) lays out five metric boxes side-by-side rather than
    # stacked, so the whole overview fits in a single glance.
    col1, col2, col3, col4, col5 = st.columns(5)
    # st.metric() renders a large label + number. The label strings
    # ("Dataset", "Rows", "Columns", etc.) are what the user sees above
    # each number in the UI.
    col1.metric("Dataset", dataset_name)
    col2.metric("Rows", f"{df.shape[0]:,}")
    col3.metric("Columns", f"{df.shape[1]:,}")
    col4.metric("Numerical Columns", len(numeric_columns))
    col5.metric("Categorical Columns", len(categorical_columns))
    # Datetime count is shown as a caption rather than a 6th metric box
    # so the layout stays readable on narrower screens.
    st.caption(f"Datetime columns detected: {len(datetime_columns)}")
# ---------------------------------------------------------
# DESCRIPTIVE STATISTICS
# ---------------------------------------------------------
def calculate_descriptive_statistics(df, numeric_columns):
    """
    Build a DataFrame of descriptive statistics (count, mean, median,
    std, min, max, quartiles, range) for every numerical column. This
    is the data used both for the on-screen table and the CSV download.
    """
    stats_rows = []
    for column in numeric_columns:
        series = df[column].dropna()
        if series.empty:
            # Skip columns that are entirely missing; there is nothing
            # meaningful to compute statistics on.
            continue
        stats_rows.append({
            "Column": column,
            "Count": series.count(),
            "Mean": series.mean(),
            "Median": series.median(),
            "Std Dev": series.std(),
            "Min": series.min(),
            "Max": series.max(),
            "Q1 (25%)": series.quantile(0.25),
            "Q3 (75%)": series.quantile(0.75),
            "Range": series.max() - series.min(),
        })
    return pd.DataFrame(stats_rows)
def render_descriptive_statistics(df, numeric_columns):
    """
    Render the "Descriptive Statistics" section: a table of summary
    statistics for every numerical column, plus a CSV download button
    for that table.
    """
    st.subheader("Descriptive Statistics")
    # Guard: statistics require at least one numerical column. Without
    # this check, the table below would simply be empty and confusing.
    if not numeric_columns:
        # st.info() renders a neutral blue message box for a condition
        # that isn't an error, just "nothing to show here."
        st.info("No numerical columns were found, so descriptive statistics are not available.")
        return
    stats_df = calculate_descriptive_statistics(df, numeric_columns)
    # st.dataframe() renders an interactive, scrollable statistics
    # table so the user can inspect every numerical column at once.
    st.dataframe(stats_df, use_container_width=True, hide_index=True)
    # Convert the statistics table to CSV bytes so it can be offered as
    # a download without writing anything to disk on the server.
    csv_bytes = stats_df.to_csv(index=False).encode("utf-8")
    # The "Download Statistics as CSV" label is the visible text on the
    # download button; clicking it saves the descriptive statistics
    # table to the user's computer.
    st.download_button(
        "Download Statistics as CSV",
        data=csv_bytes,
        file_name="descriptive_statistics.csv",
        mime="text/csv",
    )
# ---------------------------------------------------------
# NUMERICAL ANALYSIS: HISTOGRAMS + BOX PLOTS
# ---------------------------------------------------------
def create_histogram(df, column):
    """Build an interactive histogram (distribution) for one numeric column."""
    # px.histogram() plots the frequency distribution of the selected
    # column, letting the user see the shape (skew, spread, modality)
    # of the data.
    figure = px.histogram(df, x=column, title=f"Distribution of {column}")
    return figure
def create_box_plot(df, column):
    """Build an interactive box plot for one numeric column (for outlier visualization)."""
    # px.box() draws quartiles, the median line, and flags points
    # outside 1.5x the interquartile range as outliers, which makes it
    # the standard chart for spotting extreme values.
    figure = px.box(df, y=column, title=f"Box Plot of {column}")
    return figure
def render_numerical_analysis(df, numeric_columns):
    """
    Render the "Distribution Analysis" and "Box Plot Analysis" sections:
    a column selector plus a histogram and box plot for the chosen
    numerical column, used to inspect distribution shape and outliers.
    """
    st.subheader("Numerical Analysis")
    if not numeric_columns:
        st.info("No numerical columns were found, so distribution and box plot charts are not available.")
        return
    # The "Select Numerical Column" string creates the visible dropdown
    # control that lets the user choose which numerical variable to
    # analyze with the histogram and box plot below.
    selected_column = st.selectbox(
        "Select Numerical Column",
        numeric_columns,
        key="numeric_analysis_column",
    )
    # Two side-by-side charts: distribution shape on the left, outlier
    # view on the right, both driven by the same selected column.
    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.markdown("**Distribution (Histogram)**")
        try:
            st.plotly_chart(create_histogram(df, selected_column), use_container_width=True)
        except Exception:
            # st.error() communicates that chart generation itself
            # failed (e.g. unusable/non-numeric values slipped through
            # detection), without exposing a raw traceback to the user.
            st.error(f"Unable to generate a histogram for '{selected_column}'.")
    with chart_col2:
        st.markdown("**Outliers (Box Plot)**")
        try:
            st.plotly_chart(create_box_plot(df, selected_column), use_container_width=True)
        except Exception:
            st.error(f"Unable to generate a box plot for '{selected_column}'.")
# ---------------------------------------------------------
# CATEGORICAL ANALYSIS
# ---------------------------------------------------------
def create_bar_chart(category_counts, column, top_n):
    """
    Build a bar chart of the top-N most frequent categories for one
    categorical column, using pre-computed value counts.
    """
    figure = px.bar(
        category_counts,
        x=column,
        y="Count",
        title=f"Top {top_n} Categories in {column}",
        text="Percentage",
    )
    # Format the percentage labels shown above each bar so they read
    # as "12.3%" instead of a raw float.
    figure.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
    return figure
def render_categorical_analysis(df, categorical_columns):
    """
    Render the "Categorical Analysis" section: category counts,
    percentages, and a bar chart of the top-N categories for a
    user-selected categorical column.
    """
    st.subheader("Categorical Analysis")
    if not categorical_columns:
        st.info("No categorical columns were found, so category counts and bar charts are not available.")
        return
    # "Select Categorical Column" creates the dropdown that lets the
    # user pick which text/category column to break down below.
    selected_column = st.selectbox(
        "Select Categorical Column",
        categorical_columns,
        key="categorical_analysis_column",
    )
    # "Number of Top Categories to Display" creates a slider so the
    # user can control how many bars appear in the chart, since some
    # categorical columns may have dozens or hundreds of unique values.
    unique_count = df[selected_column].nunique(dropna=True)
    if unique_count == 0:
        st.warning(f"'{selected_column}' has no non-missing values to analyze.")
        return
    max_top_n = min(unique_count, 50)
    top_n = st.slider(
        "Number of Top Categories to Display",
        min_value=1,
        max_value=max_top_n,
        value=min(10, max_top_n),
        key="categorical_top_n",
    )
    # Compute counts and percentages for the top-N categories.
    value_counts = df[selected_column].value_counts(dropna=True).head(top_n)
    category_counts = pd.DataFrame({
        selected_column: value_counts.index.astype(str),
        "Count": value_counts.values,
    })
    category_counts["Percentage"] = (category_counts["Count"] / df[selected_column].notna().sum()) * 100
    # Table view: exact counts and percentages for the selected top-N
    # categories, useful for readers who want precise numbers.
    st.dataframe(category_counts, use_container_width=True, hide_index=True)
    # Chart view: the same data as a bar chart for quick visual
    # comparison between categories.
    try:
        st.plotly_chart(
            create_bar_chart(category_counts, selected_column, top_n),
            use_container_width=True,
        )
    except Exception:
        st.error(f"Unable to generate a bar chart for '{selected_column}'.")
# ---------------------------------------------------------
# CORRELATION ANALYSIS
# ---------------------------------------------------------
def create_correlation_heatmap(df, numeric_columns):
    """Build a correlation heatmap across all numerical columns."""
    # .corr() computes pairwise Pearson correlation coefficients
    # between every pair of numerical columns, producing a square
    # matrix of values between -1 and 1.
    correlation_matrix = df[numeric_columns].corr()
    figure = px.imshow(
        correlation_matrix,
        text_auto=".2f",
        color_continuous_scale="RdBu_r",
        zmin=-1,
        zmax=1,
        title="Correlation Heatmap",
    )
    return figure, correlation_matrix
def render_correlation_analysis(df, numeric_columns):
    """
    Render the "Correlation Analysis" section: a correlation matrix
    table and heatmap showing how numerical columns relate to one
    another.
    """
    st.subheader("Correlation Analysis")
    # Correlation requires at least two numerical columns to be
    # meaningful; with fewer than two there is nothing to correlate.
    if len(numeric_columns) < 2:
        st.info("At least two numerical columns are required to calculate correlations.")
        return
    try:
        heatmap_figure, correlation_matrix = create_correlation_heatmap(df, numeric_columns)
    except Exception:
        st.error("Unable to calculate correlations for the numerical columns in this dataset.")
        return
    st.plotly_chart(heatmap_figure, use_container_width=True)
    # Expander keeps the raw numeric matrix available without cluttering
    # the page for users who only want the visual heatmap.
    with st.expander("View Correlation Matrix (Table)"):
        st.dataframe(correlation_matrix, use_container_width=True)
# ---------------------------------------------------------
# SCATTER PLOT ANALYSIS
# ---------------------------------------------------------
def create_scatter_plot(df, x_column, y_column, group_column=None):
    """
    Build an interactive scatter plot of x_column vs y_column, with an
    optional categorical column used to color-code the points by
    group.
    """
    figure = px.scatter(
        df,
        x=x_column,
        y=y_column,
        color=group_column if group_column else None,
        title=f"{y_column} vs {x_column}",
    )
    return figure
def render_scatter_analysis(df, numeric_columns, categorical_columns):
    """
    Render the "Scatter Plot Analysis" section: X/Y axis selectors, an
    optional grouping column, and the resulting interactive scatter
    plot showing the relationship between two numerical variables.
    """
    st.subheader("Scatter Plot Analysis")
    # A scatter plot needs at least two numerical columns: one for each
    # axis.
    if len(numeric_columns) < 2:
        st.info("At least two numerical columns are required to build a scatter plot.")
        return
    axis_col1, axis_col2, axis_col3 = st.columns(3)
    with axis_col1:
        # "Select X-Axis Column" creates the dropdown that sets which
        # numerical variable is plotted along the horizontal axis.
        x_column = st.selectbox("Select X-Axis Column", numeric_columns, index=0, key="scatter_x")
    with axis_col2:
        # "Select Y-Axis Column" creates the dropdown for the vertical
        # axis; defaults to the second numeric column so X and Y are
        # not the same column by default.
        default_y_index = 1 if len(numeric_columns) > 1 else 0
        y_column = st.selectbox("Select Y-Axis Column", numeric_columns, index=default_y_index, key="scatter_y")
    with axis_col3:
        # "Group By (Optional)" lets the user color-code points by a
        # categorical column, e.g. to compare groups within the same
        # scatter plot. "None" means no grouping is applied.
        group_options = ["None"] + categorical_columns
        group_selection = st.selectbox("Group By (Optional)", group_options, key="scatter_group")
        group_column = None if group_selection == "None" else group_selection
    try:
        st.plotly_chart(
            create_scatter_plot(df, x_column, y_column, group_column),
            use_container_width=True,
        )
    except Exception:
        st.error(f"Unable to generate a scatter plot for '{x_column}' vs '{y_column}'.")
# ---------------------------------------------------------
# TIME-SERIES ANALYSIS
# ---------------------------------------------------------
# Maps each user-facing aggregation label to the Pandas resample rule
# that implements it. Centralizing this mapping keeps the UI strings
# and the resampling logic in sync.
AGGREGATION_RULES = {
    # These aliases are broadly supported across current pandas releases.
    "Day": "D",
    "Week": "W",
    "Month": "MS",
    "Quarter": "QS",
    "Year": "YS",
}
def create_time_series_chart(df, date_column, value_column, aggregation_label):
    """
    Build a line chart of value_column over time (date_column), after
    aggregating (summing) values at the requested time granularity
    (day/week/month/quarter/year).
    """
    # Work on a copy limited to the two relevant columns so the
    # aggregation below doesn't affect the caller's DataFrame.
    working_df = df[[date_column, value_column]].copy()
    # Convert the date column to real datetime values; invalid dates
    # become NaT (Not a Time) rather than raising an error.
    working_df[date_column] = pd.to_datetime(working_df[date_column], errors="coerce")
    # Drop rows where either the date or the value is missing/invalid,
    # since neither can be plotted on a time-series line chart.
    working_df = working_df.dropna(subset=[date_column, value_column])
    if working_df.empty:
        raise ValueError("No valid date/value pairs are available for this combination.")
    # Aggregate (sum) the value column at the chosen time granularity.
    # .resample() requires a DatetimeIndex, so the date column is set
    # as the index first.
    resample_rule = AGGREGATION_RULES[aggregation_label]
    aggregated = (
        working_df.set_index(date_column)[value_column]
        .resample(resample_rule)
        .sum()
        .reset_index()
    )
    figure = px.line(
        aggregated,
        x=date_column,
        y=value_column,
        title=f"{value_column} Over Time ({aggregation_label}ly Aggregation)",
        markers=True,
    )
    return figure
def render_time_series_analysis(df, datetime_columns, numeric_columns):
    """
    Render the "Time-Series Analysis" section: datetime and numeric
    column selectors, an aggregation-level control, and the resulting
    line chart of the numeric value over time.
    """
    st.subheader("Time-Series Analysis")
    # Time-series analysis needs at least one datetime column and at
    # least one numerical column to plot against it.
    if not datetime_columns:
        st.info("No datetime columns were detected, so time-series charts are not available.")
        return
    if not numeric_columns:
        st.info("No numerical columns were found, so there is no value to plot over time.")
        return
    control_col1, control_col2, control_col3 = st.columns(3)
    with control_col1:
        # "Select Date Column" creates the dropdown that chooses which
        # detected datetime column defines the time axis.
        date_column = st.selectbox("Select Date Column", datetime_columns, key="ts_date_column")
    with control_col2:
        # "Select Value Column" creates the dropdown that chooses which
        # numerical column is aggregated and plotted over time.
        value_column = st.selectbox("Select Value Column", numeric_columns, key="ts_value_column")
    with control_col3:
        # "Aggregate By" creates the dropdown that controls the time
        # granularity (day/week/month/quarter/year) used to group and
        # sum the value column before plotting.
        aggregation_label = st.selectbox(
            "Aggregate By",
            list(AGGREGATION_RULES.keys()),
            index=2,
            key="ts_aggregation",
        )
    try:
        chart = create_time_series_chart(df, date_column, value_column, aggregation_label)
        st.plotly_chart(chart, use_container_width=True)
    except ValueError as error:
        st.warning(str(error))
    except Exception:
        st.error(f"Unable to generate a time-series chart for '{value_column}' over '{date_column}'.")
# ---------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------
def render_analytics(df, dataset_name="Uploaded Dataset"):
    """
    Public entry point for Part 3: Analytics & Charts.
    The main app.py router should call this function AFTER a real
    DataFrame has been produced by Part 1 (Data Upload) and reviewed by
    Part 2 (Data Quality Report), passing that same DataFrame in as
    `df`. This function never generates or substitutes sample data; if
    `df` is missing or empty, it shows guidance instead of a chart.
    Example integration from app.py:
        from analytics import render_analytics
        render_analytics(st.session_state["dataframe"], st.session_state["dataset_name"])
    """
    # -------------------------------------------------------------
    # GUARD: no dataset uploaded yet, or df is not usable.
    # -------------------------------------------------------------
    if df is None:
        # st.info() (not st.error()) because "nothing uploaded yet" is
        # an expected state, not a failure.
        st.info("No dataset has been uploaded yet. Please upload a file in the Data Upload section first.")
        return
    if not isinstance(df, pd.DataFrame):
        st.error("The provided data is not a valid table and cannot be analyzed.")
        return
    if df.empty or df.shape[1] == 0:
        st.warning("The uploaded dataset is empty, so there is nothing to analyze yet.")
        return
    # Replace infinite values with NaN up front so downstream stats and
    # charts (which do not handle +/-inf gracefully) behave predictably.
    df = df.replace([np.inf, -np.inf], np.nan)
    # -------------------------------------------------------------
    # COLUMN DETECTION: classify columns once, reuse everywhere below.
    # -------------------------------------------------------------
    numeric_columns = detect_numeric_columns(df)
    categorical_columns = detect_categorical_columns(df)
    datetime_columns = detect_datetime_columns(df)
    # -------------------------------------------------------------
    # RENDER EACH SECTION IN ORDER.
    # Each section function contains its own guard clauses, so a
    # dataset missing one column type (e.g. no dates) still renders
    # every other applicable section normally.
    # -------------------------------------------------------------
    render_dataset_overview(df, dataset_name, numeric_columns, categorical_columns, datetime_columns)
    st.divider()
    render_descriptive_statistics(df, numeric_columns)
    st.divider()
    render_numerical_analysis(df, numeric_columns)
    st.divider()
    render_categorical_analysis(df, categorical_columns)
    st.divider()
    render_correlation_analysis(df, numeric_columns)
    st.divider()
    render_scatter_analysis(df, numeric_columns, categorical_columns)
    st.divider()
    render_time_series_analysis(df, datetime_columns, numeric_columns)
# -----------------------------------------------------------------------
# STANDALONE TEST MODE
# -----------------------------------------------------------------------
# This block allows analytics.py to be run directly (streamlit run
# analytics.py) for isolated testing, WITHOUT faking data as part of the
# normal render_analytics() flow. It only activates when a file is
# uploaded here directly, and is skipped entirely when this module is
# imported by app.py.


# ============================================================================
# UNIFIED APPLICATION INTEGRATION
# ============================================================================
# The upload module stores one UploadedDataset in Streamlit session state.
# The quality report and analytics modules receive the same DataFrame directly,
# which avoids duplicated loaders and prevents stale module-to-module copies.

def get_current_dataset():
    """Return the current UploadedDataset stored by Part 1, or None."""
    return get_uploaded_dataset()


def run_data_upload_page():
    """Render Part 1: Data Upload."""
    st.header("Data Upload")
    st.caption("Part 1: Upload and inspect your dataset.")

    uploaded_file = st.file_uploader(
        "Choose a file",
        type=SUPPORTED_EXTENSIONS,
        key="main_dataset_uploader",
    )
    st.caption(
        "Supported formats: " +
        ", ".join(ext.upper() for ext in SUPPORTED_EXTENSIONS)
    )

    if uploaded_file is None:
        # Clear stale data when the uploader is emptied.
        clear_uploaded_dataset()
        st.info("Upload a file to get started.")
        return

    try:
        # getvalue() reads the upload without depending on a file-pointer position.
        raw_bytes = uploaded_file.getvalue()
        dataset = process_upload(uploaded_file.name, raw_bytes)
    except UploadError as error:
        clear_uploaded_dataset()
        st.error(str(error))
        return
    except MemoryError:
        clear_uploaded_dataset()
        st.error("This file is too large to process on this computer.")
        return
    except Exception:
        clear_uploaded_dataset()
        logger.exception("Unexpected upload failure")
        st.error(UNEXPECTED_ERROR_MESSAGE)
        return

    # A header-only dataset is valid for loading but not useful for analysis.
    if dataset.dataframe.shape[0] == 0:
        clear_uploaded_dataset()
        st.warning("This dataset has no rows - only column headers were found.")
        return

    # Avoid reparsing the same file on every Streamlit rerun.
    current = get_current_dataset()
    if current is None or (
        current.name != dataset.name
        or current.file_size_bytes != dataset.file_size_bytes
    ):
        store_uploaded_dataset(dataset)
        # A new dataset must invalidate the Part 2 cache.
        st.session_state.pop("dq_report_cache", None)
        st.session_state.pop("dq_report_fingerprint", None)

    dataset = get_current_dataset()

    st.success(f"'{dataset.name}' was uploaded and read successfully.")
    display_dataset_information(dataset)

    if dataset.image is not None:
        display_image_preview(dataset.image)

    display_dataset_preview(dataset.dataframe)
    display_column_information(dataset.dataframe)
    display_basic_summary(dataset.dataframe)


def run_quality_report_page():
    """Render Part 2: Data Quality Report using Part 1's DataFrame."""
    dataset = get_current_dataset()

    if dataset is None:
        st.header("Data Quality Report")
        st.info("Upload a dataset in Part 1 before opening the Data Quality Report.")
        return

    display_quality_report(dataset.dataframe, dataset.name)


def run_analytics_page():
    """Render Part 3: Analytics & Charts using Part 1's DataFrame."""
    dataset = get_current_dataset()

    if dataset is None:
        st.header("Analytics & Charts")
        st.info("Upload a dataset in Part 1 before opening Analytics & Charts.")
        return

    render_analytics(dataset.dataframe, dataset.name)


def main():
    """Start the three-part Streamlit application."""
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title(APP_TITLE)
    st.caption("Data Upload → Data Quality Report → Analytics & Charts")

    page = st.sidebar.radio(
        "Navigation",
        ["1. Data Upload", "2. Data Quality Report", "3. Analytics & Charts"],
        key="main_navigation",
    )

    if page == "1. Data Upload":
        run_data_upload_page()
    elif page == "2. Data Quality Report":
        run_quality_report_page()
    else:
        run_analytics_page()


if __name__ == "__main__":
    main()
