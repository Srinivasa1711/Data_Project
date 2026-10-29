"""
demo_app.py
Standalone test harness for Part 5. In the full platform, the main app passes
the DataFrame and the results of Parts 1-4 straight into render_exportable_report().
Run with: streamlit run demo_app.py
"""
# Streamlit draws the page; pandas reads the uploaded CSV file.
import pandas as pd
import streamlit as st
from exportable_report import render_exportable_report

# Must be the first Streamlit call. Sets the browser tab title and page layout.
st.set_page_config(page_title="Exportable Report - Standalone Test", layout="wide")

# Stand-in for the Data Upload part: lets you test Part 5 with your own CSV.
# Only the DataFrame is passed on, so quality score and recommendations are
# reported as "not available" instead of being invented.
uploaded = st.sidebar.file_uploader("Upload a CSV file to test Part 5", type=["csv"])
df = None
if uploaded is not None:
    try:
        df = pd.read_csv(uploaded)
    except Exception as exc:  # noqa: BLE001 - shown to the user
        # Tell the user why the file could not be read.
        st.sidebar.error(f"Could not read the CSV file: {exc}")

render_exportable_report(df, quality_results=None, analytics_results=None, recommendations=None)
