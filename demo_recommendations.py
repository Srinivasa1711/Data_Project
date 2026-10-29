"""
demo_recommendations.py -- stand-alone test harness for Part 4 (Recommendations).

Replaces the old app.py harness, which also imported analytics.py (not part of
this package). Run with: streamlit run demo_recommendations.py
"""
import pandas as pd
import streamlit as st
from recommendations_engine import render_recommendations

# Must be the first Streamlit call.
st.set_page_config(page_title="Recommendations - Standalone Test", layout="wide")

uploaded = st.sidebar.file_uploader("Upload a dataset", type=["csv", "xlsx"])
if uploaded is None:
    st.info("Upload a CSV or Excel file in the sidebar to begin.")
else:
    try:
        data = (pd.read_csv(uploaded) if uploaded.name.endswith(".csv")
                else pd.read_excel(uploaded))
    except Exception as error:  # noqa: BLE001 - shown to the user
        st.error(f"The file could not be read: {error}")
    else:
        render_recommendations(data)
