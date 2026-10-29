"""
exportable_report.py
====================
Part 5 - Exportable Report (Streamlit interface).

Call render_exportable_report(df, quality_results, analytics_results,
recommendations) from the main platform app. All heavy lifting (building the
content, PDF, Excel, CSV) lives in report_builder.py; this file only draws the
interface and wires the controls to those functions.
"""

# ---------------------------------------------------------------------------
# IMPORTS
# ---------------------------------------------------------------------------
# Streamlit draws the interface; pandas is used only to check the input type.
import pandas as pd
import streamlit as st

# Report engine: content building, validation constants and the export dispatcher.
from report_builder import (
    CSV_TABLE_OPTIONS,
    DEFAULT_REPORT_TITLE,
    REPORT_FORMATS,
    export_report,
    generate_report,
)


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def _md(text):
    # Escape Markdown characters so column names or findings that contain
    # symbols such as * _ or $ are shown literally instead of being formatted.
    escaped = str(text)
    for ch in ("\\", "`", "*", "_", "[", "]", "<", ">", "#", "|", "$", "~"):
        escaped = escaped.replace(ch, "\\" + ch)
    return escaped


def _show_table(frame, empty_message):
    # Show a results table, or an explanatory message when it has no rows.
    if frame is None or frame.empty:
        # Tell the user that this check produced no rows, so an empty table is
        # understood as "nothing found" rather than "something failed".
        st.info(empty_message)
    else:
        # Display the computed table so the user can review it before export.
        st.dataframe(frame, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# REPORT PREVIEW
# ---------------------------------------------------------------------------
def render_report_preview(content):
    """Show every report section on screen, using the same content the exports use."""
    summary, quality, analytics = content["summary"], content["quality"], content["analytics"]
    recs = content["recommendations"]

    # ---------------------------------------------------------------------
    # REPORT PREVIEW HEADING
    # ---------------------------------------------------------------------
    # Display the report preview heading so the user knows the sections below
    # are the exact content that will be written into the exported report.
    st.header("Report Preview")

    # ---------------------------------------------------------------------
    # REPORT OVERVIEW
    # ---------------------------------------------------------------------
    # Show the report title and generation time (taken from datetime.now()
    # when the content was built) followed by the executive summary, which is
    # assembled only from findings computed on the user's DataFrame.
    st.subheader("Report Overview")
    st.markdown(f"**{_md(content['title'])}**")
    st.caption(f"Generated on {content['generated_at_str']}")
    for line in content["executive_summary"]:
        st.markdown(f"- {_md(line)}")

    # ---------------------------------------------------------------------
    # DATASET SUMMARY
    # ---------------------------------------------------------------------
    # Present the size and structure of the analysed DataFrame (rows, columns,
    # missing values, duplicates) as headline numbers plus a detail table.
    st.subheader("Dataset Summary")
    metric_cols = st.columns(4)
    metric_cols[0].metric("Rows", f"{summary['rows']:,}")
    metric_cols[1].metric("Columns", f"{summary['columns']:,}")
    metric_cols[2].metric("Missing Values", f"{summary['missing_total']:,}")
    metric_cols[3].metric("Duplicate Rows", "N/A" if summary["duplicate_rows"] is None else f"{summary['duplicate_rows']:,}")
    st.dataframe(summary["table"], use_container_width=True, hide_index=True)

    # ---------------------------------------------------------------------
    # DATA QUALITY SUMMARY
    # ---------------------------------------------------------------------
    # Summarise the data-quality checks. The score is shown only when the Data
    # Quality Report supplied one; otherwise the user is told it is missing.
    st.subheader("Data Quality Summary")
    if quality["score"]:
        st.metric("Overall Quality Score (from Data Quality Report)", quality["score"])
    else:
        # Make clear that no score is displayed because none was provided,
        # rather than inventing a number.
        st.info("No overall quality score was supplied, so none is shown.")
    _show_table(quality["findings_table"], "No data-quality findings are available.")

    # Expandable detail sections, one per quality dimension, so the preview
    # stays compact but every table that goes into the report is reviewable.
    with st.expander("Missing Values by Column"):
        _show_table(quality["missing_table"], "No missing values were found.")
    with st.expander("Duplicate Rows"):
        if quality["duplicate_rows"] is None:
            st.warning("Duplicate rows could not be counted for this dataset.")
        else:
            st.write(f"{quality['duplicate_rows']:,} duplicate row(s) ({quality['duplicate_pct']}% of rows). "
                     "Duplicates are reported only; nothing is deleted.")
    with st.expander("Potential Outliers (IQR method)"):
        st.caption("These are statistical flags and are not automatically incorrect values.")
        _show_table(quality["outlier_table"], "No potential outliers were found.")
    with st.expander("Invalid Values"):
        _show_table(quality["invalid_table"], "No invalid values were found by the checks performed.")
    with st.expander("Data Types"):
        _show_table(quality["dtype_table"], "No column information is available.")
    with st.expander("Consistency Findings"):
        _show_table(quality["consistency_table"], "No categorical inconsistencies were detected.")

    # ---------------------------------------------------------------------
    # ANALYTICS SUMMARY
    # ---------------------------------------------------------------------
    # Summarise the analytics (descriptive statistics, categories,
    # correlations, dates and charts) computed from the DataFrame.
    st.subheader("Analytics Summary")
    st.markdown("**Numerical Summary**")
    _show_table(analytics["numeric_stats"], "The dataset has no numerical columns to summarise.")
    st.markdown("**Categorical Summary**")
    _show_table(analytics["categorical_summary"], "The dataset has no categorical columns to summarise.")
    st.markdown("**Correlation Analysis**")
    _show_table(analytics["correlation_pairs"], "At least two numerical columns are needed for correlations.")
    if len(analytics["datetime_summary"]):
        st.markdown("**Date/Time Columns**")
        st.dataframe(analytics["datetime_summary"], use_container_width=True, hide_index=True)
    if analytics["supplied_findings"]:
        st.markdown("**Findings from the Analytics Section**")
        for text in analytics["supplied_findings"]:
            st.markdown(f"- {_md(text)}")

    # Show each chart that will be placed in the PDF. Only charts that were
    # actually rendered to an image are listed.
    st.markdown("**Charts**")
    if content["charts"]:
        for chart in content["charts"]:
            st.image(chart["png"], caption=chart["title"])
    else:
        # Explain that no chart images exist for this report, so the user
        # does not expect charts in the PDF export.
        st.info("No charts are included in this report.")

    # ---------------------------------------------------------------------
    # RECOMMENDATIONS
    # ---------------------------------------------------------------------
    # List the recommendations supplied by Part 4 (priority, category, issue,
    # finding, recommendation, suggested action) exactly as they will be exported.
    st.subheader("Recommendations")
    if recs["available"]:
        for _, rec in recs["table"].iterrows():
            st.markdown(
                f"**{_md(rec['Priority'])} Priority** - Category: {_md(rec['Category'])}  \n"
                f"Issue: {_md(rec['Issue'])}  \n"
                f"Finding: {_md(rec['Finding'])}  \n"
                f"Recommendation: {_md(rec['Recommendation'])}  \n"
                f"Suggested action: {_md(rec['Suggested Action'])}")
    else:
        # Inform the user that recommendation data was not available so the
        # report can clearly identify the missing section.
        st.info("Recommendation results are not currently available.")

    # ---------------------------------------------------------------------
    # RECOMMENDED ACTIONS
    # ---------------------------------------------------------------------
    # Show the next steps derived from the actual findings above.
    st.subheader("Recommended Actions")
    for number, action in enumerate(content["actions"], start=1):
        st.markdown(f"{number}. {_md(action)}")

    # Show notes that state which information was unavailable.
    if content["notes"]:
        with st.expander("Report Notes"):
            for note in content["notes"]:
                st.markdown(f"- {_md(note)}")


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT FOR THE PLATFORM
# ---------------------------------------------------------------------------
def render_exportable_report(df, quality_results=None, analytics_results=None, recommendations=None):
    """Draw the complete Exportable Report page. Call this from the main app."""

    # ---------------------------------------------------------------------
    # PAGE TITLE
    # ---------------------------------------------------------------------
    # Display the main report title so the user knows they are viewing the
    # Exportable Report section of the platform.
    st.title("Exportable Report")

    # ---------------------------------------------------------------------
    # PRIVACY NOTICE
    # ---------------------------------------------------------------------
    # Reassure the user that the report is built locally from summaries and
    # findings; no dataset rows are sent anywhere.
    st.caption("Reports are generated locally on this computer. They contain summaries and findings, not dataset rows.")

    # ---------------------------------------------------------------------
    # EMPTY / INVALID DATA GUARD
    # ---------------------------------------------------------------------
    # Stop early when there is no usable DataFrame (nothing uploaded, or an
    # empty file) so the rest of the page never runs on missing data.
    if not isinstance(df, pd.DataFrame) or df.shape[0] == 0 or df.shape[1] == 0:
        st.warning("No report data is available.")
        st.info("Upload a dataset (with at least one row and one column) in the Data Upload section first.")
        return

    # ---------------------------------------------------------------------
    # REPORT SETTINGS
    # ---------------------------------------------------------------------
    st.header("Report Settings")

    # Allow the user to provide a custom title for the generated report so the
    # exported document can be identified easily. The text is printed on the
    # preview, the PDF and the Excel Summary sheet.
    report_title = st.text_input("Report Title", value=DEFAULT_REPORT_TITLE)

    # Allow the user to select which report format they want. The value picks
    # which Python function runs: ReportLab (PDF), OpenPyXL (Excel) or Pandas (CSV).
    report_format = st.selectbox("Select Report Format", REPORT_FORMATS)

    # CSV holds one table per file, so ask which results table to export.
    # Only shown for CSV because the other formats include every section.
    csv_table = CSV_TABLE_OPTIONS[0]
    if report_format == "CSV":
        csv_table = st.selectbox("Select CSV Table", CSV_TABLE_OPTIONS)

    # Charts are drawn as images and embedded in the PDF; the option controls
    # whether the preview and PDF include them.
    include_charts = st.checkbox("Include charts", value=True)

    # Privacy control: category names come straight from the dataset, so the
    # most frequent value of each categorical column is excluded unless chosen.
    include_category_values = st.checkbox("Include most frequent category values (shows data values)", value=False)

    # ---------------------------------------------------------------------
    # BUILD THE REPORT CONTENT
    # ---------------------------------------------------------------------
    # Build the report from the DataFrame and the results of the earlier parts.
    # The DataFrame is only read, never changed. Errors are shown, not hidden.
    try:
        # Tell the user the analysis is running; large datasets can take a moment.
        with st.spinner("Preparing report preview..."):
            content = generate_report(
                df, quality_results, analytics_results, recommendations,
                report_title=report_title,
                include_charts=include_charts,
                include_category_values=include_category_values,
            )
    except Exception as exc:  # noqa: BLE001 - reported to the user
        # Explain that the content could not be built and include the reason.
        st.error(f"The report could not be prepared: {type(exc).__name__}: {exc}")
        return

    # Show the preview so the user can review everything before exporting.
    render_report_preview(content)

    # ---------------------------------------------------------------------
    # REPORT GENERATION
    # ---------------------------------------------------------------------
    # A lightweight fingerprint of the current settings and dataset. It lets the
    # page detect that a previously generated file is out of date.
    fingerprint = (report_format, csv_table, content["title"], include_charts, include_category_values,
                   df.shape, tuple(map(str, df.columns)), content["summary"]["missing_total"])

    st.header("Generate and Download")

    # Provide a button that allows the user to explicitly generate the selected
    # report after reviewing the findings. It uses the content built above (the
    # same data shown in the preview) and writes a file in memory in the chosen
    # format. When clicked, the file is stored so the download button appears.
    if st.button("Generate Report", type="primary"):
        try:
            # Tell the user the file is being created.
            with st.spinner(f"Generating {report_format} report..."):
                data, file_name, mime = export_report(report_format, content, csv_table)
            st.session_state["part5_report"] = {"data": data, "file_name": file_name, "mime": mime, "fingerprint": fingerprint}
            # Confirm that a real file was created and is ready to download.
            st.success("Report generated successfully.")
        except ImportError as exc:
            # A required package (reportlab or openpyxl) is not installed.
            st.error(f"A required package is missing: {exc}")
        except ValueError as exc:
            # Invalid format or an empty table selected for CSV.
            st.error(f"The report could not be generated: {exc}")
        except Exception as exc:  # noqa: BLE001 - reported to the user
            # Any other PDF/Excel/CSV generation problem, shown with its type.
            st.error(f"Report generation failed ({type(exc).__name__}): {exc}")

    # ---------------------------------------------------------------------
    # DOWNLOAD REPORT
    # ---------------------------------------------------------------------
    saved = st.session_state.get("part5_report")
    if saved and saved["fingerprint"] == fingerprint:
        # Display the file size so the user knows the file contains data.
        st.caption(f"Ready: {saved['file_name']} ({len(saved['data']) / 1024:,.1f} KB)")
        # Provide a download button so the user can save the generated report
        # to their local computer. The bytes are the real file created above.
        st.download_button(
            label="Download Report",
            data=saved["data"],
            file_name=saved["file_name"],
            mime=saved["mime"],
        )
    elif saved:
        # The settings or dataset changed after the last export, so the stored
        # file would no longer match the preview.
        st.info("Settings or data changed since the last export. Click Generate Report to create an updated file.")
