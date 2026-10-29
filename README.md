<div align="center">

# Clarifile

### Data Quality & Analytics Platform

**Upload data. Understand its quality. Explore patterns. Get actionable recommendations. Export a report.**

[![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-Web%20App-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)

[Repository](https://github.com/Srinivasa1711/Data_Project) · [Quick Start](#quick-start) · [Features](#key-features) · [Project Structure](#project-structure)

</div>

---

## Overview

**Clarifile** is an interactive, Python-based data quality and exploratory analytics application built with **Streamlit**. It brings together file ingestion, data validation, quality assessment, visual analysis, rule-based recommendations, and downloadable reports in one workflow.

Instead of manually inspecting a dataset for missing entries, duplicate records, inconsistent values, or unusual distributions, users can upload a file and review structured findings through a browser-based interface. The reporting modules help turn those findings into shareable deliverables.

The platform is useful for exploratory data analysis, data cleaning preparation, and communicating data quality issues before downstream analytics or modeling. **The recommendations are derived from transparent rules and statistical checks, not from a generative AI model.**

## Key Features

| Capability | What it does |
| --- | --- |
| **File ingestion** | Accepts CSV, Excel (`.xlsx`, `.xls`), PDF, TXT, and JPG/JPEG files, with extension/content validation and error handling. |
| **Data profiling** | Summarizes dataset dimensions, column types, missing values, and other structural characteristics. |
| **Quality assessment** | Checks for missing/placeholder values, duplicates, inconsistent data types, potentially invalid values, and outliers where applicable. |
| **Analytics and charts** | Supports numerical summaries, distributions, categorical exploration, and correlation analysis for suitable tabular data. |
| **Prioritized recommendations** | Creates explainable, actionable findings labeled High, Medium, Low, or Informational. |
| **Exportable reports** | Builds data-driven reports and supports PDF, Excel, and CSV exports, including selectable report sections. |
| **Modular development** | Includes stand-alone demos for recommendations and report exports, plus a report-generation test script. |

> **Format note:** CSV and Excel datasets provide the most complete tabular analytics experience. PDF, TXT, and image uploads are supported by the ingestion layer, but the insights available depend on the data that can be extracted from each file.

## How It Works

```text
       Upload a supported file
                 |
                 v
       Validate and parse input
                 |
                 v
        Profile dataset structure
                 |
                 v
      Run data quality checks
                 |
                 v
      Explore analytics and charts
                 |
                 v
   Generate prioritized recommendations
                 |
                 v
        Export PDF / Excel / CSV
```

### Data quality checks

The analysis and reporting code includes checks such as:

- **Completeness:** null or blank values and missing-value percentages.
- **Uniqueness:** repeated rows and potentially problematic duplicate values.
- **Consistency:** mixed types, numeric-looking text, and inconsistent categorical values.
- **Validity:** placeholder values and context-sensitive invalid-value indicators.
- **Outliers:** potential anomalies in numeric fields, including interquartile-range (IQR) checks.
- **Relationships:** statistical summaries and correlations where sufficient numeric data is available.

Quality indicators and recommendation priorities are **heuristics**, not universal judgments about whether a dataset is correct. Review findings against the dataset's business or research context before modifying values.

## Tech Stack

| Layer | Technologies |
| --- | --- |
| Application and interface | Python, Streamlit |
| Data processing | pandas, NumPy |
| Visualizations | Plotly, Matplotlib |
| Spreadsheet support | openpyxl, xlrd |
| Document and image handling | pypdf, Pillow |
| Report generation | ReportLab, pandas, openpyxl |
| Source control | Git, GitHub |

All dependencies for the documented main workflow are listed in [`requirements.txt`](requirements.txt).

## Quick Start

### Prerequisites

- A working **Python 3** installation (Python 3.11 is a reasonable local development choice).
- **pip** and **Git**.
- A terminal such as PowerShell, Command Prompt, macOS Terminal, or a Linux shell.

### 1. Clone the repository

```bash
git clone https://github.com/Srinivasa1711/Data_Project.git
cd Data_Project
```

### 2. Create and activate a virtual environment

Create the environment:

```bash
python -m venv .venv
```

Activate it on **Windows (PowerShell)**:

```powershell
.venv\Scripts\Activate.ps1
```

Or on **macOS/Linux**:

```bash
source .venv/bin/activate
```

### 3. Install dependencies

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 4. Start Clarifile

```bash
python -m streamlit run Data-Software-Merged1.2.3.py
```

Streamlit normally opens the app in your default browser at **http://localhost:8501**. If the browser does not open automatically, use the local URL displayed in the terminal.

**Windows shortcut:** `start_app.bat` installs the requirements and launches the same Streamlit entry point. It can be run from the repository folder.

> Run the application with `streamlit run` (or `python -m streamlit run`), rather than executing the Streamlit script directly with the editor's normal Python Run button.

## Using the Application

1. **Upload:** Choose a supported file and inspect its preview or extracted information.
2. **Review quality:** Examine available checks for missing data, duplicates, types, invalid entries, and potential outliers.
3. **Analyze:** Explore the relevant numerical or categorical summaries and visualizations.
4. **Act on findings:** Review prioritized recommendations and suggested next steps.
5. **Export:** Produce an available PDF, Excel, or CSV report for documentation or sharing.

The precise sections shown can vary with file type and the data available to analyze.

## Project Structure

```text
Data_Project/
├── Data-Software-Merged1.2.3.py   # Main Streamlit application entry point
├── recommendations_engine.py      # Rule-based recommendations and priority labels
├── exportable_report.py           # Report-generation Streamlit interface
├── report_builder.py              # Report content and PDF/Excel/CSV exporters
├── platform_helpers.py            # Shared data and column utilities
├── main_app.py                    # Alternate app implementation (includes image sorting)
├── demo_recommendations.py        # Stand-alone recommendations demo
├── demo_app.py                    # Stand-alone report export demo
├── test_part5.py                  # Tests for report generation and exports
├── requirements.txt               # Python dependencies
├── start_app.bat                  # Windows launcher for the main app
├── config.toml                    # Streamlit server configuration (see note below)
├── .gitignore                     # Local environment/cache exclusions
└── README.md                      # Project documentation
```

### Important modules

**`recommendations_engine.py`** evaluates the dataset and assembles explainable findings, each with a priority, issue, recommendation, suggested action, affected columns, and rule information.

**`report_builder.py`** creates report contents from the dataset and available analysis results. It provides exporters for PDF, Excel, and CSV, and includes safeguards such as preventing spreadsheet formula injection in exported text values.

**`exportable_report.py`** provides the user-facing reporting interface, while **`platform_helpers.py`** centralizes commonly used dataframe utilities.

**`main_app.py`** is a separate implementation that includes image sorting with an optional TensorFlow-based classifier. TensorFlow is **not** included in the current `requirements.txt`; use the documented merged application as the primary launcher unless you specifically intend to develop that alternative feature.

## Run Individual Modules

To try the recommendation engine independently:

```bash
python -m streamlit run demo_recommendations.py
```

This demonstration accepts **CSV or XLSX** uploads.

To try the report export interface independently:

```bash
python -m streamlit run demo_app.py
```

This demonstration accepts **CSV** uploads. When prior data-quality or analytics results have not been supplied, the report builder does not fabricate them.

## Testing

Run the included report/export test script:

```bash
python test_part5.py
```

`test_part5.py` exercises report construction and export paths with synthetic data, including assertions for quality findings, recommendations, charts, and document generation. It does not constitute an end-to-end test of the entire Streamlit interface.

## Configuration and Upload Limits

The repository includes a `config.toml` file with a Streamlit server setting intended to permit uploads up to **1 GB**. **Streamlit normally reads a project configuration file from `.streamlit/config.toml`, not from the repository root.** If you want that setting applied, place the file at:

```text
Data_Project/.streamlit/config.toml
```

with content such as:

```toml
[server]
maxUploadSize = 1024
```

**Important:** the upload validation code in `Data-Software-Merged1.2.3.py` also defines `MAX_FILE_SIZE_MB = 200`. Increasing the Streamlit server limit alone will **not** bypass the app's separate 200 MB check. If you intentionally support larger files, review both settings and the available system memory.

## Data Privacy and Limitations

- The application runs in the Streamlit process you start; uploaded data is handled by that running instance. **Do not assume that a hosted deployment provides the same privacy or access controls as a local run.**
- The report builder primarily exports summaries and analysis findings rather than raw dataset rows. Some values may be included if the relevant report option is enabled.
- Detection rules can flag legitimate unusual values as potential issues. A reported outlier is not necessarily an error.
- This repository is a development project, not a certified data-validation system. Test it with representative, non-sensitive data before operational use.
- Export capabilities and analysis depth depend on which modules and dataset fields are available for the selected workflow.

## Potential Enhancements

The following are **ideas for future development**, not claims about the current implementation:

- A single unified navigation and state-management layer across all five modules.
- More automated end-to-end tests and CI checks.
- Customizable quality-rule thresholds and user-defined validation rules.
- Sample datasets and screenshots showing the full workflow.
- Deployment documentation and authentication for hosted use.

## Author

**Srinivasa Chowdary Eluri**  
GitHub: [@Srinivasa1711](https://github.com/Srinivasa1711)

Contributions, bug reports, and improvement ideas are welcome through the repository's [Issues](https://github.com/Srinivasa1711/Data_Project/issues) section.

---

<div align="center">

*Clarifile — making data quality easier to inspect, explain, and communicate.*

</div>
