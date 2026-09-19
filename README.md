# Data_Project

# Data Quality & Analytics Platform

A tool that lets someone upload a dataset or a batch of images and instantly
get a useful overview of it — no coding required. This is the MVP (Minimum
Viable Product) stage: the first two working features out of a planned
five-feature platform.

## What it does

### 1. Data Upload
Upload a CSV, Excel, PDF, TXT, or image file and instantly see:
- Row and column counts
- A preview of the first 50 rows
- Data type and missing-value count per column
- Summary statistics (min, max, average, etc.)

### 2. Image Sorting by Category
Upload multiple images at once (e.g., an entire folder from a hard drive)
and the app automatically groups them by what's actually shown in each
photo — using a pretrained image classification model (MobileNetV2),
not the filename. Tested successfully sorting cats, dogs, and other
animals into specific categories automatically.

## Tech Stack

- **Python** — core language
- **Streamlit** — builds the web interface
- **Pandas** — reads and analyzes tabular data
- **TensorFlow / Keras (MobileNetV2)** — pretrained image classification model
- **Pillow (PIL)** — image handling
- **pypdf** — extracts text from PDF uploads

## Project Files

| File | Purpose |
|---|---|
| `main_app.py` | The combined app — both features, with a sidebar to switch between them |
| `start_app.bat` | Double-click to auto-install packages and launch the app |

## How to Run It

### First-time setup

TensorFlow can be picky about Python versions on Windows. If you're on
Python 3.12, you may hit a `DLL load failed` error when using Image
Sorting. The fix that worked for this project: use **Python 3.11**
in a dedicated virtual environment.

1. Install Python 3.11.9 specifically (the last version of 3.11 with a
   Windows installer): `https://www.python.org/downloads/release/python-3119/`
   → download the **Windows installer (64-bit)**.
   During install, check **"Add Python to environment variables."**

2. Open a terminal in the project folder and create a virtual environment:
   ```
   py -3.11 -m venv venv311
   ```

3. Activate it:
   ```
   venv311\Scripts\activate
   ```
   *(If PowerShell blocks this with a "running scripts is disabled" error,
   run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` first,
   then try activating again.)*

4. Install all required packages:
   ```
   pip install streamlit pandas openpyxl Pillow pypdf tensorflow numpy
   ```

### Running the app

With the virtual environment activated:
```
streamlit run main_app.py
```

This opens the app automatically in your browser at `http://localhost:8501`.

**Important:** Always run this with the `streamlit run` command in the
terminal — never with VS Code's green "Run" button. The Run button
executes the file as a plain Python script, which will not work for a
Streamlit app and produces a `missing ScriptRunContext` warning or a
"file does not exist" error if run from the wrong folder.

## Known Limitations (Current MVP Stage)

- **Runs locally only** — currently only works on the developer's own
  computer. Not yet deployed as a public website (a future step would be
  hosting it via Streamlit Community Cloud so others can use it without
  installing anything).
- **Image classification uses general-purpose categories** — the model
  was trained on everyday objects (ImageNet), so it's accurate on common
  animals/objects (e.g., correctly identified a Golden Retriever and a
  Tabby cat) but can misclassify visually similar but less common items
  (e.g., mistook a rat for a mink in testing).
- **No data quality checks yet** — the next planned feature (Data Quality
  Report) will flag missing values, duplicates, and outliers with an
  overall quality score. Not yet built.
- **No analytics/charts yet** — statistical summaries, distributions, and
  visualizations are a separate planned feature, not yet built.
- **No permanent storage** — uploaded files are not saved; each session
  starts fresh. This is intentional at this stage to avoid handling
  sensitive data before proper security/storage is in place.

## Planned Next Steps

1. Data Quality Report (missing values, duplicates, outliers, quality score)
2. Analytics & Charts (statistical summaries, distributions, correlations)
3. Recommendations (explain issues found, suggest fixes)
4. Exportable Report (downloadable data-quality/analytics report)
5. Deployment (make it accessible online, not just on local machines)

## Troubleshooting Log (for reference)

Issues hit and resolved while building this MVP:

- `ModuleNotFoundError: No module named 'streamlit'` → package wasn't
  installed; fixed with `pip install streamlit`.
- `Error: Invalid value: File does not exist` → terminal was in the wrong
  folder; fixed by navigating to the correct project directory first.
- `missing ScriptRunContext` warning → caused by using VS Code's Run
  button instead of the `streamlit run` command.
- `ImportError: DLL load failed` (TensorFlow) → caused by a Python 3.12 +
  TensorFlow compatibility issue on Windows; fixed by creating a Python
  3.11 virtual environment specifically for this project.
- PowerShell blocked `venv311\Scripts\activate` → fixed with
  `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`.
