@echo off
REM ------------------------------------------------------------------
REM Double-click this file to launch the Data Quality & Analytics app.
REM It automatically:
REM   1. Goes into this file's own folder (wherever you put it)
REM   2. Installs every required package
REM   3. Runs the app, which opens your browser automatically
REM
REM NOTE: If Image Sorting shows a TensorFlow DLL error, that is a
REM Windows system issue, not something this script can fix. Install
REM the Microsoft Visual C++ Redistributable from:
REM https://aka.ms/vs/17/release/vc_redist.x64.exe
REM then restart your computer. Data Upload will work either way.
REM ------------------------------------------------------------------

cd /d "%~dp0"

echo Checking required packages...
pip install streamlit pandas openpyxl Pillow pypdf tensorflow numpy --quiet

echo Starting the app... your browser will open automatically.
streamlit run main_app.py

pause
