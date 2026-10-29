@echo off
REM Double-click to install what is needed and start Clarifile.
REM If you use a virtual environment (venv311), activate it first in a terminal instead.
cd /d "%~dp0"
echo Checking required packages...
python -m pip install -r requirements.txt --quiet
echo Starting Clarifile... your browser will open automatically.
python -m streamlit run Data-Software-Merged1.2.3.py
pause
