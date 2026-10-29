# Clarifile - Data Quality & Analytics Platform

Upload a data file, get a quality score, charts, plain-English recommendations and a downloadable report.

## Files (keep ALL of them in the same folder)
| File | What it is |
|---|---|
| Data-Software-Merged1.2.3.py | Main app: landing page, Data Upload, Data Quality Report, Analytics & Charts, Recommendations, Export Report |
| recommendations_engine.py | Part 4 - Recommendations |
| exportable_report.py | Part 5 - Exportable Report screen |
| report_builder.py | Part 5 - builds the report and exports PDF / Excel / CSV |
| platform_helpers.py | Small shared helpers |
| requirements.txt | Python packages needed |
| test_part5.py | Automated test for report_builder.py: `python test_part5.py` |
| demo_app.py | Run Part 5 on its own: `streamlit run demo_app.py` |
| demo_recommendations.py | Run Part 4 on its own: `streamlit run demo_recommendations.py` |
| .streamlit/config.toml | Raises the upload limit to 1 GB |
| .gitignore | Keeps virtual environments, caches and secrets out of GitHub |

## Run
    pip install -r requirements.txt
    streamlit run Data-Software-Merged1.2.3.py

Always start with `streamlit run` in a terminal (not VS Code's Run button).
Stop the old server with Ctrl+C before restarting after any code change.

## Put it on GitHub
    git init
    git add .
    git commit -m "Clarifile: first working version"
Check that `venv311/` does not appear in `git status` before you push.
