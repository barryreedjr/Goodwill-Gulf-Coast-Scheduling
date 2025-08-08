
# Goodwill Scheduler (Prototype)

A simple web app to generate 4-week retail schedules from an Excel template.
Built with [Streamlit](https://streamlit.io/).

## Features
- Upload Excel input and download a 4-week schedule workbook
- Per-employee shift length (Open starts 8:00, Mid 10:00, Close ends 20:30)
- Weekend rotation (each employee gets one Sat+Sun off within 4 weeks)
- Calendar-style sheets per week (Employee x Sunday–Saturday)
- Color-coded cells by role and print-friendly (landscape, 1-page width, repeating header)

## Quick Start (Local)
```bash
pip install -r requirements.txt
streamlit run app_streamlit.py
```
Open the local URL, upload `schedule_input_template_v2.xlsx`, optionally set a Sunday start date (YYYY-MM-DD), and click **Generate**.

## Deploy on Streamlit Community Cloud
1. Push this repo to GitHub with these files at the root:
   - `app_streamlit.py`
   - `requirements.txt`
   - `schedule_input_template_v2.xlsx`
   - `README.md`
2. Go to Streamlit Community Cloud → **New app** → select your repo, branch, and `app_streamlit.py`.
3. Click **Deploy**. Share the app URL (you can restrict by email/domain).

## Customize
- Edit role colors in `app_streamlit.py` (search for `default_role_colors`).
- Adjust coverage counts in `coverage_template` tab of the Excel.
- Change shift anchors/times in `shift_anchors` tab.
- Business rules (min/max shift hours, gaps, max days/week) in `rules` tab.
