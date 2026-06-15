# Karad Division — Daily Defaulter Report

Streamlit app to generate the daily Delivery Productivity + DSS Usage Defaulter Report for Karad Division, India Post.

**Live app:** https://karad-defaulter-report.streamlit.app *(update after deployment)*

---

## How to use (daily workflow)

1. Open the app
2. Set dates:
   - **Cumulative From** — defaults to 1st of current month (change only if needed)
   - **Cumulative To** — defaults to yesterday
   - **Daily date** — defaults to yesterday
3. Drop the 4 CSV files in their respective slots (file names don't matter):
   - ① Cumulative Delivery Productivity CSV
   - ② Daily Delivery Productivity CSV
   - ③ Cumulative DSS Usage CSV
   - ④ Daily DSS Usage CSV
4. Click **Generate Defaulter Report**
5. Review the summary table → click **Download**

Report date is automatically set to `daily date + 1 day` (today).

---

## Files in this repo

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `report_builder.py` | All Excel generation logic |
| `Office_Master_File.xlsx` | Office master (288 offices) — update here when structure changes |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme |

---

## Updating the master file

When office structure changes (new offices, discontinued offices):
1. Prepare the updated `Office_Master_File.xlsx`
2. Upload it to this repo (replace the existing file)
3. The app picks it up automatically on next run (cached for 1 hour)

OR use the **Upload new master file** button in the app sidebar — but note that this only applies for the current session. For it to persist, you still need to push to the repo.

---

## Deployment (Streamlit Cloud)

1. Fork / push this repo to `dokaradmmu/karad-defaulter-report`
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. New app → select this repo → `main` branch → `app.py`
4. Deploy

No secrets or environment variables needed.

---

## Permanently excluded offices

- Shenawadi B.O
- Yeralwadi B.O

These are hardcoded in `report_builder.py` and will never appear in any report.

---

## KPI threshold

**90.00%** — offices below this in the **cumulative** column are listed as defaulters.
