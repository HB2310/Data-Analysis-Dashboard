# Data Trend Analysis & Visualization Dashboard

A Python-based data dashboard that ingests CSV datasets, performs data cleaning and preprocessing, and generates interactive visualizations.

## Features

- **Data Ingestion** — Upload any CSV; auto-detects date columns
- **Trend Analysis** — Line charts with configurable rolling averages
- **Outlier Detection** — IQR method with visual flagging and report
- **Statistical Summary** — Mean, std, min, max, skew, kurtosis, outlier %
- **Correlation Heatmap** — Pearson correlation matrix
- **Categorical Breakdown** — Group by any string column
- **Export Formats:**
  - CSV (raw data)
  - Excel (.xlsx) — 3 sheets: Raw Data, Summary Stats, PowerBI_Ready
  - Power BI Package (.zip) — CSV + schema + DAX measures guide
  - PDF Report — Charts + stats table embedded

## Setup

```bash
# 1. Clone / copy the project folder
cd dashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
streamlit run app.py
```

The dashboard opens at http://localhost:8501

## Using Your Own Data

Upload any CSV with:
- A date/time column (named "date", "Date", "timestamp", etc.)
- One or more numeric columns
- Optional: string columns for categorical breakdowns

### Example CSV structure:
```
date,revenue,users,cost,category,region
2024-01-01,142.5,1200,65.3,Product A,North
2024-01-08,138.2,1340,63.1,Product B,South
...
```

## Power BI Integration

After exporting the Power BI Package (.zip):
1. Unzip it
2. Open Power BI Desktop
3. Home → Get Data → Text/CSV → select `PowerBI_Ready.csv`
4. Set column types (see README.txt inside the zip)
5. Copy the DAX measures from README.txt into your model
6. Build visuals

Alternatively use the Excel export → PowerBI_Ready sheet directly.

## Project Structure

```
dashboard/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

## Tech Stack

| Library      | Purpose                          |
|-------------|----------------------------------|
| Streamlit    | Web UI framework                 |
| Pandas       | Data cleaning & preprocessing    |
| NumPy        | Numerical operations             |
| Plotly       | Interactive visualizations       |
| OpenPyXL     | Excel file generation            |
| ReportLab    | PDF report generation            |
| Kaleido      | Chart → PNG for PDF embedding    |
