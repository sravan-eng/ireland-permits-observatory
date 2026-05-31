# 🇮🇪 Ireland Work Permits Observatory

Interactive data visualisation website covering **17 years of Irish employment permit data** (2009–2026), built directly from official DETE Excel files.

## 🔑 Key feature — real data pipeline

Data is extracted **directly from the official DETE xlsx files** using an automated Python script. No manual data entry.

```
data/raw/*.xlsx  →  scripts/extract-data.py  →  src/data/permits.js + companies.js
```

## 🚀 Quick start in VS Code

1. Install the **Live Server** extension
2. Right-click `index.html` → **Open with Live Server**

## 🔄 To update with new xlsx files

1. Drop new xlsx files into `data/raw/`
2. Run: `python3 scripts/extract-data.py`
3. Refresh browser — data updates automatically

## 📁 Project structure

```
ireland-permits/
├── index.html                    # Full website — all sections
├── src/
│   ├── styles.css                # Light theme design system
│   ├── main.js                   # Charts, search engine, interactions
│   └── data/
│       ├── permits.js            # AUTO-GENERATED — sectors, counties, nationalities
│       └── companies.js          # AUTO-GENERATED — 4,224 companies
├── scripts/
│   └── extract-data.py          # Data pipeline — reads xlsx → generates JS
├── data/
│   └── raw/                     # Place your xlsx files here
│       ├── employment-permits-by-sector-2026.xlsx
│       ├── employment-permits-by-county-2026.xlsx
│       ├── employment-permits-by-nationality-2026.xlsx
│       └── employment-permits-issued-to-companies-2026.xlsx
└── README.md
```

## 📦 Python dependencies (for data pipeline only)

```bash
pip install pandas openpyxl
python3 scripts/extract-data.py
```

## 🔗 Data source

Official Irish Government open data — PSI Licence:
👉 https://enterprise.gov.ie/en/what-we-do/workplace-and-skills/employment-permits/statistics/

## ©️ Copyright

© 2026 Sravan Bathini. All rights reserved.
Data published under the PSI Licence by the Department of Enterprise, Tourism & Employment, Ireland.
This is an independent project — not affiliated with or endorsed by the Irish Government.
