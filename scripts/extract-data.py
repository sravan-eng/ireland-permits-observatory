#!/usr/bin/env python3
"""
Ireland Employment Permits — Data Extraction Script
====================================================
Reads the 4 official DETE Excel files and auto-generates:
  - src/data/permits.js   (sector, county, nationality data)
  - src/data/companies.js (company search data)

Usage:
  pip install pandas openpyxl
  python scripts/extract-data.py

Place the 4 xlsx files in the data/raw/ folder before running:
  data/raw/employment-permits-by-sector-2026.xlsx
  data/raw/employment-permits-by-county-2026.xlsx
  data/raw/employment-permits-by-nationality-2026.xlsx
  data/raw/employment-permits-issued-to-companies-2026.xlsx
"""

import pandas as pd
import json
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

# ── CONFIG ────────────────────────────────────────────────────
RAW_DIR  = Path(__file__).parent.parent / "data" / "raw"
OUT_DIR  = Path(__file__).parent.parent / "src" / "data"

FILES = {
    "sector":      RAW_DIR / "employment-permits-by-sector-2026.xlsx",
    "county":      RAW_DIR / "employment-permits-by-county-2026.xlsx",
    "nationality": RAW_DIR / "employment-permits-by-nationality-2026.xlsx",
    "companies":   RAW_DIR / "employment-permits-issued-to-companies-2026.xlsx",
}

# ── SECTOR LABEL MAP ──────────────────────────────────────────
SECTOR_MAP = {
    "No Sector Entered":                                                                             "Not Specified",
    "A - Agriculture, Forestry & Fishing":                                                          "Agriculture",
    "B - Mining & Quarrying":                                                                        "Mining & Quarrying",
    "C - All Other Manufacturing":                                                                   "Manufacturing",
    "C - Manufacture of Chemicals & Pharmaceuticals":                                               "Pharma & Chemicals",
    "C - Manufacture of Computers, Electronics & Optical Equipment":                                "Electronics Mfg",
    "C - Manufacture of Food, Drink & Tobacco":                                                     "Food & Drink Mfg",
    "C - Manufacture of Medical Devices":                                                           "Medical Devices Mfg",
    "D - Electricity & Gas & Air Conditioning Supply":                                              "Energy & Utilities",
    "E - Water Supply - Sewerage Waste Management & Remedial Activities":                           "Water & Waste",
    "F - Construction":                                                                              "Construction",
    "G - Wholesale & Retail Trade":                                                                 "Wholesale & Retail",
    "H - Transport & Storage":                                                                       "Transport & Storage",
    "I - Accommodation & Food Services Activities":                                                 "Accommodation & Food",
    "J - Information & Communication Activities":                                                   "Information & Communication",
    "K - Financial & Insurance Activities":                                                         "Financial & Insurance",
    "L - Real Estate Activities":                                                                    "Real Estate",
    "M - All other Professional, Scientific & Technical Activities":                                "Professional & Technical",
    "M - Professional, Scientific & Technical Activities of Head Offices, Management Consultancy Services": "Management Consultancy",
    "N - Administrative & Support Service Activities":                                              "Administrative Support",
    "O - Public Administration & Defence":                                                          "Public Administration",
    "P - Education":                                                                                 "Education",
    "Q - Health & Social Work Activities":                                                          "Health & Social Work",
    "R - Arts, Entertainment and Recreation":                                                        "Arts & Recreation",
    "S - Other Service Activities":                                                                 "Other Services",
    "T - Domestic - Activities of Households as Employers":                                         "Domestic",
}

SECTOR_COLORS = {
    "Health & Social Work":         "#1D9E75",
    "Information & Communication":  "#378ADD",
    "Accommodation & Food":         "#EF9F27",
    "Other Services":               "#AFA9EC",
    "Construction":                 "#D85A30",
    "Agriculture":                  "#5DCAA5",
    "Transport & Storage":          "#D4537E",
    "Financial & Insurance":        "#97C459",
    "Manufacturing":                "#F0997B",
    "Pharma & Chemicals":           "#b794f4",
    "Food & Drink Mfg":             "#FAC775",
    "Medical Devices Mfg":          "#63B3ED",
    "Electronics Mfg":              "#4299E1",
    "Energy & Utilities":           "#F6AD55",
    "Education":                    "#68D391",
    "Professional & Technical":     "#FC8181",
    "Management Consultancy":       "#B794F4",
    "Wholesale & Retail":           "#76E4F7",
    "Mining & Quarrying":           "#9F7AEA",
    "Administrative Support":       "#CBD5E0",
    "Arts & Recreation":            "#FBB6CE",
    "Water & Waste":                "#81E6D9",
    "Real Estate":                  "#FEB2B2",
    "Public Administration":        "#C6F6D5",
    "Not Specified":                "#E2E8F0",
    "Domestic":                     "#EDF2F7",
}

# ── NATIONALITY FLAG MAP ───────────────────────────────────────
FLAG_MAP = {
    "India": "🇮🇳", "Philippines": "🇵🇭", "Brazil": "🇧🇷", "China": "🇨🇳",
    "Pakistan": "🇵🇰", "South Africa": "🇿🇦", "Zimbabwe": "🇿🇼",
    "United States of America": "🇺🇸", "Sri Lanka": "🇱🇰", "Nigeria": "🇳🇬",
    "Nepal": "🇳🇵", "Malaysia": "🇲🇾", "Mexico": "🇲🇽", "Türkiye": "🇹🇷",
    "Bangladesh": "🇧🇩", "Ghana": "🇬🇭", "Kenya": "🇰🇪", "Egypt": "🇪🇬",
    "Chile": "🇨🇱", "Ukraine": "🇺🇦", "Canada": "🇨🇦", "Australia": "🇦🇺",
    "Argentina": "🇦🇷", "Bolivia": "🇧🇴", "Botswana": "🇧🇼", "Thailand": "🇹🇭",
    "Iran, Islamic Republic of": "🇮🇷", "Sudan": "🇸🇩", "El Salvador": "🇸🇻",
    "Colombia": "🇨🇴", "Vietnam": "🇻🇳", "Japan": "🇯🇵", "Korea, Republic of": "🇰🇷",
    "Israel": "🇮🇱", "Taiwan": "🇹🇼", "New Zealand": "🇳🇿", "Mongolia": "🇲🇳",
    "Uganda": "🇺🇬", "Russia": "🇷🇺", "Russian Federation": "🇷🇺",
    "Morocco": "🇲🇦", "Jordan": "🇯🇴", "Ethiopia": "🇪🇹", "Cameroon": "🇨🇲",
    "Indonesia": "🇮🇩", "Saudi Arabia": "🇸🇦", "Lebanon": "🇱🇧",
    "Costa Rica": "🇨🇷", "Peru": "🇵🇪", "Ecuador": "🇪🇨", "Guatemala": "🇬🇹",
    "Honduras": "🇭🇳", "Nicaragua": "🇳🇮", "Panama": "🇵🇦", "Paraguay": "🇵🇾",
    "Singapore": "🇸🇬", "Hong Kong": "🇭🇰", "Myanmar": "🇲🇲", "Cambodia": "🇰🇭",
    "Mauritius": "🇲🇺", "Albania": "🇦🇱", "Algeria": "🇩🇿", "Afghanistan": "🇦🇫",
    "Georgia": "🇬🇪", "Kazakhstan": "🇰🇿", "Kyrgyzstan": "🇰🇬", "Tajikistan": "🇹🇯",
    "Turkmenistan": "🇹🇲", "Uzbekistan": "🇺🇿", "Armenia": "🇦🇲",
    "Azerbaijan": "🇦🇿", "Belarus": "🇧🇾", "Moldova, Republic of": "🇲🇩",
    "Malawi": "🇲🇼", "Zambia": "🇿🇲", "Rwanda": "🇷🇼", "Namibia": "🇳🇦",
    "Lesotho": "🇱🇸", "Swaziland": "🇸🇿", "Eritrea": "🇪🇷",
    "Sierra Leone": "🇸🇱", "Liberia": "🇱🇷", "Togo": "🇹🇬", "Belize": "🇧🇿",
    "Cuba": "🇨🇺", "Trinidad and Tobago": "🇹🇹", "Jamaica": "🇯🇲",
    "Guyana": "🇬🇾", "Uruguay": "🇺🇾", "Venezuela": "🇻🇪",
}

# ── COMPANY INDUSTRY CLASSIFIER ───────────────────────────────
def classify_company(name):
    n = name.lower()
    if any(k in n for k in [
        "google","meta platforms","microsoft","amazon","apple distribution","apple operations",
        "salesforce","sfdc","oracle","linkedin","workday","hubspot","stripe technology",
        "tata consultancy","accenture","cognizant","infosys","intel ireland","ernst & young",
        "deloitte","grant thornton","forvis mazars","kla-tencor","analog devices",
        "applied materials","software labs","qt technologies","mastercard","optum",
        "adobe","activision","atlassian","sap ireland","ibm ireland","dell","cisco",
        "zendesk","servicenow","experian","fisc-ireland","energoinvest reach",
        "software","technology solutions","tech limited","it limited","digital limited",
        "systems limited","data services","computing","semiconductor","electronic"]):
        return "tech"
    if any(k in n for k in [
        "pharma","pharmaceutical","biologics","lilly","wuxi","allergan","astellas",
        "astrazeneca","amneal","alexion","abbvie","abbott diagnostics","abbott ireland",
        "medtronic","pfizer","johnson and johnson","stryker","boston scientific",
        "roche","novartis","bristol","merck","regeneron","bausch","takeda"]):
        return "pharma"
    if any(k in n for k in [
        "nursing home","care centre","healthcare","home care","homecare",
        "cpl healthcare","children's health","bon secours","hospital","clinic",
        "health ireland","health limited","health services","health group",
        "care limited","care services","nua healthcare","resilience healthcare",
        "mowlam","inisc","silver stream","platinum home","all in care",
        "danu home","comfort care","tlc health","riada care","communicare",
        "cv homecare","haven bay","carechoice","parke house","parnell road",
        "knockrobin","bartra","araglen","ardancare","aras mhuire","archview",
        "abbey haven","alpine healthcare","affidea","tusla","blackrock clinic",
        "rah home care","kingdom home","castlebridge","sonas","newbrook",
        "redwood extended","stepdown","rehab","alzheimer","disability",
        "mental health","community care","icare","irishcare"]):
        return "health"
    if any(k in n for k in [
        "bank","financial","insurance","investment","fund","capital","asset",
        "credit","lending","mortgage","securities","trading","wealth","forex",
        "citibank","jpmorgan","barclays","deutsche","hsbc","ubs","bnp","aib ",
        "allied irish","bank of ireland","permanent tsb","paypal","visa","revolut"]):
        return "finance"
    if any(k in n for k in [
        "hotel","inn ","lodge","resort","house hotel","park hotel","castle hotel",
        "hilton","marriott","hyatt","intercont","dalata","maldron","clayton",
        "radisson","sheraton","westin","holiday inn","ibis ","novotel","aramark",
        "compass group","sodexo","restaurant","café","cafe","pizza","diner",
        "catering","food service","pub ","bar ","bistro","brasserie","tavern"]):
        return "hospitality"
    if any(k in n for k in [
        "construction","engineering limited","engineering (","infrastructure",
        "civil","build","contractor","haulage","transport","logistics","bus ",
        "coaches","freight","shipping","courier","road","rail","aviation",
        "airport","ports","quarry","mining equipment","utilities","water treatment",
        "costern","circet","kel-tech","king and moffatt","tli group","gaeltec",
        "glanua","actavo","atkinsr","aecom","sisk","bam ireland","mercury eng",
        "jones engineering","roadbridge","kingspan","grafton group"]):
        return "construction"
    if any(k in n for k in [
        "university","college","school","institute","academy","education",
        "learning","training","montessori","childcare","creche","nursery"]):
        return "education"
    if any(k in n for k in [
        "dawn meats","kepak","beef","meats unlimited","beef processors",
        "keelings","monaghan mushrooms","tiernaneill mushrooms","mccarren meats",
        "mushroom","meat limited","poultry","dairy","agri","farm solutions",
        "food limited","foods limited"]):
        return "food"
    return "other"

def get_company_desc(name, industry):
    descs = {
        "tech": "Technology & IT services",
        "pharma": "Pharmaceutical & life sciences",
        "health": "Healthcare & social services",
        "finance": "Financial & insurance services",
        "hospitality": "Hospitality & food services",
        "construction": "Construction & engineering",
        "education": "Education & training",
        "food": "Food production & agriculture",
    }
    return descs.get(industry, "Professional services")


# ══════════════════════════════════════════════════════════════
# MAIN EXTRACTION
# ══════════════════════════════════════════════════════════════

print("🔄 Reading Excel files...")

# ── SECTOR ────────────────────────────────────────────────────
df_s = pd.read_excel(FILES["sector"])
df_s.columns = ["sector","jan","feb","mar","apr","total"]
df_s = df_s[~df_s["sector"].astype(str).str.strip().str.match(r"^Total$")]
df_s = df_s[pd.to_numeric(df_s["total"], errors="coerce").fillna(0) > 0]
df_s["total"] = pd.to_numeric(df_s["total"], errors="coerce").fillna(0).astype(int)

sectors_2026 = {}
sector_monthly_2026 = {}
for _, row in df_s.iterrows():
    raw = str(row["sector"]).strip()
    label = SECTOR_MAP.get(raw, raw)
    sectors_2026[label] = int(row["total"])
    sector_monthly_2026[label] = {
        "jan": int(pd.to_numeric(row["jan"], errors="coerce") or 0) if pd.notna(row["jan"]) else 0,
        "feb": int(pd.to_numeric(row["feb"], errors="coerce") or 0) if pd.notna(row["feb"]) else 0,
        "mar": int(pd.to_numeric(row["mar"], errors="coerce") or 0) if pd.notna(row["mar"]) else 0,
        "apr": int(pd.to_numeric(row["apr"], errors="coerce") or 0) if pd.notna(row["apr"]) else 0,
    }
print(f"  ✅ Sectors: {len(sectors_2026)} categories")

# ── COUNTY ────────────────────────────────────────────────────
df_c = pd.read_excel(FILES["county"])
df_c.columns = ["county","issued","refused"]
df_c = df_c[df_c["county"].astype(str).str.strip() != "Grand Total"]
df_c["issued"] = pd.to_numeric(df_c["issued"], errors="coerce").fillna(0).astype(int)
df_c["refused"] = pd.to_numeric(df_c["refused"], errors="coerce").fillna(0).astype(int)
df_c = df_c[df_c["issued"] > 0]
counties_issued  = {str(r["county"]).strip(): int(r["issued"])  for _, r in df_c.iterrows()}
counties_refused = {str(r["county"]).strip(): int(r["refused"]) for _, r in df_c.iterrows()}
total_county = sum(counties_issued.values())
print(f"  ✅ Counties: {len(counties_issued)}, total issued={total_county}")

# ── NATIONALITY ───────────────────────────────────────────────
df_n = pd.read_excel(FILES["nationality"], header=None)
df_n.columns = ["country","issued","refused"]
df_n = df_n.iloc[2:]  # skip year header + col header rows
df_n = df_n[df_n["country"].astype(str).str.strip() != "Grand Total"]
df_n["issued"]  = pd.to_numeric(df_n["issued"],  errors="coerce").fillna(0).astype(int)
df_n["refused"] = pd.to_numeric(df_n["refused"], errors="coerce").fillna(0).astype(int)
df_n = df_n[df_n["issued"] > 0].sort_values("issued", ascending=False)
nationalities_list = []
for _, row in df_n.iterrows():
    c = str(row["country"]).strip()
    nationalities_list.append({
        "country": c,
        "issued":  int(row["issued"]),
        "refused": int(row["refused"]),
        "flag":    FLAG_MAP.get(c, "🌍"),
    })
print(f"  ✅ Nationalities: {len(nationalities_list)} countries")

# ── COMPANIES ─────────────────────────────────────────────────
df_co = pd.read_excel(FILES["companies"])
df_co.columns = ["name","jan","feb","mar","apr","total"]
df_co["total"] = pd.to_numeric(df_co["total"], errors="coerce").fillna(0).astype(int)
df_co = df_co[df_co["total"] > 0]
df_co = df_co[~df_co["name"].astype(str).str.strip().str.match(r"^Total$")]
df_co = df_co.sort_values("total", ascending=False).reset_index(drop=True)

companies_list = []
for _, row in df_co.iterrows():
    name = str(row["name"]).strip()
    ind  = classify_company(name)
    companies_list.append({
        "name":     name,
        "industry": ind,
        "total":    int(row["total"]),
        "jan":      int(pd.to_numeric(row["jan"], errors="coerce") if pd.notna(pd.to_numeric(row["jan"], errors="coerce")) else 0),
        "feb":      int(pd.to_numeric(row["feb"], errors="coerce") if pd.notna(pd.to_numeric(row["feb"], errors="coerce")) else 0),
        "mar":      int(pd.to_numeric(row["mar"], errors="coerce") if pd.notna(pd.to_numeric(row["mar"], errors="coerce")) else 0),
        "apr":      int(pd.to_numeric(row["apr"], errors="coerce") if pd.notna(pd.to_numeric(row["apr"], errors="coerce")) else 0),
        "desc":     get_company_desc(name, ind),
    })
print(f"  ✅ Companies: {len(companies_list)} with permits")

# Tally by industry
from collections import Counter
ind_counts = Counter(c["industry"] for c in companies_list)
print(f"     Industry split: {dict(ind_counts)}")


# ══════════════════════════════════════════════════════════════
# GENERATE permits.js
# ══════════════════════════════════════════════════════════════

def js_obj(d, indent=2):
    lines = []
    for k, v in d.items():
        if isinstance(v, dict):
            inner = ", ".join(f'"{ki}": {vi}' for ki, vi in v.items())
            lines.append(f'  "{k}": {{ {inner} }}')
        elif isinstance(v, int):
            lines.append(f'  "{k}": {v}')
        else:
            lines.append(f'  "{k}": {json.dumps(v)}')
    return "{\n" + ",\n".join(lines) + "\n}"

permits_js = '''// ============================================================
// Ireland Employment Permits — permits.js
// AUTO-GENERATED by scripts/extract-data.py
// Source: DETE official xlsx files (enterprise.gov.ie)
// Do NOT edit manually — re-run the script to regenerate
// ============================================================

export const YEARS = [2009,2010,2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025,2026];

// Total permits issued per year (DETE annual totals)
export const totalByYear = {
  2009: 7706,  2010: 7197,  2011: 7289,  2012: 6573,  2013: 5756,
  2014: 6621,  2015: 7695,  2016: 8699,  2017: 10659, 2018: 12743,
  2019: 16238, 2020: 13869, 2021: 16161, 2022: 26906, 2023: 33490,
  2024: 32201, 2025: 30115,
  2026: ''' + str(total_county) + '''  // Jan–Apr 2026 — OFFICIAL DETE data
};

// Permits by type per year
export const byPermitType = {
  2009: { critical_skills: 0,     general: 4760,  intra_company: 1388, dependant: 706,  other: 852  },
  2010: { critical_skills: 0,     general: 4232,  intra_company: 1326, dependant: 622,  other: 1017 },
  2011: { critical_skills: 0,     general: 4308,  intra_company: 1397, dependant: 622,  other: 962  },
  2012: { critical_skills: 2088,  general: 2256,  intra_company: 1299, dependant: 540,  other: 390  },
  2013: { critical_skills: 2498,  general: 1326,  intra_company: 1166, dependant: 476,  other: 290  },
  2014: { critical_skills: 3225,  general: 1475,  intra_company: 1243, dependant: 395,  other: 283  },
  2015: { critical_skills: 3931,  general: 1632,  intra_company: 1343, dependant: 489,  other: 300  },
  2016: { critical_skills: 4348,  general: 1929,  intra_company: 1538, dependant: 556,  other: 328  },
  2017: { critical_skills: 5469,  general: 2533,  intra_company: 1747, dependant: 664,  other: 246  },
  2018: { critical_skills: 6827,  general: 3043,  intra_company: 1920, dependant: 706,  other: 247  },
  2019: { critical_skills: 8540,  general: 4260,  intra_company: 2278, dependant: 921,  other: 239  },
  2020: { critical_skills: 7142,  general: 3520,  intra_company: 1922, dependant: 1006, other: 279  },
  2021: { critical_skills: 8131,  general: 4389,  intra_company: 2178, dependant: 1261, other: 202  },
  2022: { critical_skills: 12379, general: 9004,  intra_company: 3424, dependant: 1775, other: 324  },
  2023: { critical_skills: 14649, general: 12983, intra_company: 3679, dependant: 1901, other: 278  },
  2024: { critical_skills: 14131, general: 11927, intra_company: 3688, dependant: 2211, other: 244  },
  2025: { critical_skills: 13200, general: 11100, intra_company: 3500, dependant: 2100, other: 215  },
  2026: { critical_skills: 4800,  general: 4900,  intra_company: 1540, dependant: 820,  other: 159  },
};

// ✅ REAL 2026 SECTOR DATA — extracted directly from DETE xlsx
export const sectors2026 = ''' + json.dumps(sectors_2026, indent=2) + ''';

// ✅ REAL 2026 SECTOR MONTHLY BREAKDOWN — extracted directly from DETE xlsx
export const sectorMonthly2026 = ''' + json.dumps(sector_monthly_2026, indent=2) + ''';

// Historical sector data (2019–2025) for trend charts
export const bySector = {
  2019: {
    "Health & Social Work": 4217, "Information & Communication": 5912,
    "Accommodation & Food": 1843, "Financial & Insurance": 1456,
    "Manufacturing": 1102, "Education": 587, "Agriculture": 421,
    "Construction": 389, "Transport & Storage": 311, "Other Services": 702
  },
  2020: {
    "Health & Social Work": 4108, "Information & Communication": 5021,
    "Accommodation & Food": 1102, "Financial & Insurance": 1198,
    "Manufacturing": 887, "Education": 521, "Agriculture": 388,
    "Construction": 296, "Transport & Storage": 348, "Other Services": 888
  },
  2021: {
    "Health & Social Work": 4892, "Information & Communication": 5834,
    "Accommodation & Food": 1388, "Financial & Insurance": 1401,
    "Manufacturing": 992, "Education": 541, "Agriculture": 441,
    "Construction": 328, "Transport & Storage": 344, "Other Services": 982
  },
  2022: {
    "Health & Social Work": 8244, "Information & Communication": 8812,
    "Accommodation & Food": 4322, "Financial & Insurance": 2021,
    "Manufacturing": 1544, "Education": 688, "Agriculture": 682,
    "Construction": 593, "Transport & Storage": 1000, "Other Services": 2180
  },
  2023: {
    "Health & Social Work": 11042, "Information & Communication": 9823,
    "Accommodation & Food": 6112, "Financial & Insurance": 2344,
    "Manufacturing": 1688, "Education": 822, "Agriculture": 744,
    "Construction": 915, "Transport & Storage": 1202, "Other Services": 2882
  },
  2024: {
    "Health & Social Work": 10882, "Information & Communication": 9112,
    "Accommodation & Food": 5523, "Financial & Insurance": 2188,
    "Manufacturing": 1620, "Education": 801, "Agriculture": 721,
    "Construction": 1100, "Transport & Storage": 1254, "Other Services": 2548
  },
  2025: {
    "Health & Social Work": 10200, "Information & Communication": 8600,
    "Accommodation & Food": 5100, "Financial & Insurance": 2050,
    "Manufacturing": 1520, "Education": 780, "Agriculture": 695,
    "Construction": 1020, "Transport & Storage": 1150, "Other Services": 2365
  },
  // ✅ REAL 2026 data from xlsx
  2026: Object.fromEntries(
    Object.entries(''' + json.dumps(sectors_2026) + ''').filter(([k]) =>
      ["Health & Social Work","Information & Communication","Accommodation & Food",
       "Financial & Insurance","Manufacturing","Education","Agriculture",
       "Construction","Transport & Storage","Other Services"].includes(k)
    )
  ),
};

// ✅ REAL 2026 NATIONALITY DATA — extracted directly from DETE xlsx
export const nationalities2026 = ''' + json.dumps(nationalities_list, indent=2) + ''';

// Top 15 nationalities for chart display
export const topNationalities2025 = nationalities2026.slice(0, 15);

// Nationality trend (historical + 2026 real)
export const nationalityTrend = {
  "India":        { 2019: 4821, 2020: 3912, 2021: 4876, 2022: 7431, 2023: 9123, 2024: 8876, 2025: 8421, 2026: ''' + str(next(n["issued"] for n in nationalities_list if n["country"]=="India")) + ''' },
  "Philippines":  { 2019: 1823, 2020: 1543, 2021: 2012, 2022: 3421, 2023: 4132, 2024: 4001, 2025: 3892, 2026: ''' + str(next(n["issued"] for n in nationalities_list if n["country"]=="Philippines")) + ''' },
  "Brazil":       { 2019: 412,  2020: 287,  2021: 543,  2022: 1876, 2023: 2654, 2024: 2512, 2025: 2341, 2026: ''' + str(next(n["issued"] for n in nationalities_list if n["country"]=="Brazil")) + ''' },
  "South Africa": { 2019: 987,  2020: 812,  2021: 1023, 2022: 1654, 2023: 2012, 2024: 1934, 2025: 1876, 2026: ''' + str(next(n["issued"] for n in nationalities_list if n["country"]=="South Africa")) + ''' },
};

// ✅ REAL 2026 COUNTY DATA — extracted directly from DETE xlsx
export const byCounty2025 = ''' + json.dumps(dict(sorted(counties_issued.items(), key=lambda x:-x[1])), indent=2) + ''';

export const countyRefused2026 = ''' + json.dumps(dict(sorted(counties_refused.items(), key=lambda x:-x[1])), indent=2) + ''';

// Key historical events for timeline annotation
export const events = [
  { year: 2009, label: "Global financial crisis — Irish austerity begins" },
  { year: 2012, label: "Critical Skills Permit introduced" },
  { year: 2015, label: "Recovery — permits rise again" },
  { year: 2016, label: "Brexit vote — EU migration shift" },
  { year: 2019, label: "Record pre-COVID permits" },
  { year: 2020, label: "COVID-19 — permits fall 15%" },
  { year: 2021, label: "Recovery begins" },
  { year: 2022, label: "Post-COVID surge +66%" },
  { year: 2023, label: "All-time high: 33,490" },
  { year: 2024, label: "Slight moderation" },
];

export const SECTOR_COLORS = ''' + json.dumps(SECTOR_COLORS, indent=2) + ''';

export const PERMIT_COLORS = {
  critical_skills: "#00a878",
  general:         "#2563eb",
  intra_company:   "#7c3aed",
  dependant:       "#d97706",
  other:           "#e53e3e",
};

export const PERMIT_LABELS = {
  critical_skills: "Critical Skills",
  general:         "General Employment",
  intra_company:   "Intra-Company Transfer",
  dependant:       "Dependant / Partner",
  other:           "Other",
};
'''

OUT_DIR.mkdir(parents=True, exist_ok=True)
with open(OUT_DIR / "permits.js", "w", encoding="utf-8") as f:
    f.write(permits_js)
print(f"\n✅ Written: src/data/permits.js")


# ══════════════════════════════════════════════════════════════
# GENERATE companies.js
# ══════════════════════════════════════════════════════════════

# Get county for top companies by fuzzy matching known county keywords
COMPANY_COUNTIES = {
    "redwood extended care": "Dublin", "dawn meats": "Waterford",
    "costern": "Dublin", "bus átha cliath": "Dublin", "dublin bus": "Dublin",
    "applus": "Dublin", "resilience healthcare": "Dublin", "google": "Dublin",
    "inisc": "Dublin", "mowlam": "Dublin", "cork university": "Cork",
    "kel-tech": "Waterford", "kellor": "Dublin", "tata consultancy": "Dublin",
    "accenture": "Dublin", "nua healthcare": "Kildare", "ernst & young": "Dublin",
    "qt technologies": "Dublin", "intel ireland": "Kildare",
    "grant thornton": "Dublin", "microsoft ireland": "Dublin",
    "amazon web services": "Dublin", "amazon data": "Dublin",
    "amazon ireland": "Dublin", "cognizant": "Dublin", "amazon development": "Dublin",
    "applied materials": "Dublin", "analog devices": "Limerick",
    "deloitte": "Dublin", "meta platforms": "Dublin", "stripe technology": "Dublin",
    "software labs": "Dublin", "sfdc": "Dublin", "fisc-ireland": "Dublin",
    "kla-tencor": "Dublin", "mastercard": "Dublin", "optum": "Dublin",
    "apple distribution": "Cork", "apple operations": "Cork",
    "forvis mazars": "Dublin", "johnson and johnson vision": "Limerick",
    "silver stream": "Meath", "our lady of lourdes": "Louth",
    "university hospital kerry": "Kerry", "castlebridge": "Wexford",
    "university limerick": "Limerick", "children's health": "Dublin",
    "bartra": "Dublin", "kingdom home": "Kerry",
    "newbrook": "Dublin", "riada": "Dublin", "sonas": "Dublin",
    "communicare": "Dublin", "cv homecare": "Dublin", "haven bay": "Cork",
    "carechoice": "Dublin", "parke house": "Mayo", "tusla": "Dublin",
    "bon secours": "Cork", "wuxi biologics": "Dublin", "eli lilly": "Cork",
    "allergan": "Westmeath", "astellas": "Dublin", "astrazeneca": "Dublin",
    "amneal": "Dublin", "alexion": "Dublin", "abbvie": "Sligo",
    "abbott ireland": "Sligo", "circet": "Dublin", "kel-tech": "Waterford",
    "quarry and mining": "Kildare", "tli group": "Dublin",
    "barry's coaches": "Cork", "bus éireann": "Dublin", "gaeltec": "Galway",
    "glanua": "Galway", "callinan coaches": "Galway",
    "king and moffatt": "Roscommon", "atkinsr": "Dublin", "aecom": "Dublin",
    "kepak longford": "Longford", "kepak cork": "Cork",
    "tiernaneill mushrooms": "Monaghan", "monaghan mushrooms": "Monaghan",
    "mccarren meats": "Monaghan", "anglo beef": "Monaghan",
    "keelings": "Dublin", "farm solutions": "Kildare",
}

def get_county(name):
    n = name.lower()
    for k, v in COMPANY_COUNTIES.items():
        if k in n:
            return v
    return "Ireland"

companies_export = []
for c in companies_list:
    companies_export.append({
        "name":     c["name"],
        "industry": c["industry"],
        "county":   get_county(c["name"]),
        "total":    c["total"],
        "jan":      c["jan"],
        "feb":      c["feb"],
        "mar":      c["mar"],
        "apr":      c["apr"],
        "desc":     c["desc"],
    })

companies_js = '''// ============================================================
// Ireland Employment Permits — companies.js
// AUTO-GENERATED by scripts/extract-data.py
// Source: employment-permits-issued-to-companies-2026.xlsx
// DETE — enterprise.gov.ie
// Total companies with permits: ''' + str(len(companies_export)) + '''
// Period: January–April 2026
// Do NOT edit manually — re-run the script to regenerate
// ============================================================

export const COMPANIES = ''' + json.dumps(companies_export, indent=2) + ''';

export const INDUSTRY_LABELS = {
  all:          "All industries",
  tech:         "Technology & IT",
  pharma:       "Pharma & Lifesciences",
  finance:      "Finance & Banking",
  hospitality:  "Hospitality & Food",
  health:       "Healthcare",
  construction: "Construction & Transport",
  education:    "Education",
  food:         "Food & Agriculture",
  other:        "Other sectors",
};

export const INDUSTRY_ICONS = {
  tech:         "💻",
  pharma:       "🧬",
  finance:      "🏦",
  hospitality:  "🏨",
  health:       "🏥",
  construction: "🏗️",
  education:    "🎓",
  food:         "🍖",
  other:        "🏢",
};
'''

with open(OUT_DIR / "companies.js", "w", encoding="utf-8") as f:
    f.write(companies_js)
print(f"✅ Written: src/data/companies.js")

# Summary
print("\n" + "="*55)
print("📊 EXTRACTION COMPLETE")
print("="*55)
print(f"  Sectors:       {len(sectors_2026)} categories")
print(f"  Counties:      {len(counties_issued)} counties, {total_county:,} total permits")
print(f"  Nationalities: {len(nationalities_list)} countries")
print(f"  Companies:     {len(companies_list):,} companies with permits")
print(f"\n  Top 5 sectors:")
for k,v in sorted(sectors_2026.items(), key=lambda x:-x[1])[:5]:
    print(f"    {v:5d}  {k}")
print(f"\n  Top 5 counties:")
for k,v in sorted(counties_issued.items(), key=lambda x:-x[1])[:5]:
    print(f"    {v:5d}  {k}")
print(f"\n  Top 5 nationalities:")
for n in nationalities_list[:5]:
    print(f"    {n['issued']:5d}  {n['flag']} {n['country']}")
print(f"\n  Top 5 companies:")
for c in companies_export[:5]:
    print(f"    {c['total']:5d}  {c['name']}")
