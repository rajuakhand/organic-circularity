
# 🌱 Organic Circularity Dashboard

### Organic Waste Tracking, Performance Monitoring & Carbon Accounting

An interactive Python and Streamlit application for organic waste tracking, performance monitoring, compost quality monitoring, and IPCC Tier 1 composting process-emissions estimates.

## 🔗 Project Links

- **Live Dashboard:** https://organic-circularity-rajuakhand.streamlit.app/
- **GitHub Repository:** https://github.com/rajuakhand/organic-circularity

## ✨ Key Features

- Organic waste tracking and performance overview
- Monthly KPI monitoring and interactive charts
- Excel-based data input and CSV reports
- Data-quality reporting
- Compost quality monitoring with Google Sheets
- IPCC Tier 1 carbon accounting with daily and monthly charts
- Optional AI-assisted analysis

## 🌍 Carbon Accounting

Generic IPCC Tier 1 composting factors:

- CH₄: 4 g CH₄/kg wet waste
- N₂O: 0.24 g N₂O/kg wet waste
- AR5 GWP100: CH₄ = 28; N₂O = 265

Forecast entries are identified separately from recorded data. These estimates are not direct gas measurements or verified avoided emissions. Net GHG benefit is not calculated because the previous disposal baseline is not independently verified. Electricity, fuel, and transport emissions are excluded.

The workbook's legacy CO₂ estimate is retained separately and is not verified carbon savings.

## 🛠️ Technology

Python, Streamlit, Pandas, Plotly, Excel, Google Sheets, and Google Gemini API.

## 🚀 Run Locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
