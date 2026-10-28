# 🌱 Organic Circularity Dashboard

### Organic Waste Tracking, Performance Monitoring & Circularity Analytics

An interactive data analytics application built with **Python** and **Streamlit** to explore organic waste tracking data, monitor monthly KPIs, and review data quality.

---

## 📑 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Technology Stack](#technology-stack)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
- [Run the Application](#run-the-application)
- [Reports](#reports)
- [Security](#security)
- [Author](#author)

---

## 📌 Overview

The Organic Circularity Dashboard provides an interactive interface for reviewing organic waste tracking information, understanding performance indicators, and examining data-quality results.

## ✨ Key Features

- 📊 Interactive dashboard for organic waste data
- 📈 Monthly KPI monitoring
- 📋 Monthly summary reports
- 🔎 Data-quality reporting
- 🤖 AI Analyst interface
- 📁 Excel-based data input and CSV reports

## 🛠️ Technology Stack

| Tool | Purpose |
|---|---|
| Python | Application logic and data processing |
| Streamlit | Interactive web dashboard |
| Pandas | Data analysis and transformation |
| Excel | Organic waste tracking data |
| CSV | Exported reports |
| Google Gemini API | AI-assisted analysis, when configured |


## 📂 Project Structure

```text
organic-circularity/
├── app.py
├── data/
├── outputs/
├── backups/
├── .streamlit/
├── .gitignore
└── README.md
```

## 🚀 Installation

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install streamlit pandas openpyxl google-genai
```

## ▶️ Run the Application

চালানোর জন্য টার্মিনালে ব্যবহার করো:

```bash
streamlit run app.py
```

## 📊 Reports & Outputs

- `data_quality_report.csv` — Data quality findings
- `monthly_kpi_report.csv` — Monthly KPI results
- `monthly_summary.csv` — Monthly summary data

## 🔐 Security

- Keep `.streamlit/secrets.toml` private.
- Never publish API keys or passwords to GitHub.
- Review reports before sharing them publicly.

## 👤 Author

**Raju Akhand**

GitHub: [@rajuakhand](https://github.com/rajuakhand)

---

*Supporting data-driven organic waste management and circularity analysis.*
