import streamlit as st
import pandas as pd
from pathlib import Path
import plotly.express as px

st.set_page_config(
    page_title="Organic Waste Circularity Dashboard",
    page_icon="♻️",
    layout="wide",
    initial_sidebar_state="expanded"
)

DEFAULT_FILE = Path("data/MRCPS Organic Waste Tracking 2026.xlsx")
SHEET_NAME = "Waste Data 2026"

REQUIRED = [
    "Date",
    "Organic Waste (kg)",
    "CO2 Avoided (kg)",
    "Compost (kg)",
    "Liquid Leachate (L)"
]

# ---------- Professional styling ----------
st.markdown("""
<style>
    :root {
        color-scheme: light;
    }

    .stApp {
        background:
            radial-gradient(ellipse at 8% 0%, rgba(199, 229, 204, .42), transparent 35%),
            radial-gradient(ellipse at 100% 18%, rgba(224, 237, 215, .55), transparent 32%),
            #f7f8f4;
        color: #20372b;
    }

    [data-testid="stHeader"] {
        background: rgba(247, 248, 244, .88);
    }

    .block-container {
        max-width: 1500px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(175deg, #123b2b 0%, #1b5038 58%, #286448 100%);
        border-right: 1px solid rgba(255,255,255,.12);
    }

    [data-testid="stSidebar"] * {
        color: #f3fff6;
    }

    [data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 12px;
        padding: 10px;
    }

    .hero {
        background: linear-gradient(115deg, #103c2b 0%, #1c6845 58%, #39845a 100%);
        padding: 34px 36px;
        border-radius: 22px;
        color: #ffffff;
        margin: 0 0 28px 0;
        box-shadow: 0 14px 32px rgba(25, 75, 49, .16);
        position: relative;
        overflow: hidden;
    }

    .hero h1 {
        color: #ffffff;
        font-size: clamp(1.8rem, 3vw, 2.5rem);
        letter-spacing: -0.7px;
        margin: 0 0 10px 0;
        line-height: 1.2;
    }

    .hero p {
        color: #e0f2e5;
        margin: 0;
        font-size: 1.02rem;
        line-height: 1.7;
        max-width: 850px;
    }

    h1, h2, h3 {
        color: #183f2d;
        letter-spacing: -0.35px;
    }

    h2 {
        margin-top: 1.2rem;
    }

    [data-testid="stMetric"] {
        background: linear-gradient(145deg, #ffffff, #fbfdf9);
        border: 1px solid #e1e9df;
        border-radius: 18px;
        padding: 21px 20px;
        min-height: 125px;
        box-shadow: 0 6px 20px rgba(29, 66, 43, .055);
        transition: transform .18s ease, box-shadow .18s ease;
    }

    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px rgba(29, 66, 43, .10);
    }

    [data-testid="stMetricLabel"] {
        color: #607768;
        font-size: .92rem;
        font-weight: 600;
    }

    [data-testid="stMetricValue"] {
        color: #17563a;
        font-size: clamp(1.35rem, 2vw, 2rem);
        font-weight: 750;
    }

    [data-testid="stPlotlyChart"] {
        background: #ffffff;
        border: 1px solid #e4ebe3;
        border-radius: 17px;
        padding: 10px 8px 4px 8px;
        box-shadow: 0 5px 18px rgba(29, 66, 43, .045);
    }

    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 8px;
        background: #eaf0e9;
        padding: 7px;
        border-radius: 13px;
    }

    [data-testid="stTabs"] button[data-baseweb="tab"] {
        background: transparent;
        border-radius: 9px;
        padding: 10px 16px;
        color: #3d5c49;
        font-weight: 600;
    }

    [data-testid="stTabs"] button[aria-selected="true"] {
        background: #ffffff;
        color: #17613d;
        box-shadow: 0 2px 7px rgba(20, 60, 38, .09);
    }

    [data-testid="stDataFrame"] {
        background: #ffffff;
        border: 1px solid #e2eae1;
        border-radius: 14px;
        overflow: hidden;
    }

    .stDownloadButton button {
        background: linear-gradient(110deg, #1b7048, #2a8054);
        color: #ffffff;
        border: 0;
        border-radius: 10px;
        padding: 10px 18px;
        font-weight: 650;
        box-shadow: 0 4px 10px rgba(29, 99, 61, .12);
    }

    .stDownloadButton button:hover {
        background: #145637;
        color: #ffffff;
        border: 0;
    }

    div[data-testid="stAlert"] {
        border-radius: 12px;
    }

    hr {
        border-color: #dfe8de;
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
    }

    .section-note {
        color: #63796b;
        font-size: .94rem;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-top: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        .hero {
            padding: 24px 21px;
            border-radius: 16px;
        }
        [data-testid="stMetric"] {
            padding: 15px;
            min-height: 105px;
        }
    }
</style>
""", unsafe_allow_html=True)


# ---------- Data loading ----------
@st.cache_data
def load_excel(file_source):
    return pd.read_excel(file_source, sheet_name=SHEET_NAME)


st.sidebar.markdown("## ♻️ Circularity Hub")
st.sidebar.caption("Organic waste • Compost • Carbon")

uploaded = st.sidebar.file_uploader(
    "Upload or replace Excel file",
    type=["xlsx"],
    help="Upload a workbook containing the 'Waste Data 2026' sheet."
)

if uploaded is not None:
    source_label = f"Uploaded: {uploaded.name}"
    try:
        df = load_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read the uploaded workbook: {e}")
        st.info(f"Check that it contains a sheet named '{SHEET_NAME}'.")
        st.stop()
elif DEFAULT_FILE.exists():
    source_label = DEFAULT_FILE.name
    try:
        df = load_excel(str(DEFAULT_FILE))
    except Exception as e:
        st.error(f"Could not read the default Excel file: {e}")
        st.stop()
else:
    st.warning("Default Excel file not found. Upload your workbook in the sidebar.")
    st.stop()


# ---------- Validate and prepare data ----------
df.columns = df.columns.astype(str).str.strip()

missing = [col for col in REQUIRED if col not in df.columns]
if missing:
    st.error("Required columns are missing: " + ", ".join(missing))
    st.write("Columns found in the workbook:", df.columns.tolist())
    st.stop()

df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)

for col in REQUIRED[1:]:
    df[col] = pd.to_numeric(df[col], errors="coerce")

invalid_dates = int(df["Date"].isna().sum())
df = df.dropna(subset=["Date"]).copy()

if df.empty:
    st.error("No valid dates were found. Please check the Date column.")
    st.stop()

df["Month"] = df["Date"].dt.to_period("M").astype(str)

# Daily measurements: aggregate across each month.
monthly = df.groupby("Month", as_index=False).agg(
    Waste_kg=("Organic Waste (kg)", "sum"),
    Existing_CO2_estimate_kg=("CO2 Avoided (kg)", "sum"),
    Days_Recorded=("Date", "nunique")
)

# Compost and leachate: monthly totals recorded on month-end rows.
production_rows = df.dropna(
    subset=["Compost (kg)", "Liquid Leachate (L)"]
)

monthly_production = production_rows.groupby("Month", as_index=False).agg(
    Compost_kg=("Compost (kg)", "sum"),
    Leachate_L=("Liquid Leachate (L)", "sum")
)

monthly = monthly.merge(monthly_production, on="Month", how="left")
monthly[["Compost_kg", "Leachate_L"]] = (
    monthly[["Compost_kg", "Leachate_L"]].fillna(0)
)

total_waste = df["Organic Waste (kg)"].sum()
total_co2 = df["CO2 Avoided (kg)"].sum()
total_compost = monthly["Compost_kg"].sum()
total_leachate = monthly["Leachate_L"].sum()
month_count = df["Month"].nunique()

# ---------- Sidebar information ----------
st.sidebar.divider()
st.sidebar.markdown("### 📁 Data source")
st.sidebar.caption(source_label)
st.sidebar.success(f"{len(df):,} valid-date rows loaded")
st.sidebar.write(f"Months found: {month_count}")

if invalid_dates:
    st.sidebar.warning(f"{invalid_dates} row(s) with invalid dates were excluded.")

st.sidebar.divider()
st.sidebar.caption(
    "CO₂ values are inherited from the workbook. "
    "They are estimates, not independently verified carbon savings."
)


# ---------- Header ----------
st.markdown("""
<div class="hero">
    <h1>♻️ Organic Waste Circularity</h1>
    <p>Waste tracking, compost production, liquid leachate and workbook-based carbon estimates.</p>
</div>
""", unsafe_allow_html=True)

st.markdown("### 📊 Performance overview")
st.markdown(
    '<p class="section-note">Summary of the data currently loaded into the dashboard.</p>',
    unsafe_allow_html=True
)

# ---------- KPI cards ----------
k1, k2, k3, k4 = st.columns(4)

k1.metric("Organic Waste", f"{total_waste:,.0f} kg")
k2.metric("Existing CO₂ Estimate", f"{total_co2:,.1f} kg CO₂e")
k3.metric("Compost Produced", f"{total_compost:,.0f} kg")
k4.metric("Liquid Leachate", f"{total_leachate:,.0f} L")

st.write("")

# ---------- Chart configuration ----------
chart_template = "plotly_white"
chart_colors = {
    "green": "#238653",
    "dark": "#24543b",
    "teal": "#2a8c82",
    "gold": "#d3a43b"
}

tab_overview, tab_data, tab_raw, tab_method = st.tabs([
    "📈 Overview",
    "📋 Monthly data",
    "🗂️ Raw data",
    "🔬 Methodology"
])

with tab_overview:
    left, right = st.columns(2)

    with left:
        st.subheader("Monthly Organic Waste")
        fig = px.bar(
            monthly,
            x="Month",
            y="Waste_kg",
            text_auto=".0f",
            labels={"Month": "Month", "Waste_kg": "Organic waste (kg)"},
            template=chart_template,
            color_discrete_sequence=[chart_colors["green"]]
        )
        fig.update_layout(
            xaxis_tickangle=-40,
            margin=dict(l=15, r=15, t=25, b=15),
            yaxis_title="Waste (kg)",
            xaxis_title=""
        )
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Monthly CO₂ Estimate")
        fig = px.line(
            monthly,
            x="Month",
            y="Existing_CO2_estimate_kg",
            markers=True,
            labels={
                "Month": "Month",
                "Existing_CO2_estimate_kg": "CO₂ estimate (kg CO₂e)"
            },
            template=chart_template,
            color_discrete_sequence=[chart_colors["dark"]]
        )
        fig.update_traces(line_width=3, marker_size=8)
        fig.update_layout(
            xaxis_tickangle=-40,
            margin=dict(l=15, r=15, t=25, b=15),
            yaxis_title="Estimated CO₂ (kg CO₂e)",
            xaxis_title=""
        )
        st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)

    with left:
        st.subheader("Monthly Compost Production")
        fig = px.bar(
            monthly,
            x="Month",
            y="Compost_kg",
            text_auto=".0f",
            labels={"Month": "Month", "Compost_kg": "Compost (kg)"},
            template=chart_template,
            color_discrete_sequence=[chart_colors["teal"]]
        )
        fig.update_layout(
            xaxis_tickangle=-40,
            margin=dict(l=15, r=15, t=25, b=15),
            yaxis_title="Compost (kg)",
            xaxis_title=""
        )
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Monthly Liquid Leachate")
        fig = px.bar(
            monthly,
            x="Month",
            y="Leachate_L",
            text_auto=".0f",
            labels={"Month": "Month", "Leachate_L": "Leachate (L)"},
            template=chart_template,
            color_discrete_sequence=[chart_colors["gold"]]
        )
        fig.update_layout(
            xaxis_tickangle=-40,
            margin=dict(l=15, r=15, t=25, b=15),
            yaxis_title="Leachate (L)",
            xaxis_title=""
        )
        st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "Note: the CO₂ chart shows the workbook's existing estimate. "
        "It should not be interpreted as verified emissions reductions."
    )

with tab_data:
    st.subheader("Monthly KPI table")

    display_monthly = monthly.rename(columns={
        "Month": "Month",
        "Waste_kg": "Organic Waste (kg)",
        "Existing_CO2_estimate_kg": "Existing CO₂ Estimate (kg CO₂e)",
        "Compost_kg": "Compost (kg)",
        "Leachate_L": "Liquid Leachate (L)",
        "Days_Recorded": "Days Recorded"
    })

    st.dataframe(
        display_monthly,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "⬇️ Download Monthly Summary CSV",
        data=display_monthly.to_csv(index=False).encode("utf-8-sig"),
        file_name="monthly_circularity_summary.csv",
        mime="text/csv"
    )

with tab_raw:
    st.subheader("Source data preview")
    st.caption("This table shows the records read from the selected Excel workbook.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download Processed Data CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name="organic_waste_processed_data.csv",
        mime="text/csv"
    )

with tab_method:
    st.subheader("Data definitions and methodology")

    st.markdown("""
**Organic waste**
- Daily values from the `Organic Waste (kg)` column are summed by month.

**Compost and liquid leachate**
- The workbook records these as monthly production totals.
- Rows containing both production values are grouped by month.
- The monthly values are summed once from the rows recorded in the workbook.

**CO₂ estimate**
- The dashboard sums the existing values in `CO2 Avoided (kg)`.
- The emission factor and calculation methodology have not been independently validated here.
- Treat this as a workbook-based estimate, not verified carbon savings or a carbon-credit claim.

**Data quality**
- Rows with invalid dates are excluded from monthly reporting.
- Required column names are checked before analysis.
- Numeric columns are converted to numeric values where possible.
""")

    st.info(
        "For scientific reporting, document and validate the emissions factor, "
        "baseline, system boundary and calculation method before claiming CO₂ reductions."
    )

st.divider()
st.caption(
    "Organic Waste Circularity Dashboard • Data-driven monitoring for waste, "
    "compost and leachate • Validate carbon methodology before external reporting."
)
