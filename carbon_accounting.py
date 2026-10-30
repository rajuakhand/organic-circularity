import pandas as pd
import plotly.express as px
import streamlit as st


def render_carbon_accounting(df):
    st.subheader("IPCC Tier 1 – Carbon Accounting")
    st.caption(
        "Daily calculations are generated automatically from the loaded Excel data. "
        "The dataset total includes both recorded and planned forecast entries."
    )

    daily = df[["Date", "Organic Waste (kg)"]].copy()
    daily["Date"] = pd.to_datetime(daily["Date"], errors="coerce").dt.normalize()
    daily["Organic Waste (kg)"] = pd.to_numeric(
        daily["Organic Waste (kg)"], errors="coerce"
    )
    daily = daily.dropna(subset=["Date", "Organic Waste (kg)"])
    daily = daily[daily["Organic Waste (kg)"] >= 0].copy()

    if daily.empty:
        st.warning("No valid waste records are available for carbon accounting.")
        return

    # IPCC Tier 1 calculations
    daily["CH4 (kg)"] = daily["Organic Waste (kg)"] * 4 / 1000
    daily["N2O (kg)"] = daily["Organic Waste (kg)"] * 0.24 / 1000
    daily["CH4 (kg CO2e)"] = daily["CH4 (kg)"] * 28
    daily["N2O (kg CO2e)"] = daily["N2O (kg)"] * 265
    daily["Total CO2e (kg)"] = (
        daily["CH4 (kg CO2e)"] + daily["N2O (kg CO2e)"]
    )

    today = pd.Timestamp.now(tz="Asia/Kolkata").tz_localize(None).normalize()
    daily["Data status"] = daily["Date"].apply(
        lambda date: "Actual / recorded" if date <= today else "Forecast / planned"
    )
    daily["Month"] = daily["Date"].dt.to_period("M").astype(str)

    actual = daily[daily["Data status"] == "Actual / recorded"]
    forecast = daily[daily["Data status"] == "Forecast / planned"]

    # Full dataset totals: actual + forecast
    waste = daily["Organic Waste (kg)"].sum()
    ch4 = daily["CH4 (kg)"].sum()
    n2o = daily["N2O (kg)"].sum()
    total = daily["Total CO2e (kg)"].sum()

    st.markdown("### Complete dataset summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total organic waste", f"{waste:,.2f} kg")
    c2.metric("Estimated CH₄", f"{ch4:,.4f} kg")
    c3.metric("Estimated N₂O", f"{n2o:,.4f} kg")
    c4.metric("Composting emissions", f"{total:,.2f} kg CO₂e")

    st.caption(
        f"Includes {len(actual)} actual/recorded rows and "
        f"{len(forecast)} forecast/planned rows."
    )

    st.markdown("### Actual and forecast breakdown")
    a1, a2, a3 = st.columns(3)
    a1.metric(
        "Recorded waste",
        f"{actual['Organic Waste (kg)'].sum():,.2f} kg"
    )
    a2.metric(
        "Recorded emissions",
        f"{actual['Total CO2e (kg)'].sum():,.2f} kg CO₂e"
    )
    a3.metric("Recorded rows", f"{len(actual):,}")

    f1, f2, f3 = st.columns(3)
    f1.metric(
        "Forecast waste",
        f"{forecast['Organic Waste (kg)'].sum():,.2f} kg"
    )
    f2.metric(
        "Forecast emissions",
        f"{forecast['Total CO2e (kg)'].sum():,.2f} kg CO₂e"
    )
    f3.metric("Forecast rows", f"{len(forecast):,}")

    st.divider()
    st.markdown("### Daily organic waste and emissions")
    daily_chart = daily.sort_values("Date")

    fig_waste = px.bar(
        daily_chart,
        x="Date",
        y="Organic Waste (kg)",
        color="Data status",
        title="Daily Organic Waste",
        labels={"Organic Waste (kg)": "Organic waste (kg)", "Date": "Date"},
        color_discrete_map={
            "Actual / recorded": "#238b45",
            "Forecast / planned": "#f0a202",
        },
    )
    st.plotly_chart(fig_waste, use_container_width=True)

    fig_gases = px.line(
        daily_chart,
        x="Date",
        y=["CH4 (kg)", "N2O (kg)"],
        title="Daily Estimated CH₄ and N₂O",
        labels={"value": "Gas quantity (kg)", "Date": "Date", "variable": "Gas"},
    )
    st.plotly_chart(fig_gases, use_container_width=True)

    fig_daily_co2e = px.line(
        daily_chart,
        x="Date",
        y="Total CO2e (kg)",
        color="Data status",
        title="Daily Composting Process Emissions",
        labels={"Total CO2e (kg)": "CO₂e (kg)", "Date": "Date"},
        color_discrete_map={
            "Actual / recorded": "#238b45",
            "Forecast / planned": "#f0a202",
        },
    )
    st.plotly_chart(fig_daily_co2e, use_container_width=True)

    st.markdown("### Monthly emissions trend")
    monthly = (
        daily.groupby(["Month", "Data status"], as_index=False)
        .agg(
            Waste_kg=("Organic Waste (kg)", "sum"),
            CH4_kg=("CH4 (kg)", "sum"),
            N2O_kg=("N2O (kg)", "sum"),
            CO2e_kg=("Total CO2e (kg)", "sum"),
        )
    )
    fig_monthly = px.bar(
        monthly,
        x="Month",
        y="CO2e_kg",
        color="Data status",
        barmode="stack",
        title="Monthly Composting Process Emissions",
        labels={"CO2e_kg": "CO₂e (kg)", "Month": "Month"},
        color_discrete_map={
            "Actual / recorded": "#238b45",
            "Forecast / planned": "#f0a202",
        },
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

    st.markdown("### Daily calculation table")
    display_daily = daily.sort_values("Date", ascending=False).drop(columns=["Month"])
    st.dataframe(
        display_daily,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Date": st.column_config.DateColumn("Date", format="DD-MM-YYYY"),
            "Organic Waste (kg)": st.column_config.NumberColumn(format="%.2f"),
            "CH4 (kg)": st.column_config.NumberColumn(format="%.5f"),
            "N2O (kg)": st.column_config.NumberColumn(format="%.5f"),
            "CH4 (kg CO2e)": st.column_config.NumberColumn(format="%.4f"),
            "N2O (kg CO2e)": st.column_config.NumberColumn(format="%.4f"),
            "Total CO2e (kg)": st.column_config.NumberColumn(format="%.4f"),
        },
    )

    st.markdown("### Calculation formulas")
    st.latex(r"CH_4(kg)=\frac{Waste(kg)\times4}{1000}")
    st.latex(r"N_2O(kg)=\frac{Waste(kg)\times0.24}{1000}")
    st.latex(r"Total\ CO_2e=(CH_4\times28)+(N_2O\times265)")

    st.markdown("### Methodology and limitations")
    st.markdown("""
- CH₄ emission factor: 4 g CH₄/kg wet waste.
- N₂O emission factor: 0.24 g N₂O/kg wet waste.
- GWP100 values (IPCC AR5): CH₄ = 28; N₂O = 265.
- This is a generic IPCC Tier 1 estimate of composting process emissions.
- These values are not direct gas measurements or verified avoided emissions.
- Forecast/planned entries are estimates and should be updated when actual records become available.
- Net GHG benefit is not calculated because the previous disposal baseline is not independently verified.
- Electricity, fuel and transport emissions are not included.
""")

    st.download_button(
        "Download daily carbon calculations CSV",
        daily.drop(columns=["Month"]).to_csv(index=False).encode("utf-8-sig"),
        file_name="daily_carbon_calculations.csv",
        mime="text/csv",
    )
