import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="LNG Cargo Exposure Dashboard",
    page_icon="🚢",
    layout="wide"
)

# =========================
# STYLING
# =========================

st.markdown("""
<style>

.block-container {
    padding-top: 1rem;
}

[data-testid="metric-container"] {
    background-color: #111827;
    border: 1px solid #293241;
    padding: 15px;
    border-radius: 12px;
}

h1,h2,h3 {
    color: white;
}

</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================

st.title("🚢 LNG Cargo Exposure Dashboard")

uploaded_file = st.file_uploader(
    "Upload COB Dashboard",
    type=["xlsx", "xlsm"]
)

if uploaded_file:

    # =========================
    # READ FILE
    # =========================

    try:

        df = pd.read_excel(
            uploaded_file,
            sheet_name="TRADESNEW",
            header=6,
            engine="openpyxl"
        )

    except Exception as e:

        st.error(f"Failed to load workbook: {e}")
        st.stop()

    df.columns = [str(c).strip() for c in df.columns]

    required_cols = [
        "FIRM",
        "TYPE",
        "CARGO#/TRADEID",
        "VOLUME",
        "EXPOSURE DATE",
        "IFRS YEAR"
    ]

    missing_cols = [
        c for c in required_cols
        if c not in df.columns
    ]

    if missing_cols:

        st.error(
            f"Missing columns: {missing_cols}"
        )

        st.write("Columns found:")
        st.write(df.columns.tolist())

        st.stop()

    # =========================
    # IFRS YEAR
    # =========================

    years = sorted(
        pd.to_numeric(
            df["IFRS YEAR"],
            errors="coerce"
        )
        .dropna()
        .astype(int)
        .unique()
    )

    selected_year = st.selectbox(
        "IFRS Year",
        years,
        index=len(years)-1
    )

    # =========================
    # DATA CLEANING
    # =========================

    data = df.copy()

    data["VOLUME"] = pd.to_numeric(
        data["VOLUME"],
        errors="coerce"
    )

    data["IFRS YEAR"] = pd.to_numeric(
        data["IFRS YEAR"],
        errors="coerce"
    )

    data["EXPOSURE DATE"] = pd.to_datetime(
        data["EXPOSURE DATE"],
        errors="coerce"
    )

    # =========================
    # FILTERS
    # =========================

    data = data[
        (
                data["FIRM"]
                .astype(str)
                .str.strip()
                .str.upper()
                == "FIRM"
        )
        &
        (
                data["TYPE"]
                .astype(str)
                .str.strip()
                .str.upper()
                == "CARGO"
        )
        &
        (data["IFRS YEAR"] == selected_year)
        &
        (
                data["CARGO#/TRADEID"]
                .astype(str)
                .str.strip()
                != ""
        )
        &
        (data["VOLUME"] != 0)
    ]

    if data.empty:

        st.warning("No Cargo exposures found.")

        st.stop()

    # =========================
    # MONTH
    # =========================

    data["MONTH"] = (
        data["EXPOSURE DATE"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    min_month = (
        data["MONTH"]
        .min()
        .replace(day=1)
    )

    max_month = (
        data["MONTH"]
        .max()
        .replace(day=1)
    )

    months = pd.date_range(
        start=min_month,
        end=max_month,
        freq="MS"
    )

    # =========================
    # PIVOT
    # =========================

    pivot = pd.pivot_table(
        data,
        index="CARGO#/TRADEID",
        columns="MONTH",
        values="VOLUME",
        aggfunc="sum"
    )

    pivot = pivot.reindex(
        columns=months
    )

    pivot = pivot / 1_000_000

    month_labels = [
        m.strftime("%b-%y")
        for m in months
    ]

    pivot.columns = month_labels

    # =========================
    # SEARCH
    # =========================

    search = st.text_input(
        "🔍 Search Cargo Reference"
    )

    if search:

        pivot = pivot[
            pivot.index.astype(str)
            .str.contains(
                search,
                case=False,
                na=False
            )
        ]

    # =========================
    # KPIs
    # =========================

    kpi_data = data.copy()

    cargo_exposure = (
            kpi_data.groupby("CARGO#/TRADEID")["VOLUME"]
            .sum()
            / 1_000_000
    )

    cargo_exposure = (
            data.groupby(
                "CARGO#/TRADEID"
            )["VOLUME"]
            .sum()
            / 1_000_000
    )

    cargo_count = len(cargo_exposure)

    long_cargo_count = int(
        (cargo_exposure > 0).sum()
    )

    short_cargo_count = int(
        (cargo_exposure < 0).sum()
    )

    long_volume = (
            data.loc[
                data["VOLUME"] > 0,
                "VOLUME"
            ].sum()
            / 1_000_000
    )

    short_volume = (
            data.loc[
                data["VOLUME"] < 0,
                "VOLUME"
            ].sum()
            / 1_000_000
    )

    net_position = (
            data["VOLUME"].sum()
            / 1_000_000
    )

    row1 = st.columns(3)

    row1[0].metric(
        "Cargo References",
        f"{cargo_count:,}"
    )

    row1[1].metric(
        "Long Cargoes",
        int(long_cargo_count)
    )

    row1[2].metric(
        "Short Cargoes",
        int(short_cargo_count)
    )

    row2 = st.columns(3)

    row2[0].metric(
        "Net Exposure",
        f"{net_position:,.2f} TBtu"
    )

    row2[1].metric(
        "Long Volume",
        f"{long_volume:,.2f} TBtu"
    )

    row2[2].metric(
        "Short Volume",
        f"{short_volume:,.2f} TBtu"
    )

    st.divider()

    # =========================
    # TABLE
    # =========================

    st.subheader("Cargo Exposure Matrix")

    display_df = pivot.copy()

    for col in display_df.columns:

        display_df[col] = display_df[col].map(
            lambda x:
            "-"
            if pd.isna(x)
            else (
                "-"
                if abs(x) < 0.00001
                else (
                    f"({abs(x):,.2f})"
                    if x < 0
                    else f"{x:,.2f}"
                )
            )
        )

    st.dataframe(
        display_df,
        use_container_width=True,
        height=700
    )

    # =========================
    # MONTHLY TOTALS
    # =========================

    monthly_totals = pd.DataFrame(
        {
            "Month": months,
            "Volume (TBtu)":
            [
                data.loc[
                    data["MONTH"] == m,
                    "VOLUME"
                ].sum() / 1_000_000
                for m in months
            ]
        }
    )

    st.divider()

    st.subheader("Monthly LNG Exposure")

    fig = px.bar(
        monthly_totals,
        x="Month",
        y="Volume (TBtu)",
        text="Volume (TBtu)",
        template="plotly_dark"
    )

    fig.update_traces(
        marker_color="#00E676",
        texttemplate="%{y:.1f}",
        textposition="outside"
    )

    fig.update_layout(
        height=500,
        xaxis_title="",
        yaxis_title="Volume (TBtu)",
        paper_bgcolor="#0E1117",
        plot_bgcolor="#0E1117"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.divider()

    st.subheader("Monthly Long / Short Breakdown")

    monthly_summary = []

    for m in months:

        month_data = data.loc[
            data["MONTH"] == m
            ].copy()

        if len(month_data) == 0:
            monthly_summary.append(
                [
                    m.strftime("%b-%y"),
                    0,
                    0,
                    0,
                    0,
                    0
                ]
            )

            continue

        cargo_month = (
                month_data.groupby(
                    "CARGO#/TRADEID"
                )["VOLUME"]
                .sum()
                / 1_000_000
        )

        long_count = int(
            (cargo_month > 0).sum()
        )

        short_count = int(
            (cargo_month < 0).sum()
        )

        long_tbtu = (
                month_data.loc[
                    month_data["VOLUME"] > 0,
                    "VOLUME"
                ].sum()
                / 1_000_000
        )

        short_tbtu = (
                month_data.loc[
                    month_data["VOLUME"] < 0,
                    "VOLUME"
                ].sum()
                / 1_000_000
        )

        net_tbtu = (
                month_data["VOLUME"].sum()
                / 1_000_000
        )

        monthly_summary.append(
            [
                m.strftime("%b-%y"),
                long_count,
                short_count,
                round(long_tbtu, 3),
                round(short_tbtu, 3),
                round(net_tbtu, 3)
            ]
        )

    monthly_summary = pd.DataFrame(
        monthly_summary,
        columns=[
            "Month",
            "Long Cargoes",
            "Short Cargoes",
            "Long TBtu",
            "Short TBtu",
            "Net TBtu"
        ]
    )

    st.dataframe(
        monthly_summary,
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # DOWNLOAD
    # =========================

    output = BytesIO()

    export_df = pivot.copy()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        export_df.to_excel(
            writer,
            sheet_name="Cargo Exposure"
        )

    st.download_button(
        "📥 Download Excel",
        output.getvalue(),
        file_name=f"Cargo_Exposure_{int(selected_year)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write(
        "Cargo refs in filtered data",
        data["CARGO#/TRADEID"].nunique()
    )

    display_data = data[
        data["EXPOSURE DATE"] >= min_month
        ]

    st.write(
        "Cargo refs after month filter",
        display_data["CARGO#/TRADEID"].nunique()
    )
