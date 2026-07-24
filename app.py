import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="LNG Cargo Dashboard",
    layout="wide"
)

st.title("🚢 LNG Cargo Dashboard")

uploaded_file = st.file_uploader(
    "Upload Dashboard File",
    type=["xlsx", "xlsm"]
)

if uploaded_file:

    df = pd.read_excel(
        uploaded_file,
        sheet_name="TRADESNEW",
        header=6,
        engine="openpyxl"
    )

    df.columns = [str(c).strip() for c in df.columns]

    years = sorted(
        pd.to_numeric(
            df["IFRS YEAR"],
            errors="coerce"
        )
        .dropna()
        .unique()
    )

    selected_year = st.selectbox(
        "Select IFRS Year",
        years,
        index=len(years)-1
    )

    cargo_col = "CARGO#/TRADEID"

    filtered = df.copy()

    filtered["VOLUME"] = pd.to_numeric(
        filtered["VOLUME"],
        errors="coerce"
    )

    filtered["IFRS YEAR"] = pd.to_numeric(
        filtered["IFRS YEAR"],
        errors="coerce"
    )

    filtered["EXPOSURE DATE"] = pd.to_datetime(
        filtered["EXPOSURE DATE"],
        errors="coerce"
    )

    filtered = filtered[
        (filtered["FIRM"].astype(str).str.strip() == "FIRM")
        &
        (filtered["TYPE"].astype(str).str.strip() == "Cargo")
        &
        (filtered["IFRS YEAR"] == selected_year)
        &
        (filtered[cargo_col].notna())
    ]

    if filtered.empty:

        st.warning("No matching cargos found.")

    else:

        filtered["MONTH"] = (
            filtered["EXPOSURE DATE"]
            .dt.to_period("M")
            .dt.to_timestamp()
        )

        all_months = pd.date_range(
            start=f"{int(selected_year)}-01-01",
            end=f"{int(selected_year)}-12-01",
            freq="MS"
        )

        pivot = pd.pivot_table(
            filtered,
            index=cargo_col,
            columns="MONTH",
            values="VOLUME",
            aggfunc="sum"
        )

        pivot = pivot.reindex(
            columns=all_months,
            fill_value=np.nan
        )

        pivot = pivot / 1_000_000

        month_names = [
            d.strftime("%b-%y")
            for d in all_months
        ]

        pivot.columns = month_names

        pivot = pivot.sort_index()

        search = st.text_input(
            "Search Cargo Reference"
        )

        if search:

            pivot = pivot[
                pivot.index.str.contains(
                    search,
                    case=False,
                    na=False
                )
            ]

        total_volume = (
            filtered["VOLUME"].sum()
            / 1_000_000
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Cargo References",
            len(pivot)
        )

        col2.metric(
            "Total LNG Volume (MM)",
            f"{total_volume:,.2f}"
        )

        col3.metric(
            "IFRS Year",
            int(selected_year)
        )

        def format_value(x):

            if pd.isna(x):
                return "-"

            if abs(x) < 0.0001:
                return "-"

            if x < 0:
                return f"({abs(x):.2f})"

            return f"{x:.2f}"

        styled = pivot.style.format(
            format_value
        )

        styled = styled.map(
            lambda v:
            "color:red;"
            if (
                pd.notna(v)
                and isinstance(v, (int,float))
                and v < 0
            )
            else (
                "color:#00ff99;"
                if (
                    pd.notna(v)
                    and isinstance(v,(int,float))
                    and v > 0
                )
                else ""
            )
        )

        st.dataframe(
            pivot.style.format(format_value),
            use_container_width=True,
            height=700
        )

        month_totals = pivot.sum(
            numeric_only=True
        )

        st.subheader("Monthly Totals")

        st.bar_chart(month_totals)

        export_df = pivot.copy()

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
        ) as writer:

            export_df.to_excel(
                writer,
                sheet_name="Cargo Dashboard"
            )

        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name=f"Cargo_Dashboard_{int(selected_year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
