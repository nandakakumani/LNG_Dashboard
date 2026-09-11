import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(
    page_title="LNG Contract Exposure Dashboard",
    page_icon="🚢",
    layout="wide",
)

# ============================================================
# STYLE
# ============================================================

st.markdown(
    """
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

    h1, h2, h3 {
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("LNG Contract Exposure Dashboard")

uploaded_file = st.file_uploader(
    "Upload COB Dashboard",
    type=["xlsx", "xlsm"],
)

if uploaded_file is None:
    st.info("Upload the COB workbook to generate the exposure table.")
    st.stop()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """Standardise text used for matching."""
    if pd.isna(value):
        return ""

    return " ".join(str(value).strip().upper().split())


def read_workbook(file):
    """Read the required worksheets."""
    trades = pd.read_excel(
        file,
        sheet_name="TRADESNEW",
        header=6,
        engine="openpyxl",
    )

    file.seek(0)

    cargo_refs = pd.read_excel(
        file,
        sheet_name="All cargo refs",
        header=None,
        engine="openpyxl",
    )

    return trades, cargo_refs


def build_reference_mapping(cargo_refs):
    """
    Build a lookup table where each cargo reference is linked
    to the contract shown in column B of 'All cargo refs'.

    Column B = contract
    Columns C onward = cargo references
    """
    mapping_rows = []

    for _, row in cargo_refs.iterrows():
        contract = row.iloc[1] if len(row) > 1 else None

        if pd.isna(contract):
            continue

        contract = str(contract).strip()

        if contract == "":
            continue

        if clean_text(contract) == "CONTRACT":
            continue

        for cargo_ref in row.iloc[2:]:
            if pd.isna(cargo_ref):
                continue

            cargo_ref = str(cargo_ref).strip()

            if cargo_ref == "":
                continue

            mapping_rows.append(
                {
                    "CARGO_REFERENCE": cargo_ref,
                    "REFERENCE_KEY": clean_text(cargo_ref),
                    "CONTRACT_FROM_LIST": contract,
                }
            )

    mapping = pd.DataFrame(mapping_rows)

    if mapping.empty:
        return mapping

    mapping = (
        mapping.drop_duplicates(
            subset=["REFERENCE_KEY"],
            keep="first",
        )
        .sort_values(
            "REFERENCE_KEY",
            key=lambda x: x.str.len(),
            ascending=False,
        )
        .reset_index(drop=True)
    )

    return mapping


def find_contract(cargo_reference, reference_mapping):
    """
    Match a TRADESNEW cargo reference to the contract list.

    Exact matching is attempted first.

    For PHYSICAL records, references may contain pricing-period
    suffixes. For example:

        YAMAL CY23-01-1
        YAMAL CY23-01-2

    These should match:

        YAMAL CY23-01

    References are therefore checked from longest to shortest.
    """
    reference_key = clean_text(cargo_reference)

    if reference_key == "":
        return None

    if reference_mapping.empty:
        return None

    exact_match = reference_mapping.loc[
        reference_mapping["REFERENCE_KEY"] == reference_key,
        "CONTRACT_FROM_LIST",
    ]

    if not exact_match.empty:
        return exact_match.iloc[0]

    for row in reference_mapping.itertuples(index=False):
        listed_reference = row.REFERENCE_KEY

        if reference_key.startswith(listed_reference + "-"):
            return row.CONTRACT_FROM_LIST

        if reference_key.startswith(listed_reference + " "):
            return row.CONTRACT_FROM_LIST

    return None


def classify_cargo_product(product):
    """Classify cargo products into the four requested sections."""
    product_key = clean_text(product)

    if product_key == "NWE":
        return "NWE"

    if product_key == "MED":
        return "MED"

    if product_key == "INDIA":
        return "India"

    if product_key == "JKTC":
        return "JKTC"

    return "Unclassified"


def classify_physical_product(product):
    """Classify physical products into HH, Oil, EU Gas or JKM."""
    product_key = clean_text(product)

    # HH section
    if product_key == "HH" or product_key.startswith("HH "):
        return "HH"

    # Oil section
    oil_terms = [
        "BFL",
        "JCC",
        "DATED BRENT",
        "BRENT",
        "DUBAI",
    ]

    if any(term in product_key for term in oil_terms):
        return "Oil"

    # EU Gas section
    eu_gas_terms = [
        "TTF",
        "THE",
        "PEG",
        "NBP",
        "ZTP",
    ]

    if any(term in product_key for term in eu_gas_terms):
        return "EU Gas"

    # JKM section
    if "JKM" in product_key:
        return "JKM"

    return "Unclassified"


def make_exposure_table(filtered_data, section_order):
    """
    Create rows for contract/product combinations with month columns,
    followed by a subtotal for each section and a grand total.
    """
    months = pd.date_range(
        start=f"{int(filtered_data['IFRS YEAR'].iloc[0])}-01-01",
        end=f"{int(filtered_data['IFRS YEAR'].iloc[0])}-12-01",
        freq="MS",
    )

    month_labels = [month.strftime("%b-%y") for month in months]

    filtered_data = filtered_data.copy()
    filtered_data["MONTH_LABEL"] = filtered_data[
        "EXPOSURE DATE"
    ].dt.strftime("%b-%y")

    detail = pd.pivot_table(
        filtered_data,
        index=["SECTION", "MATCHED CONTRACT", "PRODUCT"],
        columns="MONTH_LABEL",
        values="VOLUME_TBTU",
        aggfunc="sum",
        fill_value=0,
    )

    detail = detail.reindex(
        columns=month_labels,
        fill_value=0,
    ).reset_index()

    output_rows = []

    for section in section_order:
        section_rows = detail.loc[
            detail["SECTION"] == section
        ].copy()

        if section_rows.empty:
            subtotal = {
                "Section": section,
                "Contract": f"{section} TOTAL",
                "Product": "",
                "Row Type": "Subtotal",
            }

            for month in month_labels:
                subtotal[month] = 0.0

            subtotal["Total"] = 0.0
            output_rows.append(subtotal)
            continue

        section_rows = section_rows.sort_values(
            ["MATCHED CONTRACT", "PRODUCT"]
        )

        for _, row in section_rows.iterrows():
            output_row = {
                "Section": section,
                "Contract": row["MATCHED CONTRACT"],
                "Product": row["PRODUCT"],
                "Row Type": "Detail",
            }

            for month in month_labels:
                output_row[month] = row[month]

            output_row["Total"] = sum(
                output_row[month]
                for month in month_labels
            )

            output_rows.append(output_row)

        subtotal = {
            "Section": section,
            "Contract": f"{section} TOTAL",
            "Product": "",
            "Row Type": "Subtotal",
        }

        for month in month_labels:
            subtotal[month] = section_rows[month].sum()

        subtotal["Total"] = sum(
            subtotal[month]
            for month in month_labels
        )

        output_rows.append(subtotal)

    grand_total = {
        "Section": "",
        "Contract": "GRAND TOTAL",
        "Product": "",
        "Row Type": "Grand Total",
    }

    for month in month_labels:
        grand_total[month] = filtered_data.loc[
            filtered_data["MONTH_LABEL"] == month,
            "VOLUME_TBTU",
        ].sum()

    grand_total["Total"] = filtered_data["VOLUME_TBTU"].sum()
    output_rows.append(grand_total)

    return pd.DataFrame(output_rows), month_labels


def format_number(value):
    """Use dashes for zero and brackets for negative values."""
    if pd.isna(value) or abs(value) < 0.00001:
        return "-"

    if value < 0:
        return f"({abs(value):,.2f})"

    return f"{value:,.2f}"


def highlight_rows(row):
    """Apply formatting to subtotal and grand-total rows."""
    if row["Row Type"] == "Grand Total":
        return [
            "background-color: #0f766e; color: white; font-weight: bold"
        ] * len(row)

    if row["Row Type"] == "Subtotal":
        return [
            "background-color: #1f2937; color: white; font-weight: bold"
        ] * len(row)

    return [""] * len(row)


# ============================================================
# READ WORKBOOK
# ============================================================

try:
    trades, cargo_refs = read_workbook(uploaded_file)
except Exception as error:
    st.error(f"Failed to read workbook: {error}")
    st.stop()

trades.columns = [
    str(column).strip()
    for column in trades.columns
]

required_columns = [
    "FIRM",
    "TYPE",
    "CONTRACT",
    "CARGO#/TRADEID",
    "PRODUCT",
    "VOLUME",
    "EXPOSURE DATE",
    "IFRS YEAR",
]

missing_columns = [
    column
    for column in required_columns
    if column not in trades.columns
]

if missing_columns:
    st.error(f"Missing TRADESNEW columns: {missing_columns}")
    st.write("Columns found:")
    st.write(trades.columns.tolist())
    st.stop()


# ============================================================
# CLEAN DATA
# ============================================================

data = trades.copy()

data["TYPE"] = data["TYPE"].map(clean_text)
data["FIRM"] = data["FIRM"].map(clean_text)
data["PRODUCT"] = data["PRODUCT"].astype(str).str.strip()
data["CARGO#/TRADEID"] = (
    data["CARGO#/TRADEID"]
    .astype(str)
    .str.strip()
)
data["TRADESNEW CONTRACT"] = (
    data["CONTRACT"]
    .astype(str)
    .str.strip()
)

data["VOLUME"] = pd.to_numeric(
    data["VOLUME"],
    errors="coerce",
)

data["EXPOSURE DATE"] = pd.to_datetime(
    data["EXPOSURE DATE"],
    errors="coerce",
)

data["IFRS YEAR"] = pd.to_numeric(
    data["IFRS YEAR"],
    errors="coerce",
)

data = data.loc[
    data["TYPE"].isin(["CARGO", "PHYSICAL"])
    & data["VOLUME"].notna()
    & data["EXPOSURE DATE"].notna()
    & data["IFRS YEAR"].notna()
    & data["CARGO#/TRADEID"].notna()
].copy()

data = data.loc[
    data["CARGO#/TRADEID"].str.lower().ne("nan")
    & data["CARGO#/TRADEID"].ne("")
].copy()

# Convert the workbook volumes to TBtu.
data["VOLUME_TBTU"] = data["VOLUME"] / 1_000_000


# ============================================================
# REFERENCE TO CONTRACT MATCHING
# ============================================================

reference_mapping = build_reference_mapping(cargo_refs)

if reference_mapping.empty:
    st.error(
        "No cargo references could be read from the "
        "'All cargo refs' worksheet."
    )
    st.stop()

with st.spinner("Matching cargo references to contracts..."):
    data["MATCHED CONTRACT"] = data["CARGO#/TRADEID"].apply(
        lambda value: find_contract(
            value,
            reference_mapping,
        )
    )


# ============================================================
# FILTER CONTROLS
# ============================================================

control1, control2, control3 = st.columns([1, 1, 2])

with control1:
    selected_type = st.radio(
        "Exposure Type",
        options=["CARGO", "PHYSICAL"],
        horizontal=True,
    )

available_years = sorted(
    data["IFRS YEAR"]
    .dropna()
    .astype(int)
    .unique()
)

if not available_years:
    st.warning("No IFRS years were found.")
    st.stop()

with control2:
    selected_year = st.selectbox(
        "IFRS Year",
        options=available_years,
        index=len(available_years) - 1,
    )

with control3:
    contract_search = st.text_input(
        "Search contract or cargo reference",
        placeholder="Enter contract or cargo reference",
    )


# ============================================================
# APPLY FILTERS AND CLASSIFICATION
# ============================================================

filtered = data.loc[
    (data["TYPE"] == selected_type)
    & (data["IFRS YEAR"] == selected_year)
].copy()

if selected_type == "CARGO":
    filtered["SECTION"] = filtered["PRODUCT"].apply(
        classify_cargo_product
    )

    section_order = [
        "NWE",
        "MED",
        "India",
        "JKTC",
    ]

else:
    filtered["SECTION"] = filtered["PRODUCT"].apply(
        classify_physical_product
    )

    section_order = [
        "HH",
        "Oil",
        "EU Gas",
        "JKM",
    ]

if contract_search:
    search_key = clean_text(contract_search)

    filtered = filtered.loc[
        filtered["MATCHED CONTRACT"]
        .fillna("")
        .map(clean_text)
        .str.contains(search_key, regex=False)
        |
        filtered["CARGO#/TRADEID"]
        .fillna("")
        .map(clean_text)
        .str.contains(search_key, regex=False)
    ].copy()


# ============================================================
# MATCHING CONTROLS
# ============================================================

unmatched_references = filtered.loc[
    filtered["MATCHED CONTRACT"].isna(),
    "CARGO#/TRADEID",
].drop_duplicates()

unclassified_products = filtered.loc[
    filtered["SECTION"] == "Unclassified",
    "PRODUCT",
].drop_duplicates()

matched = filtered.loc[
    filtered["MATCHED CONTRACT"].notna()
    & filtered["SECTION"].isin(section_order)
].copy()

if matched.empty:
    st.warning(
        "No matching exposures were found for the selected "
        "type and IFRS year."
    )

    if not unmatched_references.empty:
        with st.expander("Unmatched cargo references"):
            st.dataframe(
                unmatched_references.to_frame(
                    name="Cargo Reference"
                ),
                use_container_width=True,
                hide_index=True,
            )

    st.stop()


# ============================================================
# KPIs
# ============================================================

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

kpi1.metric(
    "Contracts",
    f"{matched['MATCHED CONTRACT'].nunique():,}",
)

kpi2.metric(
    "Cargo References",
    f"{matched['CARGO#/TRADEID'].nunique():,}",
)

kpi3.metric(
    "Products",
    f"{matched['PRODUCT'].nunique():,}",
)

kpi4.metric(
    "Total Exposure",
    f"{matched['VOLUME_TBTU'].sum():,.2f} TBtu",
)

st.divider()


# ============================================================
# BUILD TABLE
# ============================================================

exposure_table, month_columns = make_exposure_table(
    matched,
    section_order,
)

st.subheader(
    f"{selected_type.title()} Contract Exposure by Product"
)

display_table = exposure_table.copy()

numeric_columns = month_columns + ["Total"]

for column in numeric_columns:
    display_table[column] = display_table[column].apply(
        format_number
    )

styled_table = (
    display_table.style
    .apply(highlight_rows, axis=1)
    .hide(axis="columns", subset=["Row Type"])
)

st.dataframe(
    styled_table,
    use_container_width=True,
    height=750,
    hide_index=True,
)


# ============================================================
# SECTION TOTALS
# ============================================================

st.subheader("Section Totals")

section_totals = (
    matched.groupby("SECTION", as_index=False)["VOLUME_TBTU"]
    .sum()
    .rename(
        columns={
            "SECTION": "Section",
            "VOLUME_TBTU": "Total Exposure (TBtu)",
        }
    )
)

section_totals["Section"] = pd.Categorical(
    section_totals["Section"],
    categories=section_order,
    ordered=True,
)

section_totals = section_totals.sort_values("Section")

st.dataframe(
    section_totals.style.format(
        {
            "Total Exposure (TBtu)": lambda x: format_number(x)
        }
    ),
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# DATA QUALITY WARNINGS
# ============================================================

if not unmatched_references.empty:
    st.warning(
        f"{len(unmatched_references):,} cargo references could not "
        "be matched to the 'All cargo refs' worksheet."
    )

    with st.expander("View unmatched cargo references"):
        unmatched_detail = filtered.loc[
            filtered["MATCHED CONTRACT"].isna(),
            [
                "CARGO#/TRADEID",
                "TRADESNEW CONTRACT",
                "TYPE",
                "PRODUCT",
                "EXPOSURE DATE",
                "VOLUME_TBTU",
            ],
        ].copy()

        st.dataframe(
            unmatched_detail,
            use_container_width=True,
            hide_index=True,
        )

if not unclassified_products.empty:
    st.warning(
        f"{len(unclassified_products):,} products could not be "
        f"classified into the requested {selected_type} sections."
    )

    with st.expander("View unclassified products"):
        st.dataframe(
            unclassified_products.to_frame(name="Product"),
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# EXCEL DOWNLOAD
# ============================================================

output = BytesIO()

export_table = exposure_table.drop(
    columns=["Row Type"]
).copy()

with pd.ExcelWriter(
    output,
    engine="xlsxwriter",
) as writer:
    export_table.to_excel(
        writer,
        sheet_name=f"{selected_type.title()} Exposure",
        index=False,
    )

    section_totals.to_excel(
        writer,
        sheet_name="Section Totals",
        index=False,
    )

    unmatched_export = filtered.loc[
        filtered["MATCHED CONTRACT"].isna()
    ].copy()

    if not unmatched_export.empty:
        unmatched_export.to_excel(
            writer,
            sheet_name="Unmatched References",
            index=False,
        )

st.download_button(
    label="📥 Download Exposure Table",
    data=output.getvalue(),
    file_name=(
        f"{selected_type.title()}_Contract_Exposure_"
        f"{int(selected_year)}.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)
