import streamlit as st
import pandas as pd
from io import BytesIO


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="LNG Contract Exposure Dashboard",
    page_icon="🚢",
    layout="wide",
)


# ============================================================
# STYLING
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


# ============================================================
# HEADER AND FILE UPLOAD
# ============================================================

st.title("LNG Contract Exposure Dashboard")

uploaded_file = st.file_uploader(
    "Upload COB Dashboard",
    type=["xlsx", "xlsm"],
)

if uploaded_file is None:
    st.info(
        "Upload the COB workbook to generate the exposure table."
    )
    st.stop()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text_series(series):
    """
    Standardise a pandas Series for matching.
    """
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


def clean_text(value):
    """
    Standardise an individual value for matching.
    """
    if pd.isna(value):
        return ""

    return " ".join(
        str(value).strip().upper().split()
    )


def match_contracts(
    reference_series,
    reference_dictionary,
):
    """
    Match cargo references using vectorised dictionary mapping.

    Matching order:
      1. Exact cargo reference
      2. Remove one final numeric suffix
      3. Remove two final numeric suffixes
      4. Remove three final numeric suffixes

    Example:

      YAMAL CY23-09-1

    can match:

      YAMAL CY23-09
    """
    reference_keys = clean_text_series(
        reference_series
    )

    matched = reference_keys.map(
        reference_dictionary
    )

    candidate = reference_keys.copy()

    for _ in range(3):
        candidate = candidate.str.replace(
            r"-\d+$",
            "",
            regex=True,
        )

        matched = matched.fillna(
            candidate.map(reference_dictionary)
        )

    return matched


@st.cache_data(show_spinner=False)
def read_workbook(file_bytes):
    """
    Read only the required workbook data.

    The result is cached so the workbook is not reopened
    every time a dashboard control changes.
    """
    trades = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="TRADESNEW",
        header=6,
        engine="openpyxl",
        usecols=[
            "FIRM",
            "TYPE",
            "CONTRACT",
            "CARGO#/TRADEID",
            "PRODUCT",
            "VOLUME",
            "EXPOSURE DATE",
            "IFRS YEAR",
        ],
    )

    cargo_refs = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name="All cargo refs",
        header=None,
        engine="openpyxl",
    )

    return trades, cargo_refs


@st.cache_data(show_spinner=False)
def build_reference_mapping(cargo_refs):
    """
    Convert the wide 'All cargo refs' worksheet into a
    cargo-reference-to-contract dictionary.

    Column B contains the contract.
    Columns C onward contain the cargo references.
    """
    mapping_rows = []

    for _, row in cargo_refs.iterrows():
        if len(row) < 3:
            continue

        contract = row.iloc[1]

        if pd.isna(contract):
            continue

        contract = str(contract).strip()

        if not contract:
            continue

        if clean_text(contract) == "CONTRACT":
            continue

        references = (
            row.iloc[2:]
            .dropna()
            .astype(str)
            .str.strip()
        )

        references = references.loc[
            references.ne("")
        ]

        for reference in references:
            mapping_rows.append(
                {
                    "REFERENCE_KEY": clean_text(reference),
                    "MATCHED CONTRACT": contract,
                }
            )

    mapping = pd.DataFrame(mapping_rows)

    if mapping.empty:
        return {}

    mapping = mapping.drop_duplicates(
        subset="REFERENCE_KEY",
        keep="first",
    )

    return dict(
        zip(
            mapping["REFERENCE_KEY"],
            mapping["MATCHED CONTRACT"],
        )
    )


@st.cache_data(show_spinner=False)
def prepare_data(trades, reference_dictionary):
    """
    Clean the data once and match cargo references to contracts.
    """
    data = trades.copy()

    data.columns = [
        str(column).strip()
        for column in data.columns
    ]

    data["FIRM"] = clean_text_series(
        data["FIRM"]
    )

    data["TYPE"] = clean_text_series(
        data["TYPE"]
    )

    data["PRODUCT"] = (
        data["PRODUCT"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data["CARGO#/TRADEID"] = (
        data["CARGO#/TRADEID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data["TRADESNEW CONTRACT"] = (
        data["CONTRACT"]
        .fillna("")
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
        data["TYPE"].isin(
            [
                "CARGO",
                "PHYSICAL",
                "FINANCIAL",
            ]
        )
        & data["VOLUME"].notna()
        & data["EXPOSURE DATE"].notna()
        & data["IFRS YEAR"].notna()
        & data["CARGO#/TRADEID"].ne("")
    ].copy()

    data["IFRS YEAR"] = (
        data["IFRS YEAR"]
        .astype(int)
    )

    # Convert volume to TBtu.
    data["VOLUME_TBTU"] = (
        data["VOLUME"] / 1_000_000
    )

    data["MATCHED CONTRACT"] = match_contracts(
        data["CARGO#/TRADEID"],
        reference_dictionary,
    )

    return data


def classify_products(
    product_series,
    exposure_type,
):
    """
    Classify products into the relevant sections.
    """
    product_keys = clean_text_series(
        product_series
    )

    section = pd.Series(
        "Unclassified",
        index=product_series.index,
        dtype="object",
    )

    if exposure_type == "CARGO":
        section.loc[
            product_keys.eq("NWE")
        ] = "NWE"

        section.loc[
            product_keys.eq("MED")
        ] = "MED"

        section.loc[
            product_keys.eq("INDIA")
        ] = "India"

        section.loc[
            product_keys.eq("JKTC")
        ] = "JKTC"

        section_order = [
            "NWE",
            "MED",
            "India",
            "JKTC",
        ]

    else:
        # Henry Hub
        hh_mask = (
            product_keys.eq("HH")
            | product_keys.str.startswith("HH ")
        )

        section.loc[
            hh_mask
        ] = "HH"

        # Oil products
        oil_mask = product_keys.str.contains(
            r"BFL|JCC|DATED BRENT|BRENT|DUBAI",
            regex=True,
            na=False,
        )

        section.loc[
            oil_mask
        ] = "Oil"

        # European gas products
        eu_gas_mask = product_keys.str.contains(
            r"TTF|THE|PEG|NBP|ZTP",
            regex=True,
            na=False,
        )

        section.loc[
            eu_gas_mask
        ] = "EU Gas"

        # JKM products
        jkm_mask = product_keys.str.contains(
            "JKM",
            regex=False,
            na=False,
        )

        section.loc[
            jkm_mask
        ] = "JKM"

        section_order = [
            "HH",
            "Oil",
            "EU Gas",
            "JKM",
        ]

    return section, section_order


def format_number(value):
    """
    Display zero as a dash and negative numbers in brackets.
    """
    if pd.isna(value):
        return "-"

    if abs(value) < 0.00001:
        return "-"

    if value < 0:
        return f"({abs(value):,.2f})"

    return f"{value:,.2f}"


def build_exposure_table(
    filtered_data,
    section_order,
    selected_year,
):
    """
    Build the contract, product and monthly exposure matrix.

    All Years:
      Columns are Jan-26, Feb-26 and so on.

    Individual year:
      Columns are Jan, Feb and so on.

    Each populated section includes a subtotal row.
    """
    table_data = filtered_data.copy()

    if selected_year == "All Years":
        minimum_date = (
            table_data["EXPOSURE DATE"]
            .min()
            .to_period("M")
            .to_timestamp()
        )

        maximum_date = (
            table_data["EXPOSURE DATE"]
            .max()
            .to_period("M")
            .to_timestamp()
        )

        months = pd.date_range(
            start=minimum_date,
            end=maximum_date,
            freq="MS",
        )

        table_data["MONTH_LABEL"] = (
            table_data["EXPOSURE DATE"]
            .dt.strftime("%b-%y")
        )

        month_labels = [
            month.strftime("%b-%y")
            for month in months
        ]

    else:
        selected_year = int(selected_year)

        months = pd.date_range(
            start=f"{selected_year}-01-01",
            end=f"{selected_year}-12-01",
            freq="MS",
        )

        table_data["MONTH_LABEL"] = (
            table_data["EXPOSURE DATE"]
            .dt.strftime("%b")
        )

        month_labels = [
            month.strftime("%b")
            for month in months
        ]

    detail = pd.pivot_table(
        table_data,
        index=[
            "SECTION",
            "MATCHED CONTRACT",
            "PRODUCT",
        ],
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
    subtotal_row_indexes = []

    for section_name in section_order:
        section_data = detail.loc[
            detail["SECTION"] == section_name
        ].copy()

        # Do not display empty sections.
        if section_data.empty:
            continue

        section_data = section_data.sort_values(
            by=[
                "MATCHED CONTRACT",
                "PRODUCT",
            ]
        )

        for _, row in section_data.iterrows():
            output_row = {
                "Section": section_name,
                "Contract": row["MATCHED CONTRACT"],
                "Product": row["PRODUCT"],
            }

            for month in month_labels:
                output_row[month] = row[month]

            output_rows.append(
                output_row
            )

        subtotal_row = {
            "Section": section_name,
            "Contract": f"{section_name} TOTAL",
            "Product": "",
        }

        for month in month_labels:
            subtotal_row[month] = (
                section_data[month].sum()
            )

        subtotal_row_indexes.append(
            len(output_rows)
        )

        output_rows.append(
            subtotal_row
        )

    output = pd.DataFrame(
        output_rows
    )

    return (
        output,
        month_labels,
        subtotal_row_indexes,
    )


def style_table(
    display_df,
    subtotal_rows,
):
    """
    Highlight subtotal rows without adding a Row Type column.
    """
    def apply_row_style(row):
        if row.name in subtotal_rows:
            return [
                (
                    "background-color: #1f2937; "
                    "color: white; "
                    "font-weight: bold; "
                    "border-top: 1px solid #6b7280;"
                )
            ] * len(row)

        return [""] * len(row)

    return display_df.style.apply(
        apply_row_style,
        axis=1,
    )


# ============================================================
# LOAD WORKBOOK
# ============================================================

file_bytes = uploaded_file.getvalue()

try:
    with st.spinner("Loading workbook..."):
        trades, cargo_refs = read_workbook(
            file_bytes
        )

except Exception as error:
    st.error(
        f"Failed to read workbook: {error}"
    )
    st.stop()


# ============================================================
# VALIDATE REQUIRED COLUMNS
# ============================================================

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
    st.error(
        f"Missing TRADESNEW columns: {missing_columns}"
    )

    st.write("Columns found:")
    st.write(trades.columns.tolist())

    st.stop()


# ============================================================
# BUILD CONTRACT MAPPING
# ============================================================

reference_dictionary = build_reference_mapping(
    cargo_refs
)

if not reference_dictionary:
    st.error(
        "No cargo references could be read from "
        "the 'All cargo refs' worksheet."
    )
    st.stop()


# ============================================================
# PREPARE DATA
# ============================================================

with st.spinner(
    "Cleaning data and matching cargo references..."
):
    data = prepare_data(
        trades,
        reference_dictionary,
    )


# ============================================================
# MAIN CONTROLS
# ============================================================

control1, control2, control3 = st.columns(
    [1.2, 1, 2]
)

with control1:
    selected_type = st.radio(
        "Exposure Type",
        options=[
            "CARGO",
            "PHYSICAL",
            "FINANCIAL",
        ],
        horizontal=True,
    )

available_years = sorted(
    data["IFRS YEAR"]
    .dropna()
    .unique()
    .tolist()
)

year_options = (
    ["All Years"]
    + available_years
)

with control2:
    selected_year = st.selectbox(
        "Year View",
        options=year_options,
        index=0,
    )

with control3:
    search = st.text_input(
        "Search contract, reference or product",
        placeholder="Enter search text",
    )


# ============================================================
# FILTER BY EXPOSURE TYPE AND YEAR
# ============================================================

filtered_before_contracts = data.loc[
    data["TYPE"] == selected_type
].copy()

if selected_year != "All Years":
    filtered_before_contracts = (
        filtered_before_contracts.loc[
            filtered_before_contracts["IFRS YEAR"]
            == int(selected_year)
        ].copy()
    )


# ============================================================
# CLASSIFY PRODUCTS
# ============================================================

(
    filtered_before_contracts["SECTION"],
    section_order,
) = classify_products(
    filtered_before_contracts["PRODUCT"],
    selected_type,
)


# ============================================================
# CONTRACT MULTISELECT
# ============================================================

available_contracts = sorted(
    filtered_before_contracts[
        "MATCHED CONTRACT"
    ]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

if not available_contracts:
    st.warning(
        "No contracts could be matched for this selection."
    )

    unmatched_selection = (
        filtered_before_contracts.loc[
            filtered_before_contracts[
                "MATCHED CONTRACT"
            ].isna()
        ]
    )

    if not unmatched_selection.empty:
        with st.expander(
            "View unmatched cargo references"
        ):
            st.dataframe(
                unmatched_selection[
                    [
                        "CARGO#/TRADEID",
                        "TRADESNEW CONTRACT",
                        "TYPE",
                        "PRODUCT",
                        "EXPOSURE DATE",
                        "IFRS YEAR",
                        "VOLUME_TBTU",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

    st.stop()


# Reset contract choices when the file, type or year changes.
contract_context = (
    f"{selected_type}_{selected_year}_{uploaded_file.name}"
)

if (
    "last_contract_context" not in st.session_state
    or st.session_state["last_contract_context"]
    != contract_context
):
    st.session_state["contract_selection"] = (
        available_contracts.copy()
    )

    st.session_state["last_contract_context"] = (
        contract_context
    )


# Remove selections that are not available in the current context.
st.session_state["contract_selection"] = [
    contract
    for contract in st.session_state.get(
        "contract_selection",
        [],
    )
    if contract in available_contracts
]


st.subheader("Contract Selection")

button1, button2, button3 = st.columns(
    [1, 1, 4]
)

with button1:
    if st.button(
        "Select all",
        use_container_width=True,
    ):
        st.session_state["contract_selection"] = (
            available_contracts.copy()
        )

        st.rerun()

with button2:
    if st.button(
        "Clear all",
        use_container_width=True,
    ):
        st.session_state["contract_selection"] = []

        st.rerun()


selected_contracts = st.multiselect(
    "Contracts",
    options=available_contracts,
    key="contract_selection",
    placeholder="Select contracts to display",
    help=(
        "All contracts are selected by default. "
        "Remove a contract to switch it off. "
        "Add it back to switch it on."
    ),
)

if not selected_contracts:
    st.warning(
        "No contracts are currently selected. "
        "Select at least one contract to display the table."
    )

    st.stop()


# ============================================================
# APPLY CONTRACT FILTER
# ============================================================

filtered = filtered_before_contracts.loc[
    filtered_before_contracts[
        "MATCHED CONTRACT"
    ].isin(selected_contracts)
].copy()


# ============================================================
# SEARCH FILTER
# ============================================================

if search:
    search_key = clean_text(search)

    contract_match = (
        clean_text_series(
            filtered["MATCHED CONTRACT"]
        )
        .str.contains(
            search_key,
            regex=False,
            na=False,
        )
    )

    reference_match = (
        clean_text_series(
            filtered["CARGO#/TRADEID"]
        )
        .str.contains(
            search_key,
            regex=False,
            na=False,
        )
    )

    product_match = (
        clean_text_series(
            filtered["PRODUCT"]
        )
        .str.contains(
            search_key,
            regex=False,
            na=False,
        )
    )

    filtered = filtered.loc[
        contract_match
        | reference_match
        | product_match
    ].copy()


# ============================================================
# DATA QUALITY
# ============================================================

unmatched_data = (
    filtered_before_contracts.loc[
        filtered_before_contracts[
            "MATCHED CONTRACT"
        ].isna()
    ].copy()
)

unclassified_data = (
    filtered_before_contracts.loc[
        filtered_before_contracts[
            "MATCHED CONTRACT"
        ].notna()
        & filtered_before_contracts[
            "SECTION"
        ].eq("Unclassified")
    ].copy()
)

matched = filtered.loc[
    filtered["MATCHED CONTRACT"].notna()
    & filtered["SECTION"].isin(
        section_order
    )
].copy()

if matched.empty:
    st.warning(
        "No matched exposure records were found "
        "for the selected filters."
    )

    st.stop()


# ============================================================
# KPIs
# ============================================================

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

kpi1.metric(
    "Selected Contracts",
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
    "Net Exposure",
    f"{matched['VOLUME_TBTU'].sum():,.2f} TBtu",
)

st.divider()


# ============================================================
# BUILD EXPOSURE TABLE
# ============================================================

(
    exposure_table,
    month_columns,
    subtotal_rows,
) = build_exposure_table(
    matched,
    section_order,
    selected_year,
)

view_name = (
    "All Years"
    if selected_year == "All Years"
    else str(selected_year)
)

st.subheader(
    f"{selected_type.title()} Exposure: {view_name}"
)


# ============================================================
# FORMAT AND DISPLAY TABLE
# ============================================================

display_table = exposure_table.copy()

for column in month_columns:
    display_table[column] = (
        display_table[column]
        .apply(format_number)
    )

styled_table = style_table(
    display_table,
    subtotal_rows,
)

st.dataframe(
    styled_table,
    use_container_width=True,
    height=750,
    hide_index=True,
)


# ============================================================
# DATA QUALITY WARNINGS
# ============================================================

if not unmatched_data.empty:
    unmatched_count = (
        unmatched_data["CARGO#/TRADEID"]
        .nunique()
    )

    st.warning(
        f"{unmatched_count:,} cargo references could not "
        "be matched to the 'All cargo refs' worksheet."
    )

    with st.expander(
        "View unmatched cargo references"
    ):
        st.dataframe(
            unmatched_data[
                [
                    "CARGO#/TRADEID",
                    "TRADESNEW CONTRACT",
                    "TYPE",
                    "PRODUCT",
                    "EXPOSURE DATE",
                    "IFRS YEAR",
                    "VOLUME_TBTU",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


if not unclassified_data.empty:
    unclassified_count = (
        unclassified_data["PRODUCT"]
        .nunique()
    )

    st.warning(
        f"{unclassified_count:,} products could not be "
        f"assigned to a {selected_type} section."
    )

    with st.expander(
        "View unclassified products"
    ):
        st.dataframe(
            unclassified_data[
                [
                    "PRODUCT",
                    "CARGO#/TRADEID",
                    "MATCHED CONTRACT",
                    "IFRS YEAR",
                    "EXPOSURE DATE",
                    "VOLUME_TBTU",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# EXCEL DOWNLOAD
# ============================================================

output = BytesIO()

with pd.ExcelWriter(
    output,
    engine="xlsxwriter",
) as writer:
    exposure_table.to_excel(
        writer,
        sheet_name="Exposure",
        index=False,
    )

    workbook = writer.book
    worksheet = writer.sheets["Exposure"]

    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#111827",
            "border": 0,
        }
    )

    number_format = workbook.add_format(
        {
            "num_format": (
                '#,##0.00;#,##0.00;-'
            ),
        }
    )

    subtotal_text_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F2937",
        }
    )

    subtotal_number_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F2937",
            "num_format": (
                '#,##0.00;#,##0.00;-'
            ),
        }
    )

    for column_number, column_name in enumerate(
        exposure_table.columns
    ):
        worksheet.write(
            0,
            column_number,
            column_name,
            header_format,
        )

    worksheet.set_column(
        0,
        0,
        14,
    )

    worksheet.set_column(
        1,
        1,
        30,
    )

    worksheet.set_column(
        2,
        2,
        20,
    )

    if month_columns:
        first_month_column = 3

        last_month_column = (
            first_month_column
            + len(month_columns)
            - 1
        )

        worksheet.set_column(
            first_month_column,
            last_month_column,
            12,
            number_format,
        )

    for subtotal_index in subtotal_rows:
        excel_row = subtotal_index + 1

        for column_number in range(
            len(exposure_table.columns)
        ):
            value = exposure_table.iloc[
                subtotal_index,
                column_number,
            ]

            if column_number < 3:
                worksheet.write(
                    excel_row,
                    column_number,
                    value,
                    subtotal_text_format,
                )

            else:
                worksheet.write_number(
                    excel_row,
                    column_number,
                    float(value),
                    subtotal_number_format,
                )

    worksheet.freeze_panes(
        1,
        3,
    )

    if not unmatched_data.empty:
        unmatched_export = unmatched_data[
            [
                "CARGO#/TRADEID",
                "TRADESNEW CONTRACT",
                "TYPE",
                "PRODUCT",
                "EXPOSURE DATE",
                "IFRS YEAR",
                "VOLUME_TBTU",
            ]
        ].copy()

        unmatched_export.to_excel(
            writer,
            sheet_name="Unmatched References",
            index=False,
        )

    if not unclassified_data.empty:
        unclassified_export = unclassified_data[
            [
                "PRODUCT",
                "CARGO#/TRADEID",
                "MATCHED CONTRACT",
                "IFRS YEAR",
                "EXPOSURE DATE",
                "VOLUME_TBTU",
            ]
        ].copy()

        unclassified_export.to_excel(
            writer,
            sheet_name="Unclassified Products",
            index=False,
        )


# ============================================================
# DOWNLOAD BUTTON
# ============================================================

download_year = (
    "All_Years"
    if selected_year == "All Years"
    else str(selected_year)
)

st.download_button(
    label="📥 Download Exposure Table",
    data=output.getvalue(),
    file_name=(
        f"{selected_type.title()}_Contract_Exposure_"
        f"{download_year}.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)
