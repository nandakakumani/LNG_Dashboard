import streamlit as st
import pandas as pd
from io import BytesIO

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="SEFE LNG Exposure Dashboard",
    page_icon="",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem;}
    h1, h2, h3 {color: white;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SEFE LNG Exposure Dashboard")
uploaded_file = st.file_uploader(
    "Upload Master Dashboard",
    type=["xlsx", "xlsm"],
)

if uploaded_file is None:
    st.info("Upload the Master Dashboard to generate the exposure table.")
    st.stop()

# ============================================================
# TEXT CLEANING
# ============================================================
def clean_text_series(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


def clean_text(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


# ============================================================
# WORKBOOK READING
# ============================================================
@st.cache_data(show_spinner=False)
def read_workbook(file_bytes):
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


# ============================================================
# CONTRACT REFERENCE MAPPING
# ============================================================
@st.cache_data(show_spinner=False)
def build_reference_mapping(cargo_refs):
    mapping_rows = []

    for _, row in cargo_refs.iterrows():
        if len(row) < 3:
            continue

        contract = row.iloc[1]
        if pd.isna(contract):
            continue

        contract = str(contract).strip()
        if not contract or clean_text(contract) == "CONTRACT":
            continue

        references = row.iloc[2:].dropna().astype(str).str.strip()
        references = references.loc[references.ne("")]

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


def match_contracts(reference_series, reference_dictionary):
    reference_keys = clean_text_series(reference_series)
    matched = reference_keys.map(reference_dictionary)
    candidate = reference_keys.copy()

    for _ in range(3):
        candidate = candidate.str.replace(r"-\d+$", "", regex=True)
        matched = matched.fillna(candidate.map(reference_dictionary))

    return matched


# ============================================================
# DATA PREPARATION
# ============================================================
@st.cache_data(show_spinner=False)
def prepare_data(trades, reference_dictionary):
    data = trades.copy()
    data.columns = [str(column).strip() for column in data.columns]

    data["FIRM"] = clean_text_series(data["FIRM"])
    data["TYPE"] = clean_text_series(data["TYPE"])
    data["PRODUCT"] = data["PRODUCT"].fillna("").astype(str).str.strip()
    data["CARGO#/TRADEID"] = (
        data["CARGO#/TRADEID"].fillna("").astype(str).str.strip()
    )
    data["TRADESNEW CONTRACT"] = (
        data["CONTRACT"].fillna("").astype(str).str.strip()
    )
    data["VOLUME"] = pd.to_numeric(data["VOLUME"], errors="coerce")
    data["EXPOSURE DATE"] = pd.to_datetime(
        data["EXPOSURE DATE"],
        errors="coerce",
    )
    data["IFRS YEAR"] = pd.to_numeric(data["IFRS YEAR"], errors="coerce")

    data = data.loc[
        data["TYPE"].isin(["CARGO", "PHYSICAL", "FINANCIAL"])
        & data["VOLUME"].notna()
        & data["EXPOSURE DATE"].notna()
        & data["IFRS YEAR"].notna()
        & data["CARGO#/TRADEID"].ne("")
    ].copy()

    data["IFRS YEAR"] = data["IFRS YEAR"].astype(int)
    data["EXPOSURE DATE YEAR"] = data["EXPOSURE DATE"].dt.year.astype(int)
    data["YEAR MISMATCH"] = (
        data["EXPOSURE DATE YEAR"] != data["IFRS YEAR"]
    )

    data["VOLUME_TBTU"] = data["VOLUME"] / 1_000_000
    data["MATCHED CONTRACT"] = match_contracts(
        data["CARGO#/TRADEID"],
        reference_dictionary,
    )

    return data


# ============================================================
# PRODUCT CLASSIFICATION
# ============================================================
def classify_products(product_series, exposure_type):
    product_keys = clean_text_series(product_series)
    section = pd.Series(
        "Unclassified",
        index=product_series.index,
        dtype="object",
    )

    if exposure_type == "CARGO":
        section.loc[product_keys.eq("NWE")] = "NWE"
        section.loc[product_keys.eq("MED")] = "MED"
        section.loc[product_keys.eq("INDIA")] = "India"
        section.loc[product_keys.eq("JKTC")] = "JKTC"
        section_order = ["NWE", "MED", "India", "JKTC"]
    else:
        hh_mask = (
            product_keys.eq("HH")
            | product_keys.str.startswith("HH ")
            | product_keys.str.startswith("HH_")
        )
        section.loc[hh_mask] = "HH"

        oil_mask = product_keys.str.contains(
            r"\bBFL\b|\bJCC\b|DATED BRENT|BRENT FUTURES|"
            r"\bBRENT\b|\bDUBAI\b|\bDFL\b",
            regex=True,
            na=False,
        )
        section.loc[oil_mask] = "Oil"

        eugas_mask = product_keys.str.contains(
            r"\bTTF\b|\bTHE\b|\bPEG\b|\bNBP\b|\bZTP\b",
            regex=True,
            na=False,
        )
        section.loc[eugas_mask] = "EUGAS"

        jkm_mask = product_keys.str.contains(
            r"\bJKM\b",
            regex=True,
            na=False,
        )
        section.loc[jkm_mask] = "JKM"
        section_order = ["HH", "Oil", "EUGAS", "JKM"]

    return section, section_order


def normalise_display_products(product_series, section_series):
    display_product = product_series.fillna("").astype(str).str.strip()
    eugas_mask = section_series.eq("EUGAS")

    display_product.loc[eugas_mask] = (
        display_product.loc[eugas_mask]
        .str.extract(r"^([A-Za-z]+)", expand=False)
        .fillna(display_product.loc[eugas_mask])
        .str.upper()
    )

    return display_product


# ============================================================
# NUMBER FORMATTING
# ============================================================
def format_number(value):
    if pd.isna(value) or abs(value) < 0.00001:
        return "-"
    if value < 0:
        return f"({abs(value):,.2f})"
    return f"{value:,.2f}"


# ============================================================
# BUILD EXPOSURE TABLE
# ============================================================
def build_exposure_table(
    filtered_data,
    section_order,
    selected_year,
):
    """
    All Years view:
        Columns are IFRS years.

    Individual IFRS year view:
        Rows are filtered by IFRS YEAR, but columns are the actual
        month and year from EXPOSURE DATE. This means IFRS 2027 can
        show columns such as Dec-26, Jan-27, and Jan-28 when those
        exposure dates exist in the selected IFRS year.
    """
    table_data = filtered_data.copy()

    if selected_year == "All Years":
        source_columns = sorted(
            table_data["IFRS YEAR"]
            .dropna()
            .astype(int)
            .unique()
            .tolist()
        )
        column_field = "IFRS YEAR"
        display_column_map = {
            year: str(year) for year in source_columns
        }
    else:
        selected_year_number = int(selected_year)
        table_data = table_data.loc[
            table_data["IFRS YEAR"] == selected_year_number
        ].copy()

        # Keep the complete exposure month and year, not just month number.
        table_data["EXPOSURE MONTH"] = (
            table_data["EXPOSURE DATE"].dt.to_period("M").dt.to_timestamp()
        )

        # Only display actual exposure months found for the selected IFRS year.
        # The columns are sorted chronologically and may fall outside that year.
        source_columns = sorted(
            table_data["EXPOSURE MONTH"]
            .dropna()
            .unique()
            .tolist()
        )
        column_field = "EXPOSURE MONTH"
        display_column_map = {
            month: pd.Timestamp(month).strftime("%b-%y")
            for month in source_columns
        }

    detail = pd.pivot_table(
        table_data,
        index=["SECTION", "MATCHED CONTRACT", "DISPLAY PRODUCT"],
        columns=column_field,
        values="VOLUME_TBTU",
        aggfunc="sum",
        fill_value=0,
    )
    detail = detail.reindex(
        columns=source_columns,
        fill_value=0,
    ).reset_index()

    output_rows = []
    subtotal_row_indexes = []

    for section_name in section_order:
        section_data = detail.loc[
            detail["SECTION"] == section_name
        ].copy()
        if section_data.empty:
            continue

        section_data = section_data.sort_values(
            by=["MATCHED CONTRACT", "DISPLAY PRODUCT"]
        )

        for _, row in section_data.iterrows():
            output_row = {
                "Section": section_name,
                "Contract": row["MATCHED CONTRACT"],
                "Product": row["DISPLAY PRODUCT"],
            }
            for source_column in source_columns:
                output_row[display_column_map[source_column]] = row[source_column]
            output_rows.append(output_row)

        subtotal_row = {
            "Section": section_name,
            "Contract": f"{section_name} TOTAL",
            "Product": "",
        }
        for source_column in source_columns:
            subtotal_row[display_column_map[source_column]] = (
                section_data[source_column].sum()
            )
        subtotal_row_indexes.append(len(output_rows))
        output_rows.append(subtotal_row)

    output = pd.DataFrame(output_rows)
    value_columns = [
        display_column_map[column] for column in source_columns
    ]
    return output, value_columns, subtotal_row_indexes
# ============================================================
# TABLE STYLING
# ============================================================
def style_table(numeric_df, display_df, subtotal_rows, value_columns):
    def apply_row_style(row):
        styles = []

        for column in display_df.columns:
            cell_style = "text-align: center; vertical-align: middle;"
            is_subtotal = row.name in subtotal_rows
            is_negative_number = (
                column in value_columns
                and pd.notna(numeric_df.loc[row.name, column])
                and numeric_df.loc[row.name, column] < 0
            )

            if is_subtotal:
                cell_style += (
                    "background-color: #1f2937; "
                    "font-weight: bold; "
                    "border-top: 2px solid #64748b; "
                    "border-bottom: 1px solid #64748b;"
                )
                cell_style += (
                    "color: #ff4b4b;"
                    if is_negative_number
                    else "color: white;"
                )
            elif is_negative_number:
                cell_style += "color: #dc2626; font-weight: 600;"

            styles.append(cell_style)

        return styles

    styled = display_df.style.apply(apply_row_style, axis=1)
    styled = styled.set_table_styles(
        [
            {
                "selector": "thead th",
                "props": [
                    ("background-color", "#111827"),
                    ("color", "#ffffff"),
                    ("font-weight", "700"),
                    ("text-align", "center"),
                    ("vertical-align", "middle"),
                    ("border-bottom", "2px solid #475569"),
                    ("padding", "10px"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("text-align", "center"),
                    ("vertical-align", "middle"),
                    ("padding", "8px"),
                ],
            },
            {"selector": "th.row_heading", "props": [("display", "none")]},
            {"selector": "th.blank", "props": [("display", "none")]},
        ]
    )

    return styled


# ============================================================
# LOAD AND VALIDATE WORKBOOK
# ============================================================
file_bytes = uploaded_file.getvalue()

try:
    with st.spinner("Loading workbook..."):
        trades, cargo_refs = read_workbook(file_bytes)
except Exception as error:
    st.error(f"Failed to read DB: {error}")
    st.stop()

trades.columns = [str(column).strip() for column in trades.columns]
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
    column for column in required_columns if column not in trades.columns
]

if missing_columns:
    st.error(f"Missing TRADESNEW columns: {missing_columns}")
    st.write("Columns found:")
    st.write(trades.columns.tolist())
    st.stop()

reference_dictionary = build_reference_mapping(cargo_refs)
if not reference_dictionary:
    st.error(
        "No cargo references could be read from the "
        "'All cargo refs' worksheet."
    )
    st.stop()

with st.spinner("Cleaning and matching cargo references..."):
    data = prepare_data(trades, reference_dictionary)

# ============================================================
# MAIN CONTROLS
# ============================================================
control1, control2 = st.columns([1.2, 1])
with control1:
    selected_type = st.radio(
        "Exposure Type",
        options=["CARGO", "PHYSICAL", "FINANCIAL"],
        horizontal=True,
    )

# The dashboard is always controlled by IFRS YEAR.
available_years = sorted(
    data["IFRS YEAR"].dropna().astype(int).unique().tolist()
)
year_options = ["All Years"] + available_years

with control2:
    selected_year = st.selectbox(
        "IFRS Year View",
        options=year_options,
        index=0,
        help=(
            "Select an IFRS year. When you drill into a year, the columns "
            "show every actual EXPOSURE DATE month found in that IFRS year, "
            "including months in adjacent calendar years."
        ),
    )

search = st.text_input(
    "Search contract, reference or product",
    placeholder="Enter search text",
)

# ============================================================
# YEAR MISMATCH CHECK
# ============================================================
year_mismatches = data.loc[data["YEAR MISMATCH"]].copy()

if not year_mismatches.empty:
    mismatch_rows = len(year_mismatches)
    mismatch_references = year_mismatches["CARGO#/TRADEID"].nunique()

    st.warning(
        f"{mismatch_rows:,} exposure rows across "
        f"{mismatch_references:,} unique references have an EXPOSURE DATE "
        "year that differs from IFRS YEAR. The exposure remains allocated "
        "to its IFRS YEAR, while the drill-down column shows the actual "
        "EXPOSURE DATE month and year."
    )

    with st.expander("View exposure year mismatches"):
        st.dataframe(
            year_mismatches[
                [
                    "CARGO#/TRADEID",
                    "TRADESNEW CONTRACT",
                    "MATCHED CONTRACT",
                    "TYPE",
                    "PRODUCT",
                    "EXPOSURE DATE",
                    "EXPOSURE DATE YEAR",
                    "IFRS YEAR",
                    "VOLUME_TBTU",
                ]
            ].sort_values(
                by=["IFRS YEAR", "EXPOSURE DATE", "CARGO#/TRADEID"]
            ),
            use_container_width=True,
            hide_index=True,
        )

# ============================================================
# FILTER BY TYPE AND IFRS YEAR
# ============================================================
filtered_before_contracts = data.loc[
    data["TYPE"] == selected_type
].copy()

if selected_year != "All Years":
    filtered_before_contracts = filtered_before_contracts.loc[
        filtered_before_contracts["IFRS YEAR"] == int(selected_year)
    ].copy()

# ============================================================
# CLASSIFY AND NORMALISE PRODUCTS
# ============================================================
(
    filtered_before_contracts["SECTION"],
    section_order,
) = classify_products(
    filtered_before_contracts["PRODUCT"],
    selected_type,
)

filtered_before_contracts["DISPLAY PRODUCT"] = normalise_display_products(
    filtered_before_contracts["PRODUCT"],
    filtered_before_contracts["SECTION"],
)

# ============================================================
# CONTRACT OPTIONS AND SESSION STATE
# ============================================================
available_contracts = sorted(
    filtered_before_contracts["MATCHED CONTRACT"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

if not available_contracts:
    st.warning("No contracts could be matched for this selection.")
    unmatched_selection = filtered_before_contracts.loc[
        filtered_before_contracts["MATCHED CONTRACT"].isna()
    ].copy()

    if not unmatched_selection.empty:
        with st.expander("View unmatched cargo references"):
            st.dataframe(
                unmatched_selection[
                    [
                        "CARGO#/TRADEID",
                        "TRADESNEW CONTRACT",
                        "TYPE",
                        "PRODUCT",
                        "EXPOSURE DATE",
                        "EXPOSURE DATE YEAR",
                        "IFRS YEAR",
                        "VOLUME_TBTU",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )
    st.stop()

contract_context = (
    f"{selected_type}_{selected_year}_IFRS_"
    f"{uploaded_file.name}_{len(file_bytes)}"
)

if st.session_state.get("last_contract_context") != contract_context:
    st.session_state["contract_selection"] = available_contracts.copy()
    st.session_state["last_contract_context"] = contract_context
    st.session_state.pop("product_selection", None)
    st.session_state.pop("last_product_context", None)
else:
    st.session_state["contract_selection"] = [
        contract
        for contract in st.session_state.get("contract_selection", [])
        if contract in available_contracts
    ]

st.subheader("Contract Selection")
contract_button1, contract_button2, _ = st.columns([1, 1, 4])

with contract_button1:
    if st.button("Select all contracts", use_container_width=True):
        st.session_state["contract_selection"] = available_contracts.copy()
        st.rerun()

with contract_button2:
    if st.button("Clear all contracts", use_container_width=True):
        st.session_state["contract_selection"] = []
        st.rerun()

selected_contracts = st.multiselect(
    "Contracts",
    options=available_contracts,
    key="contract_selection",
    placeholder="Select contracts to display",
    help="All contracts are selected by default. Remove a contract to exclude it.",
)

if not selected_contracts:
    st.warning("No contracts are selected. Select at least one contract.")
    st.stop()

filtered_after_contracts = filtered_before_contracts.loc[
    filtered_before_contracts["MATCHED CONTRACT"].isin(selected_contracts)
].copy()

# ============================================================
# PRODUCT OPTIONS AND SESSION STATE
# ============================================================
available_products = sorted(
    filtered_after_contracts["DISPLAY PRODUCT"]
    .dropna()
    .astype(str)
    .loc[lambda values: values.str.strip().ne("")]
    .unique()
    .tolist()
)

if not available_products:
    st.warning("No classified products were found for the selected contracts.")
    st.stop()

product_context = (
    f"{contract_context}_{'|'.join(sorted(selected_contracts))}"
)

if st.session_state.get("last_product_context") != product_context:
    st.session_state["product_selection"] = available_products.copy()
    st.session_state["last_product_context"] = product_context
else:
    st.session_state["product_selection"] = [
        product
        for product in st.session_state.get("product_selection", [])
        if product in available_products
    ]

st.subheader("Product Selection")
product_button1, product_button2, _ = st.columns([1, 1, 4])

with product_button1:
    if st.button("Select all products", use_container_width=True):
        st.session_state["product_selection"] = available_products.copy()
        st.rerun()

with product_button2:
    if st.button("Clear all products", use_container_width=True):
        st.session_state["product_selection"] = []
        st.rerun()

selected_products = st.multiselect(
    "Products",
    options=available_products,
    key="product_selection",
    placeholder="Select products to display",
    help="EUGAS products are consolidated to TTF, THE, PEG, NBP, or ZTP.",
)

if not selected_products:
    st.warning("No products are selected. Select at least one product.")
    st.stop()

filtered = filtered_after_contracts.loc[
    filtered_after_contracts["DISPLAY PRODUCT"].isin(selected_products)
].copy()

# ============================================================
# SEARCH FILTER
# ============================================================
if search:
    search_key = clean_text(search)
    contract_match = clean_text_series(
        filtered["MATCHED CONTRACT"]
    ).str.contains(search_key, regex=False, na=False)
    reference_match = clean_text_series(
        filtered["CARGO#/TRADEID"]
    ).str.contains(search_key, regex=False, na=False)
    product_match = clean_text_series(
        filtered["DISPLAY PRODUCT"]
    ).str.contains(search_key, regex=False, na=False)

    filtered = filtered.loc[
        contract_match | reference_match | product_match
    ].copy()

# ============================================================
# DATA QUALITY
# ============================================================
unmatched_data = filtered_before_contracts.loc[
    filtered_before_contracts["MATCHED CONTRACT"].isna()
].copy()

unclassified_data = filtered_before_contracts.loc[
    filtered_before_contracts["MATCHED CONTRACT"].notna()
    & filtered_before_contracts["SECTION"].eq("Unclassified")
].copy()

matched = filtered.loc[
    filtered["MATCHED CONTRACT"].notna()
    & filtered["SECTION"].isin(section_order)
].copy()

if matched.empty:
    st.warning("No matched exposure records were found for the selected filters.")
    st.stop()

# ============================================================
# BUILD AND DISPLAY EXPOSURE TABLE
# ============================================================
(
    exposure_table,
    value_columns,
    subtotal_rows,
) = build_exposure_table(
    matched,
    section_order,
    selected_year,
)

view_name = "All Years" if selected_year == "All Years" else str(selected_year)

st.subheader(
    f"{selected_type.title()} Exposure: {view_name} (IFRS Year basis)"
)

numeric_table = exposure_table.copy()
display_table = exposure_table.copy()

for column in value_columns:
    display_table[column] = display_table[column].apply(format_number)

styled_table = style_table(
    numeric_df=numeric_table,
    display_df=display_table,
    subtotal_rows=subtotal_rows,
    value_columns=value_columns,
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
    unmatched_count = unmatched_data["CARGO#/TRADEID"].nunique()
    st.warning(
        f"{unmatched_count:,} cargo references could not be matched to "
        "the 'All cargo refs' worksheet."
    )

    with st.expander("View unmatched cargo references"):
        st.dataframe(
            unmatched_data[
                [
                    "CARGO#/TRADEID",
                    "TRADESNEW CONTRACT",
                    "TYPE",
                    "PRODUCT",
                    "EXPOSURE DATE",
                    "EXPOSURE DATE YEAR",
                    "IFRS YEAR",
                    "VOLUME_TBTU",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

if not unclassified_data.empty:
    unclassified_count = unclassified_data["PRODUCT"].nunique()
    st.warning(
        f"{unclassified_count:,} products could not be assigned to a "
        f"{selected_type} section."
    )

    with st.expander("View unclassified products"):
        st.dataframe(
            unclassified_data[
                [
                    "PRODUCT",
                    "CARGO#/TRADEID",
                    "MATCHED CONTRACT",
                    "IFRS YEAR",
                    "EXPOSURE DATE",
                    "EXPOSURE DATE YEAR",
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

with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
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
            "align": "center",
        }
    )
    number_format = workbook.add_format(
        {"num_format": '#,##0.00;[Red](#,##0.00)'}
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
            "num_format": '#,##0.00;[Red](#,##0.00)',
        }
    )

    for column_number, column_name in enumerate(exposure_table.columns):
        worksheet.write(0, column_number, column_name, header_format)

    worksheet.set_column(0, 0, 14)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 18)

    if value_columns:
        first_value_column = 3
        last_value_column = first_value_column + len(value_columns) - 1
        worksheet.set_column(
            first_value_column,
            last_value_column,
            13,
            number_format,
        )

    for subtotal_index in subtotal_rows:
        excel_row = subtotal_index + 1

        for column_number in range(len(exposure_table.columns)):
            value = exposure_table.iloc[subtotal_index, column_number]

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

    worksheet.freeze_panes(1, 3)

    if not year_mismatches.empty:
        mismatch_export = year_mismatches[
            [
                "CARGO#/TRADEID",
                "TRADESNEW CONTRACT",
                "MATCHED CONTRACT",
                "TYPE",
                "PRODUCT",
                "EXPOSURE DATE",
                "EXPOSURE DATE YEAR",
                "IFRS YEAR",
                "VOLUME_TBTU",
            ]
        ].copy()
        mismatch_export.to_excel(
            writer,
            sheet_name="Year Mismatches",
            index=False,
        )

    if not unmatched_data.empty:
        unmatched_export = unmatched_data[
            [
                "CARGO#/TRADEID",
                "TRADESNEW CONTRACT",
                "TYPE",
                "PRODUCT",
                "EXPOSURE DATE",
                "EXPOSURE DATE YEAR",
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
                "EXPOSURE DATE YEAR",
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
    "All_Years" if selected_year == "All Years" else str(selected_year)
)

st.download_button(
    label="Download Exposure Table",
    data=output.getvalue(),
    file_name=(
        f"{selected_type.title()}_Contract_Exposure_"
        f"{download_year}_IFRS_Year.xlsx"
    ),
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
)
