
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

from dotenv import load_dotenv
from google import genai
import os
import json


# ==================================================
# Gemini Configuration
# ==================================================

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")

if api_key:
    client = genai.Client(api_key=api_key)
else:
    client = None


# ==================================================
# Streamlit Configuration
# ==================================================

st.set_page_config(
    page_title="AI Business Reporting Assistant",
    page_icon="📊",
    layout="wide"
)

st.markdown(
    """
    <style>

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 3rem;
        padding-right: 3rem;
    }

    </style>
    """,
    unsafe_allow_html=True
)

st.title("📊 AI Business Reporting Assistant")

st.caption(
    "AI-powered business analytics and reporting dashboard"
)

st.write(
    "Upload your business data to explore performance, "
    "identify trends and generate AI-powered insights."
)


# ==================================================
# File Upload
# ==================================================

uploaded_file = st.file_uploader(
    "Upload your Excel file",
    type=["xlsx"]
)

if uploaded_file is None:

    st.info(
        "📁 Please upload an Excel file to start the analysis."
    )

    st.stop()

if uploaded_file is not None:

# ==================================================
# Read Excel
# ==================================================

    df = pd.read_excel(uploaded_file)

# ==================================================
# Column Name Mapping
# ==================================================

column_mapping = {
    "Date": [
        "Date",
        "date",
        "Transaction Date",
        "Order Date"
    ],

    "Department": [
        "Department",
        "department",
        "Dept",
        "Business Unit"
    ],

    "Product": [
        "Product",
        "product",
        "Product Name",
        "Item",
        "Item Name"
    ],

    "Sales": [
        "Sales",
        "sales",
        "Revenue",
        "Total Sales",
        "Net Sales"
    ],

    "Cost": [
        "Cost",
        "cost",
        "Total Cost",
        "Product Cost"
    ],

    "Units": [
        "Units",
        "units",
        "Quantity",
        "Qty",
        "Units Sold"
    ],

    "Inventory": [
        "Inventory",
        "inventory",
        "Stock",
        "Stock Level",
        "Inventory Level"
    ]
}

# ==================================================
# Automatic Column Mapping
# ==================================================

rename_columns = {}

for standard_name, possible_names in column_mapping.items():

    for column in df.columns:

        if str(column).strip() in possible_names:

            rename_columns[column] = standard_name

            break


df = df.rename(
    columns=rename_columns
)

# ==================================================
# Data Validation
# ==================================================

required_columns = [
    "Date",
    "Department",
    "Product",
    "Sales",
    "Cost",
    "Units",
    "Inventory"
]


missing_columns = [
    column
    for column in required_columns
    if column not in df.columns
]


if missing_columns:

    st.error(
        "❌ The uploaded Excel file is missing "
        "required columns."
    )

    st.write(
        "Missing columns:"
    )

    st.write(
        missing_columns
    )

    st.stop()


# ==================================================
# Show Column Mapping
# ==================================================

if rename_columns:

    st.info(
        "ℹ️ Some column names were automatically "
        "mapped to the standard dataset format."
    )

    mapping_display = pd.DataFrame(
        [
            {
                "Original Column": original,
                "Standard Column": standard
            }
            for original, standard
            in rename_columns.items()
            if original != standard
        ]
    )

    if not mapping_display.empty:

        st.dataframe(
            mapping_display,
            use_container_width=True,
            hide_index=True
        )
    # ==================================================
    # Numeric Data Validation
    # ==================================================

    numeric_columns = [
        "Sales",
        "Cost",
        "Units",
        "Inventory"
    ]


    invalid_numeric_columns = []


    for column in numeric_columns:

        converted = pd.to_numeric(
            df[column],
            errors="coerce"
        )

        invalid_values = (
        converted.isna()
        & df[column].notna()
    )

    if invalid_values.any():

        invalid_numeric_columns.append(
            column
        )


    if invalid_numeric_columns:

        st.error(
            "❌ Invalid numeric data detected."
        )

        st.write(
            "The following columns contain "
            "non-numeric values:"
        )

        st.write(
            invalid_numeric_columns
        )

        st.stop()


    # ==================================================
    # Negative Value Validation
    # ==================================================

    negative_columns = []


    for column in numeric_columns:

        if (df[column] < 0).any():

            negative_columns.append(
                column
            )


    if negative_columns:

        st.warning(
            "⚠️ Negative values were found "
            "in the uploaded data."
        )

        st.write(
            "Columns containing negative values:"
        )

        st.write(
            negative_columns
        )

    # ==================================================
    # Missing Value Validation
    # ==================================================

    missing_values = (
        df[required_columns]
        .isna()
        .sum()
    )


    missing_values = (
        missing_values[
            missing_values > 0
        ]
    )


    if not missing_values.empty:

        st.warning(
            "⚠️ Missing values were detected "
            "in the uploaded dataset."
        )

        st.dataframe(
            missing_values.rename(
                "Missing Values"
            )
        )

    # ==================================================
    # Date Validation
    # ==================================================

    original_dates = df["Date"].copy()

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    invalid_date_mask = (
        df["Date"].isna()
        & original_dates.notna()
    )

    if invalid_date_mask.any():

        invalid_date_rows = df.loc[
            invalid_date_mask
        ].copy()

        invalid_date_rows.insert(
            0,
            "Excel Row",
            invalid_date_rows.index + 2
        )

        invalid_date_rows.insert(
            1,
            "Invalid Date",
            original_dates.loc[
                invalid_date_mask
            ].values
        )

        st.error(
            "❌ Invalid date values were detected "
            "in the Date column."
        )

        st.write(
            "The following rows contain invalid dates:"
        )

        st.dataframe(
            invalid_date_rows[
                ["Excel Row", "Invalid Date"]
            ],
            use_container_width=True,
            hide_index=True
        )

        st.stop()
    # ==================================================
    # Data Quality Report
    # ==================================================

    st.markdown("---")

    st.subheader("📋 Data Quality Report")

    # Total rows
    total_rows = len(df)

    # Total columns
    total_columns = len(df.columns)

    # Missing values
    total_missing = int(
        df[required_columns]
        .isna()
        .sum()
        .sum()
    )

    # Duplicate rows
    total_duplicates = int(
        df.duplicated().sum()
    )

    if total_duplicates > 0:

        st.warning(
            f"⚠️ {total_duplicates} duplicate "
            "row(s) detected."
        )

        duplicate_rows = df[
            df.duplicated(keep=False)
        ].copy()

        duplicate_rows.insert(
            0,
            "Excel Row",
            duplicate_rows.index + 2
        )

        with st.expander("🔍 View Duplicate Rows"):

            st.dataframe(
                duplicate_rows,
                use_container_width=True,
                hide_index=True
            )
    
    # Invalid dates
    invalid_dates = int(
        df["Date"].isna().sum()
    )


    quality_col1, quality_col2, quality_col3, quality_col4, quality_col5 = st.columns(5)


    quality_col1.metric(
        "Rows",
        f"{total_rows:,}"
    )

    quality_col2.metric(
        "Columns",
        f"{total_columns:,}"
    )

    quality_col3.metric(
        "Missing Values",
        f"{total_missing:,}"
    )

    quality_col4.metric(
        "Duplicate Rows",
        f"{total_duplicates:,}"
    )

    quality_col5.metric(
        "Invalid Dates",
        f"{invalid_dates:,}"
    )

    # ==================================================
    # Data Quality Score
    # ==================================================

    total_cells = (
        total_rows * len(required_columns)
    )

    if total_cells > 0:

        missing_ratio = (
            total_missing / total_cells
        )

        duplicate_ratio = (
            total_duplicates / total_rows
        )

        quality_score = (
            1
            - missing_ratio
            - duplicate_ratio
        ) * 100

        quality_score = max(
            0,
            min(100, quality_score)
        )

    else:

        quality_score = 0


    st.metric(
        "Overall Data Quality Score",
        f"{quality_score:.1f}%"
    )

    if quality_score >= 95:

        st.success(
            "✅ Dataset quality is excellent."
        )

    elif quality_score >= 80:

        st.warning(
            "⚠️ Dataset quality is acceptable, "
            "but some issues should be reviewed."
        )

    else:

        st.error(
            "❌ Dataset quality requires attention."
        )

    # ==================================================
    # Dataset Information
    # ==================================================
    st.success(
        "✅ Dataset validation completed successfully."
    )
    info_col1, info_col2, info_col3 = st.columns(3)


    info_col1.metric(
        "Rows",
        f"{len(df):,}"
    )


    info_col2.metric(
        "Columns",
        f"{len(df.columns):,}"
    )


    info_col3.metric(
        "Date Range",
        f"{df['Date'].min():%d.%m.%y} → "
        f"{df['Date'].max():%d.%m.%y}"
    )


    st.success("Excel file uploaded successfully!")


    # ==================================================
    # Sidebar Filters
    # ==================================================

    st.sidebar.title("⚙️ Dashboard Controls")

    st.sidebar.markdown("---")

    st.sidebar.subheader("Filters")

    departments = ["All"] + sorted(
        df["Department"].unique().tolist()
    )

    selected_department = st.sidebar.selectbox(
        "Department",
        departments
    )
    
    st.sidebar.markdown("---")

    st.sidebar.caption(
        "AI Business Reporting Assistant"
    )

    st.sidebar.caption(
        "Data analysis powered by Python & Pandas"
    )

    st.sidebar.caption(
        "AI insights powered by Gemini"
    )

    # Apply Department Filter

    if selected_department != "All":

        filtered_df = df[
            df["Department"] == selected_department
        ]

    else:

        filtered_df = df.copy()


    # ==================================================
    # KPI Calculations
    # ==================================================

    total_sales = filtered_df["Sales"].sum()

    total_cost = filtered_df["Cost"].sum()

    total_profit = (
        total_sales
        -
        total_cost
    )


    if total_sales != 0:

        profit_margin = (
            total_profit
            /
            total_sales
        ) * 100

    else:

        profit_margin = 0


    total_units = filtered_df["Units"].sum()

    total_inventory = filtered_df["Inventory"].sum()


    # ==================================================
    # Business KPIs
    # ==================================================

    st.subheader("📈 Business Performance")


    col1, col2, col3, col4 = st.columns(4)


    col1.metric(
        "💰 Total Sales",
        f"{total_sales:,.0f} TL"
    )

    col2.metric(
        "💸 Total Cost",
        f"{total_cost:,.0f} TL"
    )

    col3.metric(
        "📊 Total Profit",
        f"{total_profit:,.0f} TL"
    )

    col4.metric(
        "📈 Profit Margin",
        f"{profit_margin:.2f}%"
    )


    col1, col2 = st.columns(2)

    col1.metric(
        "📦 Units Sold",
        f"{total_units:,}"
    )

    col2.metric(
        "🏭 Inventory",
        f"{total_inventory:,}"
    )


    # ==================================================
    # Sales Analysis
    # ==================================================

    st.subheader("📊 Sales Performance")


    # --------------------------------------------------
    # Sales by Department & Sales by Product
    # --------------------------------------------------

    department_sales = (
        filtered_df
        .groupby("Department")["Sales"]
        .sum()
        .reset_index()
    )

    product_sales = (
            filtered_df
            .groupby("Product")["Sales"]
            .sum()
            .reset_index()
            .sort_values(
                "Sales",
                ascending=False
            )
        )
    
    col1, col2 = st.columns(2)


    with col1:

        fig_department = px.bar(
            department_sales,
            x="Department",
            y="Sales",
            title="Department Sales"
        )

        st.plotly_chart(
            fig_department,
            use_container_width=True
        )


    with col2:

        fig_product = px.bar(
            product_sales,
            x="Product",
            y="Sales",
            title="Product Sales"
        )

        st.plotly_chart(
            fig_product,
            use_container_width=True
        )


    # --------------------------------------------------
    # Sales Trend
    # --------------------------------------------------

    daily_sales_chart = (
        filtered_df
        .groupby("Date")["Sales"]
        .sum()
        .reset_index()
    )


    fig_time = px.line(
        daily_sales_chart,
        x="Date",
        y="Sales",
        title="Daily Sales Trend"
    )


    st.plotly_chart(
        fig_time,
        use_container_width=True
    )

   
    # ==================================================
    # Automated Performance Analysis
    # ==================================================

    st.subheader("🔎 Automated Performance Analysis")


    # --------------------------------------------------
    # Department Performance
    # --------------------------------------------------

    department_performance = (
        filtered_df
        .groupby("Department")["Sales"]
        .sum()
        .sort_values(
            ascending=False
        )
    )


    top_department = (
        department_performance.index[0]
    )


    top_department_sales = (
        department_performance.iloc[0]
    )


    # Detailed Department Analysis

    department_analysis = (
        filtered_df
        .groupby("Department")
        .agg(
            Sales=("Sales", "sum"),
            Cost=("Cost", "sum"),
            Units=("Units", "sum"),
            Inventory=("Inventory", "sum")
        )
    )


    department_analysis["Profit"] = (
        department_analysis["Sales"]
        -
        department_analysis["Cost"]
    )


    department_analysis["Profit Margin"] = (
        department_analysis["Profit"]
        /
        department_analysis["Sales"]
    ) * 100


    if total_sales != 0:

        department_analysis["Sales Share"] = (
            department_analysis["Sales"]
            /
            total_sales
        ) * 100

    else:

        department_analysis["Sales Share"] = 0


    # --------------------------------------------------
    # Product Performance
    # --------------------------------------------------

    product_performance = (
        filtered_df
        .groupby("Product")["Sales"]
        .sum()
        .sort_values(
            ascending=False
        )
    )


    top_product = (
        product_performance.index[0]
    )


    top_product_sales = (
        product_performance.iloc[0]
    )


    # Detailed Product Analysis

    product_analysis = (
        filtered_df
        .groupby("Product")
        .agg(
            Sales=("Sales", "sum"),
            Cost=("Cost", "sum"),
            Units=("Units", "sum"),
            Inventory=("Inventory", "sum")
        )
    )


    product_analysis["Profit"] = (
        product_analysis["Sales"]
        -
        product_analysis["Cost"]
    )


    product_analysis["Profit Margin"] = (
        product_analysis["Profit"]
        /
        product_analysis["Sales"]
    ) * 100


    if total_sales != 0:

        product_analysis["Sales Share"] = (
            product_analysis["Sales"]
            /
            total_sales
        ) * 100

    else:

        product_analysis["Sales Share"] = 0


    # Lowest Margin Product

    product_analysis = product_analysis.dropna(
        subset=["Profit Margin"]
    )


    lowest_margin_product = (
        product_analysis[
            "Profit Margin"
        ].idxmin()
    )


    lowest_margin_value = (
        product_analysis.loc[
            lowest_margin_product,
            "Profit Margin"
        ]
    )


    # --------------------------------------------------
    # Best Sales Day
    # --------------------------------------------------

    daily_sales = (
        filtered_df
        .groupby("Date")["Sales"]
        .sum()
        .sort_values(
            ascending=False
        )
    )


    best_sales_day = (
        daily_sales.index[0]
    )


    best_sales_day_value = (
        daily_sales.iloc[0]
    )


    # --------------------------------------------------
    # Monthly Sales Analysis
    # --------------------------------------------------

    monthly_sales = (
        filtered_df
        .groupby(
            filtered_df["Date"].dt.to_period("M")
        )["Sales"]
        .sum()
    )


    monthly_sales.index = (
        monthly_sales.index.astype(str)
    )

    # ==================================================
    # Advanced Department Analytics
    # ==================================================

    st.markdown("---")

    st.subheader("🏢 Department Performance")

    department_display = department_analysis.copy()

    department_display = department_display.reset_index()

    department_display["Sales"] = (
        department_display["Sales"]
        .round(0)
    )

    department_display["Profit"] = (
        department_display["Profit"]
        .round(0)
    )

    department_display["Profit Margin"] = (
        department_display["Profit Margin"]
        .round(2)
    )

    department_display["Sales Share"] = (
        department_display["Sales Share"]
        .round(2)
    )

    st.dataframe(
        department_display[
            [
                "Department",
                "Sales",
                "Profit",
                "Profit Margin",
                "Sales Share"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )


    fig_department_profit = px.bar(
        department_display,
        x="Department",
        y="Profit",
        title="Department Profit"
    )

    st.plotly_chart(
        fig_department_profit,
        use_container_width=True
    )


    # ==================================================
    # Advanced Product Analytics
    # ==================================================

    st.markdown("---")

    st.subheader("📦 Product Performance")

    product_display = product_analysis.copy()

    product_display = product_display.reset_index()

    product_display["Sales"] = (
        product_display["Sales"]
        .round(0)
    )

    product_display["Profit"] = (
        product_display["Profit"]
        .round(0)
    )

    product_display["Profit Margin"] = (
        product_display["Profit Margin"]
        .round(2)
    )

    product_display["Sales Share"] = (
        product_display["Sales Share"]
        .round(2)
    )

    st.dataframe(
        product_display[
            [
                "Product",
                "Sales",
                "Profit",
                "Profit Margin",
                "Sales Share"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )


    low_margin_products = (
        product_display
        .sort_values(
            "Profit Margin",
            ascending=True
        )
        .head(5)
    )


    fig_low_margin = px.bar(
        low_margin_products,
        x="Profit Margin",
        y="Product",
        orientation="h",
        title="Lowest Profit Margin Products"
    )

    st.plotly_chart(
        fig_low_margin,
        use_container_width=True
    )


    # ==================================================
    # Inventory Analytics
    # ==================================================

    st.markdown("---")

    st.subheader("📦 Inventory Analysis")

    inventory_analysis = (
        filtered_df
        .groupby("Product")
        .agg(
            Units_Sold=("Units", "sum"),
            Inventory=("Inventory", "sum")
        )
        .reset_index()
    )

    inventory_analysis["Inventory / Units"] = (
        inventory_analysis["Inventory"]
        /
        inventory_analysis["Units_Sold"].replace(0, pd.NA)
    )

    inventory_analysis = (
        inventory_analysis
        .sort_values(
            "Inventory / Units",
            ascending=False
        )
    )

    st.dataframe(
        inventory_analysis.round(2),
        use_container_width=True,
        hide_index=True
    )


    # ==================================================
    # Monthly Sales Analytics
    # ==================================================

    st.markdown("---")

    st.subheader("📅 Monthly Sales Performance")

    monthly_sales_df = (
        monthly_sales
        .reset_index()
    )

    monthly_sales_df.columns = [
        "Month",
        "Sales"
    ]

    fig_monthly = px.line(
        monthly_sales_df,
        x="Month",
        y="Sales",
        markers=True,
        title="Monthly Sales Trend"
    )

    st.plotly_chart(
        fig_monthly,
        use_container_width=True
    )


    monthly_sales_df["MoM Change"] = (
        monthly_sales_df["Sales"]
        .pct_change()
        * 100
    )

    monthly_sales_df["MoM Change"] = (
        monthly_sales_df["MoM Change"]
        .round(2)
    )

    st.dataframe(
        monthly_sales_df,
        use_container_width=True,
        hide_index=True
    )

    # ==================================================
    # Automated Analysis Display
    # ==================================================

    col1, col2 = st.columns(2)


    with col1:

        st.info(
            f"🏆 **Top Department:** "
            f"{top_department} "
            f"({top_department_sales:,.0f} TL)"
        )


        st.info(
            f"📦 **Top Product:** "
            f"{top_product} "
            f"({top_product_sales:,.0f} TL)"
        )


    with col2:

        st.warning(
            f"⚠️ **Lowest Margin Product:** "
            f"{lowest_margin_product} "
            f"({lowest_margin_value:.2f}%)"
        )


        st.success(
            f"📅 **Best Sales Day:** "
            f"{best_sales_day.strftime('%Y-%m-%d')} "
            f"({best_sales_day_value:,.0f} TL)"
        )

    # ==================================================
    # Excel Report Export
    # ==================================================

    def create_excel_report(
        df,
        department_analysis,
        product_analysis,
        ai_insights=""
    ):

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            # ==================================================
            # KPI Summary
            # ==================================================

            total_sales = df["Sales"].sum()
            total_cost = df["Cost"].sum()
            total_profit = total_sales - total_cost

            total_units = df["Units"].sum()
            total_inventory = df["Inventory"].sum()

            profit_margin = (
                total_profit / total_sales * 100
                if total_sales != 0
                else 0
            )

            kpi_summary = pd.DataFrame({
                "Metric": [
                    "Total Sales",
                    "Total Cost",
                    "Total Profit",
                    "Profit Margin",
                    "Total Units Sold",
                    "Total Inventory"
                ],

                "Value": [
                    total_sales,
                    total_cost,
                    total_profit,
                    f"{profit_margin:.2f}%",
                    total_units,
                    total_inventory
                ]
            })

            kpi_summary.to_excel(
                writer,
                sheet_name="KPI Summary",
                index=False
            )

            # ==================================================
            # Dataset
            # ==================================================

            df.to_excel(
                writer,
                sheet_name="Dataset",
                index=False
            )

            # ==================================================
            # Department Analysis
            # ==================================================

            department_analysis.to_excel(
                writer,
                sheet_name="Department Analysis"
            )

            # ==================================================
            # Product Analysis
            # ==================================================

            product_analysis.to_excel(
                writer,
                sheet_name="Product Analysis"
            )

            # ==================================================
            # AI Business Insights
            # ==================================================

            if ai_insights:

                ai_lines = ai_insights.splitlines()

                ai_report = pd.DataFrame({
                    "AI Business Insights": ai_lines
                })

                ai_report.to_excel(
                    writer,
                    sheet_name="AI Business Insights",
                    index=False
                )

            # ==================================================
            # Monthly Analysis
            # ==================================================

            monthly_analysis = (
                df.assign(
                    Month=df["Date"]
                    .dt.to_period("M")
                    .astype(str)
                )
                .groupby("Month")
                .agg(
                    Sales=("Sales", "sum"),
                    Cost=("Cost", "sum"),
                    Units=("Units", "sum")
                )
                .reset_index()
            )

            monthly_analysis["Profit"] = (
                monthly_analysis["Sales"]
                - monthly_analysis["Cost"]
            )

            monthly_analysis["Profit Margin"] = (
                monthly_analysis["Profit"]
                / monthly_analysis["Sales"]
                * 100
            )

            monthly_analysis.to_excel(
                writer,
                sheet_name="Monthly Analysis",
                index=False
            )

            # ==================================================
            # Inventory Analysis
            # ==================================================

            inventory_analysis = (
                df.groupby("Product")
                .agg(
                    Total_Units_Sold=("Units", "sum"),
                    Total_Inventory=("Inventory", "sum")
                )
                .reset_index()
            )

            inventory_analysis["Inventory / Units Sold"] = (
                inventory_analysis["Total_Inventory"]
                / inventory_analysis["Total_Units_Sold"]
                .replace(0, pd.NA)
            )

            inventory_analysis.to_excel(
                writer,
                sheet_name="Inventory Analysis",
                index=False
            )

            # ==================================================
            # Professional Excel Formatting
            # ==================================================

            workbook = writer.book

            for worksheet in workbook.worksheets:

                # Header formatting
                for cell in worksheet[1]:

                    cell.font = Font(
                        bold=True
                    )

                    cell.alignment = Alignment(
                        horizontal="center"
                    )

                # Automatic column width
                for column_cells in worksheet.columns:

                    max_length = 0

                    column_letter = get_column_letter(
                        column_cells[0].column
                    )

                    for cell in column_cells:

                        if cell.value is not None:

                            cell_length = len(
                                str(cell.value)
                            )

                            max_length = max(
                                max_length,
                                cell_length
                            )

                    worksheet.column_dimensions[
                        column_letter
                    ].width = min(
                        max_length + 2,
                        40
                    )

                # Freeze header row
                worksheet.freeze_panes = "A2"


            # ==================================================
            # KPI Summary Formatting
            # ==================================================

            kpi_sheet = workbook["KPI Summary"]

            kpi_sheet.column_dimensions["A"].width = 25
            kpi_sheet.column_dimensions["B"].width = 20

            for row in kpi_sheet.iter_rows(
                min_row=2,
                max_row=kpi_sheet.max_row,
                min_col=2,
                max_col=2
            ):

                for cell in row:

                    if cell.row in [2, 3, 4]:

                        cell.number_format = '#,##0.00'

                    elif cell.row in [6, 7]:

                        cell.number_format = '#,##0'

            # Move KPI Summary to first position
            workbook._sheets.insert(
                0,
                workbook._sheets.pop(
                    workbook.sheetnames.index(
                        "KPI Summary"
                    )
                )
            )

            # ==================================================
            # Excel Charts
            # ==================================================

            # --------------------------------------------------
            # Monthly Sales & Profit Chart
            # --------------------------------------------------

            monthly_sheet = workbook["Monthly Analysis"]

            sales_profit_chart = LineChart()

            sales_profit_chart.title = "Monthly Sales & Profit"
            sales_profit_chart.y_axis.title = "Amount"
            sales_profit_chart.x_axis.title = "Month"

            data = Reference(
                monthly_sheet,
                min_col=2,
                max_col=5,
                min_row=1,
                max_row=monthly_sheet.max_row
            )

            categories = Reference(
                monthly_sheet,
                min_col=1,
                min_row=2,
                max_row=monthly_sheet.max_row
            )

            sales_profit_chart.add_data(
                data,
                titles_from_data=True
            )

            sales_profit_chart.set_categories(
                categories
            )

            sales_profit_chart.height = 8
            sales_profit_chart.width = 16

            monthly_sheet.add_chart(
                sales_profit_chart,
                "H2"
            )


            # --------------------------------------------------
            # Product Inventory Chart
            # --------------------------------------------------

            inventory_sheet = workbook["Inventory Analysis"]

            inventory_chart = BarChart()

            inventory_chart.type = "bar"
            inventory_chart.style = 10

            inventory_chart.title = "Inventory by Product"
            inventory_chart.y_axis.title = "Product"
            inventory_chart.x_axis.title = "Inventory"

            inventory_data = Reference(
                inventory_sheet,
                min_col=3,
                max_col=3,
                min_row=1,
                max_row=inventory_sheet.max_row
            )

            inventory_categories = Reference(
                inventory_sheet,
                min_col=1,
                min_row=2,
                max_row=inventory_sheet.max_row
            )

            inventory_chart.add_data(
                inventory_data,
                titles_from_data=True
            )

            inventory_chart.set_categories(
                inventory_categories
            )

            inventory_chart.height = 10
            inventory_chart.width = 16

            inventory_sheet.add_chart(
                inventory_chart,
                "F2"
            )

        output.seek(0)

        return output

    # ==================================================
    # Export Business Report
    # ==================================================

    st.markdown("---")

    st.subheader("📥 Export Business Report")

    report_file = create_excel_report(
    df,
    department_analysis,
    product_analysis,
    ai_insights=st.session_state.get(
        "ai_response",
        ""
    )
)
    st.download_button(
        label="📥 Download Excel Report",
        data=report_file,
        file_name="business_analysis_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ==================================================
    # AI Analysis Summary
    # ==================================================

    st.markdown("---")
    analysis_summary = {

        "overall": {

            "total_sales": float(
                total_sales
            ),

            "total_cost": float(
                total_cost
            ),

            "total_profit": float(
                total_profit
            ),

            "profit_margin": float(
                profit_margin
            ),

            "total_units": int(
                total_units
            ),

            "total_inventory": int(
                total_inventory
            )
        },


        "top_performers": {

            "top_department":
                top_department,

            "top_department_sales":
                float(
                    top_department_sales
                ),

            "top_product":
                top_product,

            "top_product_sales":
                float(
                    top_product_sales
                ),

            "lowest_margin_product":
                lowest_margin_product,

            "lowest_margin_value":
                float(
                    lowest_margin_value
                ),

            "best_sales_day":
                str(
                    best_sales_day.strftime(
                        "%Y-%m-%d"
                    )
                ),

            "best_sales_day_sales":
                float(
                    best_sales_day_value
                )
        },


        "department_analysis":
            department_analysis
            .round(2)
            .reset_index()
            .to_dict(
                orient="records"
            ),


        "product_analysis":
            product_analysis
            .round(2)
            .reset_index()
            .to_dict(
                orient="records"
            ),


        "monthly_sales":
            monthly_sales
            .round(2)
            .to_dict()
    }


    # ==================================================
    # AI Analysis Data
    # ==================================================

    st.subheader("AI Analysis Data")

    st.json(
        analysis_summary
    )


    # ==================================================
    # Prepare Data for Gemini
    # ==================================================

    analysis_json = json.dumps(
        analysis_summary,
        ensure_ascii=False,
        indent=2
    )


    # ==================================================
    # Gemini Prompt
    # ==================================================

    prompt = f"""
You are a business intelligence analyst.

Analyze the verified business data provided below.

Your task is to produce a concise management-level business report.

IMPORTANT RULES:

1. Use only the data provided below.
2. Do not invent numerical values.
3. Do not claim that a correlation or causal relationship exists unless it is directly supported by the data.
4. When suggesting possible explanations, clearly label them as hypotheses.
5. Do not treat inventory levels as a financial risk unless the provided data supports that conclusion.
6. All numerical values must come from the provided data.
7. Focus on actionable business insights.
8. Keep the report concise and professional.

Structure your response using exactly these sections:

## Executive Summary

Summarize the overall business performance in 3-5 bullet points.

## Department Performance

Identify important differences between departments using sales, profit,
profit margin and sales share.

## Product Performance

Identify important product-level performance differences.

## Sales Trend

Analyze the monthly sales data and identify notable patterns.

## Potential Risks

Identify data-supported risks or areas that require further investigation.

## Recommended Actions

Provide 3-5 practical recommendations based strictly on the available data.

Remember:
- Observations must be based on the data.
- Hypotheses must be clearly identified as hypotheses.
- Do not invent facts.

Business data:

{analysis_json}
"""

# ==================================================
# AI Business Assistant
# ==================================================

st.markdown("---")

st.subheader("💬 Ask AI About Your Business Data")

st.write(
    "Ask a natural-language question about the "
    "business data and receive an AI-powered answer."
)

st.markdown("#### Quick Questions")

quick_question = st.selectbox(
    "Choose a question",
    [
        "Select a question...",
        "Which department has the highest sales?",
        "Which product has the lowest profit margin?",
        "Which product generates the highest sales?",
        "What is the overall profit margin?",
        "Which department generates the highest profit?",
        "Which products have the highest inventory levels?"
    ]
)

user_question = st.text_input(
    "Your question",
    value="" if quick_question == "Select a question..." else quick_question,
    placeholder="e.g. Which product has the lowest profit margin?"
)


ask_ai = st.button(
    "💬 Ask AI",
    type="primary"
)


if ask_ai:

    if not user_question.strip():

        st.warning(
            "Please enter a question first."
        )

    elif client is None:

        st.error(
            "Gemini API key not found. "
            "Please check your .env file."
        )

    else:

        with st.spinner(
            "AI is analyzing your question..."
        ):

            try:

                interaction = client.interactions.create(
                    model="gemini-3.6-flash",
                    input=(
                            f"""
                        You are an AI business analyst.

                        Your task is to answer the user's business question
                        using ONLY the verified metrics provided below.

                        Rules:
                        - Do not invent or assume data.
                        - Do not use external information.
                        - Base every numerical statement on the provided metrics.
                        - Keep the answer concise and business-focused.
                        - Explain the reasoning briefly when useful.
                        - If the available data is insufficient, clearly state that.

                        Verified Business Metrics:
                        {prompt}

                        User Question:
                        {user_question}

                        Provide a clear answer suitable for a business manager.
                        """
                        )
                )

                ai_answer = interaction.output_text

                st.markdown("---")

                st.markdown("### 🤖 AI Answer")

                st.info(ai_answer)

            except Exception as e:

                error_message = str(e)

                if "429" in error_message:

                    st.warning(
                        "⚠️ The Gemini free-tier request "
                        "limit has been reached. "
                        "Please try again later."
                    )

                else:

                    st.error(
                        f"Gemini API error: {e}"
                    )

# ==================================================
# Gemini AI Analysis
# ==================================================
st.markdown("---")
st.subheader("🤖 AI Business Insights")

st.write(
    "Generate an AI-powered management report "
    "based on the verified business metrics above."
)


if "ai_response" not in st.session_state:
    st.session_state.ai_response = None


generate_report = st.button(
    "🤖 Generate AI Report",
    type="primary"
)


if generate_report:

    if client is None:

        st.error(
            "Gemini API key not found. "
            "Please check your .env file."
        )

    else:

        with st.spinner(
            "Gemini is analyzing the business data..."
        ):

            try:

                interaction = client.interactions.create(
                    model="gemini-3.6-flash",
                    input=prompt
                )

                st.session_state.ai_response = (
                    interaction.output_text
                )

            except Exception as e:

                error_message = str(e)

                if "429" in error_message:

                    st.warning(
                        "⚠️ The Gemini free-tier request "
                        "limit has been reached. "
                        "Please try again later."
                    )

                else:

                    st.error(
                        f"Gemini API error: {e}"
                    )


if st.session_state.ai_response:

    st.markdown(
        st.session_state.ai_response
    )

    # ==================================================
    # Dataset Preview
    # ==================================================
    st.markdown("---")
    st.subheader("📋 Dataset Preview")


    st.write(
        f"Rows: {filtered_df.shape[0]} | "
        f"Columns: {filtered_df.shape[1]}"
    )


    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True
    )

