import base64
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
import altair as alt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ---------------- PAGE SETTINGS ----------------
st.set_page_config(
    page_title="AutoInsight | Car Sales SmartBot",
    page_icon="🚗",
    layout="wide"
)

# ---------------- CUSTOM DESIGN ----------------
car_background = base64.b64encode(
    Path("assets/car-background.jpg").read_bytes()
).decode("ascii")
custom_css = """
<style>
    .stApp {
        background: linear-gradient(180deg, #fff 0%, #eef3f9 100%);
    }
    [data-testid="stSidebar"] {
        background: #101b33;
    }
    [data-testid="stSidebar"] * {
        color: #f3f6ff;
    }
    .hero {
        padding: 26px 28px;
        border-radius: 20px;
        background:
            linear-gradient(90deg, #101b33 0%, #174a86 43%, rgba(16, 27, 51, 0.12) 100%),
            url("data:image/jpeg;base64,__CAR_BACKGROUND__") right center / 58% auto no-repeat,
            linear-gradient(115deg, #101b33, #174a86);
        color: white;
        margin-bottom: 20px;
        min-height: 145px;
    }
    .hero h1 {
        color: white;
        margin: 0;
        font-size: 2.2rem;
    }
    .hero p {
        color: #d8e7ff;
        margin: 8px 0 0 0;
        font-size: 1rem;
    }
    .metric-card {
        background: white;
        padding: 18px 20px;
        border-radius: 16px;
        border: 1px solid #e1e8f2;
        box-shadow: 0 5px 18px rgba(16, 27, 51, 0.06);
        min-height: 115px;
    }
    .metric-label {
        color: #66758c;
        font-size: 0.88rem;
        margin-bottom: 8px;
    }
    .metric-value {
        color: #132344;
        font-size: 1.65rem;
        font-weight: 750;
    }
    .section-title {
        color: #132344;
        font-size: 1.25rem;
        font-weight: 700;
        margin: 12px 0 6px 0;
    }
    .st-emotion-cache-vyz0sr .stDataFrameGlideDataEditor
     { border:2px solid rgb(255, 75, 75); }

    .small-note {
        color: #66758c;
        font-size: 0.85rem;
    }
    div[data-testid="stChatMessage"] {
        background: white;
        border: 1px solid #e1e8f2;
        border-radius: 14px;
        padding: 10px;
    }
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
    }
    .stFileUploader button,
    .stFileUploader [data-testid="stFileUploaderDropzone"] {
        background-color: rgb(255, 75, 75);
        border-color: #fff;
        color: #ffffff;
        border-radius: 10px;
    }
    .stFileUploader button:hover,
    .stFileUploader [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #0f3565;
        border-color: #0f3565;
        color: #ffffff;
    }
</style>
""".replace("__CAR_BACKGROUND__", car_background)
st.markdown(custom_css, unsafe_allow_html=True)

DEFAULT_FILE_NAME = "India_Car_Sales_150x16.xlsx"
REQUIRED_COLUMNS = {
    "Region",
    "Car_Manufacturer",
    "Fuel_Type",
    "Total_Sales",
    "Sale_ID",
    "Customer_Rating",
    "Car_Model",
    "Units_Sold",
    "Sale_Date",
    "Car_Segment",
}

@st.cache_data
def load_default_data(file_name: str):
    data = pd.read_excel(file_name, sheet_name="Car Sales Data")
    return prepare_data(data)


def prepare_data(data):
    data = data.copy()
    data.columns = [str(col).strip() for col in data.columns]
    column_aliases = {
        "Manufacturer": "Car_Manufacturer",
        "manufacturer": "Car_Manufacturer",
        "Car Manufacturer": "Car_Manufacturer",
        "Fuel": "Fuel_Type",
        "Fuel Type": "Fuel_Type",
        "Total Sale": "Total_Sales",
        "Total Sales": "Total_Sales",
        "Sale ID": "Sale_ID",
        "Sales ID": "Sale_ID",
        "Customer Rating": "Customer_Rating",
        "Car Model": "Car_Model",
        "Units Sold": "Units_Sold",
        "Sale Date": "Sale_Date",
        "Car Segment": "Car_Segment",
    }
    data = data.rename(columns=column_aliases)

    missing = sorted(REQUIRED_COLUMNS - set(data.columns))
    if missing:
        raise ValueError(
            "Uploaded dataset is missing required columns: " + ", ".join(missing)
        )

    for col in ["Total_Sales", "Units_Sold", "Sale_ID"]:
        data[col] = pd.to_numeric(data[col], errors="coerce")
    data["Sale_Date"] = pd.to_datetime(data["Sale_Date"], errors="coerce")
    return data


def load_uploaded_data(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.lower().endswith(".csv"):
        data = pd.read_csv(uploaded_file)
    else:
        uploaded_file.seek(0)
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_name = "Car Sales Data" if "Car Sales Data" in excel_file.sheet_names else excel_file.sheet_names[0]
        uploaded_file.seek(0)
        data = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    return prepare_data(data)


uploaded_file = st.sidebar.file_uploader(
    "Upload dataset",
    type=["csv", "xlsx", "xls"],
    help="Replace the default sample file with your own sales dataset."
)

if uploaded_file is not None:
    try:
        all_data = load_uploaded_data(uploaded_file)
        st.sidebar.success(f"Loaded dataset: {uploaded_file.name}")
    except Exception as exc:
        st.sidebar.error(f"Unable to load uploaded file: {exc}")
        st.stop()
else:
    try:
        all_data = load_default_data(DEFAULT_FILE_NAME)
    except Exception:
        st.error(
            "Excel file not found. Keep 'India_Car_Sales_150x16.xlsx' "
            "in the same folder as app.py or upload a dataset from the sidebar."
        )
        st.stop()

# ---------------- SIDEBAR FILTERS ----------------
st.sidebar.markdown("## 🚗 AutoInsight")
st.sidebar.caption("Car Sales Analytics + SmartBot")
st.sidebar.markdown("---")

regions = sorted(all_data["Region"].dropna().unique().tolist())
manufacturers = sorted(all_data["Car_Manufacturer"].dropna().unique().tolist())
fuels = sorted(all_data["Fuel_Type"].dropna().unique().tolist())

selected_regions = st.sidebar.multiselect("Region", regions, default=regions)
selected_manufacturers = st.sidebar.multiselect(
    "Manufacturer", manufacturers, default=manufacturers
)
selected_fuels = st.sidebar.multiselect("Fuel type", fuels, default=fuels)

filtered = all_data[
    all_data["Region"].isin(selected_regions)
    & all_data["Car_Manufacturer"].isin(selected_manufacturers)
    & all_data["Fuel_Type"].isin(selected_fuels)
].copy()

st.sidebar.markdown("---")
st.sidebar.caption("Tip: Use filters to explore part of the dataset.")

# ---------------- HEADER ----------------
st.markdown("""
<div class="hero">
    <h1>🚘 AutoInsight</h1>
    <p>Car Sales Dashboard & SmartBot · Explore sales data with simple questions</p>
</div>
""", unsafe_allow_html=True)
st.caption(
    'This is a portfolio project built with Python, Pandas, and Streamlit. '
)

if filtered.empty:
    st.warning("No rows match those filters. Please select at least one option in each filter.")
    st.stop()

# ---------------- KPI CARDS ----------------
total_sales = filtered["Total_Sales"].sum()
total_orders = filtered["Sale_ID"].nunique()
total_customers = filtered["Customer_Rating"].notna().sum()
total_car_models = filtered["Car_Model"].nunique()

def money(value):
    return "₹{:,.0f}".format(value)

k1, k2, k3, k4 = st.columns(4)
cards = [
    (k1, "💰 Total Sales", money(total_sales)),
    (k2, "🧾 Total Orders", f"{total_orders:,.0f}"),
    (k3, "👥 Total Customers", f"{total_customers:,.0f}"),
    (k4, "🚘 Total Car Models", f"{total_car_models:,.0f}"),
]
for col, label, value in cards:
    with col:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">{label}</div>'
            f'<div class="metric-value">{value}</div></div>',
            unsafe_allow_html=True
        )

st.write("")
tab_dashboard, tab_chat, tab_data = st.tabs(
    ["📊 Sales Dashboard", "💬 Ask SmartBot", "📋 Explore Data"]
)

# ---------------- DASHBOARD ----------------
with tab_dashboard:
    st.markdown('<div class="section-title">Sales Overview</div>', unsafe_allow_html=True)
    left, right = st.columns(2)

    with left:
        st.markdown("**Total sales by manufacturer**")
        by_make = (
            filtered.groupby("Car_Manufacturer", as_index=False)["Total_Sales"]
            .sum().sort_values("Total_Sales", ascending=False)
        )
        manufacturer_chart = (
            alt.Chart(by_make)
            .mark_bar()
            .encode(
                x=alt.X("Car_Manufacturer:N", sort=alt.SortField("Total_Sales", order="descending"), title="Manufacturer"),
                y=alt.Y("Total_Sales:Q", title="Total Sales"),
                color=alt.Color("Car_Manufacturer:N", legend=None, scale=alt.Scale(scheme="category10")),
                tooltip=[
                    alt.Tooltip("Car_Manufacturer:N", title="Manufacturer"),
                    alt.Tooltip("Total_Sales:Q", title="Total Sales", format=",.0f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(manufacturer_chart, use_container_width=True)

    with right:
        st.markdown("**Units sold by fuel type**")
        by_fuel = (
            filtered.groupby("Fuel_Type", as_index=False)["Units_Sold"]
            .sum().sort_values("Units_Sold", ascending=False)
        )
        fuel_chart = (
            alt.Chart(by_fuel)
            .mark_bar()
            .encode(
                x=alt.X("Fuel_Type:N", sort=alt.SortField("Units_Sold", order="descending"), title="Fuel Type"),
                y=alt.Y("Units_Sold:Q", title="Units Sold"),
                color=alt.Color("Fuel_Type:N", legend=None, scale=alt.Scale(scheme="set2")),
                tooltip=[
                    alt.Tooltip("Fuel_Type:N", title="Fuel Type"),
                    alt.Tooltip("Units_Sold:Q", title="Units Sold", format=",.0f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(fuel_chart, use_container_width=True)

    left2, right2 = st.columns(2)
    with left2:
        st.markdown("**Sales by region**")
        by_region = (
            filtered.groupby("Region", as_index=False)["Total_Sales"]
            .sum().sort_values("Total_Sales", ascending=False)
        )
        region_donut = (
            alt.Chart(by_region)
            .mark_arc(innerRadius=60)
            .encode(
                theta=alt.Theta("Total_Sales:Q"),
                color=alt.Color("Region:N", scale=alt.Scale(scheme="category10")),
                tooltip=["Region:N", alt.Tooltip("Total_Sales:Q", format=",.0f")],
            )
            .properties(height=300)
        )
        st.altair_chart(region_donut, use_container_width=True)

    with right2:
        st.markdown("**Sales by car segment**")
        by_segment = (
            filtered.groupby("Car_Segment", as_index=False)["Units_Sold"]
            .sum().sort_values("Units_Sold", ascending=False)
        )
        segment_chart = (
            alt.Chart(by_segment)
            .mark_line(point=True)
            .encode(
                x=alt.X("Car_Segment:N", title="Car Segment"),
                y=alt.Y("Units_Sold:Q", title="Units Sold"),
                color=alt.Color("Car_Segment:N", legend=None, scale=alt.Scale(scheme="category10")),
                tooltip=[
                    alt.Tooltip("Car_Segment:N", title="Car Segment"),
                    alt.Tooltip("Units_Sold:Q", title="Units Sold", format=",.0f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(segment_chart, use_container_width=True)

    st.markdown('<div class="section-title">Latest sales records</div>', unsafe_allow_html=True)
    latest_records = (
        filtered.sort_values("Sale_Date", ascending=False)
        .head(8)
        .reset_index(drop=True)
    )
    styled_latest_records = latest_records.style.apply(
        lambda row: [
            "background-color: #eef3f9" if row.name % 2 else "background-color: #ffffff"
            for _ in row
        ],
        axis=1,
    )
    st.dataframe(
        styled_latest_records,
        use_container_width=True,
        hide_index=True
    )

# ---------------- CHATBOT ----------------
def answer_question(question, data):
    q = question.lower().strip()
    if not q:
        return "Please type a question."

    if "total sales" in q or "overall sales" in q or q == "sales":
        return f"Total sales for the selected data are {money(data['Total_Sales'].sum())}."

    if "total orders" in q or "total order" in q or "orders" in q:
        return f"Total orders in the selected data are {data['Sale_ID'].nunique():,.0f}."

    if "total customers" in q or "customer count" in q or "customers" in q:
        return f"Total customers in the selected data are {data['Customer_Rating'].notna().sum():,.0f}."

    if "total car models" in q or "total model of car" in q or "car models" in q or "models of car" in q:
        return f"Total car models in the selected data are {data['Car_Model'].nunique():,.0f}."

    if "units sold" in q or "total units" in q or "how many cars" in q:
        return f"Total units sold are {data['Units_Sold'].sum():,.0f}."

    if "average rating" in q or "customer rating" in q:
        return f"The average customer rating is {data['Customer_Rating'].mean():.2f} out of 5."

    if "average discount" in q or "discount" in q and "average" in q:
        return f"The average discount is {data['Discount_Percent'].mean():.2f}%."

    if "average price" in q or "average unit price" in q:
        return f"The average unit price is {money(data['Unit_Price'].mean())}."

    if ("highest" in q or "top" in q or "best" in q) and ("sales" in q or "manufacturer" in q or "brand" in q):
        totals = data.groupby("Car_Manufacturer")["Total_Sales"].sum().sort_values(ascending=False)
        if len(totals):
            return f"{totals.index[0]} has the highest total sales: {money(totals.iloc[0])}."

    if ("highest" in q or "top" in q or "best" in q) and "region" in q:
        totals = data.groupby("Region")["Total_Sales"].sum().sort_values(ascending=False)
        if len(totals):
            return f"{totals.index[0]} has the highest sales by region: {money(totals.iloc[0])}."

    if ("most" in q or "highest" in q or "top" in q) and "fuel" in q:
        totals = data.groupby("Fuel_Type")["Units_Sold"].sum().sort_values(ascending=False)
        if len(totals):
            return f"{totals.index[0]} has the most units sold: {totals.iloc[0]:,.0f} units."

    # Find a category name typed by the user, e.g. Hyundai, Tamil Nadu, Diesel.
    searchable_columns = [
        "Car_Manufacturer", "Car_Model", "Car_Segment", "Region",
        "State", "City", "Fuel_Type", "Sales_Executive", "Domestic_Export"
    ]
    for column in searchable_columns:
        for value in data[column].dropna().astype(str).unique():
            if value.lower() in q:
                matches = data[data[column].astype(str).str.lower() == value.lower()]
                return (
                    f"For {value}: total sales are {money(matches['Total_Sales'].sum())}, "
                    f"units sold are {matches['Units_Sold'].sum():,.0f}, "
                    f"and there are {len(matches)} records."
                )

    return (
        "I couldn't match that question yet. Try: 'What are total sales?', "
        "'How many units sold?', 'Which manufacturer has the highest sales?', "
        "'What is the average rating?', or ask about a state, region, fuel type, "
        "manufacturer, car model, city, or sales executive."
    )

with tab_chat:
    st.markdown('<div class="section-title">Ask your car sales data</div>', unsafe_allow_html=True)
    st.caption("This beginner chatbot searches and calculates answers from the selected Excel rows.")

    example_questions = [
        "What are total sales?",
        "How many orders?",
        "How many customers?",
        "How many car models?",
        "Which manufacturer has the highest sales?",
        "What is the average rating?",
        "Show sales for Hyundai",
        "Which fuel type has the most units sold?"
    ]
    chosen_example = st.selectbox("Choose an example question (optional)", [""] + example_questions)

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.write(message["content"])

    question = st.chat_input("Ask about manufacturers, regions, fuel types, sales...")
    if chosen_example and st.button("Ask selected example"):
        question = chosen_example

    if question:
        st.session_state.chat_history.append({"role": "user", "content": question})
        reply = answer_question(question, filtered)
        st.session_state.chat_history.append({"role": "assistant", "content": reply})
        st.rerun()

    if st.button("Clear chat"):
        st.session_state.chat_history = []
        st.rerun()

def generate_filtered_pdf(data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, title="AutoInsight Car Sales Data")
    styles = getSampleStyleSheet()
    elements = [
        Paragraph("AutoInsight Car Sales Data", styles["Title"]),
        Spacer(1, 18),
        Paragraph(f"{len(data)} filtered records", styles["BodyText"]),
        Spacer(1, 12),
    ]

    columns = list(data.columns)
    rows = [columns] + data.astype(str).where(~data.isna(), "").values.tolist()
    table = Table(rows)
    table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#101b33")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.7, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
        ])
    )
    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()

# ---------------- DATA TABLE ----------------
with tab_data:
    st.markdown('<div class="section-title">Explore the Excel data</div>', unsafe_allow_html=True)
    st.caption(f"Showing {len(filtered)} of {len(all_data)} records after applying sidebar filters.")
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    csv_button, pdf_button = st.columns(2)
    with csv_button:
        st.download_button(
            "⬇️ Download filtered data as CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name="filtered_car_sales.csv",
            mime="text/csv"
        )
    with pdf_button:
        st.download_button(
            "📄 Download filtered data as PDF",
            data=generate_filtered_pdf(filtered),
            file_name="filtered_car_sales.pdf",
            mime="application/pdf"
        )

st.markdown(
    '<p class="small-note">Portfolio project · Built with Python, Pandas and Streamlit</p>',
    unsafe_allow_html=True
)
