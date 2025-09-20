# report.py - Client Asset Report Generator with Financial Year Grouping
import os
import sys
import base64
import re
import traceback
from io import BytesIO
from datetime import date
import zipfile

import numpy as np
import streamlit as st
import pandas as pd
from jinja2 import Template

# ------------------ Streamlit Setup ------------------
st.set_page_config(page_title="Client Asset Report Generator", layout="wide")
st.title("📊 Client Asset Report Generator (4-Page PDF)")

# ✅ Startup message (so Cloud shows UI instantly)
st.success("✅ App started successfully — waiting for file upload...")

# ------------------ Template Directory ------------------
if getattr(sys, 'frozen', False):
    TEMPLATE_DIR = sys._MEIPASS  # PyInstaller bundle
else:
    TEMPLATE_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------ Helpers ------------------
def load_template(filename: str) -> Template:
    path = os.path.join(TEMPLATE_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        return Template(f.read())

def clean_number(x):
    try:
        x = str(x).replace("₹", "").replace(",", "").strip()
        return float(x) if x not in ["", "nan", "None"] else 0.0
    except Exception:
        return 0.0

def fig_to_base64_png(fig) -> str:
    import matplotlib.pyplot as plt
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150, transparent=True)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

def format_indian_currency(amount):
    amount = float(amount)
    if amount == 0:
        return "₹ 0.00"

    amount_str = f"{amount:.2f}"
    if '.' in amount_str:
        integer_part, decimal_part = amount_str.split('.')
    else:
        integer_part, decimal_part = amount_str, "00"

    if len(integer_part) <= 3:
        formatted = integer_part
    else:
        last_three = integer_part[-3:]
        remaining = integer_part[:-3]
        groups = []
        while len(remaining) > 2:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.append(remaining)
        groups.reverse()
        formatted = ','.join(groups) + ',' + last_three

    return f"₹ {formatted}.{decimal_part}"

def extract_client_name_from_filename(filename):
    filename_no_ext = os.path.splitext(filename)[0]
    parts = filename_no_ext.split("_")
    if len(parts) >= 4:
        client_name_parts = parts[2:-1]
        return " ".join(client_name_parts).title()
    return "Client"

def get_appreciation_text(excel_file):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        ws = wb["Performance Report"]

        last_row = ws.max_row
        while not ws.cell(row=last_row, column=1).value:
            last_row -= 1

        absolute_appreciation = ws.cell(row=last_row, column=3).value
        percentage_appreciation = ws.cell(row=last_row, column=4).value

        abs_text = ""
        pct_text = ""

        if absolute_appreciation and absolute_appreciation != 0:
            if isinstance(absolute_appreciation, (int, float)):
                abs_text = format_indian_currency(absolute_appreciation)
            else:
                abs_text = str(absolute_appreciation)

        if percentage_appreciation and percentage_appreciation != 0:
            if isinstance(percentage_appreciation, (int, float)):
                pct_text = f"{percentage_appreciation:.2%}"
            else:
                pct_text = str(percentage_appreciation)

        if abs_text or pct_text:
            return (
                "<p style='font-style: italic; color: black; text-align: center;'>"
                "YOUR PORTFOLIO INVESTMENT VALUE HAS INCREASED BY "
                f"<span style='color:green;'>{abs_text}</span> OR "
                f"<span style='color:green;'>{pct_text}</span> "
                "ON ACCOUNT OF MARKET APPRECIATION."
                "</p>"
            )
        else:
            return "<i>YOUR PORTFOLIO INVESTMENT PERFORMANCE SUMMARY</i>"

    except Exception as e:
        st.error(f"Error reading appreciation data: {e}")
        return "<i>YOUR PORTFOLIO INVESTMENT PERFORMANCE SUMMARY</i>"

# ------------------ Chart Generators ------------------
def generate_performance_chart(excel_file):
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker
    import matplotlib.dates as mdates
    from matplotlib import font_manager as fm

    try:
        cap_df = pd.read_excel(excel_file, sheet_name="Capital Contribution", usecols="A,B,D")
        cap_df.columns = ["Date", "Amount Added", "Amount Withdrawn"]
        cap_df["Amount Added"] = cap_df["Amount Added"].apply(lambda x: 0 if pd.isna(x) else x)
        cap_df["Amount Withdrawn"] = cap_df["Amount Withdrawn"].apply(lambda x: 0 if pd.isna(x) else x)
        cap_df = cap_df[(cap_df["Amount Added"] != 0) | (cap_df["Amount Withdrawn"] != 0)]
        cap_df["Date"] = pd.to_datetime(cap_df["Date"])

        def get_financial_year(dt):
            if dt.month >= 4:
                return f"FY {dt.year}-{str(dt.year + 1)[-2:]}"
            else:
                return f"FY {dt.year - 1}-{str(dt.year)[-2:]}"

        cap_df["Financial Year"] = cap_df["Date"].apply(get_financial_year)

        grouped_df = cap_df.groupby("Financial Year").agg({
            "Date": "max",
            "Amount Added": "sum",
            "Amount Withdrawn": "sum"
        }).reset_index()
        grouped_df["Net Amount"] = grouped_df["Amount Added"] - grouped_df["Amount Withdrawn"]
        grouped_df = grouped_df[grouped_df["Net Amount"] != 0]
        grouped_df = grouped_df.sort_values("Date")
        grouped_df["Cumulative"] = grouped_df["Net Amount"].cumsum()

        perf_df = pd.read_excel(excel_file, sheet_name="Performance Report", usecols="A:B")
        perf_df.columns = ["Date", "Portfolio Value"]
        perf_df["Date"] = pd.to_datetime(perf_df["Date"])
        latest_row = perf_df.iloc[-1]
        latest_date = latest_row["Date"]
        latest_value = latest_row["Portfolio Value"]

        chart_df = grouped_df[["Date", "Cumulative"]].copy()
        chart_df = pd.concat([chart_df, pd.DataFrame({"Date": [latest_date], "Cumulative": [latest_value]})])
        chart_df = chart_df.sort_values("Date").reset_index(drop=True)

        plt.rcParams['font.family'] = 'Open Sans'
        fig, ax = plt.subplots(figsize=(9, 6))
        ax.yaxis.grid(True, linestyle="--", alpha=0.3)

        plt.fill_between(chart_df["Date"], chart_df["Cumulative"], color="#FFA366", alpha=0.9)
        plt.plot(chart_df["Date"], chart_df["Cumulative"], marker="o", color="#CC5500", linewidth=2)

        # ... (keep the rest of your plotting logic unchanged)

        return fig_to_base64_png(fig)

    except Exception as e:
        st.error(f"Error generating performance chart: {e}")
        st.error(f"Detailed error: {traceback.format_exc()}")
        return ""

def generate_comparison_chart(excel_file):
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker
    import openpyxl

    try:
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        ws = wb["Performance Report"]

        # ... (keep your comparison chart logic unchanged)

        return fig_to_base64_png(fig)

    except Exception as e:
        st.error(f"Error generating comparison chart: {e}")
        st.error(f"Detailed error: {traceback.format_exc()}")
        return ""

# ------------------ PDF Builder ------------------
def build_client_pdf_bytes(excel_file, client_name: str, report_dt: date, client_name_font_size: int = 45) -> bytes:
    from weasyprint import HTML
    perf_chart_b64 = generate_performance_chart(excel_file)
    comp_chart_b64 = generate_comparison_chart(excel_file)
    appreciation_text = get_appreciation_text(excel_file)
    if not perf_chart_b64 or not comp_chart_b64:
        raise Exception("Failed to generate charts")
    full_html = build_4page_html(client_name, report_dt, perf_chart_b64, comp_chart_b64, appreciation_text, client_name_font_size)
    pdf_bytes = HTML(string=full_html, base_url=TEMPLATE_DIR).write_pdf()
    return pdf_bytes

# ------------------ UI ------------------
uploaded_file = st.file_uploader("Upload Client Excel File", type=["xlsx"])
report_date = st.date_input("📅 Select Report Date", value=date.today())
client_name_font_size = st.slider("Client Name Font Size (Cover Page)", 20, 80, 45, 1)

st.caption(f"Templates folder: {TEMPLATE_DIR}")
st.write("---")

if uploaded_file:
    try:
        client_name = extract_client_name_from_filename(uploaded_file.name)
        st.success(f"✅ Client identified: {client_name}")

        appreciation_text = get_appreciation_text(uploaded_file)
        st.info(f"📈 Appreciation Text: {appreciation_text}")

        pdf_bytes = build_client_pdf_bytes(uploaded_file, client_name, report_date, client_name_font_size)

        st.download_button(
            "📥 Download Client Report (PDF)",
            data=pdf_bytes,
            file_name=f"{client_name}_Asset_Report.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error generating report: {e}")
        st.error(f"Detailed error: {traceback.format_exc()}")
