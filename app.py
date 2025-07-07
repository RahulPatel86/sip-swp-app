import streamlit as st
import pandas as pd
from num2words import num2words
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# --- SIP Calculation ---
def calculate_sip(principal, annual_rate, years):
    monthly_rate = annual_rate / 12
    months = years * 12
    data = []
    opening = 0
    for month in range(1, months + 1):
        interest = (opening + principal) * monthly_rate
        closing = opening + principal + interest
        data.append({
            'Month': month,
            'Opening Balance (₹)': round(opening, 2),
            'Monthly Investment (₹)': principal,
            'Interest Earned (₹)': round(interest, 2),
            'Closing Balance (₹)': round(closing, 2)
        })
        opening = closing
    df = pd.DataFrame(data)
    return df, principal * months, round(opening, 2)

# --- SWP Calculation ---
def calculate_swp(initial_corpus, withdrawal, annual_rate, years):
    monthly_rate = annual_rate / 12
    months = years * 12
    data = []
    opening = initial_corpus
    for month in range(1, months + 1):
        interest = opening * monthly_rate
        closing = opening + interest - withdrawal
        data.append({
            'Month': month,
            'Opening Balance (₹)': round(opening, 2),
            'Interest Earned (₹)': round(interest, 2),
            'Monthly Withdrawal (₹)': withdrawal,
            'Closing Balance (₹)': round(closing, 2)
        })
        if closing <= 0:
            break
        opening = closing
    df = pd.DataFrame(data)
    return df, round(opening, 2)

# --- Excel Export Function ---
def generate_excel(sip_df, swp_df, summary_df):
    output = BytesIO()
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "SIP Calculation"
    for r in dataframe_to_rows(sip_df, index=False, header=True):
        ws1.append(r)

    chart = LineChart()
    chart.title = "SIP Closing Balance"
    chart.y_axis.title = "₹"
    chart.x_axis.title = "Month"
    data = Reference(ws1, min_col=5, min_row=1, max_row=ws1.max_row)
    cats = Reference(ws1, min_col=1, min_row=2, max_row=ws1.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws1.add_chart(chart, f"A{ws1.max_row + 3}")

    ws2 = wb.create_sheet("SWP Plan")
    for r in dataframe_to_rows(swp_df, index=False, header=True):
        ws2.append(r)

    ws3 = wb.create_sheet("Summary")
    for r in dataframe_to_rows(summary_df, index=False, header=True):
        ws3.append(r)

    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.title("📊 SIP + SWP Excel Generator")

with st.form("inputs"):
    st.subheader("SIP Details")
    sip_amt = st.number_input("Monthly SIP Investment (₹)", value=30000)
    sip_years = st.number_input("SIP Duration (Years)", value=8, step=1)
    sip_return = st.number_input("SIP Annual Return (%)", value=12.0)

    st.subheader("SWP Details")
    swp_amt = st.number_input("Monthly SWP Withdrawal (₹)", value=40000)
    swp_years = st.number_input("SWP Duration (Years)", value=25, step=1)
    swp_return = st.number_input("SWP Annual Return (%)", value=8.0)

    submitted = st.form_submit_button("Generate Excel")

if submitted:
    sip_df, sip_invested, sip_final = calculate_sip(sip_amt, sip_return / 100, sip_years)
    swp_df, swp_remaining = calculate_swp(sip_final, swp_amt, swp_return / 100, swp_years)

    sip_final_words = num2words(sip_final, lang='en_IN', to='currency').replace("euro", "rupees").replace("cents", "paise")
    swp_remaining_words = num2words(swp_remaining, lang='en_IN', to='currency').replace("euro", "rupees").replace("cents", "paise")

    summary_df = pd.DataFrame({
        "Metric": [
            "SIP Monthly Investment", "SIP Duration", "SIP Annual Return",
            "Total SIP Invested", "Final Corpus (SIP)", "Corpus in Words",
            "SWP Monthly Withdrawal", "SWP Duration", "SWP Annual Return",
            "Remaining Corpus after SWP", "Remaining Corpus in Words"
        ],
        "Value": [
            f"₹{sip_amt:,.2f}", f"{sip_years} years", f"{sip_return:.2f}%",
            f"₹{sip_invested:,.2f}", f"₹{sip_final:,.2f}", sip_final_words,
            f"₹{swp_amt:,.2f}", f"{swp_years} years", f"{swp_return:.2f}%",
            f"₹{swp_remaining:,.2f}", swp_remaining_words
        ]
    })

    excel_file = generate_excel(sip_df, swp_df, summary_df)
    st.success("✅ Excel file generated!")
    st.download_button("📥 Download Excel", excel_file, file_name="SIP_SWP_Report.xlsx")

