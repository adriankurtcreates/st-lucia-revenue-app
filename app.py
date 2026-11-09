import streamlit as st
import pandas as pd
from modules.data_fetcher import fetch_macro_forecasts
from modules.compliance_refunds import calculate_net_vat_bridge

st.set_page_config(page_title="LCA Inland Revenue Projections", layout="wide")
st.title("🇱🇨 IRD Revenue Projection Model (Streamlit Version)")

st.sidebar.header("1. Upload Current Year Outturns")
uploaded_file = st.sidebar.file_uploader("Upload CSV/Excel with baseline FY collections", type=["csv", "xlsx"])

st.sidebar.header("2. Compliance & Refund Adjustments")
vat_refund_rate = st.sidebar.slider("Expected VAT Refund Claims (% of Gross)", 0.0, 30.0, 12.5) / 100
audit_recovery = st.sidebar.slider("Audit Detection/Disallowance Efficiency", 0.0, 50.0, 15.0) / 100

# Main Panel layout
col1, col2 = st.columns(2)
with col1:
    st.subheader("Automated External Macro Data")
    fetched_inflation = fetch_macro_forecasts()
    st.json(fetched_inflation)

with col2:
    st.subheader("Scenario Comparison Grid")
    # Interactive grid displaying Baseline vs Stress Scenarios
