import streamlit as st
import pandas as pd
import numpy as np
from modules.engine import calendarise_growth, calculate_nominal_growth, project_current_tax_head, run_arrears_stock_flow
from modules.compliance_refunds import calculate_net_vat_bridge, adjust_compliance_gap
from modules.report_gen import generate_excel_report, generate_docx_brief

# 1. Page Configuration & Custom Styling
st.set_page_config(page_title="Saint Lucia IRD Revenue Model", page_icon="🇱🇨", layout="wide")

st.markdown("""
    <style>
    .main-header { font-size:26px; font-weight:bold; color:#1B365D; margin-bottom: 5px; }
    .sub-header { font-size:14px; color:#555555; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🇱🇨 Saint Lucia Inland Revenue Department</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Medium-Term Revenue Framework Projections Framework — Production Dashboard</div>', unsafe_allow_html=True)

# 2. Sidebar Configuration (Controls & Core Parameters)
st.sidebar.header("📋 1. Upload Baseline Data")
uploaded_file = st.sidebar.file_uploader("Upload CSV/Excel with baseline FY collections", type=["csv", "xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("📈 2. Scenario Settings")
selected_scenario = st.sidebar.selectbox(
    "Macroeconomic Scenario Profile",
    options=["Baseline", "Downside", "Upside", "Climate Shock"]
)

st.sidebar.markdown("---")
st.sidebar.header("🛡️ 3. Compliance & VAT Refund Controls")
vat_refund_claim_rate = st.sidebar.slider(
    "Expected VAT Refund Claims (% of Gross VAT)",
    min_value=0.0, max_value=30.0, value=12.5, step=0.5
) / 100.0

vat_audit_efficiency = st.sidebar.slider(
    "VAT Refund Audit Recovery Efficiency (%)",
    min_value=0.0, max_value=50.0, value=15.0, step=1.0
) / 100.0

compliance_intervention_gain = st.sidebar.slider(
    "Targeted Efficiency Gained via Enforcement (p.p.)",
    min_value=0.0, max_value=5.0, value=1.0, step=0.1
) / 100.0

# 3. Macroeconomic Framework Configuration (IMF/World Bank Baselines)
macro_data = {
    "Baseline": {
        "real_gdp": [0.02, 0.017, 0.016, 0.016, 0.015],
        "inflation": [0.021, 0.02, 0.02, 0.02, 0.02],
        "arrivals": [0.02, 0.02, 0.02, 0.02, 0.02]
    },
    "Downside": {
        "real_gdp": [0.02, 0.005, 0.01, 0.013, 0.015],
        "inflation": [0.021, 0.03, 0.025, 0.022, 0.02],
        "arrivals": [0.02, -0.04, 0.0, 0.015, 0.02]
    },
    "Upside": {
        "real_gdp": [0.02, 0.025, 0.023, 0.02, 0.018],
        "inflation": [0.021, 0.02, 0.02, 0.02, 0.02],
        "arrivals": [0.02, 0.05, 0.04, 0.03, 0.03]
    },
    "Climate Shock": {
        "real_gdp": [0.02, -0.04, 0.04, 0.025, 0.016],
        "inflation": [0.021, 0.035, 0.025, 0.02, 0.02],
        "arrivals": [0.02, -0.20, 0.15, 0.05, 0.02]
    }
}

scen_macro = macro_data[selected_scenario]

# Process Calendarisation via Weights (75/25 rule)
fy_real_gdp_27_28 = calendarise_growth(scen_macro["real_gdp"][0], scen_macro["real_gdp"][1])
fy_inflation_27_28 = calendarise_growth(scen_macro["inflation"][0], scen_macro["inflation"][1])
nom_gdp_27_28 = calculate_nominal_growth(fy_real_gdp_27_28, fy_inflation_27_28)

fy_real_gdp_28_29 = calendarise_growth(scen_macro["real_gdp"][1], scen_macro["real_gdp"][2])
fy_inflation_28_29 = calendarise_growth(scen_macro["inflation"][1], scen_macro["inflation"][2])
nom_gdp_28_29 = calculate_nominal_growth(fy_real_gdp_28_29, fy_inflation_28_29)

fy_real_gdp_29_30 = calendarise_growth(scen_macro["real_gdp"][2], scen_macro["real_gdp"][3])
fy_inflation_29_30 = calendarise_growth(scen_macro["inflation"][2], scen_macro["inflation"][3])
nom_gdp_29_30 = calculate_nominal_growth(fy_real_gdp_29_30, fy_inflation_29_30)

drivers_map = {
    "Nominal GDP": {2027: nom_gdp_27_28, 2028: nom_gdp_28_29, 2029: nom_gdp_29_30},
    "Consumption": {2027: nom_gdp_27_28, 2028: nom_gdp_28_29, 2029: nom_gdp_29_30},
    "Wage bill": {2027: nom_gdp_27_28, 2028: nom_gdp_28_29, 2029: nom_gdp_29_30},
    "Lagged profits": {2027: 0.040399, 2028: nom_gdp_27_28, 2029: nom_gdp_28_29},
    "Flat": {2027: 0.0, 2028: 0.0, 2029: 0.0}
}

# 4. Handle Incoming Spreadsheet Data
tax_heads_data = []
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.xlsx'):
            user_data = pd.read_excel(uploaded_file)
        else:
            user_data = pd.read_csv(uploaded_file)
            
        user_data.columns = [c.strip().title() for c in user_data.columns]
        
        for idx, row in user_data.iterrows():
            th_name = str(row.get("Tax Head", ""))
            driver = "Nominal GDP"
            th_type = "Current"
            
            if "PERSONAL" in th_name.upper():
                driver = "Wage bill"
                th_type = "Arrears" if "ARREARS" in th_name.upper() else "Current"
            elif "CORPORATE" in th_name.upper():
                driver = "Lagged profits"
                th_type = "Arrears" if "ARREARS" in th_name.upper() else "Current"
            elif "VAT" in th_name.upper() or "VALUE ADDED" in th_name.upper():
                driver = "Consumption"
                th_type = "VAT"
            elif "LICENCE" in th_name.upper() or "FLAT" in th_name.upper():
                driver = "Flat"
                
            tax_heads_data.append({
                "Code": str(row.get("Code", f"9999 {idx}")),
                "Tax Head": th_name,
                "Type": th_type,
                "Baseline": float(row.get("Baseline", 0)),
                "Driver": driver,
                "Buoyancy": 1.0
            })
    except Exception as e:
        st.error(f"Error parsing structural layout: {e}")
        st.stop()
else:
    st.info("💡 Upload your baseline file via the sidebar. Displaying baseline reference layout below:")
    tax_heads_data = [
        {"Code": "3111 001", "Tax Head": "Personal Income Tax", "Type": "Current", "Baseline": 142506159, "Driver": "Wage bill", "Buoyancy": 1.0},
        {"Code": "3111 002", "Tax Head": "Arrears on Personal Income Tax", "Type": "Arrears", "Baseline": 29652094, "Driver": "Flat", "Buoyancy": 0.0},
        {"Code": "3112 001", "Tax Head": "Corporate Income Tax", "Type": "Current", "Baseline": 167348098, "Driver": "Lagged profits", "Buoyancy": 1.0},
        {"Code": "3141 001", "Tax Head": "Value Added Tax (VAT)", "Type": "VAT", "Baseline": 237125609, "Driver": "Consumption", "Buoyancy": 1.0}
    ]

# 5. Execute Projections Calculations Loops
output_rows = []
total_b, total_y1, total_y2, total_y3 = 0.0, 0.0, 0.0, 0.0

arrears_config = {
    "3111 002": {"opening": 148260470, "base_col": 0.20, "write_off": 0.02, "accrual": 0.22888}
}

for th in tax_heads_data:
    b_val = th["Baseline"]
    dr = th["Driver"]
    
    if th["Type"] == "Current":
        y1 = project_current_tax_head(b_val, drivers_map[dr][2027], th["Buoyancy"])
        y2 = project_current_tax_head(y1, drivers_map[dr][2028], th["Buoyancy"])
        y3 = project_current_tax_head(y2, drivers_map[dr][2029], th["Buoyancy"])
        
        if "PERSONAL" in th["Tax Head"].upper():
            y1 = adjust_compliance_gap(y1, baseline_tax_gap=0.25, policy_efficiency_gain=compliance_intervention_gain)
            y2 = adjust_compliance_gap(y2, baseline_tax_gap=0.25, policy_efficiency_gain=compliance_intervention_gain)
            y3 = adjust_compliance_gap(y3, baseline_tax_gap=0.25, policy_efficiency_gain=compliance_intervention_gain)
            
    elif th["Type"] == "VAT":
        raw_y1 = project_current_tax_head(b_val, drivers_map[dr][2027], th["Buoyancy"])
        raw_y2 = project_current_tax_head(raw_y1, drivers_map[dr][2028], th["Buoyancy"])
        raw_y3 = project_current_tax_head(raw_y2, drivers_map[dr][2029], th["Buoyancy"])
        
        y1 = calculate_net_vat_bridge(raw_y1, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]
        y2 = calculate_net_vat_bridge(raw_y2, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]
        y3 = calculate_net_vat_bridge(raw_y3, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]
        
    else:
        cfg = arrears_config.get(th["Code"], {"opening": b_val*5, "base_col": 0.20, "write_off": 0.02, "accrual": 0.20})
        res1 = run_arrears_stock_flow(cfg["opening"], cfg["base_col"], cfg["write_off"], cfg["accrual"], current_tax_collections=b_val)
        y1 = res1["collections"]
        res2 = run_arrears_stock_flow(res1["closing_stock"], cfg["base_col"], cfg["write_off"], cfg["accrual"], current_tax_collections=y1)
        y2 = res2["collections"]
        res3 = run_arrears_stock_flow(res2["closing_stock"], cfg["base_col"], cfg["write_off"], cfg["accrual"], current_tax_collections=y2)
        y3 = res3["collections"]

    cagr = np.power((y3 / b_val), 1/3) - 1.0 if b_val > 0 else 0.0
    
    total_b += b_val
    total_y1 += y1
    total_y2 += y2
    total_y3 += y3
    
    output_rows.append({
        "Code": th["Code"], "Tax Head": th["Tax Head"], "Baseline 2026/27": b_val,
        "Projected 2027/28": y1, "Projected 2028/29": y2, "Projected 2029/30": y3,
        "Avg Annual Growth": cagr
    })

grand_total_cagr = np.power((total_y3 / total_b), 1/3) - 1.0 if total_b > 0 else 0.0
output_rows.append({
    "Code": "", "Tax Head": "GRAND TOTAL REVENUE", "Baseline 2026/27": total_b,
    "Projected 2027/28": total_y1, "Projected 2028/29": total_y2, "Projected 2029/30": total_y3,
    "Avg Annual Growth": grand_total_cagr
})

final_df = pd.DataFrame(output_rows)

# 6. Display Analytics Display Layout
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label=f"Projected Total Yield ({selected_scenario}) FY27/28", value=f"EC$ {total_y1:,.0f}")
with col2:
    st.metric(label="FY2029/30 Horizon Yield", value=f"EC$ {total_y3:,.0f}")
with col3:
    st.metric(label="Horizon Integrated CAGR", value=f"{grand_total_cagr*100:.2f}%")

st.subheader("📊 Structured Projections Evaluation Matrix")
formatted_df = final_df.copy()
# 6. Display Analytics Display Layout
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label=f"Projected Total Yield ({selected_scenario}) FY27/28", value=f"EC$ {total_y1:,.0f}")
with col2:
    st.metric(label="FY2029/30 Horizon Yield", value=f"EC$ {total_y3:,.0f}")
with col3:
    st.metric(label="Horizon Integrated CAGR", value=f"{grand_total_cagr*100:.2f}%")

st.subheader("📊 Structured Projections Evaluation Matrix")
formatted_df = final_df.copy()

for col in ["Baseline 2026/27", "Projected 2027/28", "Projected 2028/29", "Projected 2029/30"]:
    formatted_df[col] = formatted_df[col].map(lambda x: f"EC$ {x:,.0f}")
formatted_df["Avg Annual Growth"] = formatted_df["Avg Annual Growth"].map(lambda x: f"{x*100:.2f}%")

st.dataframe(formatted_df, use_container_width=True)

# 7. Document Downloads Section
st.subheader("📥 Export Institutional Analysis Documents")
btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    excel_path = generate_excel_report(final_df, selected_scenario)
    with open(excel_path, "rb") as f:
        st.download_button(
            label="Download Excel Worksheet Summary Report",
            data=f, 
            file_name=f"IRD_Projections_Sheet_{selected_scenario}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with btn_col2:
    summary_stats = {"baseline_total": total_b, "year3_total": total_y3}
    docx_path = generate_docx_brief(summary_stats, selected_scenario)
    with open(docx_path, "rb") as f:
        st.download_button(
            label="Download Executive Briefing Memorandum (.docx)",
            data=f, 
            file_name=f"IRD_Methodological_Brief_{selected_scenario}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
