import numpy as np
import pandas as pd
import streamlit as st

from modules.engine import (
    calendarise_growth,
    calculate_nominal_growth,
    project_current_tax_head,
    run_arrears_stock_flow,
)
from modules.compliance_refunds import calculate_net_vat_bridge, adjust_compliance_gap
from modules.report_gen import generate_excel_report, generate_docx_brief

# ---------------------------------------------------------------------------
# 1. Page Configuration & Custom Styling
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Saint Lucia IRD Revenue Model", page_icon="🇱🇨", layout="wide")

st.markdown("""
    <style>
    .main-header { font-size:26px; font-weight:bold; color:#1B365D; margin-bottom: 5px; }
    .sub-header { font-size:14px; color:#555555; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🇱🇨 Saint Lucia Inland Revenue Department</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Medium-Term Revenue Framework Projections Framework — Production Dashboard</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# 2. Sidebar Configuration
# ---------------------------------------------------------------------------
st.sidebar.header("📋 1. Upload Baseline Data")
uploaded_file = st.sidebar.file_uploader("Upload CSV/Excel with baseline FY collections", type=["csv", "xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("📈 2. Scenario Settings")
selected_scenario = st.sidebar.selectbox(
    "Macroeconomic Scenario Profile",
    options=["Baseline", "Downside", "Upside", "Climate Shock"],
)

st.sidebar.markdown("---")
st.sidebar.header("🛡️ 3. Compliance & VAT Refund Controls")
vat_refund_claim_rate = st.sidebar.slider(
    "Expected VAT Refund Claims (% of Gross VAT)",
    min_value=0.0, max_value=30.0, value=12.5, step=0.5,
) / 100.0

vat_audit_efficiency = st.sidebar.slider(
    "VAT Refund Audit Recovery Efficiency (%)",
    min_value=0.0, max_value=50.0, value=15.0, step=1.0,
) / 100.0

compliance_intervention_gain = st.sidebar.slider(
    "Targeted Efficiency Gained via Enforcement (p.p.)",
    min_value=0.0, max_value=5.0, value=1.0, step=0.1,
) / 100.0

# ---------------------------------------------------------------------------
# 3. Macroeconomic Framework Configuration (IMF/World Bank Baselines)
# ---------------------------------------------------------------------------
macro_data = {
    "Baseline": {
        "real_gdp": [0.02, 0.017, 0.016, 0.016, 0.015],
        "inflation": [0.021, 0.02, 0.02, 0.02, 0.02],
        "arrivals": [0.02, 0.02, 0.02, 0.02, 0.02],
    },
    "Downside": {
        "real_gdp": [0.02, 0.005, 0.01, 0.013, 0.015],
        "inflation": [0.021, 0.03, 0.025, 0.022, 0.02],
        "arrivals": [0.02, -0.04, 0.0, 0.015, 0.02],
    },
    "Upside": {
        "real_gdp": [0.02, 0.025, 0.023, 0.02, 0.018],
        "inflation": [0.021, 0.02, 0.02, 0.02, 0.02],
        "arrivals": [0.02, 0.05, 0.04, 0.03, 0.03],
    },
    "Climate Shock": {
        "real_gdp": [0.02, -0.04, 0.04, 0.025, 0.016],
        "inflation": [0.021, 0.035, 0.025, 0.02, 0.02],
        "arrivals": [0.02, -0.20, 0.15, 0.05, 0.02],
    },
}

scen_macro = macro_data[selected_scenario]

# FIX: build drivers_map with a loop so years are consistent and no copy-paste drift
FY_YEARS = [2027, 2028, 2029]
fy_nominal = {}
for i, fy_year in enumerate(FY_YEARS):
    fy_real = calendarise_growth(scen_macro["real_gdp"][i], scen_macro["real_gdp"][i + 1])
    fy_infl = calendarise_growth(scen_macro["inflation"][i], scen_macro["inflation"][i + 1])
    fy_nominal[fy_year] = calculate_nominal_growth(fy_real, fy_infl)

drivers_map = {
    "Nominal GDP": dict(fy_nominal),
    "Consumption": dict(fy_nominal),
    "Wage bill": dict(fy_nominal),
    # FIX: Lagged profits = prior FY nominal GDP growth, with a documented first-year seed
    # equal to the FY2026/27 nominal growth implied by the first CY pair.
    "Lagged profits": {
        FY_YEARS[0]: fy_nominal[FY_YEARS[0]],  # seed with same-year nominal growth (documented)
        FY_YEARS[1]: fy_nominal[FY_YEARS[0]],
        FY_YEARS[2]: fy_nominal[FY_YEARS[1]],
    },
    "Flat": {y: 0.0 for y in FY_YEARS},
}

# ---------------------------------------------------------------------------
# 4. Upload parsing helpers
# ---------------------------------------------------------------------------
REQUIRED_COLS = {"Code", "Tax Head", "Baseline"}


def _read_any(uploaded):
    """Tolerant reader: handles xlsx, and CSV in utf-8/cp1252/latin-1 with , ; or tab."""
    name = uploaded.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded)
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        for sep in (",", ";", "\t"):
            try:
                uploaded.seek(0)
                df = pd.read_csv(uploaded, encoding=enc, sep=sep)
                if df.shape[1] > 1:
                    return df
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
    uploaded.seek(0)
    return pd.read_csv(uploaded)


def _to_float(val, default=0.0):
    """Coerce any uploaded cell to float without ever raising."""
    if pd.isna(val):
        return default
    if isinstance(val, (int, float, np.integer, np.floating)):
        return float(val)
    s = str(val).strip().replace(",", "").replace("EC$", "").replace("$", "")
    if s in {"", "-", "—", "n/a", "N/A", "na", "None"}:
        return default
    try:
        return float(s)
    except ValueError:
        return default


def _classify_tax_head(th_name: str):
    """Return (type, driver) using ordered, explicit rules. Order matters."""
    name = th_name.upper()
    is_arrears = "ARREARS" in name

    if "VALUE ADDED" in name or "VAT" in name:
        # FIX: VAT arrears should be treated as arrears, not as VAT bridge
        return ("Arrears" if is_arrears else "VAT", "Consumption")
    if "PERSONAL" in name:
        return ("Arrears" if is_arrears else "Current", "Wage bill")
    if "CORPORATE" in name:
        return ("Arrears" if is_arrears else "Current", "Lagged profits")
    if "LICENCE" in name or "LICENSE" in name or "FLAT" in name:
        return ("Arrears" if is_arrears else "Current", "Flat")
    return ("Arrears" if is_arrears else "Current", "Nominal GDP")


# ---------------------------------------------------------------------------
# 5. Handle Incoming Spreadsheet Data
# ---------------------------------------------------------------------------
tax_heads_data = []

if uploaded_file is not None:
    try:
        raw = _read_any(uploaded_file)
        raw.columns = [c.strip() for c in raw.columns]

        # FIX: case-insensitive column match, then canonicalise names
        col_lookup = {c.lower(): c for c in raw.columns}
        canonical = {}
        for want in REQUIRED_COLS:
            if want.lower() in col_lookup:
                canonical[col_lookup[want.lower()]] = want
            else:
                st.error(
                    f"Uploaded file is missing required column '{want}'. "
                    f"Columns found: {list(raw.columns)}. "
                    f"Required: {sorted(REQUIRED_COLS)}."
                )
                st.stop()
        raw = raw.rename(columns=canonical)

        for idx, row in raw.iterrows():
            th_name = str(row.get("Tax Head", "")).strip()
            if not th_name:
                continue

            th_type, driver = _classify_tax_head(th_name)

            tax_heads_data.append({
                "Code": str(row.get("Code", f"UNMAPPED-{idx:03d}")).strip(),
                "Tax Head": th_name,
                "Type": th_type,
                "Baseline": _to_float(row.get("Baseline", 0)),
                "Driver": driver,
                "Buoyancy": 1.0,
            })

        if not tax_heads_data:
            st.error("Uploaded file parsed successfully but contained no tax head rows.")
            st.stop()

    except Exception as e:
        # FIX: surface the exception type so the user can diagnose the upload
        st.error(f"Error parsing uploaded file: {type(e).__name__}: {e}")
        st.stop()
else:
    st.info("💡 Upload your baseline file via the sidebar. Displaying baseline reference layout below:")
    tax_heads_data = [
        {"Code": "3111 001", "Tax Head": "Personal Income Tax", "Type": "Current", "Baseline": 142506159, "Driver": "Wage bill", "Buoyancy": 1.0},
        {"Code": "3111 002", "Tax Head": "Arrears on Personal Income Tax", "Type": "Arrears", "Baseline": 29652094, "Driver": "Flat", "Buoyancy": 0.0},
        {"Code": "3112 001", "Tax Head": "Corporate Income Tax", "Type": "Current", "Baseline": 167348098, "Driver": "Lagged profits", "Buoyancy": 1.0},
        {"Code": "3141 001", "Tax Head": "Value Added Tax (VAT)", "Type": "VAT", "Baseline": 237125609, "Driver": "Consumption", "Buoyancy": 1.0},
    ]

# ---------------------------------------------------------------------------
# 6. Execute Projections
# ---------------------------------------------------------------------------
arrears_config = {
    "3111 002": {"opening": 148260470, "base_col": 0.20, "write_off": 0.02, "accrual": 0.22888},
}

# Map each arrears code to its underlying current-tax code (for accrual basis)
ARREARS_LINK = {
    "3111 002": "3111 001",  # PIT arrears accrue off PIT current collections
}

# FIX: apply compliance gain ONCE to the baseline, then project forward normally.
# Previously the gain was compounded across all three years.
current_code_revenue = {}  # code -> {year: value}
output_rows = []
total_b, total_y1, total_y2, total_y3 = 0.0, 0.0, 0.0, 0.0

# Pass 1: project all CURRENT and VAT rows first, so arrears can reference them
for th in tax_heads_data:
    if th["Type"] == "Arrears":
        continue

    b_val = th["Baseline"]
    dr = th["Driver"]

    if th["Type"] == "Current":
        # FIX: apply compliance gain once on the baseline level, not on every year
        if "PERSONAL" in th["Tax Head"].upper():
            adjusted_baseline = adjust_compliance_gap(
                b_val,
                baseline_tax_gap=0.25,
                policy_efficiency_gain=compliance_intervention_gain,
            )
        else:
            adjusted_baseline = b_val

        y1 = project_current_tax_head(adjusted_baseline, drivers_map[dr][FY_YEARS[0]], th["Buoyancy"])
        y2 = project_current_tax_head(y1, drivers_map[dr][FY_YEARS[1]], th["Buoyancy"])
        y3 = project_current_tax_head(y2, drivers_map[dr][FY_YEARS[2]], th["Buoyancy"])

    elif th["Type"] == "VAT":
        raw_y1 = project_current_tax_head(b_val, drivers_map[dr][FY_YEARS[0]], th["Buoyancy"])
        raw_y2 = project_current_tax_head(raw_y1, drivers_map[dr][FY_YEARS[1]], th["Buoyancy"])
        raw_y3 = project_current_tax_head(raw_y2, drivers_map[dr][FY_YEARS[2]], th["Buoyancy"])

        y1 = calculate_net_vat_bridge(raw_y1, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]
        y2 = calculate_net_vat_bridge(raw_y2, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]
        y3 = calculate_net_vat_bridge(raw_y3, vat_refund_claim_rate, vat_audit_efficiency)["net_vat"]

    else:
        continue  # unreachable, safety

    current_code_revenue[th["Code"]] = {
        "baseline": b_val, "y1": y1, "y2": y2, "y3": y3,
        "row": {
            "Code": th["Code"], "Tax Head": th["Tax Head"], "Baseline 2026/27": b_val,
            "Projected 2027/28": y1, "Projected 2028/29": y2, "Projected 2029/30": y3,
        },
    }

# Pass 2: project arrears using the correct current-tax accrual basis
for th in tax_heads_data:
    if th["Type"] != "Arrears":
        continue

    code = th["Code"]
    b_val = th["Baseline"]

    if code in arrears_config:
        cfg = arrears_config[code]
    else:
        cfg = {"opening": b_val * 5, "base_col": 0.20, "write_off": 0.02, "accrual": 0.20}
        st.warning(
            f"Arrears row '{th['Tax Head']}' (Code {code}) has no configured opening stock. "
            f"Using 5× baseline as a placeholder — supply real arrears data for audit-grade output."
        )

    linked_code = ARREARS_LINK.get(code)
    linked_current = current_code_revenue.get(linked_code)

    # FIX: pass the underlying CURRENT tax head collections, not prior-year arrears collections
    if linked_current is not None:
        accr1, accr2, accr3 = linked_current["y1"], linked_current["y2"], linked_current["y3"]
    else:
        if linked_code:
            st.warning(f"Arrears code {code} links to current code {linked_code}, which was not "
                       f"found in the upload. Using baseline as accrual basis.")
        accr1 = accr2 = accr3 = b_val

    res1 = run_arrears_stock_flow(cfg["opening"], cfg["base_col"], cfg["write_off"],
                                  cfg["accrual"], current_tax_collections=accr1)
    y1 = res1["collections"]
    res2 = run_arrears_stock_flow(res1["closing_stock"], cfg["base_col"], cfg["write_off"],
                                  cfg["accrual"], current_tax_collections=accr2)
    y2 = res2["collections"]
    res3 = run_arrears_stock_flow(res2["closing_stock"], cfg["base_col"], cfg["write_off"],
                                  cfg["accrual"], current_tax_collections=accr3)
    y3 = res3["collections"]

    output_rows.append({
        "Code": code, "Tax Head": th["Tax Head"], "Baseline 2026/27": b_val,
        "Projected 2027/28": y1, "Projected 2028/29": y2, "Projected 2029/30": y3,
        "Avg Annual Growth": (np.power((y3 / b_val), 1 / 3) - 1.0) if b_val > 0 else 0.0,
    })
    total_b += b_val
    total_y1 += y1
    total_y2 += y2
    total_y3 += y3

# Merge current / VAT rows into output
for code, payload in current_code_revenue.items():
    row = payload["row"]
    b_val = payload["baseline"]
    y3 = payload["y3"]
    row["Avg Annual Growth"] = (np.power((y3 / b_val), 1 / 3) - 1.0) if b_val > 0 else 0.0
    output_rows.append(row)
    total_b += b_val
    total_y1 += payload["y1"]
    total_y2 += payload["y2"]
    total_y3 += payload["y3"]

grand_total_cagr = (np.power((total_y3 / total_b), 1 / 3) - 1.0) if total_b > 0 else 0.0
output_rows.append({
    "Code": "", "Tax Head": "GRAND TOTAL REVENUE", "Baseline 2026/27": total_b,
    "Projected 2027/28": total_y1, "Projected 2028/29": total_y2, "Projected 2029/30": total_y3,
    "Avg Annual Growth": grand_total_cagr,
})

final_df = pd.DataFrame(output_rows)

# ---------------------------------------------------------------------------
# 7. Display Analytics (single block — previous duplicate removed)
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# 8. Document Downloads (cached, in-memory)
# ---------------------------------------------------------------------------
st.subheader("📥 Export Institutional Analysis Documents")


@st.cache_data(show_spinner=False)
def _excel_bytes(_df, scenario, key):
    # key forces cache invalidation when df/scenario changes; _df excluded from hashing
    return generate_excel_report(_df, scenario).getvalue()


@st.cache_data(show_spinner=False)
def _docx_bytes(baseline_total, year3_total, scenario, key):
    return generate_docx_brief(
        {"baseline_total": baseline_total, "year3_total": year3_total},
        scenario,
    ).getvalue()


# Cache key derived from the actual rendered numbers
_df_key = float(final_df["Projected 2029/30"].sum())

btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    excel_bytes = _excel_bytes(final_df, selected_scenario, _df_key)
    st.download_button(
        label="Download Excel Worksheet Summary Report",
        data=excel_bytes,
        file_name=f"IRD_Projections_Sheet_{selected_scenario}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with btn_col2:
    docx_bytes = _docx_bytes(total_b, total_y3, selected_scenario, _df_key)
    st.download_button(
        label="Download Executive Briefing Memorandum (.docx)",
        data=docx_bytes,
        file_name=f"IRD_Methodological_Brief_{selected_scenario}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
