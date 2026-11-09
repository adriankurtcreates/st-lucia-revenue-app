import requests
import streamlit as st

WB_URL = (
    "https://api.worldbank.org/v2/country/LCA/indicator/FP.CPI.TOTL.ZG"
    "?format=json&per_page=100"
)

# Documented IMF WEO / MPO April 2026 baseline CPI inflation for Saint Lucia
IMF_WEO_FALLBACK = {
    "2026": 0.021,
    "2027": 0.020,
    "2028": 0.020,
    "2029": 0.020,
    "2030": 0.020,
}


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_macro_forecasts():
    """
    Fetch Saint Lucia CPI inflation from the World Bank API.
    Falls back to the IMF WEO baseline if the API is unavailable or malformed.
    """
    try:
        response = requests.get(WB_URL, timeout=10)
        response.raise_for_status()
        payload = response.json()
        # World Bank v2 format: [metadata, [{"date": "2023", "value": 2.1, ...}, ...]]
        if not isinstance(payload, list) or len(payload) < 2 or not payload[1]:
            return dict(IMF_WEO_FALLBACK)
        data_points = {
            item["date"]: float(item["value"]) / 100.0
            for item in payload[1]
            if item.get("value") is not None
        }
        return data_points if data_points else dict(IMF_WEO_FALLBACK)
    except Exception:
        return dict(IMF_WEO_FALLBACK)
