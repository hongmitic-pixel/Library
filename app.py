import streamlit as st
import pandas as pd
import io
import os

from docxtpl import DocxTemplate


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# CONFIG
# ============================================================

# Nếu có logo thật, đặt file:
# logo_line_base.png
# cùng thư mục với app.py

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
# URL Google Sheet xuất bản dưới dạng CSV chuẩn xác của bạn:

GOOGLE_SHEET_URL = "https://google.com"


# ============================================================
# BACKUP DATA
# ============================================================

BACKUP_DATA = {
    "city": [
        "Adelanto",
        "Agoura Hills",
        "Alameda",
        "Albany",
        "Alhambra",
        "Aliso Viejo",
        "Los Angeles",
        "San Francisco",
        "San Jose",
        "San Diego",
        "Santa Ana",
    ],

    "building_code": [
        "2025/2026 California Building Code (CBC) - Adelanto City Amendments.",
        "2025/2026 California Building Code (CBC) - Agoura Hills City Amendments.",
        "2025/2026 California Building Code (CBC) - Alameda City Amendments.",
        "2025/2026 California Building Code (CBC) - Albany City Amendments.",
        "2025/2026 California Building Code (CBC) - Alhambra City Amendments.",
        "2025/2026 California Building Code (CBC) - Aliso Viejo City Amendments.",
        "2025/2026 California Building Code (CBC) - Los Angeles City Amendments.",
        "2025/2026 California Building Code (CBC) - San Francisco City Amendments.",
        "2025/2026 California Building Code (CBC) - San Jose City Amendments.",
        "2025/2026 California Building Code (CBC) - San Diego City Amendments.",
        "2025/2026 California Building Code (CBC) - Santa Ana City Amendments.",
    ],

    "drainage_civil_specs": [
        "City of Adelanto Public Works Design Manual Standard Specs.",
        "City of Agoura Hills Public Works Design Manual Standard Specs.",
        "City of Alameda Public Works Design Manual Standard Specs.",
        "City of Albany Public Works Design Manual Standard Specs.",
        "City of Alhambra Public Works Design Manual Standard Specs.",
        "City of Aliso Viejo Public Works Design Manual Standard Specs.",
        "City of Los Angeles Public Works Design Manual Standard Specs.",
        "City of San Francisco Public Works Design Manual Standard Specs.",
        "City of San Jose Public Works Design Manual Standard Specs.",
        "City of San Diego Public Works Design Manual Standard Specs.",
        "City of Santa Ana Public Works Design Manual Standard Specs.",
    ],

    "low_impact_development": [
        "Adelanto Municipal Stormwater Management - LID Rules.",
        "Agoura Hills Municipal Stormwater Management - LID Rules.",
        "Alameda Municipal Stormwater Management - LID Rules.",
        "Albany Municipal Stormwater Management - LID Rules.",
        "Alhambra Municipal Stormwater Management - LID Rules.",
        "Aliso Viejo Municipal Stormwater Management - LID Rules.",
        "Los Angeles Municipal Stormwater Management - LID Rules.",
        "San Francisco Municipal Stormwater Management - LID Rules.",
        "San Jose Municipal Stormwater Management - LID Rules.",
        "San Diego Municipal Stormwater Management - LID Rules.",
        "Santa Ana Municipal Stormwater Management - LID Rules.",
    ],

    "local_permit_agency": [
        "City of Adelanto Development Services Division.",
        "City of Agoura Hills Development Services Division.",
        "City of Alameda Development Services Division.",
        "City of Albany Development Services Division.",
        "City of Alhambra Development Services Division.",
        "City of Aliso Viejo Development Services Division.",
        "City of Los Angeles Development Services Division.",
        "City of San Francisco Development Services Division.",
        "City of San Jose Development Services Division.",
        "City of San Diego Development Services Division.",
        "City of Santa Ana Development Services Division.",
    ]
}
# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data(ttl=300)
def load_data():

    backup_df = pd.DataFrame(BACKUP_DATA)

    if not GOOGLE_SHEET_URL:
        return backup_df

    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # Chuẩn hóa tên cột nếu Google Sheet có khoảng trắng
        rename_map = {
            "drainage civil specs": "drainage_civil_specs",
            "drainage_civil_specs": "drainage_civil_specs",
            "building code": "building_code",
            "building_code": "building_code",
            "low impact development": "low_impact_development",
            "low_impact_development": "low_impact_development",
            "local permit agency": "local_permit_agency",
            "local_permit_agency": "local_permit_agency",
            "city": "city",
        }

        df = df.rename(columns=rename_map)

        required_columns = [
            "city",
            "building_code",
            "drainage_civil_specs",
            "low_impact_development",
            "local_permit_agency",
        ]

        if all(col in df.columns for col in required_columns):
            df = df[required_columns].copy()
            df["city"] = df["city"].astype(str).str.strip()
            return df

    except Exception:
        pass

    return backup_df


df_cities = load_data()


# ============================================================
# SEARCH FUNCTION
# ============================================================

def search_city(search_text):

    if not search_text:
        return None

    search_text = str(search_text).strip().lower()

    if not search_text:
        return None

    # Ưu tiên city match
    for _, row in df_cities.iterrows():
        city = str(row["city"]).strip()
        if city.lower() in search_text:
            return row

    # Nếu người dùng nhập một phần tên thành phố
    for _, row in df_cities.iterrows():
        city = str(row["city"]).strip()
        if search_text in city.lower():
            return row

    return None


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
    /* GLOBAL */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background: #FFFFFF !important;
    }
    
    [data-testid="stHeader"], header, footer, [data-testid="stToolbar"], #MainMenu {
        display: none !important;
    }
    
    /* REMOVE DEFAULT STREAMLIT SPACING */
    .block-container {
        padding: 0 !important;
        max-width: none !important;
    }
    </style>
    """,
    unsafe_html=True
)

