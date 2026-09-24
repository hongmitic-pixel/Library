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
    @import url('https://fonts.googleapis.com/css2?family=Onest:wght@400;500;600;700&display=swap');

    /* GLOBAL */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background: #FFFFFF !important;
        font-family: 'Onest', sans-serif !important;
    }

    [data-testid="stHeader"], header, footer, [data-testid="stToolbar"], #MainMenu {
        display: none !important;
    }

    /* TOP / BOTTOM GRADIENT BARS */
    .lb-top-bar {
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 68px;
        background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%);
        z-index: 1000;
    }

    /* small brand mark placeholder inside the top bar */
    .lb-top-bar-mark {
        position: absolute;
        left: 2.5vw;
        top: 22px;
        width: 126px;
        height: 25px;
        background: #FFFFFF;
    }

    .lb-bottom-bar {
        position: fixed;
        bottom: 0; left: 0; right: 0;
        height: 35px;
        background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%);
        z-index: 1000;
    }

    /* REMOVE DEFAULT STREAMLIT SPACING */
    .block-container {
        padding: 130px 3rem 80px 3rem !important;
        max-width: 900px !important;
        margin: 0 auto !important;
    }

    /* HERO HEADER */
    .lb-hero {
        display: flex;
        align-items: center;
        gap: 2rem;
        margin-bottom: 2.5rem;
    }

    .lb-hero img {
        width: 160px;
        height: auto;
    }

    .lb-hero-title {
        font-family: 'Onest', sans-serif;
        font-weight: 400;
        font-size: 4rem;
        line-height: 1.15;
        color: #89EDF3;
        text-shadow: 0px 4px 4px rgba(0, 0, 0, 0.25);
        margin: 0;
    }

    /* SEARCH PILL */
    .st-key-lb_search_wrap {
        position: relative;
    }

    div[data-testid="stTextInput"] input {
        border-radius: 70px !important;
        border: 2px solid #595656 !important;
        height: 70px !important;
        padding: 0 60px 0 28px !important;
        font-size: 1.05rem !important;
        font-family: 'Onest', sans-serif !important;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: #23749F !important;
        box-shadow: none !important;
    }

    /* search button square, sits inside the right edge of the pill */
    .lb-search-btn {
        position: absolute;
        right: 16px;
        top: 50%;
        transform: translateY(-50%);
        width: 40px;
        height: 40px;
        background: #ADADAD;
        pointer-events: none;
    }

    /* RESULT CARD CONTAINER (mint) */
    .lb-result-panel {
        background: #F3FFFE;
        border-radius: 30px;
        padding: 2.2rem 2.4rem;
        margin-top: 1.8rem;
    }

    .lb-result-city {
        font-size: 1.4rem;
        font-weight: 600;
        color: #133C55;
        margin-bottom: 1.2rem;
    }

    .result-card {
        border: 1px solid #E5E7EB;
        border-radius: 14px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 0.9rem;
        background: #FFFFFF;
    }

    .result-label {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #6B7280;
        margin-bottom: 0.3rem;
    }

    .result-value {
        font-size: 0.98rem;
        color: #111827;
    }
    </style>

    <div class="lb-top-bar">
        <div class="lb-top-bar-mark"></div>
    </div>
    <div class="lb-bottom-bar"></div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

logo_html = (
    f'<img src="{LOGO_FILE}" />'
    if os.path.exists(LOGO_FILE)
    else '<div style="width:160px;"></div>'
)

st.markdown(
    f"""
    <div class="lb-hero">
        {logo_html}
        <div class="lb-hero-title">QUICK SEARCH</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SEARCH BOX
# ============================================================

search_wrap = st.container(key="lb_search_wrap")
with search_wrap:
    query = st.text_input(
        "Tên thành phố",
        placeholder="Nhập tên thành phố...",
        label_visibility="collapsed",
    )
    st.markdown('<div class="lb-search-btn"></div>', unsafe_allow_html=True)

result = search_city(query) if query else None


# ============================================================
# RESULTS
# ============================================================

if query and result is None:
    st.markdown(
        f'<div class="lb-result-panel">Không tìm thấy thành phố phù hợp với: <b>{query}</b></div>',
        unsafe_allow_html=True,
    )

elif result is not None:
    fields = [
        ("Building Code", result["building_code"]),
        ("Drainage / Civil Specs", result["drainage_civil_specs"]),
        ("Low Impact Development (LID)", result["low_impact_development"]),
        ("Local Permit Agency", result["local_permit_agency"]),
    ]

    cards_html = "".join(
        f"""
        <div class="result-card">
            <div class="result-label">{label}</div>
            <div class="result-value">{value}</div>
        </div>
        """
        for label, value in fields
    )

    st.markdown(
        f"""
        <div class="lb-result-panel">
            <div class="lb-result-city">📍 {result['city']}</div>
            {cards_html}
        </div>
        """,
        unsafe_allow_html=True,
    )
