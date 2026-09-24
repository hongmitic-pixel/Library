import streamlit as st
import pandas as pd
import io
import os

from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt, RGBColor


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
# EXPORT TO WORD (.docx)
# ============================================================

def build_docx(result_row):
    """Tạo file Word (.docx) từ dữ liệu 1 thành phố, trả về bytes để tải xuống."""

    doc = Document()

    # Tiêu đề
    title = doc.add_heading(f"{result_row['city']}", level=1)
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x13, 0x3C, 0x55)

    subtitle = doc.add_paragraph()
    subtitle_run = subtitle.add_run("LINE BASE — Quick Search Report")
    subtitle_run.italic = True

    fields = [
        ("Building Code", result_row["building_code"]),
        ("Drainage / Civil Specs", result_row["drainage_civil_specs"]),
        ("Low Impact Development (LID)", result_row["low_impact_development"]),
        ("Local Permit Agency", result_row["local_permit_agency"]),
    ]

    for label, value in fields:
        heading = doc.add_heading(label, level=2)
        for run in heading.runs:
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x59, 0x56, 0x56)

        doc.add_paragraph(str(value))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


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

    /* dots inside the top bar (real UI, top-left) */
    .lb-top-bar-dots {
        position: absolute;
        left: 18px;
        top: 0;
        height: 68px;
        display: flex;
        align-items: center;
        gap: 9px;
    }

    .lb-top-bar-dots span {
        width: 15px;
        height: 15px;
        border-radius: 50%;
        display: inline-block;
    }

    .lb-dot-1 { background: #FFFFFF; }
    .lb-dot-2 { background: #9AFFFD; }
    .lb-dot-3 { background: #F2C94C; }

    .lb-bottom-bar {
        position: fixed;
        bottom: 0; left: 0; right: 0;
        height: 35px;
        background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%);
        z-index: 1000;
    }

    /* REMOVE DEFAULT STREAMLIT SPACING */
    .block-container {
        padding: 140px 3rem 80px 3rem !important;
        max-width: 640px !important;
        margin: 0 auto !important;
    }

    /* BRAND ROW (logo + wordmark, left aligned) */
    .lb-brand-row {
        position: fixed;
        top: 85px;  /* Đặt vị trí nằm ngay dưới thanh top bar (thanh top bar cao 68px) */
        left: 3rem; /* Căn lề trái tương đương với lề của trang */
        z-index: 999;
        display: flex;
        align-items: center;
        gap: 0.6rem;
        margin-bottom: 0rem;
    }

    .lb-brand-row img {
        height: 70px;
        width: auto;
        object-fit: contain;
    }

    /* HERO TITLE (centered, gradient fill + light stroke) */
    .lb-hero-title {
        font-family: 'Onest', sans-serif;
        font-weight: 800;
        font-size: 4rem;
        line-height: 1.2;
        text-align: center;
        margin: 0 0 1.6rem 0;
        background: linear-gradient(180deg, #BFF0F5 0%, #2E8FC0 100%);
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent;
        -webkit-text-stroke: 1px rgba(255, 255, 255, 0.55);
        text-shadow: 0px 2px 3px rgba(0, 0, 0, 0.12);
    }

/* SEARCH ROW — bo tròn, viền xanh đen, có icon kính lúp */
    div[data-testid="stTextInput"] {
        max-width: 380px;
        margin: 0 auto;
    }

    div[data-testid="stTextInput"] input {
        border-radius: 30px !important; /* Bo tròn góc nhiều hơn (kiểu viên thuốc) */
        border: 2px solid #123A54 !important; /* Viền màu xanh đen */
        height: 54px !important;
        padding: 0 52px 0 20px !important;
        font-size: 0.95rem !important;
        font-family: 'Onest', sans-serif !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        /* Icon kính lúp ở bên phải */
        background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='36' height='36'><circle cx='18' cy='18' r='16' fill='%23F3FFFE' stroke='%23123A54' stroke-width='1.5'/><circle cx='16' cy='16' r='5.5' fill='none' stroke='%23123A54' stroke-width='2'/><line x1='20' y1='20' x2='24.5' y2='24.5' stroke='%23123A54' stroke-width='2' stroke-linecap='round'/></svg>");
        background-repeat: no-repeat;
        background-position: right 10px center;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: #23749F !important;
        box-shadow: 0 0 8px rgba(35, 116, 159, 0.25) !important; /* Hiệu ứng phát sáng nhẹ khi bấm vào */
    }

    /* Xóa màu nền xanh tự động của trình duyệt khi autofill */
    div[data-testid="stTextInput"] input:-webkit-autofill,
    div[data-testid="stTextInput"] input:-webkit-autofill:focus {
        -webkit-box-shadow: 0 0 0 1000px #FFFFFF inset !important;
        -webkit-text-fill-color: #111827 !important;
    }

    /* RESULT CARD CONTAINER (mint) */
    .lb-result-panel {
        background: #F3FFFE;
        border-radius: 30px;
        padding: 2.2rem 2.4rem 1.3rem 2.4rem;
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

    /* DOWNLOAD BUTTON */
    div[data-testid="stDownloadButton"] {
        display: flex;
        justify-content: center;
        margin-top: -0.6rem;
    }

    div[data-testid="stDownloadButton"] button {
        background: #23749F !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 30px !important;
        padding: 0.6rem 1.8rem !important;
        font-family: 'Onest', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
        box-shadow: 0px 2px 6px rgba(35, 116, 159, 0.35) !important;
    }

    div[data-testid="stDownloadButton"] button:hover {
        background: #1B5C7E !important;
        color: #FFFFFF !important;
    }
    </style>

    <div class="lb-top-bar">
        <div class="lb-top-bar-dots">
            <span class="lb-dot-1"></span>
            <span class="lb-dot-2"></span>
            <span class="lb-dot-3"></span>
        </div>
    </div>
    <div class="lb-bottom-bar"></div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

def _logo_data_uri(path):
    """Đọc file ảnh và trả về data URI base64 — cách duy nhất đáng tin cậy
    để nhúng ảnh cục bộ vào HTML thô trong Streamlit (đường dẫn tương đối
    kiểu <img src="file.png"> không được serve ra web)."""
    import base64
    try:
        ext = os.path.splitext(path)[1].lstrip(".").lower() or "png"
        mime = "jpeg" if ext == "jpg" else ext
        with open(path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
        return f"data:image/{mime};base64,{encoded}"
    except Exception:
        return None


logo_uri = _logo_data_uri(LOGO_FILE) if os.path.exists(LOGO_FILE) else None

logo_html = (
    f'<img src="{logo_uri}" />'
    if logo_uri
    else '<div style="width:42px;height:42px;"></div>'
)

st.markdown(
    f"""
    <div class="lb-brand-row">
        {logo_html}
    </div>
    <div class="lb-hero-title">QUICK SEARCH</div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SEARCH BOX
# ============================================================

query = st.text_input(
    "Tên thành phố",
    placeholder="Nhập tên thành phố...",
    label_visibility="collapsed",
)

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

    docx_buffer = build_docx(result)

    st.download_button(
        label="⬇️ Tải xuống (.docx)",
        data=docx_buffer,
        file_name=f"{result['city'].replace(' ', '_')}_specs.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
