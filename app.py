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
#
# Nếu muốn lấy dữ liệu từ Google Sheet:
#
# 1. Mở Google Sheet
# 2. File → Share → Publish to web
# 3. Chọn CSV
# 4. Dán URL vào đây
#
# Ví dụ:
#
# GOOGLE_SHEET_URL = (
#     "GOOGLE_SHEET_URL = (
    "https://google.com"
    "2PACX-1vR1jFQhkmLwgj24-5qzY_gbogXBPkzcRoogwPJyOYmMKteDYDrBa2qstGaH4q8_0d1Ffmg4sFEwRP0i/pub?output=csv"
)
"
# )
#
# Nếu chưa có Google Sheet thì để None.
#



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

/* ==========================================================
   GLOBAL
   ========================================================== */

html,
body,
[data-testid="stAppViewContainer"],
[data-testid="stApp"] {

    background: #FFFFFF !important;
}

[data-testid="stHeader"] {
    display: none !important;
}

header {
    display: none !important;
}

footer {
    display: none !important;
}

[data-testid="stToolbar"] {
    display: none !important;
}

#MainMenu {
    display: none !important;
}


/* ==========================================================
   REMOVE DEFAULT STREAMLIT SPACING
   ========================================================== */

.block-container {

    padding-top: 0 !important;
    padding-bottom: 0 !important;
    padding-left: 0 !important;
    padding-right: 0 !important;

    max-width: none !important;
}


/* ==========================================================
   MAIN FIGMA CANVAS
   1440 × 1024
   ========================================================== */

.figma-page {

    position: relative;

    width: 100%;
    min-height: 100vh;

    background: #FFFFFF;

    overflow: hidden;
}


/* ==========================================================
   TOP GRADIENT BAR
   Figma:
   width 1441.82
   height 68.41
   ========================================================== */

.top-gradient {

    position: absolute;

    left: 0;
    top: 0;

    width: 100%;
    height: 68.41px;

    background:
        linear-gradient(
            90deg,
            #72DDE4 24%,
            #23749F 67%
        );
}


/* ==========================================================
   BOTTOM GRADIENT BAR
   ========================================================== */

.bottom-gradient {

    position: absolute;

    left: 0;
    bottom: 0;

    width: 100%;
    height: 35.16px;

    background:
        linear-gradient(
            90deg,
            #72DDE4 24%,
            #23749F 67%
        );
}


/* ==========================================================
   LOGO
   ========================================================== */

.logo-container {

    position: absolute;

    left: 65.46px;
    top: 88.93px;

    width: 262px;
    height: 130px;

    display: flex;

    align-items: center;
    justify-content: center;
}


.logo-container img {

    width: 262px;
    height: 130px;

    object-fit: contain;
}


.logo-fallback {

    font-family:
        Arial,
        sans-serif;

    font-size: 25px;

    font-weight: 800;

    color: #23749F;

    letter-spacing: -1px;
}


/* ==========================================================
   QUICK SEARCH TITLE
   ========================================================== */

.quick-search-title {

    position: absolute;

    left: 50%;

    top: 218px;

    transform:
        translateX(-50%);

    width: 707.60px;

    height: 118.70px;

    font-family:
        'Onest',
        Arial,
        sans-serif;

    font-size: 80px;

    font-weight: 400;

    line-height: 1.5;

    text-align: center;

    white-space: nowrap;

    color: #89EDF3;

    text-shadow:
        0px 4px 4px
        rgba(0, 0, 0, 0.25);
}


/* ==========================================================
   SEARCH AREA
   ========================================================== */

.search-area {

    position: absolute;

    left: 50%;

    top: 344.52px;

    transform:
        translateX(-50%);

    width: 585.73px;

    z-index: 20;
}


/* ==========================================================
   STREAMLIT TEXT INPUT
   ========================================================== */

.search-area
[data-testid="stTextInput"] {

    width: 585.73px !important;

    margin: 0 !important;

    padding: 0 !important;
}


.search-area
[data-testid="stTextInput"] > div {

    width: 585.73px !important;

    margin: 0 !important;

    padding: 0 !important;
}


/* hide label */

.search-area
[data-testid="stTextInput"] label {

    display: none !important;
}


/* input outer */

.search-area
[data-baseweb="input"] {

    width: 585.73px !important;

    height: 70px !important;

    box-sizing: border-box !important;

    background: #FFFFFF !important;

    border:
        2px solid #595656 !important;

    border-radius: 70px !important;

    box-shadow: none !important;

    transition: all 0.2s ease !important;
}


/* focus */

.search-area
[data-baseweb="input"]:focus-within {

    border:
        2px solid #23749F !important;

    box-shadow:
        0px 4px 12px
        rgba(35, 116, 159, 0.12) !important;
}


/* actual input */

.search-area
input {

    height: 66px !important;

    box-sizing: border-box !important;

    border: none !important;

    outline: none !important;

    background: transparent !important;

    color: #333333 !important;

    font-family:
        Arial,
        sans-serif !important;

    font-size: 17px !important;

    padding-left: 30px !important;

    padding-right: 70px !important;
}


/* placeholder */

.search-area
input::placeholder {

    color: #AAAAAA !important;

    opacity: 1 !important;
}


/* search icon */

.search-icon {

    position: absolute;

    right: 20px;

    top: 15px;

    width: 40px;
    height: 40px;

    pointer-events: none;

    z-index: 50;
}


.search-icon svg {

    width: 40px;
    height: 40px;
}


/* ==========================================================
   RESULT BOX
   Figma Desktop 3:
   643.99 × 474.73
   left 396.07
   top 447.81
   ========================================================== */

.result-area {

    position: absolute;

    left: 50%;

    top: 447.81px;

    transform:
        translateX(-50%);

    width: 643.99px;

    min-height: 474.73px;

    box-sizing: border-box;

    padding: 36px;

    background: #F3FFFE;

    border-radius: 30px;

    z-index: 10;

    box-shadow:
        0px 8px 30px
        rgba(35, 116, 159, 0.04);
}


/* result title */

.result-title {

    font-family:
        'Onest',
        Arial,
        sans-serif;

    font-size: 25px;

    font-weight: 700;

    color: #23749F;

    margin-bottom: 18px;
}


/* divider */

.result-divider {

    width: 100%;

    height: 1px;

    background: #D9E9EA;

    margin-bottom: 22px;
}


/* result section */

.result-section {

    margin-bottom: 21px;
}


.result-section-title {

    font-family:
        Arial,
        sans-serif;

    font-size: 15px;

    font-weight: 700;

    color: #23749F;

    margin-bottom: 7px;

    letter-spacing: 0.3px;
}


.result-text {

    font-family:
        Arial,
        sans-serif;

    font-size: 15px;

    line-height: 1.55;

    color: #4B5563;
}


/* ==========================================================
   EXPORT BUTTON
   ========================================================== */

.export-area {

    width: 100%;

    display: flex;

    justify-content: center;

    margin-top: 18px;
}


.export-area button {

    background: #23749F !important;

    color: #FFFFFF !important;

    border: none !important;

    border-radius: 8px !important;

    padding:
        11px 25px !important;

    font-size: 14px !important;

    font-weight: 600 !important;

    cursor: pointer !important;
}


.export-area button:hover {

    background: #185A7C !important;
}


/* ==========================================================
   ERROR
   ========================================================== */

.error-box {

    position: absolute;

    left: 50%;

    top: 447.81px;

    transform:
        translateX(-50%);

    width: 643.99px;

    box-sizing: border-box;

    padding: 25px 30px;

    background: #FFF4F5;

    border:
        1px solid #F4C7CC;

    border-radius: 20px;

    color: #A33A45;

    font-family:
        Arial,
        sans-serif;

    font-size: 15px;

    text-align: center;
}


/* ==========================================================
   RESPONSIVE
   ========================================================== */

@media (max-width: 1100px) {

    .quick-search-title {

        font-size: 60px;

        width: 650px;
    }

    .logo-container {

        left: 35px;
    }
}


@media (max-width: 800px) {

    .figma-page {

        min-width: 700px;
    }

    .quick-search-title {

        font-size: 48px;

        width: 600px;

        top: 230px;
    }

    .search-area {

        width: 520px;
    }

    .search-area
    [data-testid="stTextInput"] {

        width: 520px !important;
    }

    .search-area
    [data-testid="stTextInput"] > div {

        width: 520px !important;
    }

    .search-area
    [data-baseweb="input"] {

        width: 520px !important;
    }
}

</style>


<!-- ONEST FONT -->

<link
    href="https://fonts.googleapis.com/css2?family=Onest:wght@400;500;600;700&display=swap"
    rel="stylesheet"
>

""",
    unsafe_allow_html=True
)


# ============================================================
# START FIGMA PAGE
# ============================================================

st.markdown(
    '<div class="figma-page">',
    unsafe_allow_html=True
)


# ============================================================
# TOP BAR
# ============================================================

st.markdown(
    '<div class="top-gradient"></div>',
    unsafe_allow_html=True
)


# ============================================================
# BOTTOM BAR
# ============================================================

st.markdown(
    '<div class="bottom-gradient"></div>',
    unsafe_allow_html=True
)


# ============================================================
# LOGO
# ============================================================

if os.path.exists(LOGO_FILE):

    st.markdown(
        '<div class="logo-container">',
        unsafe_allow_html=True
    )

    st.image(
        LOGO_FILE,
        width=262
    )

    st.markdown(
        '</div>',
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <div class="logo-container">
            <div class="logo-fallback">
                LINE BASE
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# QUICK SEARCH TITLE
# ============================================================

st.markdown(
    """
    <div class="quick-search-title">
        QUICK SEARCH
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SEARCH INPUT
# ============================================================

st.markdown(
    '<div class="search-area">',
    unsafe_allow_html=True
)

search_text = st.text_input(
    "Search",
    placeholder="Nhập địa chỉ dự án hoặc tên thành phố...",
    label_visibility="collapsed",
    key="search_input"
)

# Search icon
st.markdown(
    """
    <div class="search-icon">

        <svg
            viewBox="0 0 40 40"
            xmlns="http://www.w3.org/2000/svg"
        >

            <circle
                cx="17"
                cy="17"
                r="9"
                fill="none"
                stroke="#ADADAD"
                stroke-width="3"
            />

            <line
                x1="24"
                y1="24"
                x2="33"
                y2="33"
                stroke="#ADADAD"
                stroke-width="3"
                stroke-linecap="round"
            />

        </svg>

    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# SEARCH
# ============================================================

result = search_city(search_text)


# ============================================================
# RESULT
# ============================================================

if search_text.strip():

    if result is not None:

        city = str(result["city"])

        building_code = str(
            result["building_code"]
        )

        drainage = str(
            result["drainage_civil_specs"]
        )

        lid = str(
            result["low_impact_development"]
        )

        permit_agency = str(
            result["local_permit_agency"]
        )


        # ----------------------------------------------------
        # RESULT BOX
        # ----------------------------------------------------

        st.markdown(
            f"""
            <div class="result-area">

                <div class="result-title">
                    CITY OF {city.upper()}
                </div>

                <div class="result-divider"></div>


                <div class="result-section">

                    <div class="result-section-title">
                        1. BUILDING CODES
                    </div>

                    <div class="result-text">
                        {building_code}
                    </div>

                </div>


                <div class="result-section">

                    <div class="result-section-title">
                        2. CIVIL &amp; DRAINAGE
                    </div>

                    <div class="result-text">
                        {drainage}
                    </div>

                    <div
                        class="result-text"
                        style="margin-top:6px;"
                    >
                        <b>LID Rules:</b> {lid}
                    </div>

                </div>


                <div class="result-section">

                    <div class="result-section-title">
                        3. LOCAL PERMIT AGENCY
                    </div>

                    <div class="result-text">
                        {permit_agency}
                    </div>

                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


        # ----------------------------------------------------
        # EXPORT SOW
        # ----------------------------------------------------

        if os.path.exists("sow_template.docx"):

            try:

                doc = DocxTemplate(
                    "sow_template.docx"
                )

                context = {

                    "PROJECT_ADDRESS":
                        search_text,

                    "CITY":
                        city,

                    "BUILDING_CODE":
                        building_code,

                    "DRAINAGE_CIVIL_SPECS":
                        drainage,

                    "LOW_IMPACT_DEVELOPMENT":
                        lid,

                    "LOCAL_PERMIT_AGENCY":
                        permit_agency,
                }

                doc.render(context)

                output = io.BytesIO()

                doc.save(output)

                output.seek(0)


                # Download button nằm bên dưới result box
                st.markdown(
                    '<div class="export-area">',
                    unsafe_allow_html=True
                )

                st.download_button(
                    label="Export Statement of Work (SOW)",

                    data=output,

                    file_name=(
                        f"SOW_"
                        f"{city.replace(' ', '_')}.docx"
                    ),

                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml.document"
                    ),

                    key=f"download_{city}"
                )

                st.markdown(
                    '</div>',
                    unsafe_allow_html=True
                )

            except Exception as e:

                st.error(
                    f"Không thể tạo file SOW: {e}"
                )


    else:

        # ----------------------------------------------------
        # NOT FOUND
        # ----------------------------------------------------

        st.markdown(
            """
            <div class="error-box">

                Không tìm thấy dữ liệu phù hợp.
                <br>
                Vui lòng nhập tên thành phố hoặc
                địa chỉ dự án tại California.

            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# END PAGE
# ============================================================

st.markdown(
    '</div>',
    unsafe_allow_html=True
)

