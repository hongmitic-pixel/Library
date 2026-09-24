import io
import os
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# -*- coding: utf-8 -*-
import io
import os
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CONFIG
# ============================================================

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
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
        "Adelanto Municipal Stormwater # -*- coding: utf-8 -*-
import io
import os
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CONFIG
# ============================================================

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
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
    ],
}


# ============================================================
# LOAD DATA
# ============================================================


# ============================================================
# CONFIG
# ============================================================

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
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
        "Adelanto Municipal Stormwater # -*- coding: utf-8 -*-
import io
import os
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CONFIG
# ============================================================

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
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
    ],
}


# ============================================================
# LOAD DATA
# ============================================================


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
        "Adelanto Municipal Stormwater # -*- coding: utf-8 -*-
import io
import os
import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="LINE BASE - Quick Search",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CONFIG
# ============================================================

LOGO_FILE = "logo_line_base.png"

# ------------------------------------------------------------
# GOOGLE SHEET
# ------------------------------------------------------------
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
    ],
}


# ============================================================
# LOAD DATA
# ============================================================


        border: 2px solid #123A54 !important;
        background: #FFFFFF !important;
        padding: 0 16px 0 46px !important;
        position: relative !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04) !important;
        display: flex !important;
        align-items: center !important;
        height: 50px !important;
    }

    div[data-testid="stTextInput"]::before {
        content: "";
        position: absolute;
        left: 16px;
        top: 50%;
        transform: translateY(-50%);
        width: 18px;
        height: 18px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='%23123A54' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'%3E%3C/circle%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'%3E%3C/line%3E%3C/svg%3E");
        background-repeat: no-repeat;
        background-position: center;
        pointer-events: none;
        z-index: 5;
    }

    div[data-testid="stTextInput"] > div {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        width: 100% !important;
        padding: 0 !important;
    }

    div[data-testid="stTextInput"] > div > div {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        width: 100% !important;
    }

    div[data-testid="stTextInput"] div[data-baseweb="input"] {
        background: transparent !important;
        border: none !important;
        width: 100% !important;
    }

    div[data-testid="stTextInput"] input {
        border: none !important;
        height: 40px !important;
        padding: 0 !important;
        font-size: 1rem !important;
        font-family: 'Onest', sans-serif !important;
        background-color: transparent !important;
        box-shadow: none !important;
        color: #111827 !important;
    }

    div[data-testid="stTextInput"] div[data-baseweb="input"] + div {
        position: absolute !important;
        right: 16px !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
        display: flex !important;
        align-items: center !important;
        margin: 0 !important;
        pointer-events: none;
        z-index: 5;
    }

    div[data-testid="stTextInput"] div[data-baseweb="input"] + div small {
        font-family: 'Onest', sans-serif !important;
        font-size: 0.78rem !important;
        color: #94A3B8 !important;
        visibility: visible !important;
    }

    div[data-testid="stTextInput"]:focus-within {
        border-color: #23749F !important;
        box-shadow: 0 0 8px rgba(35, 116, 159, 0.25) !important;
    }

    div[data-testid="stTextInput"] input:-webkit-autofill,
    div[data-testid="stTextInput"] input:-webkit-autofill:focus {
        -webkit-box-shadow: 0 0 0 1000px #FFFFFF inset !important;
        -webkit-text-fill-color: #111827 !important;
    }

    .lb-result-panel {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 24px;
        margin-top: 2rem;
    }

    .lb-result-city {
        font-size: 1.4rem;
        font-weight: 700;
        color: #133C55;
        margin-bottom: 1rem;
    }

    .result-card {
        margin-bottom: 1rem;
    }

    .result-label {
        font-weight: 600;
        color: #595656;
        font-size: 0.95rem;
    }

    .result-value {
        color: #111827;
        font-size: 1rem;
    }

    div[data-testid="stDownloadButton"] {
        margin-top: 1.5rem !important;
    }
    
    div[data-testid="stDownloadButton"] button {
        border-radius: 8px !important;
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #133C55 !important;
        font-weight: 600 !important;
        font-family: 'Onest', sans-serif !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.2s ease-in-out;
    }

    div[data-testid="stDownloadButton"] button:hover {
        background-color: #133C55 !important;
        color: #FFFFFF !important;
        border-color: #133C55 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HEADER
# ============================================================


def _logo_data_uri(path):
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
    <div class="lb-top-bar">
        <div class="lb-top-bar-dots">
            <span class="lb-dot-1"></span>
            <span class="lb-dot-2"></span>
            <span class="lb-dot-3"></span>
        </div>
    </div>
    <div class="lb-bottom-bar"></div>
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
        (
            "Low Impact Development (LID)",
            result["low_impact_development"],
        ),
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
            <div class="lb-result-city">&#128205; {result['city']}</div>
            {cards_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

    docx_buffer = build_docx(result)
    st.download_button(
        label="📥 Download",
        data=docx_buffer,
        file_name=f"{result['city'].replace(' ', '_')}_specs.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
