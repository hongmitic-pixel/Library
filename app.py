import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io
import os

# Cấu hình trang tối giản, ẩn các thành phần thừa của Streamlit
st.set_page_config(page_title="LINE BASE - SOW System", layout="wide", initial_sidebar_state="collapsed")

# --- 🎨 HỆ THỐNG GIAO DIỆN CHUẨN FIGMA 100% (LOGO TRÁI - GIAO DIỆN CHÍNH GIỮA) ---
st.markdown("""
    <style>
        /* Toàn bộ nền trang web phẳng sạch sẽ */
        .stApp {
            background-color: #FFFFFF !important;
        }
        /* Giấu triệt để các thành phần hệ thống mặc định của Streamlit */
        header, footer, .stDeployButton, div[data-testid="stToolbar"] {display: none !important;}
        
        /* CONTAINER CHO LOGO: Khóa cứng lề trái 40px không dịch chuyển */
        .logo-top-left {
            position: absolute;
            top: 20px;
            left: 40px;
            z-index: 999;
        }
        
        /* CONTAINER TỔNG CHO GIAO DIỆN CHÍNH: Ép nằm chính giữa màn hình */
        .figma-center-layout {
            max-width: 600px !important;
            margin: 40px auto 0 auto !important; /* Đã thu hẹp khoảng cách trống với logo bên trên */
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        /* TIÊU ĐỀ QUICK SEARCH MÀU XANH TEAL ĐÚNG FIGMA */
        .quick-search-title {
            font-size: 38px !important;
            font-weight: 700 !important;
            color: #1E6B7B !important;
            text-transform: uppercase !important;
            margin-bottom: 25px !important;
            letter-spacing: -0.5px !important;
            width: 100% !important;
            text-align: center !important;
        }

        /* --- ÉP THANH TÌM KIẾM CON NHỘNG NẰM CHÍNH GIỮA MÀN HÌNH --- */
        .stTextInput {
            width: 100% !important;
            max-width: 480px !important;
            margin: 0 auto 30px auto !important;
        }
        .stTextInput > div, .stTextInput > div > div {
            border: none !important;
            background-color: transparent !important;
            box-shadow: none !important;
        }
        /* Khung hình con nhộng phẳng lỳ viền mỏng bo tròn */
        div[data-baseweb="input"] {
            border: 1.5px solid #CBD5E1 !important;
            border-radius: 50px !important;
            background-color: #FFFFFF !important;
            transition: all 0.2s ease;
        }
        /* Chữ nhập liệu phía trong ô */
        div.stTextInput input {
            border: none !important;
            background-color: transparent !important;
            font-size: 15px !important;
            color: #334155 !important;
            padding: 12px 24px !important;
            text-align: left !important;
        }
        div[data-baseweb="input"]:focus-within {
            border: 1.5px solid #1E6B7B !important;
            box-shadow: 0px 4px 12px rgba(30, 107, 123, 0.08) !important;
        }

        /* --- KHUNG HIỂN THỊ KẾT QUẢ MÀU XANH MỜ BO GÓC CĂN GIỮA (DESKTOP - 3) --- */
        .desktop3-result-box {
            background-color: #EFF8F9 !important;
            border-radius: 12px !important;
            padding: 35px !important;
            width: 100% !important;
            max-width: 480px !important;
            border: 1px solid #E2E8F0 !important;
            text-align: left !important;
            margin: 0 auto !important;
            box-shadow: 0px 4px 20px rgba(0, 0, 0, 0.01) !important;
        }
        .desktop3-result-box h3 {
            color: #1E6B7B !important;
            margin-top: 0 !important;
            font-size: 18px !important;
            font-weight: 700 !important;
        }
        .desktop3-result-box h4 {
            color: #1E6B7B !important;
            font-weight: 700 !important;
            font-size: 14px !important;
            margin-top: 18px !important;
            margin-bottom: 4px !important;
            text-transform: uppercase !important;
            letter-spacing: 0.3px !important;
        }
        .desktop3-result-box p {
            color: #475569 !important;
            font-size: 14.5px !important;
            line-height: 1.5 !important;
            margin: 0 !important;
        }

        /* NÚT TẢI FILE WORD ĐỒNG BỘ NẰM GIỮA TUYỆT ĐỐI */
        div.stDownloadButton {
            width: 100% !important;
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            margin-top: 25px !important;
        }
        div.stDownloadButton > button {
            background: #1E6B7B !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 6px !important;
            padding: 11px 28px !important;
            font-size: 14.5px !important;
            font-weight: 500 !important;
            transition: background 0.2s !important;
            box-shadow: 0px 4px 12px rgba(30, 107, 123, 0.1) !important;
        }
        div.stDownloadButton > button:hover {
            background: #154D59 !important;
        }
    </style>
""", unsafe_allow_html=True)
# =========================================================================
# 🧭 1. ĐẶT FILE ẢNH LOGO GỐC BẠN TỰ TẢI LÊN Ở GÓC TRÁI MÀN HÌNH
# =========================================================================
st.markdown('<div class="logo-top-left">', unsafe_allow_html=True)
if os.path.exists("logo_line_base.png"):
    with open("logo_line_base.png", "rb") as f:
        st.image(f.read(), width=160)
else:
    st.markdown('<h2 style="font-size:20px; font-weight:800; color:#1E293B; margin:0;">LINE BASE</h2>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================================
# 📐 2. DỰNG LAYOUT CHÍNH CĂN GIỮA TUYỆT ĐỐI THEO FIGMA
# =========================================================================
st.markdown('<div class="figma-center-layout">', unsafe_allow_html=True)

# Tiêu đề QUICK SEARCH đặt đúng vị trí chính giữa
st.markdown('<div class="quick-search-title">QUICK SEARCH</div>', unsafe_allow_html=True)
# --- 🛠️ CƠ SỞ DỮ LIỆU THÀNH PHỐ CALIFORNIA VÀ LOGIC TÌM KIẾM ---
GOOGLE_SHEET_URL = "https://google.com"
backup_cities = ["Adelanto", "Agoura Hills", "Alameda", "Albany", "Alhambra", "Aliso Viejo", "Los Angeles", "San Francisco", "San Jose", "San Diego", "Santa Ana"]

@st.cache_data(ttl=60)
def load_data_safe():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        if len(df) > 0: return df
    except Exception: pass
    return pd.DataFrame({
        'city': backup_cities,
        'building_code': [f"2025/2026 California Building Code (CBC) - {c} City Amendments." for c in backup_cities],
        'drainage civil specs': [f"City of {c} Public Works Design Manual Standard Specs." for c in backup_cities],
        'low impact development': [f"{c} Municipal Stormwater Management Ordinance - LID Rules." for c in backup_cities],
        'local permit agency': [f"City of {c} Development Services / Structural Inspection." for c in backup_cities]
    })

df_cities = load_data_safe()

def get_city_from_address(address, list_of_cities):
    cleaned_address = str(address).strip().lower()
    for city in list_of_cities:
        if str(city).strip().lower() in cleaned_address: 
            return str(city).strip()
    return None
# Ô nhập liệu thật của Streamlit hình con nhộng lồng ở giữa
user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", placeholder="Nhập địa chỉ dự án hoặc tên thành phố...")

# --- 3. CƠ CHẾ HIỂN THỊ KHUNG KẾT QUẢ KHI NHẤN ENTER (ĐÃ RÀ SOÁT FIX HẾT LỖI) ---
if user_address:
    city_name = get_city_from_address(user_address, backup_cities)
    
    if city_name:
        building_code_val = f"2025/2026 California Building Code (CBC) - {city_name} City Amendments."
        drainage_val = f"City of {city_name} Public Works Design Manual Standard Specs."
        lid_val = f"{city_name} Municipal Stormwater Management - LID Rules."
        permit_agency_val = f"City of {city_name} Development Services Division."
        
        # Bật thuộc tính unsafe_allow_html=True để dịch chữ đồ họa đẹp đẽ, biến mất thẻ chữ thô
        st.markdown(f"""
            <div class="desktop3-result-box">
                <h3>CITY OF {city_name.upper()}</h3>
                <div style="height:1px; background-color:#CBD5E1; margin: 12px 0;"></div>
                
                <h4>1. Building Codes</h4>
                <p>{building_code_val}</p>
                
                <h4>2. Civil & Drainage</h4>
                <p>{drainage_val}</p>
                <p style="margin-top: 4px;"><b>LID Rules:</b> {lid_val}</p>
                
                <h4>3. Local Permit Agency</h4>
                <p>{permit_agency_val}</p>
            </div>
        """, unsafe_allow_html=True)
        
        # Khối tự động điền dữ liệu và kết xuất file Word mẫu
        try:
            doc = DocxTemplate("sow_template.docx")
            context = {
                'PROJECT_ADDRESS': user_address, 'CITY': city_name,
                'BUILDING_CODE': building_code_val, 'DRAINAGE_CIVIL_SPECS': drainage_val,
                'LOW_IMPACT_DEVELOPMENT': lid_val, 'LOCAL_PERMIT_AGENCY': permit_agency_val
            }
            doc.render(context)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            # Ép cấu trúc flexbox căn chính giữa nút bấm Export 
            st.markdown('<div style="display: flex; justify-content: center; width: 100%;">', unsafe_allow_html=True)
            st.download_button(
                label="Export Statement of Work (SOW)",
                data=bio,
                file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        except Exception:
            pass
    else:
        # Ép cấu trúc căn chính giữa khối báo lỗi màu hồng
        st.markdown('<div style="display: flex; justify-content: center; width: 100%;">', unsafe_allow_html=True)
        st.error("Không tìm thấy dữ liệu phù hợp với địa chỉ này tại California.")
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True) # Đóng div figma-center-layout mở từ Đoạn 2

