import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io
import os

# Cấu hình trang tối giản, ẩn các thành phần thừa của Streamlit
st.set_page_config(page_title="LINE BASE - SOW System", layout="wide", initial_sidebar_state="collapsed")

# --- 🎨 HỆ THỐNG GIAO DIỆN CHUẨN FIGMA 100% CĂN GIỮA TUYỆT ĐỐI ---
st.markdown("""
    <style>
        /* Toàn bộ nền trang web phẳng sạch sẽ */
        .stApp {
            background-color: #FFFFFF !important;
        }
        /* Giấu triệt để các thành phần hệ thống mặc định của Streamlit */
        header, footer, .stDeployButton, div[data-testid="stToolbar"] {display: none !important;}
        
        /* CONTAINER TỔNG: Căn giữa toàn bộ giao diện theo Figma */
        .figma-layout {
            max-width: 800px;
            margin: 0 auto !important;
            padding: 40px 20px 120px 20px !important;
            text-align: center !important;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        
        /* ĐỊNH DẠNG VÙNG CHỨA LOGO HÌNH ẢNH CĂN GIỮA */
        .logo-image-container {
            margin-bottom: 25px;
            width: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        
        /* TIÊU ĐỀ QUICK SEARCH MÀU XANH TEAL THEO FIGMA */
        .quick-search-title {
            font-size: 36px;
            font-weight: 700;
            color: #1E6B7B;
            text-transform: uppercase;
            margin-bottom: 25px;
            letter-spacing: -0.5px;
            width: 100%;
            text-align: center;
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
            width: 100%;
            max-width: 500px;
            border: 1px solid #E2E8F0;
            text-align: left !important;
            margin: 0 auto !important;
            box-shadow: 0px 4px 20px rgba(0, 0, 0, 0.01);
        }
        .desktop3-result-box h3 {
            color: #1E6B7B !important;
            margin-top: 0;
            font-size: 18px;
            font-weight: 700;
            text-align: left !important;
        }
        .desktop3-result-box h4 {
            color: #1E6B7B !important;
            font-weight: 700;
            font-size: 13.5px;
            margin-top: 15px;
            margin-bottom: 4px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .desktop3-result-box p {
            color: #475569 !important;
            font-size: 14px;
            line-height: 1.5;
            margin: 0;
        }

        /* NÚT TẢI FILE WORD ĐỒNG BỘ NẰM GIỮA */
        div.stDownloadButton {
            width: 100%;
            text-align: center !important;
            margin-top: 20px;
        }
        div.stDownloadButton > button {
            background: #1E6B7B !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 6px !important;
            padding: 10px 24px !important;
            font-size: 14px !important;
            font-weight: 500 !important;
            transition: background 0.2s;
            display: inline-block;
        }
        div.stDownloadButton > button:hover {
            background: #154D59 !important;
        }
    </style>
""", unsafe_allow_html=True)
# --- 📐 MỞ KHUNG CHỨA LAYOUT CĂN GIỮA ---
st.markdown('<div class="figma-layout">', unsafe_allow_html=True)

# 1. KHU VỰC HIỂN THỊ LOGO CĂN GIỮA TỰ ĐỘNG THÔNG MINH
st.markdown('<div class="logo-image-container">', unsafe_allow_html=True)
if os.path.exists("logo_line_base.png"):
    st.image("logo_line_base.png", width=160)
else:
    # Dự phòng thông minh: Tự vẽ lại logo gốc bằng CSS ở giữa nếu bạn chưa upload ảnh lên GitHub thành công
    st.markdown("""
        <div style="display: flex; align-items: center; gap: 10px;">
            <div style="width: 50px; height: 50px; border-radius: 12px; background: #FFFFFF; border: 3px solid transparent; background-image: linear-gradient(#fff, #fff), linear-gradient(135deg, #38bdf8, #1d4ed8, #0ea5e9); background-origin: border-box; background-clip: content-box, border-box; display: flex; align-items: center; justify-content: center;">
                <span style="font-family: 'Arial Black', sans-serif; font-size: 24px; font-weight: 900; background: linear-gradient(180deg, #38bdf8 30%, #1e40af 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -2px; position: relative; left: -1px;">LB</span>
            </div>
            <div style="font-family: sans-serif; font-size: 22px; font-weight: 800; color: #1E293B; letter-spacing: -0.5px; margin-top: 4px;">LINE BASE</div>
        </div>
    """, unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# 2. Tiêu đề QUICK SEARCH đặt đúng vị trí chính giữa
st.markdown('<div class="quick-search-title">QUICK SEARCH</div>', unsafe_allow_html=True)
# --- 🛠️ CƠ SỞ DỮ LIỆU THÀNH PHỐ CALIFORNIA ---
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
# 3. Ô nhập liệu thật của Streamlit lồng vào vị trí thiết kế hình con nhộng ở giữa
user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", placeholder="Nhập địa chỉ dự án hoặc tên thành phố...")

# --- 4. CƠ CHẾ HIỂN THỊ KHUNG KẾT QUẢ KHI NHẤN ENTER (SỬA LỖI HIỂN THỊ CHỮ HTML THÔ) ---
if user_address:
    city_name = get_city_from_address(user_address, backup_cities)
    
    if city_name:
        building_code_val = f"2025/2026 California Building Code (CBC) - {city_name} City Amendments."
        drainage_val = f"City of {city_name} Public Works Design Manual Standard Specs."
        lid_val = f"{city_name} Municipal Stormwater Management - LID Rules."
        permit_agency_val = f"City of {city_name} Development Services Division."
        
        # 👉 FIX LỖI TẬN GỐC: Sử dụng st.markdown kèm unsafe_allow_html=True để hiển thị chuẩn xác các thẻ nội dung
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
        
        # Khối tự động điền và kết xuất file Word mẫu
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
            
            st.download_button(
                label="Export Statement of Work (SOW)",
                data=bio,
                file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception:
            pass
    else:
        st.error("Không tìm thấy dữ liệu phù hợp với địa chỉ này tại California.")

st.markdown('</div>', unsafe_allow_html=True) # Đóng div figma-layout
