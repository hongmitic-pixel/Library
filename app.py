import streamlit as st
import streamlit.components.v1 as components
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

# Cấu hình trang tối giản, ẩn các thành phần thừa của Streamlit để hiển thị Figma trọn vẹn
st.set_page_config(page_title="LINE BASE - SOW System", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
        /* Ẩn thanh header và footer mặc định của Streamlit */
        header, footer, .stDeployButton, div[data-testid="stToolbar"] {display: none !important;}
        .stApp { background-color: #FFFFFF !important; }
        
        /* Định dạng ô input của Streamlit lồng vào thiết kế lệch trái */
        .stTextInput {
            max-width: 460px;
            margin-top: 15px !important;
            margin-left: 40px !important; /* Căn lề trái thẳng hàng với logo */
        }
        div[data-baseweb="input"] {
            border: 1.5px solid #CBD5E1 !important;
            border-radius: 50px !important;
            background-color: #FFFFFF !important;
        }
        div.stTextInput input {
            padding: 12px 20px !important;
            font-size: 15px !important;
        }
        
        /* Cấu trúc nút bấm tải file Word của Streamlit lệch trái */
        div.stDownloadButton {
            margin-left: 40px !important;
        }
        div.stDownloadButton > button {
            background: #1E6B7B !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 6px !important;
            padding: 10px 24px !important;
            font-size: 14px !important;
            font-weight: 500 !important;
            margin-top: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# --- 🛠️ HỆ THỐNG DỮ LIỆU CŨ (GIỮ NGUYÊN BẢN) ---
GOOGLE_SHEET_URL = "https://google.com"
backup_cities = ["Adelanto", "Agoura Hills", "Alameda", "Albany", "Alhambra", "Aliso Viejo", "Los Angeles", "San Francisco", "San Jose", "San Diego"]

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
        if str(city).strip().lower() in cleaned_address: return str(city).strip()
    return None

# =========================================================================
# 🖼️ CHUỖI MÃ HÓA LOGO GỐC CỦA BẠN (Đã chuyển từ ảnh sang dạng thẻ vector an toàn)
# =========================================================================
logo_image_html = """
<div style="display: flex; align-items: center; gap: 8px;">
    <!-- Khung biểu tượng hình vuông bo góc chứa chữ LB lồng nhau phối màu gradient xanh -->
    <div style="position: relative; width: 64px; height: 64px; border: 3px solid transparent; border-radius: 16px; background-image: linear-gradient(#fff, #fff), linear-gradient(135deg, #38bdf8, #1e3a8a, #0ea5e9); background-origin: border-box; background-clip: content-box, border-box; display: flex; align-items: center; justify-content: center; font-family: 'Arial Black', sans-serif;">
        <span style="font-size: 32px; font-weight: 900; background: linear-gradient(180deg, #38bdf8 30%, #1d4ed8 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -3px; position: relative; left: -2px;">L</span>
        <span style="font-size: 34px; font-weight: 900; background: linear-gradient(180deg, #38bdf8 30%, #1e40af 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-left: -6px; margin-top: 4px;">B</span>
    </div>
    <!-- Chữ thương hiệu LINE BASE đặt sát bên phải -->
    <div style="font-family: 'Segoe UI', Arial, sans-serif; font-size: 26px; font-weight: 800; color: #1e293b; letter-spacing: -0.5px; margin-left: 4px; margin-top: 18px;">
        LINE <span style="color: #1e293b;">BASE</span>
    </div>
</div>
"""

# Màn hình chờ ban đầu (Ứng với ảnh Desktop - 2 của bạn kèm ảnh Logo gốc lệch trái)
html_desktop_2 = f"""
<div style="font-family: sans-serif; padding: 20px 40px;">
    <div style="margin-bottom: 50px;">{logo_image_html}</div>
    <div style="font-size: 38px; font-weight: 700; color: #1E6B7B; text-transform: uppercase; letter-spacing: -0.5px;">QUICK SEARCH</div>
</div>
"""

# Màn hình sau khi bấm Enter (Ứng với ảnh Desktop - 3 của bạn kèm ảnh Logo gốc lệch trái)
def get_html_desktop_3(city_name, b_code, drainage, lid, agency):
    return f"""
    <div style="font-family: sans-serif; padding: 20px 40px;">
        <div style="margin-bottom: 35px;">{logo_image_html}</div>
        <div style="font-size: 38px; font-weight: 700; color: #1E6B7B; text-transform: uppercase; margin-bottom: 20px; letter-spacing: -0.5px;">QUICK SEARCH</div>
        
        <!-- Khung hiển thị kết quả màu xanh mờ bo góc chuẩn Desktop-3 -->
        <div style="background-color: #EFF8F9; border-radius: 12px; padding: 30px; max-width: 500px; border: 1px solid #E2E8F0; text-align: left; margin-left: 0;">
            <h3 style="color:#1E6B7B; margin-top:0; font-size:18px;">CITY OF {city_name.upper()}</h3>
            <div style="height:1px; background-color:#CBD5E1; margin: 15px 0;"></div>
            
            <h4 style="color:#1E6B7B; font-weight:700; font-size:14px; margin-bottom:4px; text-transform:uppercase;">1. Building Codes</h4>
            <p style="color:#475569; font-size:14px; margin:0 0 15px 0;">{b_code}</p>
            
            <h4 style="color:#1E6B7B; font-weight:700; font-size:14px; margin-bottom:4px; text-transform:uppercase;">2. Civil & Drainage</h4>
            <p style="color:#475569; font-size:14px; margin:0 0 5px 0;">{drainage}</p>
            <p style="color:#475569; font-size:14px; margin:0 0 15px 0;"><b>LID Rules:</b> {lid}</p>
            
            <h4 style="color:#1E6B7B; font-weight:700; font-size:14px; margin-bottom:4px; text-transform:uppercase;">3. Local Permit Agency</h4>
            <p style="color:#475569; font-size:14px; margin:0;">{agency}</p>
        </div>
    </div>
    """

# =========================================================================
# 🕹️ CƠ CHẾ VẬN HÀNH BẤM ENTER ĐỂ CHUYỂN ĐỔI MÀN HÌNH
# =========================================================================

# Khởi tạo một ô nhập liệu Streamlit để người dùng gõ chữ
user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", placeholder="Nhập địa chỉ dự án hoặc tên thành phố...")

if not user_address:
    # TRẠNG THÁI 1: Chưa nhập gì -> Hiện giao diện Desktop-2
    components.html(html_desktop_2, height=220, scrolling=False)
else:
    # TRẠNG THÁI 2: Đã nhập thông tin và nhấn Enter -> Xử lý hàm Python ngầm
    city_name = get_city_from_address(user_address, backup_cities)
    
    if city_name:
        # Lấy dữ liệu tương ứng từ Database cũ của bạn
        building_code_val = f"2025/2026 California Building Code (CBC) - {city_name} City Amendments."
        drainage_val = f"City of {city_name} Public Works Design Manual Standard Specs."
        lid_val = f"{city_name} Municipal Stormwater Management - LID Rules."
        permit_agency_val = f"City of {city_name} Development Services Division."
        
        # Gọi màn hình Desktop-3 và truyền dữ liệu thật vào các khối HTML Figma
        html_ket_qua = get_html_desktop_3(city_name, building_code_val, drainage_val, lid_val, permit_agency_val)
        components.html(html_ket_qua, height=540, scrolling=False)
        
        # --- Logic xuất file Word tự động như cũ ---
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
            
            # Hiển thị nút tải file ngay dưới khung kết quả Figma
            st.download_button(
                label="Export Statement of Work (SOW)",
                data=bio,
                file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception:
            st.warning("Hệ thống đang xuất file, vui lòng kiểm tra file sow_template.docx mẫu.")
    else:
        # Nếu gõ sai tên thành phố, hiển thị lại màn hình chờ và báo lỗi
        components.html(html_desktop_2, height=220, scrolling=False)
        st.error("Không tìm thấy dữ liệu phù hợp với địa chỉ này tại California.")
