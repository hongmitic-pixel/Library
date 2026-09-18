import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Hệ thống xuất SOW tự động", layout="centered")
st.title("🏗️ Hệ Thống Tra Cứu & Xuất SOW Tự Động")
st.write("Dành riêng cho dự án Xây dựng Dân dụng & Hạ tầng Civil tại California")

# --- ĐƯỜNG LINK TRỤC DỮ LIỆU THÔ CHUẨN XÁC TỪ TRANG XUẤT BẢN CSV CỦA ANH ---
GOOGLE_SHEET_URL = "https://google.com"

@st.cache_data(ttl=600)  # Tự động đồng bộ sau mỗi 10 phút nếu anh sửa file Sheets
def load_data_from_sheets():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        return df
    except Exception as e:
        st.error(f"Không thể kết nối tới kho dữ liệu Google Sheets từ máy chủ đám mây. Lỗi: {e}")
        return None

# Nạp dữ liệu bảng tính từ Google Sheets của anh
df_cities = load_data_from_sheets()

# Hàm bóc tách địa chỉ đa tầng thông minh
def get_city_from_address(address, df_city_list):
    # Cơ chế 1: Thuật toán quét chữ trực tiếp thông minh (Bất chấp viết hoa viết thường)
    for city in df_city_list:
        city_str = str(city).strip()
        if city_str.lower() in address.lower():
            return city_str

    # Cơ chế 2: Vệ tinh định vị dự phòng
    try:
        optimized_address = address
        if "california" not in address.lower() and "ca" not in address.lower():
            optimized_address += ", CA"
        if "usa" not in address.lower():
            optimized_address += ", USA"
            
        geolocator = Nominatim(user_agent="ca_civil_sow_generator_cloud_ultimate_final_v35")
        location = geolocator.geocode(optimized_address, addressdetails=True, timeout=10)
        
        if location and 'address' in location.raw:
            address_details = location.raw['address']
            possible_places = [
                address_details.get('city'), address_details.get('town'),
                address_details.get('suburb'), address_details.get('village'),
                address_details.get('municipality'), address_details.get('county')
            ]
            for place in possible_places:
                if place and any(df_city_list.astype(str).str.lower() == place.lower()):
                    for original_city in df_city_list:
                        if str(original_city).strip().lower() == place.lower():
                            return str(original_city).strip()
    except Exception:
        pass
    return None

# Giao diện người dùng
user_address = st.text_input("📍 Nhập địa chỉ dự án tại California:", placeholder="Ví dụ: 70610 Camellia Court, Rancho Mirage, CA 92270")

if st.button("🔍 Tra cứu & Chuẩn bị SOW"):
    if df_cities is None:
        st.error("Lỗi: Hệ thống đám mây chưa kết nối được dữ liệu nguồn Google Sheets.")
    elif user_address:
        with st.spinner("Hệ thống đám mây đang bóc tách địa chỉ dự án..."):
            
            # 👉 THUẬT TOÁN ĐÃ NÂNG CẤP: Tự động mò cột Thành phố thông minh bất chấp viết hoa/thường/tiếng Việt
            city_col = None
            for col in df_cities.columns:
                col_clean = str(col).strip().lower()
                if 'city' in col_clean or 'thành phố' in col_clean or 'thanh pho' in col_clean:
                    city_col = col
                    break
            
            if city_col is None:
                # Nếu quét từ khóa thất bại, hệ thống tự bốc luôn cột đầu tiên (Cột A) làm mặc định để bảo hiểm vĩnh viễn
                city_col = df_cities.columns[0]
                
            # Tiến hành bóc tách địa chỉ dựa trên cột đã dò tìm được
            city_name = get_city_from_address(user_address, df_cities[city_col])
            
            if city_name:
                match = df_cities[df_cities[city_col].astype(str).str.strip().str.lower() == city_name.lower()]
                
                if not match.empty:
                    city_info = match.iloc[0]
                    
                    # Bộ lọc thông minh tự quét từ khóa trong tiêu đề
                    def get_column_value(keywords):
                        for col in df_cities.columns:
                            col_lower = str(col).lower()
                            if any(kw in col_lower for kw in keywords):
                                return city_info[col]
                        return "N/A"

                    building_code_val = get_column_value(['building', 'structure', 'code', 'xây dựng'])
                    drainage_val = get_column_value(['drainage', 'civil', 'spec', 'thoát nước'])
                    lid_val = get_column_value(['low impact', 'lid', 'stormwater', 'nước mưa'])
                    permit_agency_val = get_column_value(['permit', 'agency', 'local', 'cấp phép'])
                    
                    st.success(f"🎯 Đã xác định được thành phố: **{city_name}**")
                    st.write("---")
                    st.markdown(f"### 🏠 1. Tiêu chuẩn Xây dựng (Building Codes)")
                    st.write(f"{building_code_val}")
                    
                    st.markdown(f"### 💧 2. Hạ tầng & Thoát nước (Civil & Drainage)")
                    st.write(f"**Thông số thoát nước:** {drainage_val}")
                    st.write(f"**Quản lý nước mưa (LID):** {lid_val}")
                    
                    st.markdown(f"### 📋 3. Pháp lý Thẩm định")
                    st.write(f"{permit_agency_val}")
                    
                    # --- XỬ LÝ ĐIỀN DATA VÀO FILE WORD SOW ---
                    try:
                        doc = DocxTemplate("sow_template.docx")
                        context = {
                            'PROJECT_ADDRESS': user_address,
                            'CITY': city_name,
                            'BUILDING_CODE': building_code_val,
                            'DRAINAGE_CIVIL_SPECS': drainage_val,
                            'LOW_IMPACT_DEVELOPMENT': lid_val,
                            'LOCAL_PERMIT_AGENCY': permit_agency_val
                        }
                        doc.render(context)
                        
                        bio = io.BytesIO()
                        doc.save(bio)
                        bio.seek(0)
                        
                        st.write("---")
                        st.subheader("🚀 Tài liệu SOW đã sẵn sàng:")
                        st.download_button(
                            label="📥 Tải file SOW (.docx) về máy",
                            data=bio,
                            file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except FileNotFoundError:
                        st.error("Không tìm thấy file mẫu 'sow_template.docx' trên GitHub. Anh hãy đảm bảo đã tải file mẫu này lên kho lưu trữ nhé.")
                    except Exception as e:
                        st.error(f"Lỗi khi khởi tạo file Word: {e}")
                else:
                    st.warning(f"Thành phố '{city_name}' hiện chưa được nạp dữ liệu kỹ thuật trên Google Sheets.")
            else:
                st.error("Không nhận diện được tên thành phố từ địa chỉ này. Anh vui lòng kiểm tra lại chính tả tên thành phố.")
    else:
        st.warning("Vui lòng gõ địa chỉ dự án vào ô tìm kiếm.")


