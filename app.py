import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Hệ thống xuất SOW tự động", layout="centered")
st.title("🏗️ Hệ Thống Tra Cứu Tự Động")
st.write("Dành riêng cho dự án Xây dựng Dân dụng & Hạ tầng Civil tại California")

# --- ĐƯỜNG LINK TRỤC DỮ LIỆU TRỰC TIẾP TỪ TRANG XUẤT BẢN CỦA ANH ---
GOOGLE_SHEET_URL = "https://google.com"

@st.cache_data(ttl=600)  # Tự động đồng bộ sau mỗi 10 phút nếu anh sửa file Sheets
def load_data_from_sheets():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        # 👉 TỰ ĐỘNG SỬA LỖI: Làm sạch khoảng trắng và ép tất cả tên cột về chữ thường hoàn toàn
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.error(f"Không thể kết nối tới kho dữ liệu Google Sheets từ máy chủ đám mây. Lỗi: {e}")
        return None

# Nạp dữ liệu bảng tính từ Google Sheets của anh
df_cities = load_data_from_sheets()

# Hàm bóc tách địa chỉ đa tầng thông minh phù hợp môi trường Cloud
def get_city_from_address(address, df_city_list):
    try:
        optimized_address = address
        if "california" not in address.lower() and "ca" not in address.lower():
            optimized_address += ", CA"
        if "usa" not in address.lower():
            optimized_address += ", USA"
            
        geolocator = Nominatim(user_agent="ca_civil_sow_generator_cloud_ultimate_final")
        location = geolocator.geocode(optimized_address, addressdetails=True, timeout=10)
        
        if location and 'address' in location.raw:
            address_details = location.raw['address']
            possible_places = [
                address_details.get('city'), address_details.get('town'),
                address_details.get('suburb'), address_details.get('village'),
                address_details.get('municipality'), address_details.get('county')
            ]
            for place in possible_places:
                if place and any(df_city_list.str.lower() == place.lower()):
                    return place
    except Exception:
        pass
        
    # Cơ chế dự phòng quét chữ thô trực tiếp
    for city in df_city_list:
        if str(city).lower() in address.lower():
            return city
    return None

# Giao diện người dùng
user_address = st.text_input("📍 Nhập địa chỉ dự án tại California:", placeholder="Ví dụ: 1992 La Cuesta Drive, Santa Ana")

if st.button("🔍 Tra cứu & Chuẩn bị SOW"):
    if df_cities is None:
        st.error("Lỗi: Hệ thống đám mây chưa kết nối được dữ liệu nguồn Google Sheets.")
    elif user_address:
        with st.spinner("Hệ thống đám mây đang bóc tách địa chỉ dự án..."):
            # 👉 ĐÃ SỬA: Gọi cột 'city' viết thường an toàn tuyệt đối theo chuẩn hóa ở trên
            city_col_name = 'city' if 'city' in df_cities.columns else df_cities.columns[0]
            
            city_name = get_city_from_address(user_address, df_cities[city_col_name])
            
            if city_name:
                match = df_cities[df_cities[city_col_name].str.lower() == city_name.lower()]
                
                if not match.empty:
                    city_info = match.iloc[0]
                    
                    # Bộ lọc thông minh tự tìm cột không phân biệt hoa thường hay khoảng trắng
                    def get_column_value(keywords):
                        for col in df_cities.columns:
                            if any(kw in col for kw in keywords):
                                return city_info[col]
                        return "N/A"

                    building_code_val = get_column_value(['building', 'structure'])
                    drainage_val = get_column_value(['drainage', 'civil', 'spec'])
                    lid_val = get_column_value(['low impact', 'lid', 'stormwater'])
                    permit_agency_val = get_column_value(['permit', 'agency', 'local'])
                    
                    st.success(f"🎯 Đã xác định được thành phố: **{city_name}**")
                    st.write("---")
                    st.markdown(f"### 🏠 1. Tiêu chuẩn Xây dựng (Building Codes)")
                    st.write(f"{building_code_val}")
                    
                    st.markdown(f"### 💧 2. Hạ tầng & Thoát nước (Civil & Drainage)")
                    st.write(f"**Thông số thoát nước:** {drainage_val}")
                    st.write(f"**Quản lý nước mưa (LID):** {lid_val}")
                    
                    st.markdown(f"### 📋 3. Pháp lý Thẩm định")
                    st.write(f"**Cơ quan cấp phép:** {permit_agency_val}")
                    
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
                st.error("Không nhận diện được tên thành phố từ địa chỉ này. Anh vui lòng nhập rõ số nhà, tên đường, bang CA.")
    else:
        st.warning("Vui lòng gõ địa chỉ dự án vào ô tìm kiếm.")


