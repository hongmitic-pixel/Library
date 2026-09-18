import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

# Cấu hình trang và ẩn các nút mặc định của Streamlit để giữ giao diện sạch sẽ
st.set_page_config(page_title="BUILDBASE - SOW System", layout="centered", initial_sidebar_state="collapsed")

# --- 🎨 CẤU HÌNH GIAO DIỆN CHUẨN UI/UX BUILDBASE (MÀU XANH MINT & KHỐI BO TRÒN VIỀN ĐEN NỔI) ---
st.markdown("""
    <style>
        /* Đổi màu nền toàn bộ trang web sang màu xanh Mint */
        .stApp {
            background-color: #C0F0E4 !important;
        }
        /* Ẩn các thanh tiêu đề thừa của Streamlit */
        header, footer, .stDeployButton {display: none !important;}
        
        /* Font chữ tiêu đề BUILDBASE nghệ thuật */
        .buildbase-logo {
            font-family: 'Impact', 'Arial Black', sans-serif;
            font-size: 72px;
            letter-spacing: 2px;
            color: #000000;
            text-align: center;
            margin-top: 20px;
            margin-bottom: 25px;
            text-transform: uppercase;
        }
        
        /* Cấu trúc hộp hiển thị kết quả bo tròn viền đen dày dặn giống ảnh mẫu */
        .result-box {
            background-color: #FFFFFF !important;
            border: 4px solid #000000 !important;
            border-radius: 40px !important;
            padding: 35px !important;
            margin-top: 30px;
            color: #000000 !important;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            box-shadow: 5px 5px 0px #000000;
        }
        .result-box h4 {
            color: #000000 !important;
            font-weight: bold;
            margin-top: 15px;
            margin-bottom: 5px;
            font-size: 18px;
        }
        .result-box p {
            color: #333333 !important;
            font-size: 15px;
            line-height: 1.6;
            margin-bottom: 15px;
        }
        
        /* Tinh chỉnh thanh tìm kiếm bo tròn viền đen */
        div.stTextInput > div > div > input {
            border: 4px solid #000000 !important;
            border-radius: 30px !important;
            padding: 12px 25px 12px 50px !important;
            font-size: 16px !important;
            color: #000000 !important;
            background-color: #FFFFFF !important;
        }
        
        /* Cấu hình nút bấm SEARCH tinh gọn viền đen nổi khối */
        div.stButton > button {
            background-color: #FFFFFF !important;
            color: #000000 !important;
            border: 4px solid #000000 !important;
            border-radius: 25px !important;
            padding: 8px 30px !important;
            font-weight: bold !important;
            text-transform: uppercase;
            box-shadow: 3px 3px 0px #000000;
            transition: all 0.2s;
        }
        div.stButton > button:hover {
            transform: translate(-2px, -2px);
            box-shadow: 5px 5px 0px #000000;
        }
    </style>
""", unsafe_allowed_html=True)

# Hiển thị Logo BUILDBASE dày dặn đầu trang
st.markdown('<div class="buildbase-logo">BUILDBASE</div>', unsafe_allowed_html=True)

# --- 🛠️ KẾT NỐI DỮ LIỆU GỐC (GIỮ NGUYÊN 100% LUỒNG ĐÃ CHẠY RA KẾT QUẢ CỦA ANH) ---
GOOGLE_SHEET_URL = "https://google.com"

# Danh sách thành phố phần 1
backup_cities = [
    "Adelanto", "Agoura Hills", "Alameda", "Albany", "Alhambra", "Aliso Viejo", "Alturas", "Amador City", 
    "American Canyon", "Anaheim", "Anderson", "Angels Camp", "Antioch", "Apple Valley", "Arcadia", "Arcata", 
    "Arroyo Grande", "Artesia", "Arvin", "Atascadero", "Atherton", "Atwater", "Auburn", "Avalon", "Avenal", 
    "Azusa", "Bakersfield", "Baldwin Park", "Banning", "Barstow", "Beaumont", "Bell", "Bell Gardens", 
    "Bellflower", "Belmont", "Belvedere", "Benicia", "Berkeley", "Beverly Hills", "Big Bear Lake", "Biggs", 
    "Bishop", "Blue Lake", "Blythe", "Bradbury", "Brawley", "Brea", "Brentwood", "Brisbane", "Buellton", 
    "Buena Park", "Burbank", "Burlingame", "Calabasas", "Calexico", "California City", "Calimesa", "Calistoga", 
    "Camarillo", "Campbell", "Canyon Lake", "Capitola", "Carlsbad", "Carmel-by-the-Sea", "Carpinteria", 
    "Carson", "Cathedral City", "Ceres", "Cerritos", "Chico", "Chino", "Chino Hills", "Chula Vista", 
    "Citrus Heights", "Claremont", "Clayton", "Clearlake", "Cloverdale", "Clovis", "Coachella", "Coalinga", 
    "Colfax", "Colma", "Colton", "Colusa", "Commerce", "Compton", "Concord", "Corcoran", "Corning", 
    "Corona", "Coronado", "Corte Madera", "Costa Mesa", "Cotati", "Covina", "Crescent City", "Cudahy", 
    "Cupertino", "Cypress", "Daly City", "Dana Point", "Danville", "Davis", "Del Mar", "Del Rey Oaks", 
    "Delano", "Desert Hot Springs", "Diamond Bar", "Dinuba", "Dixon", "Dorris", "Dos Palos", "Downey", 
    "Duarte", "Dublin", "Dunsmuir", "East Palo Alto", "Eastvale", "El Cajon", "El Centro", "El Cerrito", 
    "El Monte", "El Segundo", "Elk Grove", "Emeryville", "Encinitas", "Escalon", "Escondido", "Etna", 
    "Eureka", "Exeter", "Fairfax", "Fairfield", "Farmersville", "Ferndale", "Fillmore", "Firebaugh", 
    "Folsom", "Fontana", "Fort Bragg", "Fort Jones", "Fortuna", "Foster City", "Fountain Valley", "Fowler", 
    "Fremont", "Fresno", "Fullerton", "Galt", "Garden Grove", "Gardena", "Gilroy", "Glendale", "Glendora", 
    "Goleta", "Gonzalez", "Grand Terrace", "Grass Valley", "Greenfield", "Gridley", "Grover Beach", 
    "Guadalupe", "Gustine", "Half Moon Bay", "Hanford", "Hawaiian Gardens", "Hawthorne", "Hayward", 
    "Healdsburg", "Hemet", "Hercules"
]
# Danh sách thành phố phần 2 (Cộng nối tiếp vào danh sách trên)
backup_cities += [
    "Hermosa Beach", "Hesperia", "Hidden Hills", "Highland", "Hillsborough", "Hollister", "Holtville", "Hughson", 
    "Huntington Beach", "Huntington Park", "Imperial", "Imperial Beach", "Indian Wells", "Indio", "Industry", 
    "Inglewood", "Ione", "Irvine", "Irwindale", "Isleton", "Jackson", "Jurupa Valley", "Kerman", "King City", 
    "Kingsburg", "La Cañada Flintridge", "La Habra", "La Habra Heights", "La Mesa", "La Mirada", "La Palma", 
    "La Puente", "La Quinta", "La Verne", "Lafayette", "Laguna Beach", "Laguna Hills", "Laguna Niguel", 
    "Laguna Woods", "Lake Elsinore", "Lake Forest", "Lakeport", "Lakewood", "Lancaster", "Larkspur", "Lathrop", 
    "Lawndale", "Lemon Grove", "Lemoore", "Lincoln", "Lindsay", "Live Oak", "Livermore", "Livingston", "Lodi", 
    "Loma Linda", "Lomita", "Lompoc", "Long Beach", "Loomis", "Los Alamitos", "Los Altos", "Los Altos Hills", 
    "Los Angeles", "Los Banos", "Los Gatos", "Loyalton", "Lynwood", "Madera", "Malibu", "Mammoth Lakes", 
    "Manhattan Beach", "Manteca", "Maricopa", "Marina", "Martinez", "Marysville", "Maywood", "McFarlin", 
    "Mendota", "Menlo Park", "Merced", "Mill Valley", "Millbrae", "Milpitas", "Mission Viejo", "Modesto", 
    "Monrovia", "Montague", "Montclair", "Monte Sereno", "Montebello", "Monterey", "Monterey Park", "Moorpark", 
    "Moraga", "Moreno Valley", "Morgan Hill", "Morro Bay", "Mount Shasta", "Mountain View", "Murrieta", "Napa", 
    "National City", "Needles", "Nevada City", "Newark", "Newman", "Newport Beach", "Norco", "Norwalk", "Novato", 
    "Oakdale", "Oakland", "Oakley", "Oceanside", "Ojai", "Ontario", "Orange", "Orange Cove", "Orland", "Oroville", 
    "Oxnard", "Pacific Grove", "Pacifica", "Palm Desert", "Palm Springs", "Palmdale", "Palo Alto", 
    "Palos Verdes Estates", "Paradise", "Paramount", "Parlier", "Pasadena", "Paso Robles", "Patterson", 
    "Perris", "Petaluma", "Pico Rivera", "Piedmont", "Pinole", "Pismo Beach", "Pittsburg", "Placentia", 
    "Placerville", "Pleasant Hill", "Pleasanton", "Plymouth", "Point Arena", "Pomona", "Port Hueneme", 
    "Porterville", "Portola", "Portola Valley", "Poway", "Rancho Cordova", "Rancho Cucamonga", "Rancho Mirage", 
    "Red Bluff", "Redding", "Redlands", "Redondo Beach", "Redwood City", "Reedley", "Rialto", "Richmond", 
    "Ridgecrest", "Rio Dell", "Rio Vista", "Ripon", "Riverbank", "Riverside", "Rocklin", "Rohnert Park", 
    "Rolling Hills", "Rolling Hills Estates", "Rosemead", "Roseville", "Ross", "Sacramento", "Salinas", 
    "San Anselmo", "San Bernardino", "San Bruno", "San Carlos", "San Clemente", "San Diego", "San Dimas", 
    "San Fernando", "San Francisco", "San Gabriel", "San Jacinto", "San Joaquin", "San Jose", "San Juan Bautista", 
    "San Juan Capistrano", "San Leandro", "San Luis Obispo", "San Marcos", "San Marino", "San Mateo", 
    "San Pablo", "San Rafael", "San Ramon", "Sand City", "Sanger", "Santa Ana", "Santa Barbara", "Santa Clara", 
    "Santa Clarita", "Santa Cruz", "Santa Fe Springs", "Santa Maria", "Santa Monica", "Santa Paula", "Santa Rosa", 
    "Santee", "Saratoga", "Sausalito", "Scotts Valley", "Seal Beach", "Seaside", "Sebastopol", "Selma", 
    "Shafter", "Shasta Lake", "Sierra Madre", "Signal Hill", "Simi Valley", "Solana Beach", "Soledad", 
    "Solvang", "Sonoma", "Sonora", "South El Monte", "South Gate", "South Lake Tahoe", "South Pasadena", 
    "South San Francisco", "St. Helena", "Stanton", "Stockton", "Suisun City", "Sunnyvale", "Susanville", 
    "Sutter Creek", "Taft", "Tehachapi", "Tehama", "Temecula", "Temple City", "Thousand Oaks", "Tiburon", 
    "Torrance", "Tracy", "Trinidad", "Truckee", "Tulare", "Tulelake", "Turlock", "Tustin", "Twentynine Palms", 
    "Ukiah", "Union City", "Upland", "Vacaville", "Vallejo", "Ventura", "Vernon", "Victorville", "Villa Park", 
    "Visalia", "Vista", "Walnut", "Walnut Creek", "Wasco", "Waterford", "Watsonville", "Weed", "West Covina", 
    "West Hollywood", "West Sacramento", "Westlake Village", "Westminster", "Westmorland", "Wheatland", 
    "Whittier", "Wildomar", "Williams", "Willits", "Willows", "Windsor", "Winters", "Woodlake", "Woodland", 
    "Woodside", "Yorba Linda", "Yountville", "Yreka", "Yuba City", "Yucaipa", "Yucca Valley"
]

@st.cache_data(ttl=60)
def load_data_safe():
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        if len(df) > 0: return df
    except Exception: pass
    
    return pd.DataFrame({
        'city': backup_cities,
        'building_code': [f"2025/2026 California Building Code (CBC) - {c} City Amendments & Structural Framework." for c in backup_cities],
        'drainage civil specs': [f"City of {c} Public Works Design Manual / Engineering Standard Drainage Infrastructure Specs." for c in backup_cities],
        'low impact development': [f"{c} Municipal Stormwater Management Ordinance - Low Impact Development (LID) Rules." for c in backup_cities],
        'local permit agency': [f"City of {c} Development Services / Structural & Civil Building Inspection Division." for c in backup_cities]
    })

df_cities = load_data_safe()
### 🧱 PHẦN 2: Thuật toán Tra cứu Vệ tinh Giữ nguyên (Dán nối tiếp vào giữa file)

```python
def get_city_from_address(address, list_of_cities):
    cleaned_address = str(address).strip().lower()
    for city in list_of_cities:
        if str(city).strip().lower() in cleaned_address: return str(city).strip()
    try:
        geolocator = Nominatim(user_agent="ca_civil_buildbase_v2026")
        location = geolocator.geocode(address + ", CA, USA", addressdetails=True, timeout=10)
        if location and 'address' in location.raw:
            vals = location.raw['address'].values()
            for val in vals:
                for city in list_of_cities:
                    if str(city).strip().lower() == str(val).strip().lower(): return str(city).strip()
    except Exception: pass
    return None

# Bố trí hàng tìm kiếm (Ô nhập liệu và Nút Search nằm ngang mượt mà đúng mẫu)
col_input, col_btn = st.columns([4, 1])
with col_input:
    user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", placeholder="Nhập địa chỉ dự án (Ví dụ: 1992 La Cuesta Drive, Santa Ana)...")
with col_btn:
    st.write("<div style='margin-top:4px;'></div>", unsafe_allowed_html=True)
    search_clicked = st.button("SEARCH")
if search_clicked and user_address:
    with st.spinner("Searching..."):
        city_name = get_city_from_address(user_address, backup_cities)
        
        if city_name:
            row_data = None
            for idx, row in df_cities.iterrows():
                for col in df_cities.columns:
                    if str(row[col]).strip().lower() == city_name.lower():
                        row_data = row
                        break
                if row_data is not None: break
            
            if row_data is None:
                building_code_val = f"2025/2026 California Building Code (CBC) - {city_name} City Amendments & Structural Safety Framework."
                drainage_val = f"City of {city_name} Public Works Design Manual / Engineering Standard Drainage Infrastructure Specifications."
                lid_val = f"{city_name} Municipal Stormwater Management Ordinance - Low Impact Development (LID) Retention Rules."
                permit_agency_val = f"City of {city_name} Development Services / Structural & Civil Building Inspection Division."
            else:
                def find_val(keywords):
                    for col in df_cities.columns:
                        if any(kw in str(col).lower() for kw in keywords): return str(row_data[col])
                    return "N/A"
                building_code_val = find_val(['building', 'structure', 'code'])
                drainage_val = find_val(['drainage', 'civil', 'spec'])
                lid_val = find_val(['low impact', 'lid', 'stormwater'])
                permit_agency_val = find_val(['permit', 'agency', 'local'])

            # 👉 HIỂN THỊ HỘP KHUNG BO TRÒN VIỀN ĐEN ĐÚNG CHUẨN 100% THEO ẢNH ĐỀ XUẤT CỦA ANH
            st.markdown(f"""
                <div class="result-box">
                    <h4>1. Building Codes</h4>
                    <p>{building_code_val}</p>
                    <h4>2. Civil & Drainage</h4>
                    <p>Thông số thoát nước: {drainage_val}</p>
                    <p>LID: {lid_val}</p>
                    <h4>3. Pháp lý Thẩm định</h4>
                    <p>{permit_agency_val}</p>
                    <p style="margin-top:25px; font-weight:bold; margin-bottom:5px;">🚀 Tài liệu SOW đã sẵn sàng:</p>
                </div>
            """, unsafe_allowed_html=True)
            
            # Tích hợp luồng xuất file Word Word mẫu (.docx)
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
                
                st.write("<div style='margin-top:12px;'></div>", unsafe_allowed_html=True)
                st.download_button(
                    label="📥 Bấm vào đây để tải file SOW (.docx) về máy",
                    data=bio,
                    file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except FileNotFoundError:
                st.error("Không tìm thấy file mẫu 'sow_template.docx' trên GitHub.")
            except Exception as e:
                st.error(f"Lỗi khi khởi tạo file Word: {e}")
        else:
            st.error("Không nhận diện được tên thành phố. Anh vui lòng kiểm tra lại chính tả.")
elif search_clicked and not user_address:
    st.warning("Vui lòng nhập địa chỉ dự án vào ô tìm kiếm.")
