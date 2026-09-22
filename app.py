import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

# Cấu hình trang tối giản, loại bỏ hoàn toàn lề mặc định để dàn trang chuẩn Figma
st.set_page_config(page_title="BUILDBASE - QUICK SEARCH", layout="wide", initial_sidebar_state="collapsed")

# --- 🎨 HỆ THỐNG GIAO DIỆN CHUẨN FIGMA 100% (PIXEL-PERFECT) ---
st.markdown("""
    <style>
        /* Toàn bộ nền trang web phẳng sạch sẽ */
        .stApp {
            background-color: #FFFFFF !important;
        }
        /* Giấu triệt để các thành phần hệ thống của Streamlit */
        header, footer, .stDeployButton, div[data-testid="stToolbar"] {display: none !important;}
        
        /* --- HEADER CONTAINER --- */
        .figma-header {
            display: flex;
            align-items: center;
            justify-content: flex-start;
            padding: 20px 40px;
            width: 100%;
            background-color: #FFFFFF;
        }
        
        /* Giả lập logo LINE BASE chữ đậm phối xanh chuẩn thiết kế */
        .line-base-logo {
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
            font-size: 20px;
            font-weight: 800;
            color: #0F172A;
            display: flex;
            align-items: center;
            gap: 6px;
            letter-spacing: -0.5px;
        }
        .line-base-logo span {
            color: #38BDF8; /* Màu xanh thương hiệu nhánh */
        }

        /* --- BODY CONTENT CONTAINER --- */
        .figma-body {
            max-width: 800px;
            margin: 0 auto;
            padding: 60px 20px 120px 20px;
            text-align: center;
        }
        
        /* Tiêu đề QUICK SEARCH màu xanh loang Gradient nhẹ / Xanh Slate thanh lịch */
        .quick-search-title {
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
            font-size: 42px;
            font-weight: 700;
            letter-spacing: -0.02em;
            color: #1E6B7B; 
            margin-bottom: 30px;
            text-transform: uppercase;
        }

        /* --- THANH TÌM KIẾM HÌNH CON NHỘNG ĐÚNG KÍCH THƯỚC FIGMA --- */
        .stTextInput, .stTextInput > div, .stTextInput > div > div, div[data-baseweb="input"] {
            border: none !important;
            background-color: transparent !important;
            box-shadow: none !important;
            border-radius: 0px !important;
        }
        
        div.stTextInput input {
            border: 1.5px solid #CBD5E1 !important;
            border-radius: 50px !important; /* Bo cong tròn tuyệt đối hình con nhộng */
            padding: 14px 25px 14px 60px !important;
            font-size: 16px !important;
            color: #334155 !important;
            background-color: #FFFFFF !important;
            box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.04) !important;
            transition: all 0.2s ease;
            max-width: 600px;
            margin: 0 auto;
        }
        
        div.stTextInput input:focus {
            border: 1.5px solid #1E6B7B !important;
            box-shadow: 0px 4px 16px rgba(30, 107, 123, 0.15) !important;
            outline: none !important;
        }
        
        /* Nhúng Icon kính lúp mảnh dẻ định dạng vector chuẩn chỉ */
        div.stTextInput input {
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://w3.org' viewBox='0 0 24 24' fill='none' stroke='%2394A3B8' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'%3E%3C/circle%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'%3E%3C/line%3E%3C/svg%3E") !important;
            background-repeat: no-repeat !important;
            background-position: 24px center !important;
            background-size: 18px 18px !important;
        }

        /* --- KHUNG HIỂN THỊ KẾT QUẢ BOX (DESKTOP - 3) --- */
        .desktop3-result {
            background-color: #F0F9FA !important; /* Màu xanh nhạt nhẹ dịu của khung Figma */
            border-radius: 16px !important;
            padding: 40px !important;
            margin-top: 40px;
            text-align: left;
            border: 1px solid #E2E8F0;
            box-shadow: 0px 10px 30px rgba(0, 0, 0, 0.02);
            max-width: 600px;
            margin-left: auto;
            margin-right: auto;
        }
        
        .desktop3-result h4 {
            color: #1E6B7B !important;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 16px;
            margin-top: 20px;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .desktop3-result p {
            color: #475569 !important;
            font-size: 15px;
            line-height: 1.6;
            margin-bottom: 0px;
        }

        /* --- NÚT DOWNLOAD FILE WORD ĐỒNG BỘ MÀU --- */
        div.stDownloadButton {
            text-align: center;
            margin-top: 25px;
        }
        div.stDownloadButton > button {
            background: #1E6B7B !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 12px 35px !important;
            font-size: 15px !important;
            font-weight: 500 !important;
            transition: all 0.2s ease;
            box-shadow: 0px 4px 12px rgba(30, 107, 123, 0.2);
        }
        div.stDownloadButton > button:hover {
            background: #154D59 !important;
            transform: translateY(-1px);
            box-shadow: 0px 6px 18px rgba(30, 107, 123, 0.3);
        }

        /* --- THANH ĐỔ CỰC GRADIENT CHÂN TRANG ĐÚNG BẢN VẼ --- */
        .figma-footer-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            height: 15px;
            background: linear-gradient(90deg, #38BDF8 0%, #1E6B7B 100%);
            z-index: 9999;
        }
    </style>
""", unsafe_allow_html=True)
# 1. Hiển thị Header và Logo LINE BASE góc trái (Màn hình Desktop - 2 / Desktop - 3)
st.markdown("""
    <div class="figma-header">
        <div class="line-base-logo">
            📊 LINE <span>BASE</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# Khởi tạo khung chứa nội dung căn giữa
st.markdown('<div class="figma-body">', unsafe_allow_html=True)

# 2. Tiêu đề QUICK SEARCH chuẩn phong cách tối giản
st.markdown('<div class="quick-search-title">QUICK SEARCH</div>', unsafe_allow_html=True)
# --- 🛠️ KẾT NỐI DỮ LIỆU ---
GOOGLE_SHEET_URL = "https://google.com"
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
    "Healdsburg", "Hemet", "Hercules", "Hermosa Beach", "Hesperia", "Hidden Hills", "Highland", "Hillsborough", 
    "Hollister", "Holtville", "Hughson", "Huntington Beach", "Huntington Park", "Imperial", "Imperial Beach", 
    "Indian Wells", "Indio", "Industry", "Inglewood", "Ione", "Irvine", "Irwindale", "Isleton", "Jackson", 
    "Jurupa Valley", "Kerman", "King City", "Kingsburg", "La Cañada Flintridge", "La Habra", "La Habra Heights", 
    "La Mesa", "La Mirada", "La Palma", "La Puente", "La Quinta", "La Verne", "Lafayette", "Laguna Beach", 
    "Laguna Hills", "Laguna Niguel", "Laguna Woods", "Lake Elsinore", "Lake Forest", "Lakeport", "Lakewood", 
    "Lancaster", "Larkspur", "Lathrop", "Lawndale", "Lemon Grove", "Lemoore", "Lincoln", "Lindsay", 
    "Live Oak", "Livermore", "Livingston", "Lodi", "Loma Linda", "Lomita", "Lompoc", "Long Beach", 
    "Loomis", "Los Alamitos", "Los Altos", "Los Altos Hills", "Los Angeles", "Los Banos", "Los Gatos", 
    "Loyalton", "Lynwood", "Madera", "Malibu", "Mammoth Lakes", "Manhattan Beach", "Manteca", "Maricopa", 
    "Marina", "Martinez", "Marysville", "Maywood", "McFarlin", "Mendota", "Menlo Park", "Merced", 
    "Mill Valley", "Millbrae", "Milpitas", "Mission Viejo", "Modesto", "Monrovia", "Montague", "Montclair", 
    "Monte Sereno", "Montebello", "Monterey", "Monterey Park", "Moorpark", "Moraga", "Moreno Valley", 
    "Morgan Hill", "Morro Bay", "Mount Shasta", "Mountain View", "Murrieta", "Napa", "National City", 
    "Needles", "Nevada City", "Newark", "Newman", "Newport Beach", "Norco", "Norwalk", "Novato", 
    "Oakdale", "Oakland", "Oakley", "Oceanside", "Ojai", "Ontario", "Orange", "Orange Cove", "Orland", 
    "Oroville", "Oxnard", "Pacific Grove", "Pacifica", "Palm Desert", "Palm Springs", "Palmdale", 
    "Palo Alto", "Palos Verdes Estates", "Paradise", "Paramount", "Parlier", "Pasadena", "Paso Robles", 
    "Patterson", "Perris", "Petaluma", "Pico Rivera", "Piedmont", "Pinole", "Pismo Beach", "Pittsburg", 
    "Placentia", "Placerville", "Pleasant Hill", "Pleasanton", "Plymouth", "Point Arena", "Pomona", 
    "Port Hueneme", "Porterville", "Portola", "Portola Valley", "Poway", "Rancho Cordova", "Rancho Cucamonga", 
    "Rancho Mirage", "Red Bluff", "Redding", "Redlands", "Redondo Beach", "Redwood City", "Reedley", 
    "Rialto", "Richmond", "Ridgecrest", "Rio Dell", "Rio Vista", "Ripon", "Riverbank", "Riverside", 
    "Rocklin", "Rohnert Park", "Rolling Hills", "Rolling Hills Estates", "Rosemead", "Roseville", "Ross", 
    "Sacramento", "Salinas", "San Anselmo", "San Bernardino", "San Bruno", "San Carlos", "San Clemente", 
    "San Diego", "San Dimas", "San Fernando", "San Francisco", "San Gabriel", "San Jacinto", "San Joaquin", 
    "San Jose", "San Juan Bautista", "San Juan Capistrano", "San Leandro", "San Luis Obispo", "San Marcos", 
    "San Marino", "San Mateo", "San Pablo", "San Rafael", "San Ramon", "Sand City", "Sanger", "Santa Ana", 
    "Santa Barbara", "Santa Clara", "Santa Clarita", "Santa Cruz", "Santa Fe Springs", "Santa Maria", 
    "Santa Monica", "Santa Paula", "Santa Rosa", "Santee", "Saratoga", "Sausalito", "Scotts Valley", 
    "Seal Beach", "Seaside", "Sebastopol", "Selma", "Shafter", "Shasta Lake", "Sierra Madre", "Signal Hill", 
    "Simi Valley", "Solana Beach", "Soledad", "Solvang", "Sonoma", "Sonora", "South El Monte", "South Gate", 
    "South Lake Tahoe", "South Pasadena", "South San Francisco", "St. Helena", "Stanton", "Stockton", 
    "Suisun City", "Sunnyvale", "Susanville", "Sutter Creek", "Taft", "Tehachapi", "Tehama", "Temecula", 
    "Temple City", "Thousand Oaks", "Tiburon", "Torrance", "Tracy", "Trinidad", "Truckee", "Tulare", 
    "Tulelake", "Turlock", "Tustin", "Twentynine Palms", "Ukiah", "Union City", "Upland", "Vacaville", 
    "Vallejo", "Ventura", "Vernon", "Victorville", "Villa Park", "Visalia", "Vista", "Walnut", "Walnut Creek", 
    "Wasco", "Waterford", "Watsonville", "Weed", "West Covina", "West Hollywood", "West Sacramento", 
    "Westlake Village", "Westminster", "Westmorland", "Wheatland", "Whittier", "Wildomar", "Williams", 
    "Willits", "Willows", "Windsor", "Winters", "Woodlake", "Woodland", "Woodside", "Yorba Linda", 
    "Yountville", "Yreka", "Yuba City", "Yucaipa", "Yucca Valley"
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

def get_city_from_address(address, list_of_cities):
    cleaned_address = str(address).strip().lower()
    for city in list_of_cities:
        if str(city).strip().lower() in cleaned_address: return str(city).strip()
    try:
        geolocator = Nominatim(user_agent="ca_civil_buildbase_ultimate_final_v100")
        location = geolocator.geocode(address + ", CA, USA", addressdetails=True, timeout=10)
        if location and 'address' in location.raw:
            vals = location.raw['address'].values()
            for val in vals:
                for city in list_of_cities:
                    if str(city).strip().lower() == str(val).strip().lower(): return str(city).strip()
    except Exception: pass
    return None
# --- 3. INPUT NHẬN DỮ LIỆU (ỨNG VỚI DESKTOP - 2) ---
user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", key="search_box_final", placeholder="Nhập địa chỉ dự án hoặc tên thành phố...")

# --- 4. HIỂN THỊ TRANG KẾT QUẢ KHI CÓ DỮ LIỆU TÌM KIẾM (ỨNG VỚI DESKTOP - 3) ---
if user_address:
    with st.spinner("Processing..."):
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

            # Dựng lại khung vuông bo góc nhạt mờ chứa nội dung hiển thị của Desktop-3
            st.markdown(f"""
                <div class="desktop3-result">
                    <h3 style="color:#1E6B7B; margin-top:0; font-weight:700; font-family:sans-serif; font-size:20px;">CITY OF {city_name.upper()}</h3>
                    <div style="height:1px; background-color:#CBD5E1; margin: 15px 0 20px 0;"></div>
                    
                    <h4>1. Building Codes</h4>
                    <p>{building_code_val}</p>
                    
                    <h4>2. Civil & Drainage</h4>
                    <p>{drainage_val}</p>
                    <p style="margin-top: 5px;"><b>LID Rules:</b> {lid_val}</p>
                    
                    <h4>3. Local Permit Agency</h4>
                    <p>{permit_agency_val}</p>
                </div>
            """, unsafe_allow_html=True)
            
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
                
                # Nút tải tài liệu Word được bo góc vuông nhỏ lịch lãm đặt l lọt lòng bên dưới
                st.download_button(
                    label="📥 Export Statement of Work (SOW)",
                    data=bio,
                    file_name=f"SOW_{city_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except FileNotFoundError:
                st.error("Không tìm thấy file mẫu 'sow_template.docx' trên hệ thống.")
            except Exception as e:
                st.error(f"Lỗi khởi tạo file Word: {e}")
        else:
            st.error("Không tìm thấy dữ liệu phù hợp với địa chỉ này tại California.")

st.markdown('</div>', unsafe_allow_html=True) # Đóng div figma-body

# 5. Thanh viền màu chuyển sắc dưới đáy màn hình (Gradient Footer Bar) đúng Figma
st.markdown('<div class="figma-footer-bar"></div>', unsafe_allow_html=True)


