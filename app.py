import streamlit as st
import streamlit.components.v1 as components
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

# Cấu hình trang tối giản, ẩn các thành phần thừa của Streamlit để nhường chỗ cho Figma
st.set_page_config(page_title="BUILDBASE - FIGMA 100%", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
        /* Ẩn thanh header và footer mặc định của Streamlit */
        header, footer, .stDeployButton, div[data-testid="stToolbar"] {display: none !important;}
        .stApp { background-color: #FFFFFF !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 📋 DÁN ĐOẠN CODE HTML/CSS BẠN CỦA FIGMA VÀO ĐÂY
# ==========================================
figma_html = <div style="width: 1440px; height: 1024px; position: relative; background: white; overflow: hidden">
  <div style="width: 585.73px; height: 70px; left: 411.46px; top: 344.52px; position: absolute; background: white; border-radius: 70px; border: 2px #595656 solid"></div>
  <div style="width: 40px; height: 40px; left: 921.06px; top: 360.41px; position: absolute; background: #ADADAD"></div>
  <div style="width: 1441.82px; height: 68.41px; left: 0px; top: 0px; position: absolute; background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%)"></div>
  <div style="width: 1441.82px; height: 35.16px; left: 1.53px; top: 988.84px; position: absolute; background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%)"></div>
  <div style="width: 126.31px; height: 25px; left: 77.33px; top: 22.95px; position: absolute; background: white"></div>
  <img style="width: 262px; height: 130px; left: 65.46px; top: 88.93px; position: absolute" src="https://placehold.co/262x130" />
  <div style="width: 707.60px; height: 118.70px; left: 407.07px; top: 229.92px; position: absolute; color: #89EDF3; font-size: 80px; font-family: Onest; font-weight: 400; line-height: 120px; word-wrap: break-word; text-shadow: 0px 4px 4px rgba(0, 0, 0, 0.25)">QUICK SEARCH</div>
</div>
<div style="width: 25px; height: 25px; position: relative; background: rgba(0, 0, 0, 0)"></div>
<div style="width: 1440px; height: 1024px; position: relative; background: white; overflow: hidden">
  <div style="width: 585.73px; height: 70px; left: 411.46px; top: 344.52px; position: absolute; background: white; border-radius: 70px; border: 2px #595656 solid"></div>
  <div style="width: 40px; height: 40px; left: 921.06px; top: 360.41px; position: absolute; background: #ADADAD"></div>
  <div style="width: 1441.82px; height: 68.41px; left: 0px; top: 0px; position: absolute; background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%)"></div>
  <div style="width: 1441.82px; height: 35.16px; left: 1.53px; top: 988.84px; position: absolute; background: linear-gradient(90deg, #72DDE4 24%, #23749F 67%)"></div>
  <div style="width: 126.31px; height: 25px; left: 77.33px; top: 22.95px; position: absolute; background: white"></div>
  <img style="width: 262px; height: 130px; left: 65.46px; top: 88.93px; position: absolute" src="https://placehold.co/262x130" />
  <div style="width: 707.60px; height: 118.70px; left: 407.07px; top: 229.92px; position: absolute; color: #89EDF3; font-size: 80px; font-family: Onest; font-weight: 400; line-height: 120px; word-wrap: break-word; text-shadow: 0px 4px 4px rgba(0, 0, 0, 0.25)">QUICK SEARCH</div>
  <div style="width: 643.99px; height: 474.73px; left: 396.07px; top: 447.81px; position: absolute; background: #F3FFFE; border-radius: 30px"></div>
</div>
<!-- Thay thế toàn bộ phần này bằng code HTML/CSS bạn copy từ Plugin Figma ra -->
<div style="width: 100%; display: flex; flex-direction: column; align-items: center; font-family: sans-serif;">
    <div style="width: 100%; padding: 20px 40px; display: flex; justify-content: flex-start;">
        <span style="font-weight: bold; font-size: 20px; color: #0F172A;">📊 LINE <span style="color: #38BDF8;">BASE</span></span>
    </div>
    <div style="margin-top: 60px; font-size: 42px; font-weight: 700; color: #1E6B7B; letter-spacing: -0.02em;">QUICK SEARCH</div>
</div>
"""
# Hiển thị giao diện chuẩn Figma lên Streamlit
components.html(figma_html, height=250, scrolling=False)


# ==========================================
# 🛠️ PHẦN XỬ LÝ SEARCH BOX & FILE WORD CỦA STREAMLIT
# ==========================================
# Giữ lại ô nhập liệu Streamlit nằm ngay dưới tiêu đề Figma để xử lý logic Python
user_address = st.text_input("Tìm kiếm...", label_visibility="collapsed", placeholder="Nhập địa chỉ dự án hoặc tên thành phố...")

# (Giữ nguyên phần xử lý dữ liệu và xuất file Word giống như các đoạn code trước của bạn ở đây...)
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


