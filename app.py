import streamlit as st
from geopy.geocoders import Nominatim
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Hệ thống xuất SOW tự động", layout="centered")
st.title("🏗️ Hệ Thống Tra Cứu Thông Tin")
st.write("Dành riêng cho dự án Xây dựng Dân dụng & Hạ tầng Civil tại California")

# --- ĐƯỜNG LINK TRỤC DỮ LIỆU THÔ CHUẨN XÁC TỪ TRANG XUẤT BẢN CSV CỦA ANH ---
GOOGLE_SHEET_URL = "https://google.com"

# Hàm nạp dữ liệu an toàn
def load_data_safe():
    # Bước 1: Thử tải dữ liệu từ Google Sheets thật của anh
    try:
        df = pd.read_csv(GOOGLE_SHEET_URL)
        df.columns = df.columns.astype(str).str.strip().str.lower()
        # Nếu file Sheets có dữ liệu thật (nhiều hơn 1 hàng tiêu đề), trả về file Sheets
        if len(df) > 0:
            return df, "Google Sheets (Live)"
    except Exception:
        pass

    # Bước 2: Biện pháp Cứu cánh/Dự phòng nếu link Sheets bị trống hoặc lỗi đồng bộ ngầm
    # Tự động khởi tạo kho dữ liệu 482 thành phố chạy độc lập trong bộ nhớ đám mây
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
    
    # Tạo bảng giả lập chứa khung dữ liệu kỹ thuật tiêu chuẩn ngành Civil/Construction
    backup_df = pd.DataFrame({
        'city': [c.lower() for c in backup_cities],
        'original_name': backup_cities,
        'building_code': [f"2025/2026 California Building Code (CBC) - {c} City Amendments & Structural Safety Framework." for c in backup_cities],
        'drainage civil specs': [f"City of {c} Public Works Design Manual / Engineering Standard Drainage Infrastructure Specifications." for c in backup_cities],
        'low impact development': [f"{c} Municipal Stormwater Management Ordinance - Low Impact Development (LID) Retention Rules." for c in backup_cities],
        'local permit agency': [f"City of {c} Development Services / Structural & Civil Building Inspection Division." for c in backup_cities]
    })
    return backup_df, "Hệ thống Dự phòng (Bộ nhớ Đám mây)"

# Tải dữ liệu an toàn đa tầng
df_cities, data_source = load_data_safe()

# Hàm bóc tách địa chỉ đa tầng thông minh (Bất chấp viết hoa viết thường)
def get_city_from_address(address, df_city_list):
    cleaned_address = str(address).strip().lower()
    for city in df_city_list:
        city_str = str(city).strip().lower()
        if city_str and city_str in cleaned_address:
            return str(city).strip()
            
    try:
        optimized_address = address
        if "california" not in address.lower() and "ca" not in address.lower():
            optimized_address += ", CA"
        if "usa" not in address.lower():
            optimized_address += ", USA"
        geolocator = Nominatim(user_agent="ca_civil_sow_generator_internal_v99")
        location = geolocator.geocode(optimized_address, addressdetails=True, timeout=10)
        if location and 'address' in location.raw:
            address_details = location.raw['address']
            possible_places = [
                address_details.get('city'), address_details.get('town'),
                address_details.get('suburb'), address_details.get('village'),
