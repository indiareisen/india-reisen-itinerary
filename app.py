import streamlit as st
import json
from datetime import datetime, timedelta
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Page config
st.set_page_config(
    page_title="India Reisen - Itinerary Generator",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with India Reisen branding
st.markdown("""
<style>
    :root {
        --gold: #D4A574;
        --magenta: #d1356f;
    }
    
    .main-header {
        color: var(--magenta);
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    
    .sub-header {
        color: var(--gold);
        font-size: 1.2rem;
        margin-bottom: 2rem;
    }
    
    .card {
        background: linear-gradient(135deg, rgba(209, 53, 111, 0.05) 0%, rgba(212, 165, 116, 0.05) 100%);
        padding: 1.5rem;
        border-left: 4px solid var(--magenta);
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    
    .section-title {
        color: var(--magenta);
        font-size: 1.5rem;
        border-bottom: 2px solid var(--gold);
        padding-bottom: 0.5rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== HOTEL DATABASE ====================
HOTEL_DATABASE = {
    "Delhi": [
        {"name": "The Oberoi Delhi", "category": "5-star", "description": "Luxury palace hotel with world-class spa and fine dining", "price_range": "$$$"},
        {"name": "ITC Maurya New Delhi", "category": "5-star", "description": "Business luxury in prestigious Diplomatic Enclave", "price_range": "$$$"},
        {"name": "Taj Palace New Delhi", "category": "5-star", "description": "Iconic heritage luxury hotel with excellent service", "price_range": "$$$"},
    ],
    "Jaipur": [
        {"name": "Rambagh Palace", "category": "5-star", "description": "Historic royal palace turned luxury hotel with heritage charm", "price_range": "$$$"},
        {"name": "Taj Amar Palace", "category": "5-star", "description": "Elegant property with stunning views of City Palace", "price_range": "$$"},
        {"name": "Alsisar Haveli", "category": "4-star", "description": "Boutique heritage property in historic old city", "price_range": "$$"},
    ],
    "Agra": [
        {"name": "The Oberoi Amar Vilas", "category": "5-star", "description": "Iconic luxury with iconic Taj Mahal views from rooms", "price_range": "$$$"},
        {"name": "ITC Mughal Agra", "category": "5-star", "description": "Mughal-inspired architecture with heritage gardens", "price_range": "$$$"},
        {"name": "Taj Hotel Agra", "category": "4-star", "description": "Heritage property opposite Taj Mahal gate", "price_range": "$$"},
    ],
    "Varanasi": [
        {"name": "Taj Ganges Varanasi", "category": "5-star", "description": "Riverside luxury overlooking the sacred Ganges with ghats access", "price_range": "$$"},
        {"name": "The Brijrama Palace", "category": "5-star", "description": "Boutique palace hotel directly on sacred ghats", "price_range": "$$$"},
        {"name": "Radisson Blu Varanasi", "category": "4-star", "description": "Contemporary comfort with Ganges riverside location", "price_range": "$$"},
    ],
    "Mumbai": [
        {"name": "Taj Palace Mumbai", "category": "5-star", "description": "Iconic beachfront luxury in historic Colaba area", "price_range": "$$$"},
        {"name": "The Oberoi Mumbai", "category": "5-star", "description": "Marine Drive waterfront elegance with ocean views", "price_range": "$$$"},
        {"name": "JW Marriott Mumbai", "category": "5-star", "description": "Business luxury in South Mumbai central location", "price_range": "$$"},
    ],
    "Goa": [
        {"name": "The Oberoi Beach Resort Goa", "category": "5-star", "description": "Beachfront paradise with water sports and activities", "price_range": "$$$"},
        {"name": "Taj Exotica Goa", "category": "5-star", "description": "All-inclusive resort with private beach and entertainment", "price_range": "$$$"},
        {"name": "Leela Goa", "category": "5-star", "description": "Portuguese-influenced luxury resort with heritage vibes", "price_range": "$$"},
    ],
    "Udaipur": [
        {"name": "Taj Lake Palace", "category": "5-star", "description": "Floating marble palace on Lake Pichola - dream destination", "price_range": "$$$"},
        {"name": "The Oberoi Udaivilas", "category": "5-star", "description": "Palatial luxury overlooking Lake Pichola with boats", "price_range": "$$$"},
        {"name": "Leela Palace Udaipur", "category": "5-star", "description": "Modern palace architecture with magnificent lake views", "price_range": "$$$"},
    ],
    "Rishikesh": [
        {"name": "The Oberoi Rishikesh", "category": "5-star", "description": "Spiritual luxury on sacred Ganges banks with yoga programs", "price_range": "$$"},
        {"name": "Ananda in the Himalayas", "category": "5-star", "description": "Wellness retreat with yoga, meditation and ayurvedic spa", "price_range": "$$$"},
        {"name": "Rishikesh Valley Resort", "category": "4-star", "description": "Peaceful ashram-style accommodation with spiritual focus", "price_range": "$"},
    ],
    "Kerala": [
        {"name": "Taj Bekal Resort & Spa", "category": "5-star", "description": "Luxury in Kerala backwaters with houseboat experiences", "price_range": "$$$"},
        {"name": "The Oberoi Zarai Kumarakom", "category": "5-star", "description": "Backwater resort with ayurvedic spa and water activities", "price_range": "$$$"},
        {"name": "Kumarakom Lake Resort", "category": "5-star", "description": "Traditional Kerala resort in lush coconut groves", "price_range": "$$"},
    ],
}

# ==================== PRE-BUILT ITINERARY TEMPLATES ====================
ITINERARY_TEMPLATES = {
    "Delhi-Jaipur-Agra": {
        "title": "Golden Triangle: Delhi, Jaipur & Agra",
        "overview": "Experience the classic Golden Triangle route spanning India's most iconic destinations. Explore the bustling streets of Delhi, the pink city charm of Jaipur, and the timeless beauty of the Taj Mahal in Agra.",
        "best_time": "October to March (pleasant weather, ideal for sightseeing)",
        "budget_estimate": "$1,200-1,800 per person for 7 days",
        "visa_info": "Indian Tourist Visa (60 days) required for most nationalities",
        "best_experiences": [
            "Watch sunset from Taj Mahal (most magical time)",
            "Explore Jaipur's City Palace and Old City bazaars",
            "Visit India Gate and Red Fort in Delhi",
            "Camel safari in Jaipur (Amer Fort approach)",
            "Street food tour in Delhi's Old Delhi"
        ],
        "packing_tips": [
            "Light cotton clothes for hot days",
            "Comfortable walking shoes (you'll walk 10,000+ steps daily)",
            "Sun protection: hat, sunscreen, sunglasses",
            "Lightweight scarf for temple visits",
            "Power bank for phone charging",
            "Adapters for Indian plugs"
        ],
        "days": [
            {"day": 1, "city": "Delhi", "temperature": "25°C (Oct-Mar)", "cultural_notes": "Delhi is the capital of India and a blend of Mughal history and modern India. From ancient monuments to bustling markets, Delhi represents India's journey through centuries.", "activities": ["Arrive and settle at hotel", "Evening stroll through Connaught Place or Khan Market", "Traditional Indian dinner at a local restaurant"], "meals": {"breakfast": "Hotel breakfast or local café", "lunch": "Mughlai cuisine at a heritage restaurant", "dinner": "Street food at Chandni Chowk or restaurant meal"}},
            {"day": 2, "city": "Delhi", "temperature": "25°C", "cultural_notes": "Delhi's historical monuments represent layers of Indian history - from Mughal grandeur to British colonial architecture.", "activities": ["Visit India Gate (British memorial, iconic landmark)", "Explore Red Fort (Mughal palace fortress)", "Walk through Jama Masjid (India's largest mosque)", "Old Delhi bazaar street food tour"], "meals": {"breakfast": "Traditional paratha with chai", "lunch": "Biryani or kebabs in Old Delhi", "dinner": "Fine dining Mughlai or international cuisine"}},
            {"day": 3, "city": "Delhi → Jaipur", "temperature": "26°C", "cultural_notes": "Travel to Jaipur (240 km, 4-5 hours by road or 4.5 hours by train). Jaipur, founded in 1727, is known as the Pink City for its distinctive pink-colored buildings.", "activities": ["Morning departure from Delhi", "Afternoon arrival in Jaipur", "Evening exploration of Jaipur's City Palace surroundings", "Sunset at Nahargarh Fort viewpoint"], "meals": {"breakfast": "Early breakfast before departure", "lunch": "En-route at highway restaurant or Jaipur arrival", "dinner": "Traditional Rajasthani dinner"}},
            {"day": 4, "city": "Jaipur", "temperature": "28°C", "cultural_notes": "Jaipur is the capital of Rajasthan, known for its grand palaces, magnificent forts, and vibrant culture. The city was designed on a grid pattern influenced by Hindu and Mughal architecture.", "activities": ["Visit Amer Fort (hilltop palace with elephant rides)", "Explore Jantar Mantar (astronomical observation site)", "Walk through City Palace complex", "Shopping at local bazaars (textiles, jewelry)"], "meals": {"breakfast": "Hotel breakfast or local café with lassi", "lunch": "Dal baati churma or paneer tikka masala", "dinner": "Gatte ki sabzi or laal maas (Rajasthani cuisine)"}},
            {"day": 5, "city": "Jaipur", "temperature": "28°C", "cultural_notes": "Jaipur's old walled city is a UNESCO World Heritage Site with pink-painted buildings, narrow lanes, and traditional bazaars.", "activities": ["Hawa Mahal (Palace of Winds) photo session", "Old City bazaar exploration and shopping", "Local market visit for traditional crafts", "Cooking class (optional, learn Rajasthani cuisine)"], "meals": {"breakfast": "Kachori with aloo sabzi", "lunch": "Street food: samosas, chaat, momo", "dinner": "Restaurant dinner or cooking class meal"}},
            {"day": 6, "city": "Jaipur → Agra", "temperature": "26°C", "cultural_notes": "Travel to Agra (240 km, 5-6 hours by road or 6 hours by train). Agra was the capital of the Mughal Empire and home to the world-famous Taj Mahal.", "activities": ["Morning departure from Jaipur", "Afternoon arrival in Agra", "Evening light show at Taj Mahal", "Sunset viewing from Mehtab Bagh"], "meals": {"breakfast": "Early breakfast before departure", "lunch": "En-route or Agra arrival meal", "dinner": "Mughlai cuisine dinner in Agra"}},
            {"day": 7, "city": "Agra", "temperature": "26°C", "cultural_notes": "Agra is home to the Taj Mahal, one of the Seven Wonders of the World, built by Mughal Emperor Shah Jahan as a monument to love.", "activities": ["Sunrise at Taj Mahal (ethereal experience)", "Tour Taj Mahal and gardens", "Visit Agra Fort (Mughal fortress)", "Local marble inlay workshop visit", "Shopping at Agra bazaars"], "meals": {"breakfast": "Early breakfast, then Taj Mahal sunrise", "lunch": "Petha (sweet) and traditional Mughlai food", "dinner": "Dinner with riverside views"}},
        ]
    },
    "Agra-Varanasi": {
        "title": "Spiritual Journey: Agra to Varanasi",
        "overview": "Journey from the Mughal grandeur of Agra to the spiritual heart of India in Varanasi. Experience two of India's most significant cultural destinations and the sacred Ganges River.",
        "best_time": "October to March (pleasant weather, spiritually auspicious)",
        "budget_estimate": "$1,500-2,200 per person for 7 days",
        "visa_info": "Indian Tourist Visa (60 days) required",
        "best_experiences": ["Sunrise aarti (prayer ritual) on Ganges ghats", "Boat ride at sunrise on the Ganges", "Varanasi evening aarti ceremony", "Taj Mahal at sunrise", "Buddhist sites visit (Sarnath)"],
        "packing_tips": ["Modest clothing for temples and spiritual sites", "Comfortable water-resistant shoes", "Shawl or scarf for temple visits", "Mosquito repellent for Varanasi", "Motion sickness medication for boat rides", "Respectful attitude towards religious practices"],
        "days": [
            {"day": 1, "city": "Agra", "temperature": "26°C", "cultural_notes": "Begin your spiritual journey in Agra, home to the Taj Mahal, a monument to eternal love built by Mughal Emperor Shah Jahan.", "activities": ["Arrival and hotel check-in", "Sunset viewing at Taj Mahal", "Evening walk in Agra town"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Local Mughlai cuisine", "dinner": "Traditional Indian dinner"}},
            {"day": 2, "city": "Agra", "temperature": "26°C", "cultural_notes": "Explore Agra's historical monuments including the Taj Mahal, Agra Fort, and local bazaars.", "activities": ["Pre-dawn departure for Taj Mahal sunrise", "Guided Taj Mahal tour", "Agra Fort exploration", "Marble inlay workshop visit"], "meals": {"breakfast": "Before Taj visit", "lunch": "Local specialties", "dinner": "Regional cuisine"}},
            {"day": 3, "city": "Agra → Varanasi", "temperature": "27°C", "cultural_notes": "Travel to Varanasi, the spiritual heart of India and one of the world's oldest cities. Flight takes 2 hours or train journey takes 14-16 hours.", "activities": ["Morning departure from Agra", "Flight or train to Varanasi", "Afternoon arrival and hotel check-in", "Evening preparation for night aarti"], "meals": {"breakfast": "Early departure meal", "lunch": "En-route meal", "dinner": "Hotel or restaurant meal"}},
            {"day": 4, "city": "Varanasi", "temperature": "27°C", "cultural_notes": "Varanasi is Hinduism's holiest city, where pilgrims come to bathe in the Ganges River. The city represents the eternal cycle of life, death, and rebirth.", "activities": ["Pre-dawn Ganges boat ride to witness sunrise aarti", "Visit to Kashi Vishwanath Temple (Golden Temple)", "Exploration of ghats (riverbanks)", "Evening Ganga Aarti ceremony"], "meals": {"breakfast": "Hotel breakfast before dawn boat ride", "lunch": "Local vegetarian cuisine", "dinner": "Dinner with Ganges views"}},
            {"day": 5, "city": "Varanasi", "temperature": "27°C", "cultural_notes": "Varanasi is known for its narrow alleyways, temples, and the sacred ritual of Ghat ceremonies that continue for thousands of years.", "activities": ["Sarnath visit (Buddhist pilgrimage site where Buddha gave his first sermon)", "Sarnath Museum exploration", "Chaukhandi Stupa visit", "Return to Varanasi for evening aarti"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Buddhist monastery vegetarian food", "dinner": "Varanasi traditional cuisine"}},
            {"day": 6, "city": "Varanasi", "temperature": "27°C", "cultural_notes": "Experience the daily spiritual life in Varanasi, one of the oldest living cities in the world.", "activities": ["Morning exploration of local bazaars", "Silk weaving workshop visit", "Traditional prayer session at temple", "Cooking class (optional)"], "meals": {"breakfast": "Local street breakfast", "lunch": "Vegetarian thali", "dinner": "Cooking class meal or restaurant"}},
            {"day": 7, "city": "Varanasi", "temperature": "27°C", "cultural_notes": "Your final day in Varanasi - take in the spiritual essence and reflect on your journey through sacred India.", "activities": ["Optional: sunrise boat ride if not done earlier", "Last-minute shopping and temple visits", "Departure preparation", "Evening departure or next day travel"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Favorite meal location revisited", "dinner": "Farewell dinner or travel meal"}},
        ]
    },
    "Mumbai-Goa": {
        "title": "Coastal Paradise: Mumbai to Goa",
        "overview": "Explore the vibrant energy of Mumbai and relax on the pristine beaches of Goa. Experience the contrast between India's bustling metropolis and its laid-back beach culture.",
        "best_time": "October to May (dry season, ideal for beaches)",
        "budget_estimate": "$1,400-2,000 per person for 7 days",
        "visa_info": "Indian Tourist Visa (60 days) required",
        "best_experiences": ["Gateway of India and Art Deco architecture", "Beach hopping in Goa", "Portuguese heritage exploration", "Spice plantation tours", "Water sports and adventure activities"],
        "packing_tips": ["Light summer clothing", "Beach essentials: swimwear, sunscreen, flip-flops", "Light rain jacket (occasional showers)", "Camera for beach photography", "Water shoes for rocky beaches", "Casual evening wear for restaurants"],
        "days": [
            {"day": 1, "city": "Mumbai", "temperature": "28°C", "cultural_notes": "Mumbai is India's financial and entertainment capital, a city of Bollywood dreams and cosmopolitan culture.", "activities": ["Arrival and hotel check-in", "Evening stroll at Marine Drive", "Dinner with sea views"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Local Mumbai cuisine", "dinner": "Seafood or international cuisine"}},
            {"day": 2, "city": "Mumbai", "temperature": "28°C", "cultural_notes": "Mumbai blends modernity with history, showcasing India's colonial heritage and contemporary aspirations.", "activities": ["Gateway of India monument visit", "Taj Palace Hotel exploration", "Elephanta Caves boat trip (UNESCO site)", "Colaba Causeway shopping"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Seafood at Colaba", "dinner": "Fine dining experience"}},
            {"day": 3, "city": "Mumbai → Goa", "temperature": "30°C", "cultural_notes": "Travel to Goa (590 km, 10-12 hours by road, 12-14 hours by train, or 1.5 hours by flight). Goa is India's smallest state with a unique Portuguese heritage.", "activities": ["Morning departure from Mumbai", "Scenic journey along coastal route", "Afternoon arrival in Goa", "Beach relaxation and sunset"], "meals": {"breakfast": "Early departure meal", "lunch": "En-route meal", "dinner": "Goan cuisine dinner"}},
            {"day": 4, "city": "Goa", "temperature": "30°C", "cultural_notes": "Goa was a Portuguese colony until 1961 and retains a unique culture blend - Indian spirituality with European architecture and cuisine.", "activities": ["Baga Beach relaxation and water sports", "Anjuna Beach exploration", "Shopping at Anjuna flea market (if Wednesday)", "Fort Aguada visit (historical fort)"], "meals": {"breakfast": "Beach café breakfast", "lunch": "Fresh seafood by beach", "dinner": "Beachside restaurant"}},
            {"day": 5, "city": "Goa", "temperature": "30°C", "cultural_notes": "Goa's heritage includes Portuguese forts, churches, temples, and a unique fusion cuisine.", "activities": ["Spice plantation tour with lunch", "Old Goa churches and heritage sites", "Basilica da Bom Jesus visit", "Local market exploration"], "meals": {"breakfast": "Hotel breakfast", "lunch": "Spice plantation lunch", "dinner": "Authentic Goan cuisine"}},
            {"day": 6, "city": "Goa", "temperature": "30°C", "cultural_notes": "Goa offers perfect beach relaxation combined with adventure activities and spiritual experiences.", "activities": ["Dolphin spotting boat ride", "Sinquerim Beach exploration", "Paragliding or water sports", "Sunset beach walk"], "meals": {"breakfast": "Beach hut breakfast", "lunch": "Beachfront restaurant", "dinner": "Festive Goan dinner"}},
            {"day": 7, "city": "Goa", "temperature": "30°C", "cultural_notes": "Your final day in paradise - soak in the beach vibes and coastal tranquility.", "activities": ["Last beach visit and swimming", "Shopping for souvenirs and spices", "Relaxation and packing", "Departure preparation"], "meals": {"breakfast": "Favorite café breakfast", "lunch": "Last beach meal", "dinner": "Farewell dinner or travel"}},
        ]
    }
}

# Initialize session state
if 'itinerary' not in st.session_state:
    st.session_state.itinerary = None
if 'generated' not in st.session_state:
    st.session_state.generated = False

# Header
col1, col2 = st.columns([1, 4])
with col1:
    st.markdown("🌸")
with col2:
    st.markdown('<div class="main-header">India Reisen</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Immersive Cultural Journeys Across India</div>', unsafe_allow_html=True)

st.divider()

# Main selection section
st.markdown('<div class="section-title">Select Your Journey</div>', unsafe_allow_html=True)

# Available routes
available_routes = list(ITINERARY_TEMPLATES.keys())
selected_route = st.selectbox(
    "Choose a route",
    available_routes,
    help="Select a pre-designed route"
)

# Customize duration
st.markdown("**Customize Your Journey**")
col1, col2 = st.columns(2)

with col1:
    include_dates = st.checkbox("Include specific travel dates")
    if include_dates:
        start_date = st.date_input("Start Date")
        num_days = st.number_input("Number of Days", 3, 30, 7)
    else:
        start_date = None
        num_days = st.number_input("Number of Days", 3, 30, 7)

with col2:
    st.markdown("**Hotel Preferences**")
    st.info("Hotels will be suggested based on cities in your itinerary")

st.divider()

# Generate button
col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    if st.button("✨ Generate Itinerary", use_container_width=True, type="primary"):
        template = ITINERARY_TEMPLATES[selected_route]
        
        # Adapt template to selected duration
        total_days = len(template["days"])
        if num_days < total_days:
            selected_days = [template["days"][0]] + template["days"][-(num_days-1):]
            for i, day in enumerate(selected_days, 1):
                day["day"] = i
            template["days"] = selected_days
        elif num_days > total_days:
            extra_days_needed = num_days - total_days
            new_days = template["days"].copy()
            for i in range(extra_days_needed):
                new_day = template["days"][(i + 1) % len(template["days"])].copy()
                new_day["day"] = len(new_days) + 1
                new_days.append(new_day)
            template["days"] = new_days
        
        st.session_state.itinerary = template
        st.session_state.generated = True
        st.session_state.selected_route = selected_route
        st.session_state.start_date = start_date
        st.session_state.num_days = num_days
        st.success("✅ Itinerary generated successfully!")

# Display generated itinerary
if st.session_state.generated and st.session_state.itinerary:
    itinerary = st.session_state.itinerary
    
    st.markdown('<div class="section-title">Your Personalized Journey</div>', unsafe_allow_html=True)
    
    # Overview
    st.markdown(f"### {itinerary.get('title', 'India Reisen Journey')}")
    st.markdown(itinerary.get('overview', ''))
    
    # Key info cards
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="card">
            <strong>📅 Best Time</strong><br>
            {itinerary.get('best_time', 'Oct-Mar')}
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="card">
            <strong>💰 Budget</strong><br>
            {itinerary.get('budget_estimate', 'Contact us')}
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="card">
            <strong>🎒 Duration</strong><br>
            {st.session_state.num_days} Days
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="card">
            <strong>📋 Visa</strong><br>
            {itinerary.get('visa_info', 'Check requirements')}
        </div>
        """, unsafe_allow_html=True)
    
    # Day-by-day breakdown
    st.markdown('<div class="section-title">Day-by-Day Itinerary</div>', unsafe_allow_html=True)
    
    for day_info in itinerary.get('days', []):
        with st.expander(f"Day {day_info.get('day')} - {day_info.get('city', 'Travel')} 🌏"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"**📍 Location:** {day_info.get('city', '')}")
                if day_info.get('temperature'):
                    st.markdown(f"**🌡️ Weather:** {day_info.get('temperature')}")
                
                st.markdown("**Cultural Significance:**")
                st.write(day_info.get('cultural_notes', ''))
            
            with col2:
                st.markdown("**🍽️ Suggested Meals:**")
                meals = day_info.get('meals', {})
                for meal_type, suggestion in meals.items():
                    st.write(f"• **{meal_type.title()}:** {suggestion}")
            
            st.markdown("**📋 Activities:**")
            for activity in day_info.get('activities', []):
                st.write(f"• {activity}")
            
            # Show hotel if available
            city = day_info.get('city')
            base_city = city.split('→')[0].strip() if '→' in city else city
            
            if base_city in HOTEL_DATABASE and HOTEL_DATABASE[base_city]:
                suggested_hotel = HOTEL_DATABASE[base_city][0]
                st.markdown(f"""
                **🏨 Suggested Hotel:**
                - **{suggested_hotel['name']}** ({suggested_hotel['category']})
                - {suggested_hotel['description']}
                - Price Range: {suggested_hotel['price_range']}
                """)
    
    # Packing tips
    st.markdown('<div class="section-title">Packing Tips</div>', unsafe_allow_html=True)
    for tip in itinerary.get('packing_tips', []):
        st.write(f"✓ {tip}")
    
    # Best experiences
    if itinerary.get('best_experiences'):
        st.markdown('<div class="section-title">Must-Do Experiences</div>', unsafe_allow_html=True)
        for exp in itinerary.get('best_experiences', []):
            st.write(f"⭐ {exp}")
    
    # Export functions
    def generate_docx_export(itinerary, route_name, start_date, num_days):
        doc = Document()
        
        # Title
        title = doc.add_paragraph()
        title_run = title.add_run("🌸 INDIA REISEN")
        title_run.font.size = Pt(28)
        title_run.font.color.rgb = RGBColor(209, 53, 111)
        title_run.bold = True
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        subtitle = doc.add_paragraph("Immersive Cultural Journey")
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_run = subtitle.runs[0]
        subtitle_run.font.color.rgb = RGBColor(212, 165, 116)
        subtitle_run.font.size = Pt(14)
        
        # Journey Info
        doc.add_paragraph(f"Journey: {route_name}").style = 'Heading 2'
        doc.add_paragraph(f"Duration: {num_days} days")
        if start_date:
            doc.add_paragraph(f"Start Date: {start_date}")
        
        # Overview
        doc.add_heading('Overview', level=2).runs[0].font.color.rgb = RGBColor(209, 53, 111)
        doc.add_paragraph(itinerary.get('overview', ''))
        
        # Day by day
        doc.add_heading('Day-by-Day Itinerary', level=2).runs[0].font.color.rgb = RGBColor(209, 53, 111)
        
        for day in itinerary.get('days', []):
            doc.add_heading(f"Day {day.get('day')} - {day.get('city')}", level=3)
            doc.add_paragraph(f"Temperature: {day.get('temperature', 'N/A')}")
            doc.add_paragraph(day.get('cultural_notes', ''))
            
            doc.add_paragraph("Activities:", style='List Bullet')
            for activity in day.get('activities', []):
                doc.add_paragraph(activity, style='List Bullet 2')
            
            meals = day.get('meals', {})
            if meals:
                doc.add_paragraph("Meals:", style='List Bullet')
                for meal_type, suggestion in meals.items():
                    doc.add_paragraph(f"{meal_type.title()}: {suggestion}", style='List Bullet 2')
            
            city = day.get('city')
            base_city = city.split('→')[0].strip() if '→' in city else city
            if base_city in HOTEL_DATABASE and HOTEL_DATABASE[base_city]:
                hotel = HOTEL_DATABASE[base_city][0]
                doc.add_paragraph(f"Suggested Hotel: {hotel['name']} - {hotel['description']}")
        
        # Packing tips
        doc.add_heading('Packing Tips', level=2).runs[0].font.color.rgb = RGBColor(209, 53, 111)
        for tip in itinerary.get('packing_tips', []):
            doc.add_paragraph(tip, style='List Bullet')
        
        # Best experiences
        if itinerary.get('best_experiences'):
            doc.add_heading('Must-Do Experiences', level=2).runs[0].font.color.rgb = RGBColor(209, 53, 111)
            for exp in itinerary.get('best_experiences', []):
                doc.add_paragraph(exp, style='List Bullet')
        
        # Footer
        doc.add_paragraph()
        footer = doc.add_paragraph("🌸 Book your immersive journey with India Reisen today!")
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    
    def generate_markdown_export(itinerary, route_name, start_date, num_days):
        md = f"""# 🌸 INDIA REISEN
## Immersive Cultural Journey

**Journey:** {route_name}  
**Duration:** {num_days} days  
"""
        if start_date:
            md += f"**Start Date:** {start_date}  \n"
        
        md += f"\n## Overview\n{itinerary.get('overview', '')}\n"
        md += f"\n## Best Time to Visit\n{itinerary.get('best_time', '')}\n"
        md += f"\n## Budget\n{itinerary.get('budget_estimate', '')}\n"
        md += f"\n## Visa Information\n{itinerary.get('visa_info', '')}\n"
        
        md += "\n## Day-by-Day Itinerary\n"
        for day in itinerary.get('days', []):
            md += f"\n### Day {day.get('day')} - {day.get('city')}\n"
            md += f"**Temperature:** {day.get('temperature', 'N/A')}  \n"
            md += f"\n{day.get('cultural_notes', '')}\n\n"
            
            md += "**Activities:**\n"
            for activity in day.get('activities', []):
                md += f"- {activity}\n"
            
            meals = day.get('meals', {})
            if meals:
                md += "\n**Meals:**\n"
                for meal_type, suggestion in meals.items():
                    md += f"- **{meal_type.title()}:** {suggestion}\n"
            
            city = day.get('city')
            base_city = city.split('→')[0].strip() if '→' in city else city
            if base_city in HOTEL_DATABASE and HOTEL_DATABASE[base_city]:
                hotel = HOTEL_DATABASE[base_city][0]
                md += f"\n**Suggested Hotel:** {hotel['name']}  \n{hotel['description']}\n"
        
        if itinerary.get('packing_tips'):
            md += "\n## Packing Tips\n"
            for tip in itinerary.get('packing_tips', []):
                md += f"- {tip}\n"
        
        if itinerary.get('best_experiences'):
            md += "\n## Must-Do Experiences\n"
            for exp in itinerary.get('best_experiences', []):
                md += f"- {exp}\n"
        
        md += "\n---\n*🌸 Book your immersive journey with India Reisen today!*\n"
        return md
    
    # Export options
    st.divider()
    st.markdown('<div class="section-title">Export Your Itinerary</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📝 Generate DOCX", use_container_width=True):
            try:
                docx_buffer = generate_docx_export(itinerary, st.session_state.selected_route, st.session_state.start_date, st.session_state.num_days)
                st.download_button(
                    label="📥 Download DOCX",
                    data=docx_buffer,
                    file_name=f"India_Reisen_Itinerary_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error: {str(e)}")
    
    with col2:
        if st.button("📋 Generate Markdown", use_container_width=True):
            md_content = generate_markdown_export(itinerary, st.session_state.selected_route, st.session_state.start_date, st.session_state.num_days)
            st.download_button(
                label="📥 Download Markdown",
                data=md_content,
                file_name=f"India_Reisen_Itinerary_{datetime.now().strftime('%Y%m%d')}.md",
                mime="text/markdown",
                use_container_width=True
            )

# Footer
st.divider()
st.markdown("""
---
**India Reisen** | Immersive Cultural Experiences  
🌐 Destinations: India • Nepal • Bhutan • Tibet • Sri Lanka  
📱 Follow us: Instagram • Facebook • LinkedIn • YouTube

*Created with ✨ for travelers seeking authentic cultural immersion*
""")
