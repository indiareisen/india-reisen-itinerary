import streamlit as st
import json
from datetime import datetime, timedelta
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib import colors

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
        {"name": "ITC Mughal Agra", "category": "5-star", "description": "Mughal-inspired architecture with heritage gardens", "price_range
