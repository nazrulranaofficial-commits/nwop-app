import streamlit as st
import pandas as pd
import re
import time
import hashlib
import json
import os
import pytz
import uuid
from io import BytesIO
from datetime import datetime, date, time as dt_time
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# --- SELENIUM IMPORTS FOR LIVE SCRAPING ---
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False

# --- BANGLADESH TIMEZONE SETUP ---
BD_TZ = pytz.timezone('Asia/Dhaka')

# --- CONFIG & CUSTOM CSS (ULTRA-MODERN MOBILE UI) ---
st.set_page_config(
    page_title="NWOP - Nazrul's Order Parser", 
    page_icon="📦", 
    layout="wide", 
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "NWOP Enterprise Edition - Developed by Nazrul Rana"
    }
)

st.markdown("""
    <style>
    /* Smooth Scrolling for Jump Links */
    html { scroll-behavior: smooth; }

    /* Full App Background */
    .stApp { background-color: var(--secondary-background-color) !important; }
    
    /* Main Content Padding */
    .block-container { padding-top: 2rem !important; padding-bottom: 3rem !important; max-width: 1000px !important; }

    /* HIDE GITHUB ICON BUT KEEP MENU */
    [data-testid="stToolbar"] a { display: none !important; } 
    footer { display: none !important; }

    /* Floating Cards for Expanders */
    [data-testid="stExpander"] {
        background-color: var(--background-color) !important;
        border-radius: 16px !important;
        border: 1px solid rgba(128,128,128,0.08) !important;
        box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.04) !important;
        margin-bottom: 12px !important;
        overflow: hidden !important;
        transition: all 0.3s ease;
    }
    [data-testid="stExpander"]:hover {
        box-shadow: 0px 8px 25px rgba(0, 0, 0, 0.08) !important;
    }
    [data-testid="stExpander"] > details > summary {
        padding: 16px !important; font-weight: 700 !important; font-size: 1.05rem !important;
    }
    
    /* Metrics / Status Cards */
    [data-testid="stMetric"] {
        background-color: var(--background-color) !important;
        border-radius: 16px !important;
        padding: 15px 10px !important;
        box-shadow: 0px 4px 15px rgba(0, 0, 0, 0.04) !important;
        text-align: center !important;
        border: 1px solid rgba(128,128,128,0.08) !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 2rem !important; font-weight: 900 !important; color: #10B981 !important; 
    }
    [data-testid="stMetricLabel"] { font-size: 0.95rem !important; font-weight: 700 !important; opacity: 0.7 !important; }

    /* Modern Buttons */
    .stButton > button {
        border-radius: 20px !important; border: none !important; padding: 10px 20px !important;
        font-weight: 800 !important; letter-spacing: 0.3px !important; width: 100% !important;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.08) !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover { transform: translateY(-2px) !important; box-shadow: 0px 6px 15px rgba(0, 0, 0, 0.12) !important; }
    .stButton > button[kind="primary"] { background: linear-gradient(135deg, #10B981, #059669) !important; color: white !important; }

    /* Fix Button inside Doubtful Card */
    .fix-btn {
        background: linear-gradient(135deg, #F59E0B, #D97706);
        color: white !important;
        padding: 6px 16px;
        border-radius: 20px;
        text-decoration: none;
        font-weight: 700;
        font-size: 0.85rem;
        box-shadow: 0 2px 5px rgba(245, 158, 11, 0.3);
        transition: 0.3s;
        display: inline-block;
        white-space: nowrap;
    }
    .fix-btn:hover {
        box-shadow: 0 4px 10px rgba(245, 158, 11, 0.4);
        transform: translateY(-2px);
    }
    
    /* Doubtful Alert Card */
    .doubt-card {
        padding: 12px 15px; 
        border-left: 5px solid #F59E0B; 
        background-color: var(--background-color); 
        border-radius: 8px; 
        margin-bottom: 10px; 
        display: flex; 
        justify-content: space-between; 
        align-items: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        border: 1px solid rgba(128,128,128,0.1);
    }

    /* Pill shaped Tabs */
    .stTabs [data-baseweb="tab-list"] { background-color: transparent !important; gap: 8px !important; padding-bottom: 10px !important; }
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(128, 128, 128, 0.08) !important; border-radius: 25px !important;
        padding: 8px 22px !important; border: none !important; font-weight: 700 !important; font-size: 0.95rem !important;
    }
    .stTabs [aria-selected="true"] { background-color: #10B981 !important; color: #FFFFFF !important; box-shadow: 0px 4px 10px rgba(16, 185, 129, 0.3) !important; }

    /* Input Fields */
    .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] {
        border-radius: 12px !important; border: 1px solid rgba(128, 128, 128, 0.15) !important;
        padding: 10px !important; background-color: var(--background-color) !important;
    }

    /* Custom Header Typography */
    .main-header-title { margin-top: 10px; color: var(--text-color); font-weight: 900; font-size: 2.2rem; }
    .welcome-text { color: gray; font-size: 1.05rem; margin-top: -10px; margin-bottom: 25px; }

    /* Login Premium Card */
    .login-card {
        background-color: var(--background-color);
        padding: 40px;
        border-radius: 25px;
        box-shadow: 0px 10px 30px rgba(0,0,0,0.06);
        text-align: center;
        border: 1px solid rgba(128,128,128,0.08);
        margin-top: 20px;
    }

    /* Raw Message Box */
    .raw-msg-box {
        background-color: rgba(16, 185, 129, 0.05);
        padding: 15px;
        border-radius: 12px;
        border: 1px dashed rgba(16, 185, 129, 0.3);
        height: 100%;
        font-size: 0.95rem;
        line-height: 1.5;
    }

    @media (max-width: 768px) {
        .main-header-title { font-size: 1.8rem; margin-top: 5px; text-align: center; }
        .welcome-text { text-align: center; }
        .stTabs [data-baseweb="tab-list"] { overflow-x: auto; white-space: nowrap; flex-wrap: nowrap; }
        .login-card { padding: 25px; }
        .doubt-card { flex-direction: column; align-items: flex-start; gap: 10px; }
        .fix-btn { width: 100%; text-align: center; }
    }
    </style>
""", unsafe_allow_html=True)

# --- PERSISTENT HISTORY & CHECKPOINT SYSTEM ---
HISTORY_FILE = "nwop_history.json"

def load_data():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            try: 
                data = json.load(f)
                if isinstance(data, list): return data, "No record yet" 
                return data.get("history", []), data.get("last_checkpoint", "No record yet")
            except: return [], "No record yet"
    return [], "No record yet"

def save_data(history_list, checkpoint):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump({"history": history_list, "last_checkpoint": checkpoint}, f, ensure_ascii=False, indent=4)

if 'task_history' not in st.session_state or 'last_checkpoint' not in st.session_state:
    hist, chk = load_data()
    st.session_state.task_history = hist
    st.session_state.last_checkpoint = chk

def log_task(task_desc):
    timestamp = datetime.now(BD_TZ).strftime("%d %b %Y, %I:%M %p")
    st.session_state.task_history.insert(0, f"✅ **{timestamp}**: {task_desc}")
    save_data(st.session_state.task_history, st.session_state.last_checkpoint)

# --- LOGIN SYSTEM ---
CORRECT_PASSWORD = "nwop" 

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    _, col_center, _ = st.columns([1, 10, 1])
    with col_center:
        st.markdown("<div class='login-card'>", unsafe_allow_html=True)
        
        if os.path.exists("logo.png"):
            st.image("logo.png", width=180)
        else:
            st.markdown("<h1 style='font-size: 70px; margin-bottom: 0;'>📦</h1>", unsafe_allow_html=True)
            
        st.markdown("<h2 style='color: var(--text-color); font-weight:900; margin-top: 15px;'>NWOP Access</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: gray; font-size: 15px; margin-top:-5px;'>By <b>Nazrul Rana</b> | WhatsApp: +880164143400</p>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        password_input = st.text_input("Master Password", type="password", label_visibility="collapsed", placeholder="Enter Master Password...")
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Unlock Dashboard", type="primary"):
            if password_input == CORRECT_PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("❌ Incorrect Password! Try again.")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# --- SESSION STATE (MEMORY) ---
if 'all_orders' not in st.session_state: st.session_state.all_orders = []
if 'ignored_messages' not in st.session_state: st.session_state.ignored_messages = []
if 'total_scanned' not in st.session_state: st.session_state.total_scanned = 0
if 'sheet_date' not in st.session_state: st.session_state.sheet_date = datetime.now(BD_TZ).strftime("%d/%m/%y")
if 'total_extracted_today' not in st.session_state: st.session_state.total_extracted_today = 0

def bn_to_en_digits(text):
    return text.translate(str.maketrans('০১২৩৪৫৬৭৮৯', '0123456789'))

PHONE_PATTERN = r'((?:\+88|88)?0\s*1\s*[3-9](?:[\s-]*\d){8})'

def check_message_status(text_en):
    has_phone = re.search(PHONE_PATTERN, text_en)
    return "valid" if has_phone else "ignored"

def get_datetime_obj(date_string, time_string):
    try:
        if ':' in date_string and '/' in time_string:
            date_string, time_string = time_string, date_string

        d_clean = date_string.strip()
        if len(d_clean.split('/')[-1]) == 2:
            try: d = datetime.strptime(d_clean, "%d/%m/%y").date()
            except: d = datetime.strptime(d_clean, "%m/%d/%y").date()
        else:
            try: d = datetime.strptime(d_clean, "%d/%m/%Y").date()
            except: d = datetime.strptime(d_clean, "%m/%d/%Y").date()
            
        t_clean = time_string.replace('\u202f', ' ').replace('\u200e', '').replace('\u200f', '').strip()
        if 'AM' not in t_clean.upper() and 'PM' not in t_clean.upper():
            if t_clean.count(':') == 2: t = datetime.strptime(t_clean, "%H:%M:%S").time()
            else: t = datetime.strptime(t_clean, "%H:%M").time()
        else:
            if t_clean.count(':') == 2: t = datetime.strptime(t_clean, "%I:%M:%S %p").time()
            else: t = datetime.strptime(t_clean, "%I:%M %p").time()
        return datetime.combine(d, t)
    except: return None

def parse_copy_paste_time(pasted_str):
    if not pasted_str: return None
    clean = pasted_str.replace('[', '').replace(']', '').replace('\u200e', '').replace('\u200f', '').strip()
    if ',' in clean:
        parts = clean.split(',', 1)
        return get_datetime_obj(parts[0].strip(), parts[1].strip())
    return None

# 🌟 SMART EXTRACTION ENGINE 🌟
def extract_order_details(msg_dict):
    text = msg_dict["text"]
    raw_text = text  
    parts = re.split(r'^\[.*?\] .*?:\s', text, maxsplit=1)
    body = parts[1] if len(parts) > 1 else text
    body_en = bn_to_en_digits(body)

    status = check_message_status(body_en)
    if status == "ignored":
        return {"status": "ignored", "Date": msg_dict["date_str"], "Time": msg_dict["time_str"], "Text": text}

    body_en = body_en.replace('<This message was edited>', ' ')
    body_en = re.sub(r'অর্ডার\s*করতে\s*[-ঃ:]*\s*', ' ', body_en, flags=re.IGNORECASE)

    phone_match = re.search(PHONE_PATTERN, body_en)
    phone = "N/A"
    if phone_match:
        raw_phone = phone_match.group(1)
        body_en = body_en.replace(raw_phone, ' ') 
        clean_phone = re.sub(r'\D', '', raw_phone)
        if clean_phone.startswith('88') and len(clean_phone) > 11: clean_phone = clean_phone[2:]
        phone = clean_phone

    price = 0
    eq_match = re.search(r'\d+\s*\+\s*\d+\s*=\s*(\d+)', body_en)
    if eq_match:
        price = int(eq_match.group(1))
        body_en = body_en.replace(eq_match.group(0), ' ') 
    else:
        price_match = re.search(r'(\d+)\s*(?:টাকা|taka|tk|/-)', body_en, re.IGNORECASE)
        if price_match:
            price = int(price_match.group(1))
            body_en = body_en.replace(price_match.group(0), ' ') 
        else:
            price_match = re.search(r'=\s*(\d+)', body_en)
            if price_match:
                price = int(price_match.group(1))
                body_en = body_en.replace(price_match.group(0), ' ') 

    qty_match = re.search(r'(\d+)\s*(?:pcs?|pis|পিস|piece|টি|টা|ta(?!\w))', body_en, re.IGNORECASE)
    quantity = int(qty_match.group(1)) if qty_match else 1
    if qty_match: body_en = body_en.replace(qty_match.group(0), ' ')

    # 🌟 ALWAYS DEFAULT TO ELECTRIC BLENDER IF GRINDER/BLENDER IS FOUND 🌟
    product = "Electric Blender"
    if not re.search(r'grind|grainder|blender', body_en, re.IGNORECASE):
        pass 

    body_en = re.sub(r'electrc|electric|electronic|blenders?|grinders?|grainders?|food\s*grind|taka|tk|টাকা', ' ', body_en, flags=re.IGNORECASE)
    body_en = body_en.replace('/-', ' ')
    body_en = re.sub(r'image omitted|<media omitted>|media omitted', ' ', body_en, flags=re.IGNORECASE)

    clean_body = body_en.replace('\n', ',').replace('=', ',').replace('।।', ',').replace('।', ',').replace('|', ',')
    clean_body = re.sub(r'(?i)(?<![a-zA-Z])dist?\.', 'জেলা ', clean_body)
    
    major_labels = ['নাম', 'name', 'nam', 'ফুল ঠিকানা', 'ঠিকানা', 'ঠীকানা', 'thikana', 'address', 'add', 'এড্রেস']
    for kw in major_labels:
        clean_body = re.sub(rf'(?<![a-zA-Z0-9\u0980-\u09FF,])({kw})', r',\1', clean_body, flags=re.IGNORECASE)
        
    address_indicators = ['থানা', 'জেলা', 'গ্রাম', 'পোস্ট', 'বাজার', 'রোড', 'সদর', 'উপজেলা', 'মোড়', 'para', 'pur', 'gram', 'thana', 'bazar', 'road', 'zilla', 'district', 'upazila', 'বিভাগ', 'ওয়ার্ড', 'ঢাকা','চট্টগ্রাম','রাজশাহী','খুলনা','বরিশাল','সিলেট','রংপুর','ময়মনসিংহ','কুমিল্লা','নোয়াখালী','ফেনী','চাঁদপুর','ব্রাহ্মণবাড়িয়া','গাজীপুর','টাঙ্গাইল','নারায়ণগঞ্জ','নরসিংদী','ফরিদপুর','মাদারীপুর','শরীয়তপুর','গোপালগঞ্জ','কিশোরগঞ্জ','সুনামগঞ্জ','হবিগঞ্জ','মৌলভীবাজার','রাঙ্গামাটি','বান্দরবান','খাগড়াছড়ি','কক্সবাজার','লক্ষ্মীপুর','ভোলা','পটুয়াখালী','বরগুনা','ঝালকাঠি','পিরোজপুর','যশোর','সাতক্ষীরা','ঝিনাইদহ','মাগুরা','নড়াইল','বাগেরহাট','কুষ্টিয়া','কুষ্টিয়া','চুয়াডাঙ্গা','মেহেরপুর','পাবনা','সিরাজগঞ্জ','বগুড়া','জয়পুরহাট','নওগাঁ','নাটোর','চাঁপাইনবাবগঞ্জ','দিনাজপুর','ঠাকুরগাঁও','পঞ্চগড়','নীলফামারী','কুড়িগ্রাম','লালমনিরহাট','গাইবান্ধা','জামালপুর','শেরপুর','নেত্রকোণা']
    for kw in address_indicators:
        clean_body = re.sub(rf'(?<![a-zA-Z0-9\u0980-\u09FF,])({kw})', r',\1', clean_body, flags=re.IGNORECASE)

    clean_body = re.sub(r'(?:মোবাইল\s*নাম্বার|মোবাইল|ফোন\s*নাম্বার|ফোন|নাম্বার|mobile\s*number|mobile|phone\s*number|phone|number)[\sঃ:=-]*', '', clean_body, flags=re.IGNORECASE)
    
    raw_chunks = [c.strip() for c in clean_body.split(',') if c.strip()]
    name, address_lines = "N/A", []
    explicit_name_found = False

    for chunk in raw_chunks:
        if not re.search(r'[a-zA-Zঅ-য়0-9]', chunk): continue
        cleaned_chunk = re.sub(r'^[+0-9\s-]+$', '', chunk).strip()
        if not cleaned_chunk: continue

        if re.match(r'^(?:নাম|name|nam)\s*[:ঃ=-]+\s*|^(?:নাম|name|nam)\s+', cleaned_chunk, re.IGNORECASE):
            name_val = re.sub(r'^(?:নাম|name|nam)\s*[:ঃ=-]*\s*', '', cleaned_chunk, flags=re.IGNORECASE).strip()
            if name_val:
                if name != "N/A" and not explicit_name_found:
                    address_lines.insert(0, name) 
                name = name_val
                explicit_name_found = True
            continue
            
        if re.match(r'^(?:ফুল\s*ঠিকানা|ঠিকানা|ঠীকানা|thikana|address|add|এড্রেস)\s*[:ঃ=-]+\s*|^(?:ফুল\s*ঠিকানা|ঠিকানা|ঠীকানা|thikana|address|add|এড্রেস)\s+', cleaned_chunk, re.IGNORECASE):
            addr_val = re.sub(r'^(?:ফুল\s*ঠিকানা|ঠিকানা|ঠীকানা|thikana|address|add|এড্রেস)\s*[:ঃ=-]*\s*', '', cleaned_chunk, flags=re.IGNORECASE).strip()
            if addr_val: address_lines.append(addr_val)
            continue
            
        cleaned_chunk = re.sub(r'^(?:নাম|name|nam|ফুল\s*ঠিকানা|ঠিকানা|ঠীকানা|thikana|address|add|এড্রেস)\s*[:ঃ=-]*\s*', '', cleaned_chunk, flags=re.IGNORECASE).strip()
        if not cleaned_chunk: continue

        if name == "N/A": name = cleaned_chunk
        else: address_lines.append(cleaned_chunk)

    if name == "N/A" and address_lines: name = address_lines.pop(0)

    addr_hints = ['বাড়ি', 'বাড়ি', 'বাড়ী', 'বাড়ী', 'তলা', 'রোড', 'road', 'house', 'হাউজ', 'ফ্ল্যাট', 'flat', 'গ্রাম', 'থানা', 'জেলা', 'উপজেলা', 'মার্কেট', 'বটতলা', 'বাজার', 'কলেজ', 'গেট', 'gate', 'মোড়', 'mor', 'স্ট্যান্ড', 'stand', 'পাড়া', 'পাড়া', 'para', 'pur', 'পুর', 'নগর', 'nagar', 'ভবন', 'bhaban', 'building', 'tower', 'টাওয়ার', 'এলাকা', 'এভেনিউ', 'avenue', 'ব্লক', 'block', 'সেকশন', 'section', 'লেন', 'lane']
    
    if name != "N/A" and not explicit_name_found and len(address_lines) > 0:
        name_is_addr = any(hint.lower() in name.lower() for hint in addr_hints)
        if name_is_addr:
            for i in range(len(address_lines)-1, -1, -1):
                candidate = address_lines[i]
                if len(candidate.split()) <= 3 and not any(hint.lower() in candidate.lower() for hint in addr_hints):
                    real_name = candidate
                    address_lines.pop(i)
                    address_lines.insert(0, name)
                    name = real_name
                    break

    address = ", ".join(address_lines) if address_lines else "N/A"
    address = re.sub(r',+', ',', address) 
    address = re.sub(r'\s*,\s*', ', ', address) 
    address = address.strip(' ,-:;') 

    expander_title = f"Order: {name} | ৳{price} | 📞 {phone} | 🕒 {msg_dict['time_str']}"

    return {
        "id": str(uuid.uuid4()),
        "status": "valid", "Date": msg_dict["date_str"], "Time": msg_dict["time_str"],
        "Name": name, "Phone Number": phone, "Address": address, "Product": product,
        "Quantity": quantity, "Price": price, "Approval": "Pending", "Note": "", "is_duplicate": False,
        "RawText": raw_text,
        "Expander_Title": expander_title
    }

# --- APP LAYOUT HEADER ---
st.markdown("<br>", unsafe_allow_html=True)
col_logo, col_title, col_logout = st.columns([1.5, 6, 2])
with col_logo:
    if os.path.exists("logo.png"): st.image("logo.png", width=110)
with col_title: 
    st.markdown("<h2 class='main-header-title'>NWOP Dashboard</h2>", unsafe_allow_html=True)
    st.markdown("<div class='welcome-text'>Welcome back, Nazrul! Here's your overview.</div>", unsafe_allow_html=True)
with col_logout: 
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Logout", type="secondary"):
        st.session_state.logged_in = False
        st.rerun()

# --- TABS ---
tab_workspace, tab_merge, tab_history, tab_settings, tab_about = st.tabs(["🚀 Workspace", "🗂️ Merge", "📜 History", "⚙️ Settings", "ℹ️ About"])

with tab_workspace:
    st.sidebar.header("🛠️ Working Mode")
    mode = st.sidebar.radio("Select Input Mode:", ["Upload Chat History", "Live Scraping (Beta)"])
    
    st.sidebar.markdown("---")
    st.sidebar.success(f"⏱️ **Last Extraction Checkpoint:**\n\n`{st.session_state.last_checkpoint}`\n\n*(Copy & use this as your next Start Time)*")

    if mode == "Upload Chat History":
        st.sidebar.header("📅 Extraction Filters")
        filter_type = st.sidebar.radio("Extract Data By:", ["All Time", "Specific Date", "Time Range (Copy-Paste)"])
        target_date_str, start_str, end_str = "", "", ""

        if filter_type == "Specific Date":
            st.sidebar.caption("WhatsApp-এ ডেট যেভাবে আছে ঠিক সেভাবেই লিখুন:")
            target_date_str = st.sidebar.text_input("Enter Exact Date (e.g. 3/4/26):", datetime.now(BD_TZ).strftime("%-m/%-d/%y"))
        elif filter_type == "Time Range (Copy-Paste)":
            st.sidebar.caption("WhatsApp থেকে ব্র্যাকেট সহ টাইম কপি করে দিন:")
            start_str = st.sidebar.text_input("Start Time:", st.session_state.last_checkpoint if st.session_state.last_checkpoint != "No record yet" else "[3/4/26, 9:21:30 PM]")
            end_str = st.sidebar.text_input("End Time:", "[3/4/26, 10:08:27 PM]")

        uploaded_file = st.file_uploader("📂 Upload WhatsApp Chat (.txt)", type="txt")

        if uploaded_file:
            if st.button("▶️ Start Extraction", type="primary", use_container_width=True):
                with st.spinner("Processing file..."):
                    try:
                        content = uploaded_file.read().decode("utf-8").replace('\u200e', '').replace('\u200f', '')
                        lines = content.split('\n')
                        
                        messages, current_msg = [], None
                        for line in lines:
                            match = re.match(r'^\[(.*?),\s(.*?)\]\s.*?:', line)
                            if match:
                                if current_msg: messages.append(current_msg)
                                date_str, time_str = match.group(1), match.group(2)
                                msg_dt = get_datetime_obj(date_str, time_str)
                                dt_obj = msg_dt.date() if msg_dt else datetime.now(BD_TZ).date()
                                current_msg = {"date_obj": dt_obj, "date_str": date_str, "time_str": time_str, "msg_dt": msg_dt, "text": line}
                            else:
                                if current_msg: current_msg["text"] += "\n" + line
                        if current_msg: messages.append(current_msg)
                            
                        filtered_messages = []
                        start_dt = parse_copy_paste_time(start_str) if filter_type == "Time Range (Copy-Paste)" else None
                        end_dt = parse_copy_paste_time(end_str) if filter_type == "Time Range (Copy-Paste)" else None
                        
                        for msg in messages:
                            if filter_type == "All Time":
                                filtered_messages.append(msg)
                            elif filter_type == "Specific Date" and target_date_str:
                                if msg["date_str"] == target_date_str.strip(): filtered_messages.append(msg)
                            elif filter_type == "Time Range (Copy-Paste)" and start_dt and end_dt and msg["msg_dt"]:
                                if start_dt <= msg["msg_dt"] <= end_dt: filtered_messages.append(msg)
                        
                        st.session_state.total_scanned = len(filtered_messages)
                        temp_orders, temp_ignored = [], []
                        phone_counts = {}
                        
                        for msg in filtered_messages:
                            data = extract_order_details(msg)
                            if data:
                                if data["status"] == "valid":
                                    del data["status"]
                                    ph = data['Phone Number']
                                    if ph != "N/A":
                                        phone_counts[ph] = phone_counts.get(ph, 0) + 1
                                        if phone_counts[ph] > 1: data["is_duplicate"] = True
                                        
                                    temp_orders.append(data)
                                elif data["status"] == "ignored":
                                    temp_ignored.append(data)
                        
                        st.session_state.ignored_messages = temp_ignored
                        if temp_orders:
                            st.session_state.all_orders = temp_orders 
                            st.session_state.sheet_date = "Time_Range_Export" if filter_type == "Time Range (Copy-Paste)" else target_date_str.replace('/', '-') if filter_type == "Specific Date" else f"Bulk_{datetime.now(BD_TZ).strftime('%d-%m-%y')}"
                            st.session_state.total_extracted_today += len(temp_orders)
                            
                            last_order = temp_orders[-1]
                            st.session_state.last_checkpoint = f"[{last_order['Date']}, {last_order['Time']}]"
                            
                            hist_src = f"Time Range ({start_str} to {end_str})" if filter_type == "Time Range (Copy-Paste)" else f"Date ({target_date_str})" if filter_type == "Specific Date" else "All Time"
                            log_task(f"Extracted {len(temp_orders)} orders via Text Upload. Source: {hist_src}. <br>📌 **Stopped at Checkpoint:** `{st.session_state.last_checkpoint}`")
                            
                            st.success(f"Success! Found {len(temp_orders)} valid orders.")
                        else:
                            st.session_state.all_orders = []
                            st.error("No valid orders found with these filters.")
                    except Exception as e:
                        st.error(f"Error reading file: {e}")

    elif mode == "Live Scraping (Beta)":
        st.header("🔴 Auto-Scroll Live Scraper")
        
        if not SELENIUM_AVAILABLE:
            st.error("⚠️ Selenium is not installed! Please install via terminal: `pip install selenium webdriver-manager`")
        else:
            target_group = st.text_input("🎯 Live Target Group Name:", "ORDER COLLECTION")
            start_time_str = st.text_input("⏱️ Scrape From Exact Time (Copy-Paste):", st.session_state.last_checkpoint if st.session_state.last_checkpoint != "No record yet" else "[3/4/26, 7:16:42 PM]")

            if st.button("🚀 Launch WhatsApp & Fetch Orders", type="primary", use_container_width=True):
                target_limit_dt = parse_copy_paste_time(start_time_str)
                if not target_limit_dt:
                    st.error("⚠️ Start Time-এর ফরম্যাট ভুল! ব্র্যাকেট সহ ঠিকমতো কপি-পেস্ট করো।")
                else:
                    with st.spinner(f"Initializing Bot... Target Time: {target_limit_dt}"):
                        try:
                            chrome_options = Options()
                            chrome_options.add_argument("--disable-gpu")
                            chrome_options.add_argument("--no-sandbox")
                            chrome_options.add_argument("--disable-dev-shm-usage")
                            
                            service = ChromeService(ChromeDriverManager().install())
                            driver = webdriver.Chrome(service=service, options=chrome_options)
                            driver.get("https://web.whatsapp.com")
                            
                            st.info("🕒 Please scan the QR code within 60 seconds!")
                            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "side")))
                            st.success("✅ Login Successful! Searching for group...")
                            time.sleep(4)

                            search_box = None
                            try: search_box = driver.execute_script('return document.querySelector("#side [contenteditable=\'true\']");')
                            except: pass
                            
                            if not search_box:
                                search_xpaths = ['//*[@id="side"]//*[@contenteditable="true"]', '//div[@title="Search input textbox"]', '//div[@data-tab="3"]']
                                for xpath in search_xpaths:
                                    try:
                                        search_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
                                        if search_box: break
                                    except: pass
                                        
                            if not search_box: raise Exception("Could not find Search Box.")
                            
                            try: search_box.click()
                            except: driver.execute_script("arguments[0].click();", search_box)
                            time.sleep(1)
                            try:
                                search_box.send_keys(Keys.CONTROL + "a")
                                search_box.send_keys(Keys.DELETE)
                            except: pass 
                            search_box.send_keys(target_group)
                            time.sleep(3)
                            
                            group_clicked = False
                            group_xpaths = [f"//span[@title='{target_group}']", f"//span[contains(@title, '{target_group}')]", f"//div[@title='{target_group}']", f"//span[text()='{target_group}']"]
                            for xpath in group_xpaths:
                                try:
                                    elem = WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.XPATH, xpath)))
                                    driver.execute_script("arguments[0].click();", elem)
                                    group_clicked = True
                                    break
                                except: pass
                                    
                            if not group_clicked: raise Exception(f"Could not find or click group: '{target_group}'.")
                            st.info("✅ Group Found! Starting AUTO-SCROLL Engine...")
                            time.sleep(3)

                            safe_target_dt = target_limit_dt.replace(second=0)
                            scroll_attempts, max_scrolls = 0, 100 
                            
                            while scroll_attempts < max_scrolls:
                                msg_elements = driver.find_elements(By.XPATH, "//div[@data-pre-plain-text]")
                                if not msg_elements: break
                                first_msg_pre_text = msg_elements[0].get_attribute("data-pre-plain-text")
                                dt_str_match = re.search(r'\[(.*?)\]', first_msg_pre_text)
                                
                                if dt_str_match:
                                    dt_str = dt_str_match.group(1)
                                    parts = [p.strip() for p in dt_str.split(',')]
                                    oldest_dt = get_datetime_obj(parts[0], parts[1]) if len(parts) >= 2 else None
                                    
                                    if oldest_dt and oldest_dt > safe_target_dt:
                                        driver.execute_script("arguments[0].scrollIntoView(true);", msg_elements[0])
                                        time.sleep(1.5) 
                                        scroll_attempts += 1
                                    else:
                                        st.success(f"🎯 Target Time Reached! Extracting data...")
                                        break
                                else: break
                            
                            if scroll_attempts >= max_scrolls: st.warning("⚠️ Scrolled heavily. Extracting loaded data.")
                            time.sleep(2)

                            msg_elements = driver.find_elements(By.XPATH, "//div[@data-pre-plain-text]")
                            filtered_messages = []
                            for el in msg_elements: 
                                try:
                                    pre_text = el.get_attribute("data-pre-plain-text")
                                    text_span = el.find_element(By.XPATH, ".//span[contains(@class, 'selectable-text')]")
                                    msg_text = text_span.text
                                    
                                    dt_str_match = re.search(r'\[(.*?)\]', pre_text)
                                    if dt_str_match:
                                        dt_str = dt_str_match.group(1)
                                        parts = [p.strip() for p in dt_str.split(',')]
                                        msg_dt = get_datetime_obj(parts[0], parts[1]) if len(parts) >= 2 else None
                                        if not msg_dt: continue
                                        if msg_dt >= safe_target_dt:
                                            dt_obj = msg_dt.date()
                                            sender = pre_text.replace(f"[{dt_str}]", "").strip()
                                            date_str = parts[0] if '/' in parts[0] else parts[1]
                                            time_str = parts[0] if ':' in parts[0] else parts[1]
                                            filtered_messages.append({"date_obj": dt_obj, "date_str": date_str, "time_str": time_str, "msg_dt": msg_dt, "text": f"[{date_str}, {time_str}] {sender} {msg_text}"})
                                except: pass
                            driver.quit()
                            
                            st.session_state.total_scanned = len(filtered_messages)
                            temp_orders, temp_ignored = [], []
                            phone_counts = {}
                            
                            for msg in filtered_messages:
                                data = extract_order_details(msg)
                                if data:
                                    if data["status"] == "valid":
                                        del data["status"]
                                        ph = data['Phone Number']
                                        if ph != "N/A":
                                            phone_counts[ph] = phone_counts.get(ph, 0) + 1
                                            if phone_counts[ph] > 1: data["is_duplicate"] = True
                                            
                                        temp_orders.append(data)
                                    elif data["status"] == "ignored":
                                        temp_ignored.append(data)
                            
                            st.session_state.ignored_messages = temp_ignored
                            if temp_orders:
                                st.session_state.all_orders = temp_orders
                                st.session_state.sheet_date = f"Live_Scrape_{datetime.now(BD_TZ).strftime('%d-%m-%y_%H%M')}"
                                st.session_state.total_extracted_today += len(temp_orders)
                                
                                last_order = temp_orders[-1]
                                st.session_state.last_checkpoint = f"[{last_order['Date']}, {last_order['Time']}]"
                                
                                log_task(f"Scraped {len(temp_orders)} orders via Live Scraper. Start Time: {start_time_str}. <br>📌 **Stopped at Checkpoint:** `{st.session_state.last_checkpoint}`")
                                st.balloons()
                            else:
                                st.session_state.all_orders = []
                                st.warning("No valid new orders found from that exact time.")
                        except Exception as e:
                            st.error(f"❌ Scraping Failed! {e}")
                            try: driver.quit() 
                            except: pass

    # --- DASHBOARD UI ---
    if st.session_state.all_orders or st.session_state.ignored_messages:
        
        suspect_keywords = ['taka', 'tk', 'টাকা', '/-', 'pice', 'pcs', 'পিস', 'blender', 'grinder', 'order', 'অর্ডার', 'thana', 'zilla', 'জেলা', 'থানা', 'গ্রাম']
        suspected_msgs, system_junk = [], []
        
        for ig in st.session_state.ignored_messages:
            text_lower = ig['Text'].lower()
            if any(kw in text_lower for kw in suspect_keywords) and len(text_lower) > 15:
                suspected_msgs.append(ig)
            else:
                system_junk.append(ig)
        
        st.markdown("<br><h4 style='text-align:center; color:gray;'>📊 Data Reconciliation Report</h4>", unsafe_allow_html=True)
        col_r1, col_r2, col_r3, col_r4 = st.columns(4)
        col_r1.metric("🔍 Total Scanned", st.session_state.total_scanned)
        col_r2.metric("📦 Valid Orders", len(st.session_state.all_orders))
        col_r3.metric("🚨 Suspects", len(suspected_msgs))
        col_r4.metric("🗑️ System/Junk", len(system_junk))
        st.markdown("---")

        if st.session_state.all_orders:
            df = pd.DataFrame(st.session_state.all_orders)
            doubtful_orders = []
            passed_checks = 0
            total_checks = len(st.session_state.all_orders) * 3
            
            valid_prices = df[df['Price'].astype(int) > 0]['Price'].astype(int)
            avg_price = valid_prices.mean() if not valid_prices.empty else 0
            
            for i, row in enumerate(st.session_state.all_orders):
                issues = []
                p_check = re.match(r'^01[3-9]\d{8}$', str(row['Phone Number']))
                pr_val = int(row['Price'])
                pr_check = pr_val > 0
                q_check = int(row['Quantity']) > 0
                
                if p_check: passed_checks += 1
                else: issues.append("Invalid Phone")
                if pr_check: passed_checks += 1
                else: issues.append("Missing Price")
                if q_check: passed_checks += 1
                else: issues.append("Invalid Quantity")
                
                if avg_price > 0 and pr_val > 0 and pr_val < (avg_price * 0.5):
                    issues.append(f"📉 Low Price Alert (Avg is ৳{int(avg_price)})")
                
                # 🌟 HIGH QUANTITY ALERT 🌟
                if int(row['Quantity']) > 10:
                    issues.append("⚠️ High Qty (>10)")
                
                if str(row['Name']).strip() == "N/A" or not str(row['Name']).strip(): issues.append("Missing Name")
                elif any(h in str(row['Name']).lower() for h in ['বাড়ি', 'বাড়ি', 'থানা', 'জেলা', 'রোড', 'road', 'গ্রাম', 'house']):
                    issues.append("Name looks like Address")

                if str(row['Address']).strip() == "N/A" or not str(row['Address']).strip(): issues.append("Missing Address")
                if row.get('is_duplicate', False): issues.append("⚠️ Duplicate Data")
                
                if issues: doubtful_orders.append({"id": row['id'], "order": row, "issues": issues})
            
            accuracy_score = round((passed_checks / total_checks) * 100, 1) if total_checks > 0 else 0
            
            st.markdown("<br>", unsafe_allow_html=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📦 Orders", len(st.session_state.all_orders))
            m2.metric("💰 Revenue", f"৳ {sum(int(o['Price']) for o in st.session_state.all_orders)}")
            m3.metric("🎯 Accuracy", f"{accuracy_score}%")
            m4.metric("📈 Session Total", st.session_state.total_extracted_today)

            if doubtful_orders:
                st.error(f"⚠️ Action Required: Found {len(doubtful_orders)} doubtful or duplicate entries!")
                with st.expander("🚨 REVIEW DOUBTFUL ENTRIES", expanded=True):
                    for dob in doubtful_orders:
                        o_id = dob['id']
                        issue_text = ', '.join(dob['issues'])
                        st.markdown(f"""
                            <div class="doubt-card">
                                <div style="flex: 1;">
                                    <strong style="color: #D97706;">Issue:</strong> {issue_text}<br>
                                    <span style="color: gray; font-size:0.9rem;"><strong>Name:</strong> {dob['order']['Name']} | <strong>Phone:</strong> {dob['order']['Phone Number']}</span>
                                </div>
                                <div>
                                    <a href="#order-{o_id}" class="fix-btn">Fix 🔨</a>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
            
            col_m1, col_m2 = st.columns([4, 1.5])
            with col_m1:
                st.markdown("### 📋 Manage Orders")
            with col_m2:
                if st.button("➕ Add Manual Order", type="secondary"):
                    new_manual_order = {
                        "id": str(uuid.uuid4()), 
                        "Date": datetime.now(BD_TZ).strftime("%d/%m/%y"),
                        "Time": datetime.now(BD_TZ).strftime("%I:%M %p"),
                        "Name": "",
                        "Phone Number": "",
                        "Address": "",
                        "Product": "Electric Blender",
                        "Quantity": 1,
                        "Price": 0,
                        "Approval": "Pending",
                        "Note": "Manual Entry",
                        "is_duplicate": False,
                        "RawText": "✍️ This order was added manually.",
                        "Expander_Title": f"✍️ Manual Order | 🕒 {datetime.now(BD_TZ).strftime('%I:%M %p')}"
                    }
                    st.session_state.all_orders.append(new_manual_order)
                    st.rerun()

            for i, row in enumerate(st.session_state.all_orders):
                if 'id' not in row: row['id'] = str(uuid.uuid4())
                o_id = row['id']
                
                dup_tag = " (⚠️ Duplicate)" if row.get('is_duplicate', False) else ""
                final_title = row.get('Expander_Title', f"Order Details | 📞 {row.get('Phone Number','')}") + dup_tag
                
                st.markdown(f'<div id="order-{o_id}" style="position: relative; top: -60px;"></div>', unsafe_allow_html=True)
                
                with st.expander(final_title, expanded=False):
                    
                    clean_raw = row.get('RawText', 'N/A').replace('\n', '<br>')
                    clean_raw = re.split(r'^\[.*?\] .*?:\s', clean_raw, maxsplit=1)[-1] 
                    
                    bubble_html = f"""
                    <div style="display: flex; flex-direction: column; align-items: flex-start; margin-bottom: 20px;">
                        <div style="background: linear-gradient(135deg, #128C7E, #075E54); color: white; padding: 12px 18px; border-radius: 18px; border-top-left-radius: 2px; max-width: 95%; box-shadow: 0px 4px 10px rgba(0,0,0,0.15);">
                            <div style="font-size: 11px; color: #DCF8C6; margin-bottom: 5px; font-weight: 600;">💬 ORIGINAL CUSTOMER MESSAGE</div>
                            <div style="font-size: 15px; line-height: 1.5; font-family: sans-serif;">{clean_raw}</div>
                        </div>
                    </div>
                    """
                    st.markdown(bubble_html, unsafe_allow_html=True)
                    
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        new_name = st.text_input("👤 Name:", row['Name'], key=f"name_{o_id}")
                        new_addr = st.text_input("🏠 Address:", row['Address'], key=f"addr_{o_id}")
                        new_phone = st.text_input("📱 Phone:", row['Phone Number'], key=f"phone_{o_id}")
                        
                        st.session_state.all_orders[i]['Name'] = new_name
                        st.session_state.all_orders[i]['Address'] = new_addr
                        st.session_state.all_orders[i]['Phone Number'] = new_phone
                        
                        # 🌟 WRITEABLE PRODUCT NAME 🌟
                        new_prod = st.text_input("📦 Item:", row['Product'], key=f"prod_{o_id}")
                        st.session_state.all_orders[i]['Product'] = new_prod
                        
                        st.session_state.all_orders[i]['Note'] = st.text_input("📝 Note:", row.get('Note', ''), key=f"note_{o_id}")
                        
                    with c2:
                        col_p, col_q = st.columns(2)
                        with col_p:
                            st.session_state.all_orders[i]['Price'] = st.number_input("💰 Price (৳):", value=int(row['Price']), min_value=0, key=f"price_{o_id}")
                        with col_q:
                            st.session_state.all_orders[i]['Quantity'] = st.number_input("⚖️ Qty:", value=int(row['Quantity']), min_value=0, key=f"qty_{o_id}")
                            
                        status_list = ["Pending", "OK", "Canceled", "Talked", "Not Picked"]
                        current_idx = status_list.index(row['Approval']) if row['Approval'] in status_list else 0
                        st.session_state.all_orders[i]['Approval'] = st.selectbox("Status:", status_list, index=current_idx, key=f"status_{o_id}")
                        
                    col_rm1, col_rm2 = st.columns([2, 1])
                    with col_rm1:
                        st.markdown(f'''<a href="tel:{st.session_state.all_orders[i]['Phone Number']}" style="display:inline-block; text-align:center; width:100%; background: linear-gradient(135deg, #10B981, #059669); color:white; padding:10px 15px; border-radius:25px; margin-top:20px; font-weight:bold; box-shadow: 0px 4px 10px rgba(16, 185, 129, 0.3); text-decoration:none;">📞 Call Customer</a>''', unsafe_allow_html=True)
                    with col_rm2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🗑️ Remove", key=f"del_{o_id}"):
                            st.session_state.all_orders = [o for o in st.session_state.all_orders if o['id'] != o_id]
                            st.rerun()

            st.markdown("---")
            filename = f"NWOP_Orders_{st.session_state.sheet_date}.xlsx"
            csv_filename = f"NWOP_Orders_{st.session_state.sheet_date}.csv"
            
            # 🌟 MODIFIED EXPORT ORDER (Date & Time at Rightmost) 🌟
            export_data = [{k:v for k,v in order.items() if k not in ['is_duplicate', 'id', 'RawText', 'Expander_Title']} for order in st.session_state.all_orders]
            export_df = pd.DataFrame(export_data)
            
            export_df['Quantity'] = pd.to_numeric(export_df['Quantity'], errors='coerce').fillna(1).astype(int)
            export_df['is_multi'] = export_df['Quantity'] > 1
            export_df = export_df.sort_values(by=['is_multi']).drop(columns=['is_multi'])
            
            export_df['SNO'] = range(1, 1 + len(export_df))
            
            csv_df = export_df.rename(columns={'SNO': 'Sl.', 'Price': 'price', 'Approval': 'approved'})
            if 'Note' not in csv_df.columns: csv_df['Note'] = ""
            csv_columns = ['Sl.', 'Name', 'Phone Number', 'Address', 'Quantity', 'Product', 'price', 'approved', 'Note', 'Date', 'Time']
            for c in csv_columns:
                if c not in csv_df.columns: csv_df[c] = ""
            csv_df = csv_df[csv_columns]
            
            # 🌟 EXCEL EXPORT COLUMNS (Date and Time Moved to End) 🌟
            export_columns = ["SNO", "Name", "Phone Number", "Address", "Quantity", "Product", "Price", "Approval", "Note", "Date", "Time"]
            for col in export_columns:
                if col not in export_df.columns: export_df[col] = ""
            export_df = export_df[export_columns]
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name="Orders")
                workbook = writer.book
                worksheet = writer.sheets['Orders']
                
                # Update DataValidation to new columns: Approval is now 'H'
                status_dv = DataValidation(type="list", formula1='"Pending,OK,Canceled,Talked,Not Picked"', allow_blank=True)
                worksheet.add_data_validation(status_dv)
                status_dv.add('H2:H10000') 
                
                pd.DataFrame({"Date": [st.session_state.sheet_date], "Total": [len(export_df)]}).to_excel(writer, index=False, sheet_name="Summary")
                for idx, prod in enumerate(st.session_state.product_list, start=1): writer.sheets['Summary'].cell(row=idx, column=5, value=prod)
                
                # Product is now 'F'
                prod_dv = DataValidation(type="list", formula1=f"Summary!$E$1:$E${len(st.session_state.product_list)}", allow_blank=True)
                worksheet.add_data_validation(prod_dv)
                prod_dv.add('F2:F10000') 
                
                header_fill = PatternFill(start_color="e6f2ff", end_color="e6f2ff", fill_type="solid")
                sno_fill = PatternFill(start_color="10B981", end_color="10B981", fill_type="solid")
                green_row_fill = PatternFill(start_color="c6efce", end_color="c6efce", fill_type="solid")
                
                for cell in worksheet[1]: cell.fill, cell.font, cell.alignment, cell.border = header_fill, Font(bold=True, color="000000"), Alignment(horizontal="center", vertical="center"), Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                    # Product is now at index 5 (Column F)
                    prod_val = str(row[5].value).strip().lower()
                    is_not_blender = (prod_val != "electric blender")
                    
                    for cell in row:
                        cell.border, cell.alignment = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')), Alignment(vertical="center")
                        if cell.column == 1: 
                            cell.fill, cell.font, cell.alignment = sno_fill, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center")
                        elif is_not_blender:
                            cell.fill = green_row_fill

                # Conditional formatting rules for Approval (Column H)
                worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"OK"'], fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), font=Font(color="006100")))
                worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"Pending"'], fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"), font=Font(color="9C5700")))
                worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"Canceled"'], fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), font=Font(color="9C0006")))
                
                for col in worksheet.columns:
                    max_len = 0
                    for cell in col:
                        try: max_len = max(max_len, len(str(cell.value)))
                        except: pass
                    worksheet.column_dimensions[col[0].column_letter].width = max_len + 2

            excel_data = output.getvalue()
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    on_click=lambda: log_task(f"Downloaded Excel file: {filename}")
                )
            
            with col_d2:
                csv_data = csv_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📊 Download CSV (For Google Sheets)",
                    data=csv_data,
                    file_name=csv_filename,
                    mime="text/csv",
                    type="secondary",
                    use_container_width=True,
                    on_click=lambda: log_task(f"Downloaded CSV file for Google Sheets: {csv_filename}")
                )

        if suspected_msgs:
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander(f"🚨 SUSPECTED MISSED ORDERS ({len(suspected_msgs)} items) - MUST CHECK!", expanded=True):
                st.warning("Ei message gulor vitore 'taka', 'thana', 'blender' er moto order word ache, kintu kono Valid Phone Number pawa jayni! Dayakore manual check korun.")
                for sm in suspected_msgs:
                    st.caption(f"🕒 {sm['Date']} - {sm['Time']}")
                    clean_display = re.split(r'^\[.*?\] .*?:\s', sm['Text'], maxsplit=1)[-1]
                    st.error(clean_display)
                    
        if system_junk:
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander(f"🗑️ System Messages / Junk ({len(system_junk)} items)", expanded=False):
                st.info("Egulo asha kora jay normal text/system message. Ete kono order info nai.")
                for jm in system_junk:
                    st.caption(f"🕒 {jm['Date']} - {jm['Time']}")
                    clean_display = re.split(r'^\[.*?\] .*?:\s', jm['Text'], maxsplit=1)[-1]
                    st.code(clean_display, language="text")

with tab_merge:
    st.header("🗂️ Excel Merger (Smart Sorter)")
    st.info("Here you can upload multiple NWOP Excel files. The app will merge them, remove duplicates, sort them perfectly by Date & Time, and create a single Master File without losing phone number leading '0's!")
    
    uploaded_excels = st.file_uploader("📂 Select multiple Excel files", type=["xlsx"], accept_multiple_files=True)
    
    if uploaded_excels and len(uploaded_excels) > 0:
        with st.spinner("Processing files..."):
            try:
                all_dfs = []
                for file in uploaded_excels:
                    df = pd.read_excel(file, sheet_name="Orders", dtype=str)
                    if 'Phone Number' in df.columns:
                        df['Phone Number'] = df['Phone Number'].fillna("N/A").apply(lambda x: str(x).replace('.0', '') if str(x).endswith('.0') else str(x))
                    all_dfs.append(df)
                
                merged_df = pd.concat(all_dfs, ignore_index=True)
                
                merged_df['sort_dt'] = merged_df.apply(lambda r: get_datetime_obj(str(r.get('Date', '')), str(r.get('Time', ''))) or datetime.min, axis=1)
                merged_df = merged_df.sort_values(by='sort_dt')
                merged_df = merged_df.drop_duplicates(subset=['Phone Number'], keep='last')
                merged_df = merged_df.drop(columns=['sort_dt'])
                
                # 🌟 SORTING MERGED FILE: QUANTITY > 1 GOES TO BOTTOM 🌟
                merged_df['Quantity'] = pd.to_numeric(merged_df['Quantity'], errors='coerce').fillna(1).astype(int)
                merged_df['is_multi'] = merged_df['Quantity'] > 1
                merged_df = merged_df.sort_values(by=['is_multi']).drop(columns=['is_multi'])
                
                merged_df['SNO'] = range(1, len(merged_df) + 1)
                
                # 🌟 REORDER EXCEL COLUMNS FOR MERGE EXPORT 🌟
                export_columns = ["SNO", "Name", "Phone Number", "Address", "Quantity", "Product", "Price", "Approval", "Note", "Date", "Time"]
                excel_df = merged_df.copy()
                for c in export_columns:
                    if c not in excel_df.columns: excel_df[c] = ""
                excel_df = excel_df[export_columns]
                
                output_merge = BytesIO()
                with pd.ExcelWriter(output_merge, engine='openpyxl') as writer:
                    excel_df.to_excel(writer, index=False, sheet_name="Orders")
                    workbook = writer.book
                    worksheet = writer.sheets['Orders']
                    
                    status_dv = DataValidation(type="list", formula1='"Pending,OK,Canceled,Talked,Not Picked"', allow_blank=True)
                    worksheet.add_data_validation(status_dv)
                    status_dv.add('H2:H10000') # Approval is now H
                    
                    pd.DataFrame({"Date Tag": [f"Merged_{datetime.now(BD_TZ).strftime('%d-%m-%y')}"], "Total Orders": [len(excel_df)]}).to_excel(writer, index=False, sheet_name="Summary")
                    for idx, prod in enumerate(st.session_state.product_list, start=1): writer.sheets['Summary'].cell(row=idx, column=5, value=prod)
                    
                    prod_dv = DataValidation(type="list", formula1=f"Summary!$E$1:$E${len(st.session_state.product_list)}", allow_blank=True)
                    worksheet.add_data_validation(prod_dv)
                    prod_dv.add('F2:F10000') # Product is now F
                    
                    header_fill = PatternFill(start_color="e6f2ff", end_color="e6f2ff", fill_type="solid")
                    sno_fill = PatternFill(start_color="10B981", end_color="10B981", fill_type="solid")
                    green_row_fill = PatternFill(start_color="c6efce", end_color="c6efce", fill_type="solid") 
                    
                    for cell in worksheet[1]: cell.fill, cell.font, cell.alignment, cell.border = header_fill, Font(bold=True, color="000000"), Alignment(horizontal="center", vertical="center"), Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                        prod_val = str(row[5].value).strip().lower() # Product is now at index 5 (Column F)
                        is_not_blender = (prod_val != "electric blender")
                        
                        for cell in row:
                            cell.border, cell.alignment = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin')), Alignment(vertical="center")
                            if cell.column == 1: 
                                cell.fill, cell.font, cell.alignment = sno_fill, Font(bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center")
                            elif is_not_blender:
                                cell.fill = green_row_fill

                    worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"OK"'], fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), font=Font(color="006100")))
                    worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"Pending"'], fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"), font=Font(color="9C5700")))
                    worksheet.conditional_formatting.add('H2:H10000', CellIsRule(operator='equal', formula=['"Canceled"'], fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), font=Font(color="9C0006")))
                    
                    for col in worksheet.columns:
                        max_len = 0
                        for cell in col:
                            try: max_len = max(max_len, len(str(cell.value)))
                            except: pass
                        worksheet.column_dimensions[col[0].column_letter].width = max_len + 2
                        
                st.success(f"✅ Successfully combined {len(uploaded_excels)} files into {len(excel_df)} unique sorted orders!")
                
                # 🌟 CSV EXPORT FORMAT (MERGED) 🌟
                csv_df = merged_df.rename(columns={'SNO': 'Sl.', 'Price': 'price', 'Approval': 'approved'})
                if 'Note' not in csv_df.columns: csv_df['Note'] = ""
                csv_columns = ['Sl.', 'Name', 'Phone Number', 'Address', 'Quantity', 'Product', 'price', 'approved', 'Note', 'Date', 'Time']
                for c in csv_columns:
                    if c not in csv_df.columns: csv_df[c] = ""
                csv_df = csv_df[csv_columns]
                
                col_md1, col_md2 = st.columns(2)
                with col_md1:
                    st.download_button(
                        label="📥 Download Master Excel",
                        data=output_merge.getvalue(),
                        file_name=f"NWOP_Master_{datetime.now(BD_TZ).strftime('%d-%m-%y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                        on_click=lambda: log_task(f"Merged {len(uploaded_excels)} files into 1 Master Excel File.")
                    )
                with col_md2:
                    st.download_button(
                        label="📊 Download CSV (For Google Sheets)",
                        data=csv_df.to_csv(index=False).encode('utf-8'),
                        file_name=f"NWOP_Master_{datetime.now(BD_TZ).strftime('%d-%m-%y')}.csv",
                        mime="text/csv",
                        type="secondary",
                        use_container_width=True,
                        on_click=lambda: log_task(f"Merged CSV Downloaded.")
                    )
            except Exception as e:
                st.error(f"Error processing files: {e}. Please make sure you uploaded valid NWOP Excel files.")

with tab_history:
    st.header("📜 Task History & Logs")
    if not st.session_state.task_history:
        st.info("No activity recorded yet.")
    else:
        for task in st.session_state.task_history: st.markdown(task, unsafe_allow_html=True)
        if st.button("Clear History", type="secondary"):
            st.session_state.task_history = []
            st.session_state.last_checkpoint = "No record yet"
            save_data([], "No record yet")
            st.rerun()

with tab_settings:
    st.header("⚙️ NWOP Settings")
    st.markdown("**Version:** NWOP v16.0 (Perfect Layout Edition)")
    st.info(f"The default master password is '{CORRECT_PASSWORD}'.")
    if st.button("Reset Memory / Clear App Data", type="secondary"):
        st.session_state.all_orders, st.session_state.ignored_messages = [], []
        st.session_state.total_extracted_today = 0
        st.session_state.total_scanned = 0
        log_task("App memory completely wiped.")
        st.rerun()

with tab_about:
    st.header("ℹ️ About Developer")
    st.markdown("---")
    
    col_a1, col_a2 = st.columns([1, 3])
    with col_a1:
        if os.path.exists("logo.png"): st.image("logo.png", width=150)
        else: st.markdown("<h1 style='font-size: 80px; margin-top: -20px;'>👨‍💻</h1>", unsafe_allow_html=True)
    with col_a2:
        st.markdown("### **Nazrul's Whatsapp Order Parser (NWOP)**")
        st.write("This application is an enterprise-grade automation tool designed to extract, parse, and manage WhatsApp orders with high accuracy, smart formatting, duplicate detection, and direct Excel compilation.")
    
    st.markdown("#### 👨‍💻 Developer Profile")
    st.markdown("""
    * **Name:** Nazrul Rana
    * **WhatsApp:** +880164143400
    * **Version:** 16.0 (Perfect Layout Edition)
    """)
    
    st.info("For any bug reports, feature requests, custom automation tools, or software development inquiries, please feel free to reach out via WhatsApp.")
