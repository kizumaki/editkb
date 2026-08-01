import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
import io
import os
import re
import random
from collections import Counter
import time
import pandas as pd
import json

# --- CẤU HÌNH TRANG CHỦ STREAMLIT (SaaS LAYOUT) ---
st.set_page_config(
    page_title="ScriptPro Enterprise - Subtitle & Script Editor",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- TIÊM CUSTOM CSS CHUẨN SAAS CAO CẤP ---
st.markdown("""
<style>
    /* Tổng thể font & nền */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Hero Banner Header */
    .hero-container {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(79, 70, 229, 0.3);
    }
    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        opacity: 0.9;
        margin-top: 0.5rem;
    }
    .badge-pro {
        background-color: rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(8px);
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        display: inline-block;
        margin-bottom: 0.8rem;
    }
    
    /* SaaS Stat Metric Cards */
    .metric-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.25rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748B;
        font-weight: 500;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0F172A;
        margin-top: 0.25rem;
    }

    /* Đổi kiểu dáng cho Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 8px 8px 0 0;
        font-weight: 600;
        padding: 0 20px;
    }
    
    /* Styled Footer */
    .saas-footer {
        text-align: center;
        padding: 2rem 0 1rem 0;
        color: #94A3B8;
        font-size: 0.85rem;
        border-top: 1px solid #E2E8F0;
        margin-top: 3rem;
    }
</style>
""", unsafe_allow_html=True)

# --- HÀM ĐỌC / GHI DATABASE DỮ LIỆU CỤC BỘ ---
NON_SPEAKER_DB_FILE = "custom_non_speakers.json"
SPEAKER_DB_FILE = "custom_speakers.json"
PHONETIC_DB_FILE = "custom_phonetics.json"

DEFAULT_SOUTH_VIETNAM_PHONETICS = {
    "MCDONALD": "Mắc-đô-nồ",
    "MCDONALD'S": "Mắc-đô-nồ",
    "LONDON": "Luân Đôn",
    "TNT": "Ti-èn-ti",
    "NOOB": "Núp",
    "GOLEM": "Gô-lùm",
    "GOOGLE": "Gú-gồ",
    "FACEBOOK": "Phây-sbúc",
    "KFC": "Ke-ép-xi",
    "STARBUCKS": "Xì-ta-bắc",
    "APPLE": "Ép-pồ",
    "BURGER": "Bơ-gơ",
    "PARIS": "Ba Lê",
    "WASHINGTON": "Hoa Thịnh Đốn",
    "CALIFORNIA": "Cả-li",
    "PRO": "Prồ",
    "LAG": "Lác",
    "BUFF": "Bớp",
    "VIP": "Vi-ai-pi",
    "FBI": "Ép-bi-ai"
}

VN_SYLLABLES = {
    "a", "ai", "an", "ang", "anh", "ao", "ap", "at", "au", "ay", "ba", "bac", "bai", "ban", "bang", "bao", 
    "bat", "bay", "be", "ben", "beng", "beo", "bi", "bic", "bien", "biet", "binh", "bo", "boc", "boi", "bon", 
    "bong", "bot", "bu", "bua", "bui", "bun", "buoc", "buon", "ca", "cac", "cai", "cam", "can", "cang", "cao", 
    "cap", "cat", "cau", "cay", "cha", "chai", "cham", "chan", "chang", "chao", "chap", "chat", "chau", "chay", 
    "che", "chen", "cheo", "chi", "chia", "chiem", "chieu", "chin", "chinh", "cho", "choc", "choi", "chon", 
    "chong", "chu", "chua", "chuan", "chuc", "chui", "chum", "chun", "chuoc", "chuon", "chura", "chuong", "co", 
    "coc", "coi", "con", "cong", "cot", "cu", "cua", "cuc", "cui", "cum", "cung", "cuoc", "cuoi", "cuon", 
    "cuong", "cuot", "da", "dac", "dai", "dam", "dan", "dang", "dao", "dap", "dat", "dau", "day", "de", "den", 
    "deo", "di", "dia", "diem", "dien", "diet", "dieu", "dinh", "do", "doc", "doi", "don", "dong", "dot", "du", 
    "dua", "duc", "dui", "dum", "dung", "duoc", "duoi", "duon", "duong", "em", "ga", "gac", "gai", "gam", "gan", 
    "gang", "gao", "gap", "gat", "gau", "gay", "ge", "gen", "gi", "gia", "giac", "giai", "giam", "gian", "giang", 
    "giao", "giap", "giat", "giau", "giay", "gio", "gioc", "gioi", "gion", "giong", "giu", "giua", "go", "goc", 
    "goi", "gon", "gong", "gu", "gua", "guc", "gui", "gum", "gung", "ha", "hac", "hai", "ham", "han", "hang", 
    "hao", "hap", "hat", "hau", "hay", "he", "hen", "heo", "hi", "hien", "hiet", "hieu", "hinh", "ho", "hoc", 
    "hoi", "hon", "hong", "hot", "hu", "hua", "huc", "hui", "hum", "hung", "huoc", "huong", "huot", "i", "ic", 
    "it", "khac", "khai", "kham", "khan", "khang", "khao", "khap", "khat", "khau", "khay", "khe", "khen", "kheo", 
    "khi", "khia", "khien", "khieu", "kho", "khoc", "khoi", "khon", "khong", "khot", "khu", "khua", "khuc", 
    "khui", "khum", "khung", "khuon", "khura", "khuong", "la", "lac", "lai", "lam", "lan", "lang", "lao", "lap", 
    "lat", "lau", "lay", "le", "len", "leo", "li", "lia", "liem", "lien", "liet", "lieu", "linh", "lo", "loc", 
    "loi", "lon", "long", "lot", "lu", "lua", "luc", "lui", "lum", "lung", "luoc", "luoi", "luon", "luong", 
    "luot", "ma", "mac", "mai", "mam", "man", "mang", "mao", "map", "mat", "mau", "may", "me", "men", "meo", 
    "mi", "mia", "mien", "mieu", "minh", "mo", "moc", "moi", "mon", "mong", "mot", "mu", "mua", "muc", "mui", 
    "mum", "mung", "muoc", "muoi", "muon", "muong", "na", "nac", "nai", "nam", "nan", "nang", "nao", "nap", 
    "nat", "nau", "nay", "ne", "nen", "neo", "ni", "nia", "niem", "nien", "niet", "nieu", "ninh", "no", "noc", 
    "noi", "non", "nong", "not", "nu", "nua", "nuc", "nui", "num", "nung", "nuoc", "nuoi", "nuon", "nuong", 
    "nga", "ngac", "ngai", "ngam", "ngan", "ngang", "ngao", "ngap", "ngat", "ngau", "ngay", "nge", "ngen", 
    "nghe", "nghen", "ngheo", "nghi", "nghia", "nghiem", "nghien", "nghiet", "nghieu", "nghinh", "ngo", "ngoc", 
    "ngoi", "ngon", "ngong", "ngot", "ngu", "ngua", "nguc", "ngui", "ngum", "ngung", "nguoc", "nguoi", "nguon", 
    "nguong", "nha", "nhac", "nhai", "nham", "nhan", "nhang", "nhao", "nhap", "nhat", "nhau", "nhay", "nhe", 
    "nhen", "nheo", "nhi", "nhia", "nhiem", "nhien", "nhiet", "nhieu", "nhinh", "nho", "nhoc", "nhoi", "nhon", 
    "nhong", "nhot", "nhu", "nhua", "nhuc", "nhui", "nhum", "nhung", "nhuoc", "nhuoi", "nhuon", "nhuong", "oa", 
    "oai", "oam", "oan", "oang", "oat", "oay", "oe", "oen", "oi", "om", "on", "ong", "ot", "pa", "pha", "phac", 
    "phai", "pham", "phan", "phang", "phao", "phap", "phat", "phau", "phay", "phe", "phen", "pheo", "phi", 
    "phia", "phiem", "phien", "phiet", "phieu", "phinh", "pho", "phoc", "phoi", "phon", "phong", "phot", "phu", 
    "phua", "phuc", "phui", "phum", "phung", "phuoc", "phuong", "phuot", "qua", "quac", "quai", "quam", "quan", 
    "quang", "quao", "quap", "quat", "quay", "que", "quen", "queo", "qui", "quie", "quien", "quieu", "quinh", 
    "quo", "quoc", "quoi", "quon", "quong", "ra", "rac", "rai", "ram", "ran", "rang", "rao", "rap", "rat", 
    "rau", "ray", "re", "ren", "reo", "ri", "ria", "rim", "rin", "rinh", "ro", "roc", "roi", "ron", "rong", 
    "rot", "ru", "rua", "ruc", "rui", "rum", "rung", "ruoc", "ruoi", "ruon", "ruong", "sa", "sac", "sai", 
    "sam", "san", "sang", "sao", "sap", "sat", "sau", "say", "se", "sen", "seo", "si", "sia", "siem", "sien", 
    "siet", "sieu", "sinh", "so", "soc", "soi", "son", "song", "sot", "su", "sua", "suc", "sui", "sum", 
    "sung", "suoc", "suoi", "suon", "suong", "ta", "tac", "tai", "tam", "tan", "tang", "tao", "tap", "tat", 
    "tau", "tay", "te", "ten", "teo", "tha", "thac", "thai", "tham", "than", "thang", "thao", "thap", "that", 
    "thau", "thay", "the", "then", "theo", "thi", "thia", "thiem", "thien", "thiet", "thieu", "thinh", "tho", 
    "thoc", "thoi", "thon", "thong", "thot", "thu", "thua", "thuc", "thui", "thum", "thung", "thuoc", "thuoi", 
    "thuon", "thuong", "thuot", "to", "toc", "toi", "ton", "tong", "tot", "tra", "trac", "trai", "tram", "tran", 
    "trang", "trao", "trap", "trat", "trau", "tray", "tre", "tren", "treo", "tri", "tria", "triem", "trien", 
    "triet", "trieu", "trinh", "tro", "troc", "troi", "tron", "trong", "trot", "tru", "trua", "truc", "trui", 
    "trum", "trung", "truoc", "truoi", "truon", "truong", "tu", "tua", "tuc", "tui", "tum", "tung", "tuoc", 
    "tuoi", "tuon", "tuong", "tuot", "va", "vac", "vai", "vam", "van", "vang", "vao", "vap", "vat", "vau", 
    "vay", "ve", "ven", "veo", "vi", "via", "viem", "vien", "viet", "vieu", "vinh", "vo", "voc", "voi", "von", 
    "vong", "vot", "vu", "vua", "vuc", "vui", "vum", "vung", "vuoc", "vuoi", "vuon", "vuong", "xa", "xac", 
    "xai", "xam", "xan", "xang", "xao", "xap", "xat", "xau", "xay", "xe", "xen", "xeo", "xi", "xia", "xiem", 
    "xien", "xiet", "xieu", "xinh", "xo", "xoc", "xoi", "xon", "xong", "xot", "xu", "xua", "xuc", "xui", 
    "xum", "xung", "xuoc", "xuoi", "xuon", "xuong", "y", "eu", "yeu"
}

def is_candidate_english_word(word):
    clean_w = word.strip(".,!?:;\"'()[]{}")
    if not clean_w or len(clean_w) <= 1 or clean_w.isdigit():
        return False
        
    lower_w = clean_w.lower()
    
    if clean_w.upper() in st.session_state['custom_phonetics']:
        return True
        
    if lower_w in VN_SYLLABLES:
        return False
        
    if any(char in lower_w for char in ['f', 'j', 'w', 'z']):
        return True
        
    eng_patterns = [
        r"(bb|cc|dd|ff|gg|ll|mm|nn|pp|rr|ss|tt|zz)",
        r"(sh|ck|th|wh|ph|gh)",
        r"(tion|ment|ness|less|ing|ed|able|ible|ally|ce|ge|ck)$",
        r"^([A-Z]{2,})$"
    ]
    for pattern in eng_patterns:
        if re.search(pattern, clean_w):
            return True
            
    if clean_w[0].isupper() and lower_w not in VN_SYLLABLES:
        return True
        
    return False

def load_json_db(filepath, default_data=None):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                return set(data) if isinstance(data, list) else data
        except Exception:
            pass
    return default_data if default_data is not None else (set() if not isinstance(default_data, dict) else {})

def save_json_db(filepath, data_container):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            if isinstance(data_container, set):
                json.dump(list(data_container), f, ensure_ascii=False, indent=2)
            else:
                json.dump(data_container, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Không thể lưu vào Database: {e}")

# Khởi tạo Session State
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0
if 'spk_input_key' not in st.session_state:
    st.session_state['spk_input_key'] = 0
if 'ns_input_key' not in st.session_state:
    st.session_state['ns_input_key'] = 0
if 'pho_input_key' not in st.session_state:
    st.session_state['pho_input_key'] = 0

if 'custom_non_speakers' not in st.session_state:
    st.session_state['custom_non_speakers'] = load_json_db(NON_SPEAKER_DB_FILE, set())

if 'custom_speakers' not in st.session_state:
    st.session_state['custom_speakers'] = load_json_db(SPEAKER_DB_FILE, set())

if 'custom_phonetics' not in st.session_state:
    loaded_pho = load_json_db(PHONETIC_DB_FILE, DEFAULT_SOUTH_VIETNAM_PHONETICS)
    merged_pho = {**DEFAULT_SOUTH_VIETNAM_PHONETICS, **loaded_pho}
    st.session_state['custom_phonetics'] = merged_pho

# --- DANH SÁCH MẶC ĐỊNH ---
DEFAULT_NON_SPEAKER_PHRASES = {
    "AND REMEMBER", "OFFICIAL DISTANCE", "GOOD NEWS FOR THEIR TEAMMATES", 
    "LL BE HONEST", "FIRST AND FOREMOST", "I SAID", "THE ONLY THING LEFT TO SETTLE", 
    "QUESTION IS", "FINALISTS", "WHISPERS", "SRT CONVERSION", 
    "WILL RED THRIVE OR WILL RED BE DEAD", "BUT REMEMBER", "THE RESULTS ARE IN", 
    "WE CHALLENGED", "I THINK", "IN THEIR DEFENSE", "THE PEAK OF HIS LIFE WAS DOING THE SPACETHING",
    "THE ROCKETS ARE BIGGER", "THE DISTANCE SHOULD BE FURTHER", "GET CRAFTY", "THAT WAS SO SICK",
    "OUT OF 100 CONTESTANTS", "THE FIRST ROUND IS BRUTAL", "YOU KNOW WHICH END GOES",
    "THE GAME IS ON", "THAT'S A GOOD THROW", "HE'S GOING FOR IT", "WE GOT THIS",
    "LAUNCH", "OH NO", "OH", "AH", "YEP", "WAIT", "YEAH", "WOO", "OKAY", "YES", "I ANH", "O BRI", "NG", "THE ONLY PROBLEM", "NOTE", "WARNING", "THINGS", "AND ON THE WAY WE CAME ACROSS THIS", 
    "THIS IS THE HIGHEST SWING IN EUROPE", "AND I SWEAR", "WHICH MEANT", "THE ONLY THING IS", 
    "HERE WE GO", "NEXT UP", "STEP 1", "STEP 2", "STEP 3", "AND STEP 3", "FIRST UP", 
    "SO THE QUESTION IS", "I WAS GROWING UP", "YOU MIGHT BE WONDERING", "UPDATE", 
    "NASHVILLE TO MIAMI", "ALL I KNOW IS", "UNLIKE JUDY", "THE GOOD NEWS IS", 
    "AER LINGUS SEAT", "THE TRUE TEST IS", "JUST AS I SUSPECTED", "LIKE I SAID", 
    "STAR REVIEW AND SAID", "I TOLD THEM ALL", "AND BEST OF ALL", "THE POINT IS", 
    "AMERICANS", "I WAS THINKING", "AND THEY GO", "FIRST OF ALL", "SECOND", 
    "ARE YOU LIKE", "AS A REMINDER", "ROUND 2", "ROUND 1", "ROUND 3", "ROUND 4", 
    "ROUND 5", "WELCOME TO ROUND 3", "THE QUESTION IS", "QUICK REMINDER", 
    "IN 2ND PLACE", "COMING UP", "FIRST STOP", "NEXT STEP", "AND THAT MEANS", 
    "HASHTAG", "SO TO BE CLEAR", "YOUR SECOND WORD", "WELCOME TO ROUND 6", 
    "BATTLE FINALE TIME", "NUMBER 1", "NUMBER 2", "BUT THE TRUTH IS", 
    "SCORE TO BEAT", "AND YOUR WINNER", "\"CRAFTY\" AND \"BETCHA\". COMING UP", 
    "NEXT ONE", "KEEP IN MIND", "AND IT SAYS", "YOU COULD SAY", "WELCOME TO ROUND 2", 
    "AND THE BEST PART", "ONTO ROUND 2", "THE RIDE WE CHOSE", "GOOD NEWS IS", 
    "BAD NEWS", "GOOD NEWS", "HE THOUGHT", "3 TEAMS REMAIN", "QUICK UPDATE", "DISTORTED", "MY FIRST QUESTION IS", "AND THE BEST PART IS", "BUT AS GORDON RAMSAY MIGHT SAY", "BUT THE GOOD NEWS IS", "LET ME JUST SAY", "BUT THE BEST PART", "I WILL SAY THOUGH", "LL SAY IS, UPDATE", "LL SAY IS", "UPDATE", "TO ALL OF YOU WATCHING WITH ME RIGHT NOW", "THIS IS THE LIFE", "I HAVE A QUESTION", "I WILL BE HONEST", "SO I CAN UPDATE MY INSTAGRAM BIO TO SAY"
}

NON_SPEAKER_PHRASES = DEFAULT_NON_SPEAKER_PHRASES.union(st.session_state['custom_non_speakers'])

TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)
ENGLISH_WORD_REGEX = re.compile(r"\b[A-Za-z][A-Za-z0-9'-]*\b")

def build_speaker_regex(custom_speakers):
    base_pattern = r"[\w\s&\.\-\(\)]+"
    if custom_speakers:
        sorted_custom = sorted(list(custom_speakers), key=len, reverse=True)
        custom_pattern = "|".join([re.escape(s) for s in sorted_custom])
        pattern_str = rf"({custom_pattern}|{base_pattern}):\s*"
    else:
        pattern_str = rf"({base_pattern}):\s*"
    return re.compile(pattern_str, re.IGNORECASE | re.UNICODE)

def scan_candidate_speakers(uploaded_file, speaker_regex):
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    candidates = Counter()
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or text.lower().startswith("srt conversion"): 
            continue
        for match in speaker_regex.finditer(text):
            speaker_name = match.group(1).strip()
            candidates[speaker_name] += 1
    return candidates

def scan_english_words_in_dialogue(uploaded_file, speaker_regex):
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    eng_found = set()
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or text.lower().startswith("srt conversion") or TIMECODE_REGEX.match(text):
            continue
        
        parts = speaker_regex.split(text)
        dialogue_content = ""
        
        if len(parts) == 1:
            dialogue_content = text
        else:
            speaker_matches = list(speaker_regex.finditer(text))
            last_idx = 0
            for m in speaker_matches:
                end = m.end()
                dialogue_content += " " + text[last_idx:m.start()]
                last_idx = end
            dialogue_content += " " + text[last_idx:]

        for match in ENGLISH_WORD_REGEX.finditer(dialogue_content):
            word = match.group(0).strip()
            if is_candidate_english_word(word):
                eng_found.add(word)
                
    return sorted(list(eng_found), key=lambda x: x.upper())

def extract_phrases_from_file(file_io, file_name):
    phrases = set()
    try:
        if file_name.endswith('.txt'):
            content = file_io.getvalue().decode("utf-8")
            phrases.update([line.strip() for line in content.split('\n') if line.strip()])
        elif file_name.endswith('.docx'):
            doc = Document(io.BytesIO(file_io.getvalue()))
            for p in doc.paragraphs:
                if p.text.strip():
                    parts = re.split(r'[,\n]', p.text)
                    phrases.update([part.strip() for part in parts if part.strip()])
        elif file_name.endswith('.xlsx'):
            df = pd.read_excel(file_io, header=None)
            for col in df.columns:
                for item in df[col].dropna():
                    parts = re.split(r'[,\n]', str(item))
                    phrases.update([part.strip() for part in parts if part.strip()])
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
    return phrases

def generate_vibrant_rgb_colors(count=200):
    colors = set()
    while len(colors) < count:
        h = random.random()
        s = 0.9; v = 0.8
        if s == 0.0: r = g = b = v
        else:
            i = int(h * 6.0); f = h * 6.0 - i; p = v * (1.0 - s); q = v * (1.0 - s * f); t = v * (1.0 - s * (1.0 - f))
            if i % 6 == 0: r, g, b = v, t, p
            elif i % 6 == 1: r, g, b = q, v, p
            elif i % 6 == 2: r, g, b = p, v, t
            elif i % 6 == 3: r, g, b = p, q, v
            elif i % 6 == 4: r, g, b = t, p, v
            else: r, g, b = v, p, q
        r, g, b = int(r * 255), int(g * 255), int(b * 255)
        colors.add((r, g, b))
    return list(colors)

FONT_COLORS_RGB_200 = generate_vibrant_rgb_colors(200)

def get_speaker_color(speaker_name, speaker_color_map, used_colors):
    if speaker_name not in speaker_color_map:
        if used_colors:
            color_object = used_colors.pop()
        else:
            r, g, b = random.choice(FONT_COLORS_RGB_200)
            color_object = RGBColor(r, g, b)
        speaker_color_map[speaker_name] = color_object
    return speaker_color_map[speaker_name]

def apply_html_and_phonetic_to_paragraph(paragraph, current_text, enable_phonetic):
    if not current_text.strip(): return
    
    phonetic_db = st.session_state['custom_phonetics']
    
    if not enable_phonetic:
        matches = list(HTML_CONTENT_REGEX.finditer(current_text))
        last_end = 0
        for match in matches:
            tag_text = match.group(2)
            start, end = match.span()
            if start > last_end: paragraph.add_run(current_text[last_end:start])
            run_html = paragraph.add_run(tag_text)
            run_html.font.bold = True
            run_html.font.italic = True
            last_end = end
        if last_end < len(current_text): paragraph.add_run(current_text[last_end:])
        return

    sorted_eng_keys = sorted(phonetic_db.keys(), key=len, reverse=True)
    if sorted_eng_keys:
        pattern_str = r"\b(" + "|".join([re.escape(k) for k in sorted_eng_keys]) + r")\b"
        eng_phonetic_regex = re.compile(pattern_str, re.IGNORECASE)
    else:
        eng_phonetic_regex = None

    if eng_phonetic_regex:
        matches = list(eng_phonetic_regex.finditer(current_text))
        last_end = 0
        for match in matches:
            eng_word_original = match.group(0)
            start, end = match.span()
            
            if start > last_end:
                paragraph.add_run(current_text[last_end:start])
                
            pho_text = phonetic_db.get(eng_word_original.upper(), eng_word_original)
            
            run_pho = paragraph.add_run(f"{pho_text} ")
            run_pho.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            run_eng = paragraph.add_run(f"({eng_word_original})")
            run_eng.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            last_end = end
            
        if last_end < len(current_text):
            paragraph.add_run(current_text[last_end:])
    else:
        paragraph.add_run(current_text)

def format_and_split_dialogue(document, text, enable_colors, enable_phonetic, speaker_color_map, used_colors, stats_counter, speaker_regex):
    parts = speaker_regex.split(text)
    TAB_STOP_POSITION = Inches(1.0)
    
    if len(parts) == 1:
        new_paragraph = document.add_paragraph()
        new_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        new_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        new_paragraph.add_run('\t')
        new_paragraph.paragraph_format.space_before = Pt(0)
        new_paragraph.paragraph_format.space_after = Pt(0)
        apply_html_and_phonetic_to_paragraph(new_paragraph, text, enable_phonetic)
        return

    speaker_matches = list(speaker_regex.finditer(text))
    last_processed_index = 0
    
    for i, match in enumerate(speaker_matches):
        speaker_full = match.group(0)
        speaker_name = match.group(1).strip()
        start, end = match.span()
        
        leading_content = text[last_processed_index:start].strip()
        if leading_content:
            continuation_paragraph = document.add_paragraph()
            continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
            continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
            continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
            continuation_paragraph.add_run('\t')
            continuation_paragraph.paragraph_format.space_before = Pt(0)
            continuation_paragraph.paragraph_format.space_after = Pt(0)
            apply_html_and_phonetic_to_paragraph(continuation_paragraph, leading_content, enable_phonetic)

        if speaker_name.upper() in NON_SPEAKER_PHRASES:
            content_block = text[start:]
            continuation_paragraph = document.add_paragraph()
            continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
            continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
            continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
            continuation_paragraph.add_run('\t')
            continuation_paragraph.paragraph_format.space_before = Pt(0)
            continuation_paragraph.paragraph_format.space_after = Pt(0)
            apply_html_and_phonetic_to_paragraph(continuation_paragraph, content_block, enable_phonetic)
            return

        stats_counter[speaker_name] += 1

        if i + 1 < len(speaker_matches):
            next_match_start = speaker_matches[i+1].start()
        else:
            next_match_start = len(text)
            
        content = text[end:next_match_start].strip()
        new_paragraph = document.add_paragraph()
        new_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        new_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        
        run_speaker = new_paragraph.add_run(speaker_full)
        run_speaker.font.bold = True
        
        if enable_colors:
            run_speaker.font.color.rgb = get_speaker_color(speaker_name, speaker_color_map, used_colors)
        
        if len(speaker_full) > 10:
             new_paragraph.add_run('\t\t')
        else:
             new_paragraph.add_run('\t')

        if content: apply_html_and_phonetic_to_paragraph(new_paragraph, content, enable_phonetic)
        new_paragraph.paragraph_format.space_before = Pt(0)
        new_paragraph.paragraph_format.space_after = Pt(0)
        last_processed_index = next_match_start
    
    remaining_content = text[last_processed_index:].strip()
    if remaining_content:
        continuation_paragraph = document.add_paragraph()
        continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        continuation_paragraph.add_run('\t')
        continuation_paragraph.paragraph_format.space_before = Pt(0)
        continuation_paragraph.paragraph_format.space_after = Pt(0)
        apply_html_and_phonetic_to_paragraph(continuation_paragraph, remaining_content, enable_phonetic)

# --- MAIN PROCESSING ---
def process_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic):
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_200]
    random.shuffle(used_colors)
    stats_counter = Counter()
    
    speaker_regex = build_speaker_regex(st.session_state['custom_speakers'])
    
    original_document = Document(io.BytesIO(uploaded_file.getvalue()))
    raw_paragraphs = [p for p in original_document.paragraphs]
    document = Document()
    
    title_text_raw = file_name_without_ext.upper()
    title_text = title_text_raw.replace("CONVERTED_", "").replace("FORMATTED_", "").replace("_EDIT", "").replace(" (GỐC)", "").strip()
    title_paragraph = document.add_paragraph(title_text)
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.runs[0].font.name = 'Times New Roman'
    title_paragraph.runs[0].font.size = Pt(20)
    title_paragraph.runs[0].bold = True
    
    unique_speakers = []
    for paragraph in raw_paragraphs:
        text = paragraph.text
        if text.lower().startswith("srt conversion"): continue 
        for match in speaker_regex.finditer(text):
            speaker_name = match.group(1).strip()
            if speaker_name.upper() not in NON_SPEAKER_PHRASES and speaker_name not in unique_speakers:
                unique_speakers.append(speaker_name)
            
    if unique_speakers:
        speaker_list_text = "VAI: " + ", ".join(unique_speakers)
        p = document.add_paragraph(speaker_list_text)
        p.runs[0].font.name = 'Times New Roman'
        p.runs[0].font.size = Pt(12)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
    
    document.add_paragraph()
    document.add_paragraph()
    start_index = len(document.paragraphs)

    total_paras = len(raw_paragraphs)
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, paragraph in enumerate(raw_paragraphs):
        if idx % max(1, total_paras // 10) == 0:
            progress = int((idx / total_paras) * 100)
            progress_bar.progress(progress)
            status_text.text(f"Đang phân tích dòng {idx}/{total_paras}...")

        text = paragraph.text.strip()
        if not text or text.upper() == title_text.upper(): continue
        if text.lower().startswith("srt conversion") or re.fullmatch(r"^\s*\d+\s*$", text): continue
            
        if TIMECODE_REGEX.match(text):
            new_paragraph = document.add_paragraph(text)
            new_paragraph.runs[0].font.bold = True
            new_paragraph.runs[0].font.name = 'Times New Roman'
            new_paragraph.runs[0].font.size = Pt(12)
            new_paragraph.paragraph_format.space_before = Pt(0)
            new_paragraph.paragraph_format.space_after = Pt(0)
        else:
            format_and_split_dialogue(document, text, enable_colors, enable_phonetic, speaker_color_map, used_colors, stats_counter, speaker_regex)
            
    progress_bar.progress(100)
    status_text.text("Định dạng hoàn tất!")
    time.sleep(0.5)
    progress_bar.empty()
    status_text.empty()
            
    for paragraph in document.paragraphs[start_index:]:
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            if run.font.size is None:
                run.font.size = Pt(12)
        
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    stats = {
        "total_speakers": len(unique_speakers),
        "total_lines": sum(stats_counter.values()),
        "top_speaker": stats_counter.most_common(1)[0] if stats_counter else ("Không có", 0)
    }
    
    return modified_file, stats

def clean_file_name_for_output(original_filename):
    name_without_ext = os.path.splitext(original_filename)[0]
    cleaned = re.sub(r'(CONVERTED_|FORMATTED_|\s*\(.*\)$|_edit$)', '', name_without_ext, flags=re.IGNORECASE).strip()
    return f"{cleaned}_edit.docx"

# --- SIDEBAR (THANH ĐIỀU HƯỚNG SAAS) ---
st.sidebar.markdown("### ⚡ Control Panel")

if st.sidebar.button("🔄 Reset phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_file', 'new_filename', 'stats']:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("#### 🎛️ Bật/Tắt Tính năng")
enable_colors = st.sidebar.toggle("🌈 Tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm giọng Nam", value=True, help="Tự động chèn phiên âm giọng Nam trước từ Tiếng Anh (ngoặc đơn + tô màu vàng)")

st.sidebar.markdown("---")
st.sidebar.markdown("#### 💾 Database Quản Lý Cụm Từ")

with st.sidebar.expander("🎭 Database Người nói (Whitelist)", expanded=False):
    manual_spk_input = st.text_area("Nhập thủ công:", height=80, key=f"spk_manual_{st.session_state['spk_input_key']}")
    upload_spk_file = st.file_uploader("Tải file (.txt, .docx, .xlsx)", type=['txt', 'docx', 'xlsx'], key=f"spk_uploader_{st.session_state['spk_input_key']}")
    
    if st.button("Lưu Người Nói", use_container_width=True):
        new_spks = set()
        if manual_spk_input:
            parts = re.split(r'[,\n]', manual_spk_input)
            new_spks.update([p.strip() for p in parts if p.strip()])
        if upload_spk_file:
            new_spks.update(extract_phrases_from_file(upload_spk_file, upload_spk_file.name))
            
        if new_spks:
            st.session_state['custom_speakers'].update(new_spks)
            save_json_db(SPEAKER_DB_FILE, st.session_state['custom_speakers'])
            st.session_state['spk_input_key'] += 1
            st.success(f"✅ Đã lưu {len(new_spks)} người nói!")
            time.sleep(1)
            st.rerun()

    if len(st.session_state['custom_speakers']) > 0:
        st.info(f"Đã lưu: **{len(st.session_state['custom_speakers'])}** người nói.")

with st.sidebar.expander("🚫 Database Từ nhiễu (Non-speaker)", expanded=False):
    manual_input = st.text_area("Nhập thủ công:", height=80, key=f"ns_manual_{st.session_state['ns_input_key']}")
    upload_non_speaker = st.file_uploader("Tải file (.txt, .docx, .xlsx)", type=['txt', 'docx', 'xlsx'], key=f"ns_uploader_{st.session_state['ns_input_key']}")
    
    if st.button("Lưu Từ Nhiễu", use_container_width=True):
        new_phrases = set()
        if manual_input:
            parts = re.split(r'[,\n]', manual_input)
            new_phrases.update([p.strip().upper() for p in parts if p.strip()])
        if upload_non_speaker:
            new_phrases.update([p.upper() for p in extract_phrases_from_file(upload_non_speaker, upload_non_speaker.name)])
            
        if new_phrases:
            st.session_state['custom_non_speakers'].update(new_phrases)
            save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
            st.session_state['ns_input_key'] += 1
            st.success(f"✅ Đã lưu {len(new_phrases)} từ nhiễu!")
            time.sleep(1)
            st.rerun()

    if len(st.session_state['custom_non_speakers']) > 0:
        st.info(f"Đã lưu: **{len(st.session_state['custom_non_speakers'])}** từ nhiễu.")

# --- HERO BANNER HEADER (SaaS STYLING) ---
st.markdown("""
<div class="hero-container">
    <div class="badge-pro">v2.5 Enterprise SaaS</div>
    <div class="hero-title">🎬 ScriptPro Enterprise Studio</div>
    <div class="hero-subtitle">Hệ thống xử lý kịch bản lồng tiếng, chuẩn hóa định dạng Word & tự động phiên âm giọng Nam thông minh.</div>
</div>
""", unsafe_allow_html=True)

# --- MÀN HÌNH CHÍNH TÁCH TABS ---
tab_script, tab_phonetic_db = st.tabs(["🎬 Xử lý & Biên tập Kịch bản", "📚 Kho Database Phiên Âm Giọng Nam"])

# ==========================================
# TAB 1: XỬ LÝ KỊCH BẢN
# ==========================================
with tab_script:
    col1, col2 = st.columns([1.6, 1])

    with col1:
        with st.container(border=True):
            st.markdown("### 📁 1. Tải lên kịch bản Word (.docx)")
            uploaded_file = st.file_uploader(
                "Kéo thả file .docx của bạn vào đây", 
                type=['docx'], 
                key=f"main_uploader_{st.session_state['uploader_key']}"
            )

        if uploaded_file is None:
            st.info("📌 **Vui lòng tải file kịch bản (.docx) ở trên để hiển thị công cụ biên tập.**")
        else:
            original_filename = uploaded_file.name
            file_name_without_ext = os.path.splitext(original_filename)[0] 
            st.success(f"📄 Đã nhận file thành công: **{original_filename}**")

            speaker_regex = build_speaker_regex(st.session_state['custom_speakers'])
            candidates = scan_candidate_speakers(uploaded_file, speaker_regex)

            detected_speakers = []
            detected_non_speakers = []

            for name, count in candidates.items():
                if name.upper() in NON_SPEAKER_PHRASES:
                    detected_non_speakers.append(f"{name} ({count} lần)")
                else:
                    detected_speakers.append(f"{name} ({count} lần)")

            with st.container(border=True):
                st.markdown("### 🔍 Soát Lỗi Nhận Diện Tên Người Nói")
                tab_spk, tab_non_spk = st.tabs(["🎭 Nhận diện là NGƯỜI NÓI", "🚫 Đang bị xem là TỪ NHIỄU"])
                
                with tab_spk:
                    if detected_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_speakers]))
                        to_move_to_ns = st.multiselect(
                            "Phát hiện từ nào bị nhận diện sai? Chọn để LƯU VÀO DATABASE TỪ NHIỄU:",
                            options=[name for name in candidates.keys() if name.upper() not in NON_SPEAKER_PHRASES],
                            key="select_to_ns"
                        )
                        if st.button("➡️ Đưa vào Database TỪ NHIỄU", type="secondary"):
                            if to_move_to_ns:
                                new_items = [item.upper() for item in to_move_to_ns]
                                st.session_state['custom_non_speakers'].update(new_items)
                                save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                                st.success(f"✅ Đã lưu {len(new_items)} từ vào Database Từ Nhiễu!")
                                time.sleep(1)
                                st.rerun()
                    else:
                        st.info("Chưa tìm thấy cụm từ người nói nào.")

                with tab_non_spk:
                    if detected_non_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_non_speakers]))
                        to_move_to_spk = st.multiselect(
                            "Từ nào thực ra là NGƯỜI NÓI? Chọn để LƯU VÀO DATABASE NGƯỜI NÓI:",
                            options=[name for name in candidates.keys() if name.upper() in NON_SPEAKER_PHRASES],
                            key="select_to_spk"
                        )
                        if st.button("➡️ Đưa vào Database NGƯỜI NÓI", type="secondary"):
                            if to_move_to_spk:
                                st.session_state['custom_speakers'].update(to_move_to_spk)
                                save_json_db(SPEAKER_DB_FILE, st.session_state['custom_speakers'])
                                for item in to_move_to_spk: st.session_state['custom_non_speakers'].discard(item.upper())
                                save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                                st.success(f"✅ Đã lưu {len(to_move_to_spk)} tên vào Database Người Nói!")
                                time.sleep(1)
                                st.rerun()
                    else:
                        st.info("Không có cụm từ nào bị loại vào danh sách từ nhiễu.")

            # --- QUẢN LÝ PHIÊN ÂM KỊCH BẢN HIỆN TẠI ---
            with st.container(border=True):
                st.markdown("### 🗣️ Từ Tiếng Anh Xuất Hiện Trong Kịch Bản")
                st.caption("Quét và điều chỉnh phiên âm riêng cho kịch bản này (Đã qua bộ lọc thông minh):")

                detected_eng_words = scan_english_words_in_dialogue(uploaded_file, speaker_regex)

                if detected_eng_words:
                    table_data = []
                    for word in detected_eng_words:
                        current_pho = st.session_state['custom_phonetics'].get(word.upper(), word)
                        table_data.append({
                            "Từ Tiếng Anh": word,
                            "Phiên âm hiện tại": current_pho,
                            "Đề xuất chỉnh sửa của bạn": current_pho,
                            "Nạp vào Database": True
                        })

                    df_eng = pd.DataFrame(table_data)

                    edited_df = st.data_editor(
                        df_eng,
                        column_config={
                            "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh gốc", disabled=True),
                            "Phiên âm hiện tại": st.column_config.TextColumn("Phiên âm gán hiện tại", disabled=True),
                            "Đề xuất chỉnh sửa của bạn": st.column_config.TextColumn("Đề xuất phiên âm mới"),
                            "Nạp vào Database": st.column_config.CheckboxColumn("Lưu Database?", default=True)
                        },
                        disabled=["Từ Tiếng Anh", "Phiên âm hiện tại"],
                        hide_index=True,
                        use_container_width=True,
                        key="phonetic_script_table"
                    )

                    if st.button("💾 Nạp chỉnh sửa kịch bản này vào Database Phiên Âm", type="secondary", use_container_width=True):
                        updated_count = 0
                        for _, row in edited_df.iterrows():
                            if row["Nạp vào Database"]:
                                eng_k = str(row["Từ Tiếng Anh"]).upper().strip()
                                pho_v = str(row["Đề xuất chỉnh sửa của bạn"]).strip()
                                if pho_v:
                                    st.session_state['custom_phonetics'][eng_k] = pho_v
                                    updated_count += 1
                        
                        save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                        st.success(f"✅ Đã cập nhật {updated_count} từ phiên âm vào Database!")
                        time.sleep(1)
                        st.rerun()
                else:
                    st.info("Không phát hiện từ Tiếng Anh / Tên riêng nước ngoài nào trong phần lời thoại kịch bản này.")

            st.markdown("---")
            if st.button("✨ 2. BẮT ĐẦU ĐỊNH DẠNG TỰ ĐỘNG", use_container_width=True, type="primary"):
                try:
                    modified_file_io, stats = process_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic)
                    new_filename = clean_file_name_for_output(original_filename)
                    
                    st.session_state['processed_file'] = modified_file_io
                    st.session_state['new_filename'] = new_filename
                    st.session_state['stats'] = stats
                    
                except Exception as e:
                    st.error(f"Đã có lỗi xảy ra: {e}")

            if 'processed_file' in st.session_state:
                st.markdown("---")
                st.download_button(
                    label="⬇️ 3. TẢI FILE KỊCH BẢN ĐÃ CHUẨN HÓA (.DOCX)",
                    data=st.session_state['processed_file'],
                    file_name=st.session_state['new_filename'],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
                st.balloons()

    with col2:
        st.markdown("### 📊 SaaS Analytics")
        if 'stats' in st.session_state:
            stats = st.session_state['stats']
            
            st.markdown(f"""
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">🎭 Tổng số Nhân vật</div>
                <div class="metric-value">{stats["total_speakers"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">💬 Tổng số Câu thoại</div>
                <div class="metric-value">{stats["total_lines"]}</div>
            </div>
            """, unsafe_allow_html=True)
            
            top_name, top_count = stats["top_speaker"]
            st.info(f"👑 **Nhân vật thoại nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
        else:
            st.info("Bảng phân tích dữ liệu kịch bản sẽ xuất hiện tại đây sau khi bạn xử lý file.")

# ==========================================
# TAB 2: KHO DATABASE PHIÊN ÂM GIỌNG NAM (TỔNG HỢP)
# ==========================================
with tab_phonetic_db:
    with st.container(border=True):
        st.subheader("📚 Từ Điển Phiên Âm Giọng Nam (Global Database)")
        st.markdown("Nơi quản lý toàn bộ kho từ vựng Tiếng Anh và các bản phiên âm giọng Nam được lưu trữ lâu dài trên hệ thống.")
        
        # 1. Thêm thủ công từ mới
        st.markdown("#### ➕ Bổ sung từ phiên âm mới vào Kho")
        c1, c2, c3 = st.columns([2, 2, 1.2])
        with c1:
            tab_add_eng = st.text_input("Từ Tiếng Anh gốc:", placeholder="VD: Burger", key=f"tab_add_eng_{st.session_state['pho_input_key']}")
        with c2:
            tab_add_pho = st.text_input("Phiên âm giọng Nam:", placeholder="VD: Bơ-gơ", key=f"tab_add_pho_{st.session_state['pho_input_key']}")
        with c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm vào Database", use_container_width=True, type="primary"):
                if tab_add_eng and tab_add_pho:
                    k = tab_add_eng.upper().strip()
                    v = tab_add_pho.strip()
                    st.session_state['custom_phonetics'][k] = v
                    save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                    st.session_state['pho_input_key'] += 1
                    st.success(f"✅ Đã thêm thành công: `{tab_add_eng}` ➔ `{tab_add_pho}`")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("Vui lòng điền đủ 2 ô!")

    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 📑 Danh sách toàn bộ Từ phiên âm đã lưu")

        search_query = st.text_input("🔍 Tìm kiếm từ Tiếng Anh hoặc Từ phiên âm:", placeholder="Gõ từ cần tìm ở đây...").strip().upper()

        all_phonetics_dict = st.session_state['custom_phonetics']
        
        if search_query:
            filtered_dict = {
                k: v for k, v in all_phonetics_dict.items() 
                if search_query in k or search_query in v.upper()
            }
        else:
            filtered_dict = all_phonetics_dict

        if filtered_dict:
            db_table_data = []
            for eng_key, pho_val in sorted(filtered_dict.items()):
                db_table_data.append({
                    "Từ Tiếng Anh": eng_key,
                    "Phiên âm giọng Nam": pho_val,
                    "Xóa khỏi Database": False
                })

            df_db = pd.DataFrame(db_table_data)

            st.caption(f"Đang hiển thị **{len(df_db)}** từ phiên âm trong hệ thống:")

            edited_db_df = st.data_editor(
                df_db,
                column_config={
                    "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh gốc (In hoa)", disabled=True),
                    "Phiên âm giọng Nam": st.column_config.TextColumn("Phiên âm giọng Nam (Sửa trực tiếp tại đây)"),
                    "Xóa khỏi Database": st.column_config.CheckboxColumn("Tích chọn để XÓA")
                },
                disabled=["Từ Tiếng Anh"],
                hide_index=True,
                use_container_width=True,
                key="global_phonetic_db_editor"
            )

            if st.button("💾 LƯU TOÀN BỘ CẬP NHẬT TRONG BẢNG", type="primary", use_container_width=True):
                new_db = {}
                deleted_count = 0
                
                if search_query:
                    for k, v in all_phonetics_dict.items():
                        if k not in filtered_dict:
                            new_db[k] = v

                for _, row in edited_db_df.iterrows():
                    eng_k = str(row["Từ Tiếng Anh"]).upper().strip()
                    pho_v = str(row["Phiên âm giọng Nam"]).strip()
                    is_delete = row["Xóa khỏi Database"]
                    
                    if is_delete:
                        deleted_count += 1
                    else:
                        if pho_v:
                            new_db[eng_k] = pho_v

                st.session_state['custom_phonetics'] = new_db
                save_json_db(PHONETIC_DB_FILE, new_db)
                st.success(f"✅ Đã lưu cập nhật thành công! (Đã xóa {deleted_count} từ)")
                time.sleep(1)
                st.rerun()
        else:
            st.info("Không tìm thấy từ phiên âm nào khớp với từ khóa tìm kiếm.")

# --- FOOTER SAAS TÙY CHỈNH ---
st.markdown("""
<div class="saas-footer">
    ScriptPro Enterprise Edition • Designed for Mai Han Team
</div>
""", unsafe_allow_html=True)
