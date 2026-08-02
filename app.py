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
from gtts import gTTS
import zipfile
from datetime import datetime

# --- CẤU HÌNH TRANG CHỦ STREAMLIT ---
st.set_page_config(
    page_title="ScriptPro Enterprise - Subtitle & Script Editor",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SIDEBAR CONTROL PANEL: CÀI ĐẶT CỦA BẠN & CHẾ ĐỘ MÀN HÌNH ---
st.sidebar.markdown("### ⚡ Control Panel")

ui_theme_choice = st.sidebar.radio(
    "Lựa chọn Skin hiển thị:",
    options=["Mai Han Standard (Mặc định)", "Enterprise Pro (Tối ưu tương phản)"],
    index=0,
    help="Chế độ 'Mai Han Standard' giữ nguyên 100% giao diện truyền thống. Chế độ 'Enterprise Pro' mang lại phong cách Studio hiện đại, rõ nét và tương phản cao."
)

if st.sidebar.button("🔄 Reset phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_docx', 'processed_ass', 'processed_srt', 'actor_zip', 'r_processed_docx', 'r_processed_ass', 'r_processed_srt', 'r_actor_zip']:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.session_state['resync_uploader_key'] += 1
    st.rerun()

# --- DYNAMIC CSS INJECTION THEO CHẾ ĐỘ ĐƯỢC CHỌN ---
if "Enterprise Pro" in ui_theme_choice:
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            color: #0F172A;
        }
        
        .hero-container {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            padding: 2.2rem 2rem;
            border-radius: 14px;
            color: #FFFFFF;
            margin-bottom: 1.8rem;
            border-left: 6px solid #38BDF8;
            box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        }
        .hero-title { font-size: 2.3rem; font-weight: 800; margin: 0; letter-spacing: -0.02em; color: #FFFFFF; }
        .hero-subtitle { font-size: 1.05rem; color: #94A3B8; margin-top: 0.4rem; font-weight: 400; }
        .badge-pro {
            background-color: #0284C7; color: #FFFFFF; padding: 4px 12px;
            border-radius: 6px; font-size: 0.75rem; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.05em; display: inline-block; margin-bottom: 0.6rem;
        }
        
        [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 12px !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05) !important;
        }
        
        .metric-card {
            background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            transition: all 0.2s ease-in-out;
        }
        .metric-card:hover { border-color: #0284C7; transform: translateY(-2px); box-shadow: 0 8px 15px rgba(0,0,0,0.06); }
        .metric-label { font-size: 0.8rem; color: #475569; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; }
        .metric-value { font-size: 1.8rem; font-weight: 800; color: #0F172A; margin-top: 0.2rem; }
        
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px; border-bottom: 2px solid #E2E8F0; padding-bottom: 2px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 48px; border-radius: 8px 8px 0 0; font-weight: 700;
            padding: 0 20px; color: #475569; font-size: 0.95rem;
        }
        .stTabs [aria-selected="true"] {
            background-color: #F1F5F9 !important; color: #0284C7 !important;
            border-bottom: 3px solid #0284C7 !important;
        }
        
        .qc-card-warning {
            background-color: #FEF2F2; border-left: 5px solid #DC2626;
            color: #991B1B; padding: 12px 16px; border-radius: 8px;
            margin-bottom: 10px; font-size: 0.92rem; font-weight: 500;
        }
        .saas-footer {
            text-align: center; padding: 2rem 0; color: #64748B;
            font-size: 0.85rem; border-top: 1px solid #E2E8F0; margin-top: 3rem; font-weight: 500;
        }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .hero-container {
            background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%);
            padding: 2.5rem 2rem; border-radius: 16px; color: white; margin-bottom: 2rem;
            box-shadow: 0 10px 25px -5px rgba(79, 70, 229, 0.3);
        }
        .hero-title { font-size: 2.4rem; font-weight: 800; margin: 0; letter-spacing: -0.02em; }
        .hero-subtitle { font-size: 1.05rem; opacity: 0.9; margin-top: 0.5rem; }
        .badge-pro {
            background-color: rgba(255, 255, 255, 0.2); backdrop-filter: blur(8px);
            padding: 4px 12px; border-radius: 9999px; font-size: 0.8rem; font-weight: 600;
            text-transform: uppercase; letter-spacing: 0.05em; display: inline-block; margin-bottom: 0.8rem;
        }
        .metric-card {
            background: white; border: 1px solid #E2E8F0; border-radius: 12px; padding: 1.25rem;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05); transition: transform 0.2s, box-shadow 0.2s;
        }
        .metric-card:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05); }
        .metric-label { font-size: 0.85rem; color: #64748B; font-weight: 500; text-transform: uppercase; }
        .metric-value { font-size: 1.8rem; font-weight: 700; color: #0F172A; margin-top: 0.25rem; }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: 2px solid #E2E8F0; }
        .stTabs [data-baseweb="tab"] { height: 50px; border-radius: 8px 8px 0 0; font-weight: 600; padding: 0 20px; }
        .qc-card-warning { background-color: #FEF2F2; border-left: 4px solid #EF4444; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; font-size: 0.9rem; }
        .saas-footer { text-align: center; padding: 2rem 0 1rem 0; color: #94A3B8; font-size: 0.85rem; border-top: 1px solid #E2E8F0; margin-top: 3rem; }
    </style>
    """, unsafe_allow_html=True)

# --- HÀM CÔNG CỤ CHUYỂN ĐỔI SRT ⇄ DOCX (PURE PYTHON) ---
TARGET_FONT = 'Times New Roman'
TARGET_SIZE = Pt(12)

def set_font_and_size(run, font_name, font_size):
    run.font.name = font_name
    run.font.size = font_size

def process_srt_to_docx(uploaded_file, file_name_without_ext):
    srt_content = uploaded_file.getvalue().decode('utf-8', errors='ignore')
    blocks = re.split(r'\n\s*\n', srt_content.strip())
    document = Document()
    document.add_heading(f"SRT Conversion: {file_name_without_ext}", level=1)

    for block in blocks:
        lines = [l.strip() for l in block.strip().split('\n') if l.strip()]
        if not lines: continue
        
        idx_str = ""; tc_str = ""; text_lines = []
        
        if lines[0].isdigit():
            idx_str = lines[0]
            if len(lines) > 1 and "-->" in lines[1]:
                tc_str = lines[1]
                text_lines = lines[2:]
            else: text_lines = lines[1:]
        elif "-->" in lines[0]:
            tc_str = lines[0]
            text_lines = lines[1:]
        else: text_lines = lines

        if idx_str:
            p_index = document.add_paragraph(idx_str)
            set_font_and_size(p_index.runs[0], TARGET_FONT, TARGET_SIZE)
            p_index.paragraph_format.space_after = Pt(0)

        if tc_str:
            p_timecode = document.add_paragraph(tc_str)
            set_font_and_size(p_timecode.runs[0], TARGET_FONT, TARGET_SIZE)
            p_timecode.paragraph_format.space_after = Pt(0)
            
        if text_lines:
            clean_text = "\n".join([re.sub(r'<[^>]*>', '', l) for l in text_lines])
            p_content = document.add_paragraph(clean_text)
            set_font_and_size(p_content.runs[0], TARGET_FONT, TARGET_SIZE)
            p_content.paragraph_format.space_after = Pt(12)

    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    return modified_file

def process_docx_to_srt(uploaded_file):
    document = Document(uploaded_file)
    lines = [p.text.strip() for p in document.paragraphs if p.text.strip() != ""]
    
    srt_content = ""
    timecode_pattern = re.compile(r'\d{2}:\d{2}:\d{2}[,.]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[,.]\d{3}')
    
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.isdigit() and i + 1 < len(lines) and timecode_pattern.search(lines[i+1]):
            if srt_content != "": srt_content += "\n"
            srt_content += line + "\n"
            srt_content += lines[i+1] + "\n"
            i += 2 
            while i < len(lines):
                if lines[i].isdigit() and i + 1 < len(lines) and timecode_pattern.search(lines[i+1]):
                    break 
                else:
                    srt_content += lines[i] + "\n"
                    i += 1
        else: i += 1
            
    return srt_content.strip().encode('utf-8')

# --- SRT TO EXCEL CONVERTER MODULE (ĐỒNG BỘ 100% DATABASE CỤM TỪ STUDIO) ---
EXCEL_COLOR_PALETTE = [
    'background-color: #ADD8E6; color: #000000',
    'background-color: #90EE90; color: #000000',
    'background-color: #FFB6C1; color: #000000',
    'background-color: #FFFFE0; color: #000000',
    'background-color: #DDA0DD; color: #000000',
    'background-color: #AFEEEE; color: #000000',
    'background-color: #F0E68C; color: #000000',
    'background-color: #FFA07A; color: #000000',
    'background-color: #E0FFFF; color: #000000',
    'background-color: #F5F5DC; color: #000000',
    'background-color: #2F4F4F; color: #FFFFFF',
    'background-color: #191970; color: #FFFFFF',
    'background-color: #006400; color: #FFFFFF',
    'background-color: #800000; color: #FFFFFF',
    'background-color: #4B0082; color: #FFFFFF',
    'background-color: #556B2F; color: #FFFFFF',
    'background-color: #8B4513; color: #FFFFFF',
    'background-color: #36454F; color: #FFFFFF',
]

def clean_dialogue_text_for_excel(text):
    text = re.sub(r'<i[^>]*>(.*?)</i[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<b[^>]*>(.*?)</b[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<u[^>]*>(.*?)</u[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r'<[^>]*>', '', text, flags=re.DOTALL)
    return re.sub(r'\s+', ' ', text).strip()

def parse_srt_to_dataframe(srt_content):
    data = []
    blocks = re.split(r'\n\s*\n', srt_content.strip())
    last_known_speaker = "Unknown"

    speaker_regex = build_speaker_regex(st.session_state.get('custom_speakers', set()))

    def append_row_and_update_state(t_start, t_end, speaker, dialogue):
        nonlocal last_known_speaker
        data.append([t_start, t_end, speaker, clean_dialogue_text_for_excel(dialogue)])
        last_known_speaker = speaker 

    for block in blocks:
        lines = block.strip().split('\n')
        if len(lines) < 3: continue

        time_line = lines[1].strip()
        time_match = re.match(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})', time_line)
        if not time_match: continue

        time_start = time_match.group(1) 
        time_end = time_match.group(2)   
        dialogue_lines = lines[2:]
        current_dialogue = ""
        block_initial_speaker = last_known_speaker

        for line in dialogue_lines:
            line = line.strip()
            if not line: continue
            segments = speaker_regex.split(line)
            i = 0
            while i < len(segments):
                segment = segments[i].strip()
                i += 1
                if not segment: continue

                if segment.endswith(':') and len(segment) > 1:
                    speaker_tag = segment[:-1].strip()
                    if is_valid_speaker_name(speaker_tag):
                        if current_dialogue:
                            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
                            append_row_and_update_state(time_start, time_end, speaker_to_use, current_dialogue)
                            current_dialogue = ""
                        
                        speaker = speaker_tag
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        if dialogue_segment:
                            append_row_and_update_state(time_start, time_end, speaker, dialogue_segment)
                        if block_initial_speaker == last_known_speaker:
                             block_initial_speaker = speaker
                    else:
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        recombined_text = segment + " " + dialogue_segment
                        current_dialogue = (current_dialogue + " " + recombined_text) if current_dialogue else recombined_text
                else:
                    current_dialogue = (current_dialogue + " " + segment) if current_dialogue else segment

        if current_dialogue:
            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
            append_row_and_update_state(time_start, time_end, speaker_to_use, current_dialogue)

    return pd.DataFrame(data, columns=['Start', 'End', 'Speaker', 'Dialogue'])

def apply_excel_styles(df):
    unique_speakers = df['Speaker'].unique()
    color_map = {
        speaker: EXCEL_COLOR_PALETTE[i % len(EXCEL_COLOR_PALETTE)]
        for i, speaker in enumerate(unique_speakers)
    }
    def highlight_speaker(row):
        color_style = color_map.get(row['Speaker'], 'background-color: #FFFFFF; color: #000000')
        return [color_style] * len(row)
    try:
        return df.style.apply(highlight_speaker, axis=1)
    except Exception:
        return df

# --- HÀM TẠO ÂM THANH PHÁT ÂM TIẾNG ANH CHUẨN ---
def generate_english_audio(text_to_speak, accent='com'):
    try:
        tts = gTTS(text=text_to_speak, lang='en', tld=accent)
        fp = io.BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        return fp
    except Exception as e:
        st.error(f"Không thể tải âm thanh: {e}")
        return None

# --- HÀM ĐỌC / GHI DATABASE DỮ LIỆU CỤC BỘ ---
NON_SPEAKER_DB_FILE = "custom_non_speakers.json"
SPEAKER_DB_FILE = "custom_speakers.json"
PHONETIC_DB_FILE = "custom_phonetics.json"
CAST_DB_FILE = "custom_cast_mapping.json"
TRACKER_DB_FILE = "dubbing_tracker.json"
RATES_DB_FILE = "payroll_rates.json"

DEFAULT_CAST_MAPPING = {
    "BRI": "TRÚC", "CHASE": "THIỆN", "PRESTON": "KHÁNH", "SCOTT": "THÔNG",
    "KEELEY": "TÚ", "DAKA": "QUANG", "STEPHEN": "QUANG", "VINCE": "QUANG",
    "JOSH": "THÔNG", "YOMI": "TÚ", "LARRY": "A.TRUNG", "CHLOE": "TÚ",
    "STEVEN": "QUANG (HOẶC THÔNG)", "LOGAN": "THÔNG", "NATHAN": "A.TRUNG",
    "RILEY": "PHỤNG", "ZHONG": "THÔNG", "MICHELLE (YOUTUBER)": "TÚ",
    "BEN AZELART": "QUANG", "ZONG": "THÔNG", "BA BRI": "THÔNG", "TYLER": "PHÁT",
    "COBY": "THIỆN", "CORY": "KHÁNH", "SPARKY": "A.TRUNG", "GARRETT": "C.DŨNG",
    "CODY": "CƯỜNG", "BETHANY": "TRÚC", "KRISTIN": "PHỤNG", "AMY": "TÚ",
    "ALLISON": "TÚ", "AUBREY": "PHỤNG", "TY JACKSON": "THÔNG", "HLV": "CƯỜNG",
    "JACKSON OLSON": "THÔNG", "DANNY": "QUANG", "KYLE": "QUANG", "BILL": "HITA", "TIM": "HITA"
}

DEFAULT_SOUTH_VIETNAM_PHONETICS = {
    "MCDONALD": "Mắc-đô-nồ", "MCDONALD'S": "Mắc-đô-nồ", "LONDON": "Luân Đôn", "TNT": "Ti-èn-ti",
    "NOOB": "Núp", "GOLEM": "Gô-lùm", "GOOGLE": "Gú-gồ", "FACEBOOK": "Phây-sbúc",
    "KFC": "Ke-ép-xi", "STARBUCKS": "Xì-ta-bắc", "APPLE": "Ép-pồ", "BURGER": "Bơ-gơ",
    "PARIS": "Ba Lê", "WASHINGTON": "Hoa Thịnh Đốn", "CALIFORNIA": "Cả-li", "PRO": "Prồ",
    "LAG": "Lác", "BUFF": "Bớp", "VIP": "Vi-ai-pi", "FBI": "Ép-bi-ai"
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
    "nat", "nau", "nay", "ne", "nen", "neo", "ni", "nia", "niem", "nien", "nieu", "ninh", "no", "noc", 
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
    if not clean_w or len(clean_w) <= 1 or clean_w.isdigit(): return False
    lower_w = clean_w.lower()
    if clean_w.upper() in st.session_state['custom_phonetics']: return True
    if lower_w in VN_SYLLABLES: return False
    if any(char in lower_w for char in ['f', 'j', 'w', 'z']): return True
        
    eng_patterns = [
        r"(bb|cc|dd|ff|gg|ll|mm|nn|pp|rr|ss|tt|zz)",
        r"(sh|ck|th|wh|ph|gh)",
        r"(tion|ment|ness|less|ing|ed|able|ible|ally|ce|ge|ck)$",
        r"^([A-Z]{2,})$"
    ]
    for pattern in eng_patterns:
        if re.search(pattern, clean_w): return True
    if clean_w[0].isupper() and lower_w not in VN_SYLLABLES: return True
    return False

def load_json_db(filepath, default_data=None):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                return set(data) if isinstance(data, list) and isinstance(default_data, set) else data
        except Exception: pass
    return default_data if default_data is not None else (set() if not isinstance(default_data, dict) else {})

def save_json_db(filepath, data_container):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            if isinstance(data_container, set): json.dump(list(data_container), f, ensure_ascii=False, indent=2)
            else: json.dump(data_container, f, ensure_ascii=False, indent=2)
    except Exception as e: st.error(f"Không thể lưu vào Database: {e}")

# Khởi tạo Session State
if 'uploader_key' not in st.session_state: st.session_state['uploader_key'] = 0
if 'resync_uploader_key' not in st.session_state: st.session_state['resync_uploader_key'] = 0
if 'spk_input_key' not in st.session_state: st.session_state['spk_input_key'] = 0
if 'ns_input_key' not in st.session_state: st.session_state['ns_input_key'] = 0
if 'pho_input_key' not in st.session_state: st.session_state['pho_input_key'] = 0
if 'cast_input_key' not in st.session_state: st.session_state['cast_input_key'] = 0

if 'custom_non_speakers' not in st.session_state: st.session_state['custom_non_speakers'] = load_json_db(NON_SPEAKER_DB_FILE, set())
if 'custom_speakers' not in st.session_state: st.session_state['custom_speakers'] = load_json_db(SPEAKER_DB_FILE, set())

if 'custom_phonetics' not in st.session_state:
    loaded_pho = load_json_db(PHONETIC_DB_FILE, DEFAULT_SOUTH_VIETNAM_PHONETICS)
    merged_pho = {**DEFAULT_SOUTH_VIETNAM_PHONETICS, **loaded_pho}
    st.session_state['custom_phonetics'] = merged_pho

if 'custom_cast_mapping' not in st.session_state:
    loaded_cast = load_json_db(CAST_DB_FILE, DEFAULT_CAST_MAPPING)
    merged_cast = {**DEFAULT_CAST_MAPPING, **loaded_cast}
    st.session_state['custom_cast_mapping'] = merged_cast

if 'dubbing_tracker' not in st.session_state: st.session_state['dubbing_tracker'] = load_json_db(TRACKER_DB_FILE, [])

if 'payroll_rates' not in st.session_state:
    default_rates = {"mode": "minute", "unit_rate": 30000}
    st.session_state['payroll_rates'] = load_json_db(RATES_DB_FILE, default_rates)

# --- SIDEBAR TÙY CHỌN BẬT/TẮT TÍNH NĂNG ---
st.sidebar.markdown("---")
st.sidebar.markdown("#### 🎛️ Bật/Tắt Tính năng")
enable_colors = st.sidebar.toggle("🌈 Tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm giọng Nam", value=True, help="Tự động chèn phiên âm giọng Nam trước từ Tiếng Anh (ngoặc đơn + tô màu vàng)")
enable_cast = st.sidebar.toggle("🎭 Phân vai lồng tiếng", value=True, help="Hiển thị thông tin diễn viên lồng tiếng ở đầu trang và lần xuất hiện đầu tiên của nhân vật")

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
            st.success(f"✅ Đã lưu {len(new_spks)} người nói!"); time.sleep(1); st.rerun()

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
            st.success(f"✅ Đã lưu {len(new_phrases)} từ nhiễu!"); time.sleep(1); st.rerun()

# --- HERO BANNER HEADER ---
st.markdown(f"""
<div class="hero-container">
    <div class="badge-pro">{ui_theme_choice}</div>
    <div class="hero-title">🎬 ScriptPro Enterprise Studio</div>
    <div class="hero-subtitle">Hệ thống xử lý kịch bản lồng tiếng, chuẩn hóa định dạng Word, phân vai & báo cáo thù lao cá nhân thông minh.</div>
</div>
""", unsafe_allow_html=True)

# --- MÀN HÌNH CHÍNH TÁCH 6 TABS ---
tab_script, tab_resync, tab_dub_tracker, tab_cast_db, tab_phonetic_db, tab_tools = st.tabs([
    "🎬 Xử lý Kịch bản Gốc", 
    "🔄 Re-Sync Kịch Bản Biên Tập",
    "📋 Theo dõi & Báo cáo Lương",
    "🎭 Bảng Phân Vai Lồng Tiếng", 
    "📚 Kho Database Phiên Âm Giọng Nam",
    "🧰 Bộ Công Cụ Chuyển Đổi"
])

# ==========================================
# TAB 1: XỬ LÝ KỊCH BẢN GỐC
# ==========================================
with tab_script:
    col1, col2 = st.columns([1.6, 1])

    with col1:
        with st.container(border=True):
            st.markdown("### 📁 Tải lên file Kịch bản Word gốc (.docx)")
            uploaded_file = st.file_uploader(
                "Kéo thả file .docx gốc của bạn vào đây", 
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

            detected_speakers_names = [name for name in candidates.keys() if name.upper() not in NON_SPEAKER_PHRASES]
            detected_non_speakers_names = [name for name in candidates.keys() if name.upper() in NON_SPEAKER_PHRASES]

            detected_speakers = [f"{name} ({candidates[name]} lần)" for name in detected_speakers_names]
            detected_non_speakers = [f"{name} ({candidates[name]} lần)" for name in detected_non_speakers_names]

            with st.container(border=True):
                st.markdown("### 🎭 Phân Vai Lồng Tiếng Cho Kịch Bản Hiện Tại")
                st.caption("Xem và gán người lồng tiếng Việt cho từng nhân vật trong file kịch bản này:")

                if detected_speakers_names:
                    cast_table_data = []
                    for spk_name in detected_speakers_names:
                        current_actor = st.session_state['custom_cast_mapping'].get(spk_name.upper(), "")
                        cast_table_data.append({
                            "Nhân vật (Tiếng Anh)": spk_name,
                            "Diễn viên Lồng tiếng (Tiếng Việt)": current_actor,
                            "Nạp vào Database": True
                        })

                    df_cast = pd.DataFrame(cast_table_data)

                    edited_cast_df = st.data_editor(
                        df_cast,
                        column_config={
                            "Nhân vật (Tiếng Anh)": st.column_config.TextColumn("Nhân vật (Kịch bản gốc)", disabled=True),
                            "Diễn viên Lồng tiếng (Tiếng Việt)": st.column_config.TextColumn("Diễn viên lồng tiếng (Sửa trực tiếp)"),
                            "Nạp vào Database": st.column_config.CheckboxColumn("Lưu Database?", default=True)
                        },
                        disabled=["Nhân vật (Tiếng Anh)"],
                        hide_index=True,
                        use_container_width=True,
                        key="script_cast_editor_table"
                    )

                    if st.button("💾 Lưu Bảng Phân Vai Kịch Bản Này Vào Database", type="secondary", use_container_width=True):
                        updated_cast_count = 0
                        for _, row in edited_cast_df.iterrows():
                            if row["Nạp vào Database"]:
                                spk_k = str(row["Nhân vật (Tiếng Anh)"]).upper().strip()
                                act_v = str(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).strip().upper()
                                if act_v:
                                    st.session_state['custom_cast_mapping'][spk_k] = act_v
                                    updated_cast_count += 1
                        save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                        st.success(f"✅ Đã lưu phân vai cho {updated_cast_count} nhân vật vào Database!")
                        time.sleep(1); st.rerun()

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
                                time.sleep(1); st.rerun()
                    else: st.info("Chưa tìm thấy cụm từ người nói nào.")

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
                                time.sleep(1); st.rerun()
                    else: st.info("Không có cụm từ nào bị loại vào danh sách từ nhiễu.")

            with st.container(border=True):
                st.markdown("### 🗣️ Từ Tiếng Anh Xuất Hiện Trong Kịch Bản")
                st.caption("Quét và điều chỉnh phiên âm riêng cho kịch bản này (Đã qua bộ lọc thông minh):")

                detected_eng_words = scan_english_words_in_dialogue(uploaded_file, speaker_regex)

                if detected_eng_words:
                    st.markdown("#### 🔊 Trình nghe phát âm chuẩn giọng bản xứ (Google US/UK)")
                    col_listen1, col_listen2, col_listen3 = st.columns([2.5, 1.5, 1.5])
                    
                    with col_listen1:
                        word_to_listen = st.selectbox("Chọn từ cần nghe phát âm:", options=detected_eng_words, key="script_listen_select")
                    with col_listen2:
                        accent_choice = st.radio("Giọng phát âm:", options=["Giọng Mỹ (US)", "Giọng Anh (UK)"], horizontal=True, key="script_accent_radio")
                    with col_listen3:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        listen_btn = st.button("🔊 Nghe phát âm", type="secondary", use_container_width=True)

                    if listen_btn and word_to_listen:
                        tld = 'com' if "Mỹ" in accent_choice else 'co.uk'
                        audio_fp = generate_english_audio(word_to_listen, accent=tld)
                        if audio_fp: st.audio(audio_fp, format="audio/mp3", autoplay=True)

                    st.markdown("---")
                    
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
                        time.sleep(1); st.rerun()
                else: st.info("Không phát hiện từ Tiếng Anh / Tên riêng nước ngoài nào trong phần lời thoại kịch bản này.")

            st.markdown("---")
            if st.button("✨ 2. BẮT ĐẦU ĐỊNH DẠNG TỰ ĐỘNG", use_container_width=True, type="primary"):
                try:
                    modified_docx, ass_f, srt_f, act_zip, stats = process_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic, enable_cast, is_resync=False)
                    
                    st.session_state['processed_docx'] = modified_docx
                    st.session_state['processed_ass'] = ass_f
                    st.session_state['processed_srt'] = srt_f
                    st.session_state['actor_zip'] = act_zip
                    st.session_state['docx_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".docx")
                    st.session_state['ass_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".ass")
                    st.session_state['srt_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".srt")
                    st.session_state['zip_name'] = clean_file_name_for_output(original_filename, tag="_KichBan_TachVai", ext=".zip")
                    st.session_state['stats'] = stats
                    
                except Exception as e: st.error(f"Đã có lỗi xảy ra: {e}")

            if 'processed_docx' in st.session_state:
                st.markdown("---")
                qc_warns = st.session_state['stats'].get("qc_warnings", [])
                if qc_warns:
                    with st.expander("🔍 BÁO CÁO CẢNH BÁO CHẤT LƯỢNG (QC & CPS CHECKER)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(qc_warns) > 10: st.info(f"...và thêm {len(qc_warns)-10} cảnh báo khác.")
                
                st.markdown("### ⬇️ 3. TẢI VỀ CÁC FILE ĐÃ XỬ LÝ HOÀN HẢO")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                with col_dl1:
                    st.download_button(
                        label="📄 FILE WORD KỊCH BẢN (.DOCX)",
                        data=st.session_state['processed_docx'],
                        file_name=st.session_state['docx_name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True
                    )
                with col_dl2:
                    st.download_button(
                        label="🎬 PHỤ ĐỀ CAO CẤP (.ASS)",
                        data=st.session_state['processed_ass'],
                        file_name=st.session_state['ass_name'],
                        mime="text/plain", use_container_width=True
                    )
                with col_dl3:
                    st.download_button(
                        label="📝 PHỤ ĐỀ CHUẨN (.SRT)",
                        data=st.session_state['processed_srt'],
                        file_name=st.session_state['srt_name'],
                        mime="text/plain", use_container_width=True
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ KỊCH BẢN TÁCH VAI RIÊNG CHO PHÒNG THU LỒNG TIẾNG")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                act_map = st.session_state['stats'].get("actor_dialogue_map", {})
                if act_map:
                    col_act1, col_act2 = st.columns([2, 1])
                    with col_act1:
                        selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(act_map.keys()))
                        if selected_actor:
                            act_buf = generate_actor_docx(st.session_state['stats']['video_title'], selected_actor, act_map[selected_actor])
                            st.download_button(
                                label=f"⬇️ TẢI FILE WORD RIÊNG CHO {selected_actor} (.DOCX)",
                                data=act_buf,
                                file_name=f"KichBan_{selected_actor}_{st.session_state['stats']['video_title']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    with col_act2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 TẢI TRỌN BỘ KỊCH BẢN TÁCH VAI (.ZIP)",
                            data=st.session_state['actor_zip'],
                            file_name=st.session_state['zip_name'],
                            mime="application/zip", type="secondary", use_container_width=True
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
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">⏱️ Độ dài Video</div>
                <div class="metric-value">{stats["video_duration_min"]} phút</div>
            </div>
            """, unsafe_allow_html=True)
            top_name, top_count = stats["top_speaker"]
            st.info(f"👑 **Nhân vật thoại nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
        else: st.info("Bảng phân tích dữ liệu kịch bản sẽ xuất hiện tại đây sau khi bạn xử lý file.")

# ==========================================
# TAB 2: RE-SYNC KỊCH BẢN ĐÃ BIÊN TẬP
# ==========================================
with tab_resync:
    col_r1, col_r2 = st.columns([1.6, 1])
    
    with col_r1:
        with st.container(border=True):
            st.markdown("### 🔄 Tải lên file Kịch bản ĐÃ BIÊN TẬP THỦ CÔNG (.docx)")
            st.caption("Dành riêng cho file kịch bản đã được team biên tập chỉnh sửa lời thoại. Hệ thống sẽ giữ nguyên 100% câu thoại mới và khôi phục lại màu chữ nhân vật, phân vai và tô highlight vàng phiên âm.")
            
            resync_file = st.file_uploader(
                "Kéo thả file .docx đã biên tập vào đây", 
                type=['docx'], 
                key=f"resync_uploader_{st.session_state['resync_uploader_key']}"
            )
            project_week_input = st.text_input("📌 Gán Tuần Dự Án cho video này:", value="Tuần 1", help="VD: Tuần 1, Tuần 2, Tuần 1 - Đợt Phim A...")
            
        if resync_file is not None:
            r_filename = resync_file.name
            r_name_no_ext = os.path.splitext(r_filename)[0]
            st.success(f"📄 Đã nhận file kịch bản biên tập: **{r_filename}**")
            
            st.markdown("---")
            if st.button("✨ 2. BẮT ĐẦU RE-SYNC & CHUẨN HÓA LẠI ĐỊNH DẠNG", use_container_width=True, type="primary", key="btn_resync_start"):
                try:
                    r_docx, r_ass, r_srt, r_zip, r_stats = process_docx(resync_file, r_name_no_ext, enable_colors, enable_phonetic, enable_cast, is_resync=True)
                    
                    st.session_state['r_processed_docx'] = r_docx
                    st.session_state['r_processed_ass'] = r_ass
                    st.session_state['r_processed_srt'] = r_srt
                    st.session_state['r_actor_zip'] = r_zip
                    st.session_state['r_docx_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".docx")
                    st.session_state['r_ass_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".ass")
                    st.session_state['r_srt_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".srt")
                    st.session_state['r_zip_name'] = clean_file_name_for_output(r_filename, tag="_KichBan_TachVai_Final", ext=".zip")
                    st.session_state['resync_stats'] = r_stats
                    
                    video_title = r_stats.get("video_title", r_name_no_ext)
                    actors_list = r_stats.get("actors_list", [])
                    actors_str = ", ".join(actors_list) if actors_list else "Chưa có thông tin"
                    actor_breakdown = r_stats.get("actor_stats_breakdown", {})
                    total_lines = r_stats.get("total_lines", 0)
                    video_dur_min = r_stats.get("video_duration_min", 1)
                    
                    today_str = time.strftime("%d/%m/%Y")
                    assigned_week = project_week_input.strip() if project_week_input else "Tuần 1"
                    
                    tracker_list = st.session_state['dubbing_tracker']
                    existing_entry = next((item for item in tracker_list if item['video_title'].upper() == video_title.upper()), None)
                    
                    curr_def_rate = st.session_state['payroll_rates'].get("unit_rate", 30000)
                    custom_actor_rates = {a.upper(): curr_def_rate for a in actors_list}
                    
                    entry_data = {
                        "video_title": video_title, "actors": actors_str, "actor_breakdown": actor_breakdown,
                        "total_lines": total_lines, "video_duration_min": video_dur_min, "date": today_str,
                        "project_week": assigned_week, "custom_actor_rates": custom_actor_rates
                    }
                    
                    if existing_entry:
                        if "custom_actor_rates" in existing_entry: entry_data["custom_actor_rates"] = existing_entry["custom_actor_rates"]
                        existing_entry.update(entry_data)
                    else: tracker_list.append(entry_data)
                    
                    st.session_state['dubbing_tracker'] = tracker_list
                    save_json_db(TRACKER_DB_FILE, tracker_list)
                    
                except Exception as e: st.error(f"Lỗi xảy ra khi Re-Sync: {e}")
                    
            if 'r_processed_docx' in st.session_state:
                st.markdown("---")
                r_qc_warns = st.session_state['resync_stats'].get("qc_warnings", [])
                if r_qc_warns:
                    with st.expander("🔍 BÁO CÁO CẢNH BÁO CHẤT LƯỢNG (QC & CPS CHECKER)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in r_qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(r_qc_warns) > 10: st.info(f"...và thêm {len(r_qc_warns)-10} cảnh báo khác.")
                
                st.markdown("### ⬇️ 3. TẢI VỀ CÁC FILE CHUẨN HOÀN HẢO (FINAL)")
                col_rdl1, col_rdl2, col_rdl3 = st.columns(3)
                with col_rdl1:
                    st.download_button(
                        label="📄 FILE WORD KỊCH BẢN (.DOCX)", data=st.session_state['r_processed_docx'],
                        file_name=st.session_state['r_docx_name'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True, key="btn_resync_dl_docx"
                    )
                with col_rdl2:
                    st.download_button(
                        label="🎬 PHỤ ĐỀ CAO CẤP (.ASS)", data=st.session_state['r_processed_ass'],
                        file_name=st.session_state['r_ass_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_ass"
                    )
                with col_rdl3:
                    st.download_button(
                        label="📝 PHỤ ĐỀ CHUẨN (.SRT)", data=st.session_state['r_processed_srt'],
                        file_name=st.session_state['r_srt_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_srt"
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ KỊCH BẢN TÁCH VAI RIÊNG CHO PHÒNG THU LỒNG TIẾNG")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                r_act_map = st.session_state['resync_stats'].get("actor_dialogue_map", {})
                if r_act_map:
                    col_ract1, col_ract2 = st.columns([2, 1])
                    with col_ract1:
                        r_selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(r_act_map.keys()), key="resync_select_actor")
                        if r_selected_actor:
                            r_act_buf = generate_actor_docx(st.session_state['resync_stats']['video_title'], r_selected_actor, r_act_map[r_selected_actor])
                            st.download_button(
                                label=f"⬇️ TẢI FILE WORD RIÊNG CHO {r_selected_actor} (.DOCX)", data=r_act_buf,
                                file_name=f"KichBan_{r_selected_actor}_{st.session_state['resync_stats']['video_title']}_Final.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="btn_dl_single_actor_resync"
                            )
                    with col_ract2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 TẢI TRỌN BỘ KỊCH BẢN TÁCH VAI (.ZIP)", data=st.session_state['r_actor_zip'],
                            file_name=st.session_state['r_zip_name'], mime="application/zip", type="secondary", use_container_width=True, key="btn_dl_zip_actor_resync"
                        )
                st.balloons()
        else: st.info("📌 **Hãy tải file kịch bản đã qua chỉnh sửa thủ công để hệ thống phục hồi lại định dạng chuẩn.**")

    with col_r2:
        st.markdown("### 📊 Re-Sync Analytics")
        if 'resync_stats' in st.session_state:
            r_stats = st.session_state['resync_stats']
            st.markdown(f"""
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">🎭 Tổng số Nhân vật</div>
                <div class="metric-value">{r_stats["total_speakers"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">💬 Tổng số Câu thoại</div>
                <div class="metric-value">{r_stats["total_lines"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">⏱️ Độ dài Video</div>
                <div class="metric-value">{r_stats["video_duration_min"]} phút</div>
            </div>
            """, unsafe_allow_html=True)
            top_name, top_count = r_stats["top_speaker"]
            st.info(f"👑 **Nhân vật thoại nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
        else: st.info("Thống kê file Re-Sync sẽ xuất hiện tại đây sau khi hoàn tất.")

# ==========================================
# TAB 3: THEO DÕI & BÁO CÁO LƯƠNG
# ==========================================
with tab_dub_tracker:
    st.subheader("📋 BÁO CÁO BẢNG TÍNH LƯƠNG LỒNG TIẾNG THEO TUẦN DỰ ÁN")
    
    with st.expander("⚙️ CẤU HÌNH ĐƠN GIÁ MẶC ĐỊNH SẢN XUẤT", expanded=False):
        c_rate1, c_rate2 = st.columns([2, 3])
        curr_mode = st.session_state['payroll_rates'].get("mode", "minute")
        mode_idx = 0 if curr_mode == "minute" else (1 if curr_mode == "line" else 2)
        
        with c_rate1:
            rate_mode_choice = st.radio(
                "Cách tính thù lao:", options=["Theo Phút video (phút)", "Theo Câu thoại (câu)", "Theo Số từ (từ)"], 
                index=mode_idx, key="radio_payroll_mode"
            )
        with c_rate2:
            default_unit_rate = st.number_input(
                "Đơn giá mặc định khi thêm video mới (VNĐ):", value=int(st.session_state['payroll_rates'].get("unit_rate", 30000)), 
                step=5000, key="num_payroll_unit_rate"
            )
            unit_label = "phút" if "Phút" in rate_mode_choice else ("câu" if "Câu" in rate_mode_choice else "từ")
            st.caption(f"👉 **Đơn giá mặc định áp dụng:** `{default_unit_rate:,.0f}` VNĐ / {unit_label}")
            
        if "Phút" in rate_mode_choice: new_mode = "minute"
        elif "Câu" in rate_mode_choice: new_mode = "line"
        else: new_mode = "word"
            
        if new_mode != st.session_state['payroll_rates'].get("mode") or default_unit_rate != st.session_state['payroll_rates'].get("unit_rate"):
            st.session_state['payroll_rates'] = {"mode": new_mode, "unit_rate": default_unit_rate}
            save_json_db(RATES_DB_FILE, st.session_state['payroll_rates'])

    current_mode = st.session_state['payroll_rates'].get("mode", "minute")
    current_rate = st.session_state['payroll_rates'].get("unit_rate", 30000)

    subtab_video, subtab_custom_rates, subtab_payroll = st.tabs([
        "📹 Bảng Theo Dõi Video Theo Tuần Dự Án", "💵 Chỉnh Đơn Giá Cá Nhân (Video x Diễn viên)", "💰 Báo Cáo Lương Chi Tiết Từng Diễn Viên"
    ])

    tracker_data = st.session_state['dubbing_tracker']

    with subtab_video:
        st.markdown("#### Bảng quản lý danh sách Video và Thù lao (Định dạng chuẩn 6 cột)")
        if tracker_data:
            formatted_editor_list = []
            for idx, item in enumerate(tracker_data, 1):
                p_week = item.get("project_week", "Tuần 1")
                v_dur_min = int(item.get("video_duration_min", 1))
                v_lines = item.get("total_lines", 0)
                v_acts = item.get("actors", "")
                bd = item.get("actor_breakdown", {})
                custom_rates = item.get("custom_actor_rates", {})
                
                raw_acts = [a.strip().upper() for a in v_acts.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                
                if current_mode == "minute":
                    v_pay = sum(v_dur_min * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_dur_min * current_rate
                    dur_str = f"{v_dur_min} phút"
                elif current_mode == "line":
                    v_pay = sum((bd.get(act, {}).get("lines", v_lines)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_lines * current_rate
                    dur_str = f"{v_lines} câu"
                else:
                    tot_words = sum(a["words"] for a in bd.values()) if bd else 0
                    v_pay = sum((bd.get(act, {}).get("words", 0)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else tot_words * current_rate
                    dur_str = f"{tot_words} từ"
                    
                formatted_editor_list.append({
                    "Stt": idx, "Tuần dự án": p_week, "Tiêu đề video": item['video_title'],
                    "Thời lượng": dur_str, "Thành tiền Video": f"{v_pay:,.0f} VNĐ", "Diễn viên": v_acts, "Xóa": False
                })

            df_editor_raw = pd.DataFrame(formatted_editor_list)

            edited_tracker_df = st.data_editor(
                df_editor_raw,
                column_config={
                    "Stt": st.column_config.NumberColumn("Stt", disabled=True, width="small"),
                    "Tuần dự án": st.column_config.TextColumn("Tuần dự án (Sửa trực tiếp)"),
                    "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video (Sửa trực tiếp)"),
                    "Thời lượng": st.column_config.TextColumn("Thời lượng", disabled=True),
                    "Thành tiền Video": st.column_config.TextColumn("Thành tiền (Tự động tính theo giá riêng)", disabled=True),
                    "Diễn viên": st.column_config.TextColumn("Diễn viên (Sửa trực tiếp)"),
                    "Xóa": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="dubbing_tracker_editor_main"
            )

            col_tr1, col_tr2 = st.columns([1, 1])
            with col_tr1:
                if st.button("💾 LƯU THAY ĐỔI TRÊN BẢNG VIDEO", type="primary", use_container_width=True):
                    new_tracker = []; deleted_cnt = 0
                    for idx_r, row in edited_tracker_df.iterrows():
                        if row["Xóa"]: deleted_cnt += 1
                        else:
                            orig_item = tracker_data[idx_r]
                            orig_item['project_week'] = str(row["Tuần dự án"]).strip()
                            orig_item['video_title'] = str(row["Tiêu đề video"]).strip()
                            
                            new_acts_raw = [a.strip().upper() for a in str(row["Diễn viên"]).split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                            orig_item['actors'] = ", ".join(new_acts_raw) if new_acts_raw else "CHƯA CÓ THÔNG TIN"
                            
                            old_bd = orig_item.get("actor_breakdown", {}); new_bd = {}
                            c_rates = orig_item.get("custom_actor_rates", {})
                            for act in new_acts_raw:
                                if act in old_bd: new_bd[act] = old_bd[act]
                                else: new_bd[act] = {"lines": 0, "words": 0}
                                if act not in c_rates: c_rates[act] = current_rate
                            
                            orig_item['actor_breakdown'] = new_bd
                            orig_item['custom_actor_rates'] = c_rates
                            new_tracker.append(orig_item)
                            
                    st.session_state['dubbing_tracker'] = new_tracker
                    save_json_db(TRACKER_DB_FILE, new_tracker)
                    st.success(f"✅ Đã lưu thay đổi & đồng bộ dữ liệu!"); time.sleep(1); st.rerun()

            st.markdown("---")
            st.markdown("#### 📑 Xem theo Bảng chia Tuần Dự Án (Có Hàng Tổng)")

            project_weeks_map = {}
            for item in tracker_data:
                pw = item.get("project_week", "Tuần 1")
                if pw not in project_weeks_map: project_weeks_map[pw] = []
                project_weeks_map[pw].append(item)

            total_studio_money = 0; excel_sheets_data = {}

            for pw_name, items in sorted(project_weeks_map.items()):
                st.markdown(f"##### 📅 {pw_name.upper()}")
                week_rows = []; week_tot_dur = 0; week_tot_money = 0
                
                for idx, item in enumerate(items, 1):
                    v_title = item['video_title']; v_dur_min = int(item.get("video_duration_min", 1))
                    v_lines = item.get("total_lines", 0); v_actors = item.get("actors", "")
                    bd = item.get("actor_breakdown", {}); custom_rates = item.get("custom_actor_rates", {})
                    
                    raw_acts = [a.strip().upper() for a in v_actors.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                    
                    if current_mode == "minute":
                        v_pay = sum(v_dur_min * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_dur_min * current_rate
                        dur_str = f"{v_dur_min} phút"
                    elif current_mode == "line":
                        v_pay = sum((bd.get(act, {}).get("lines", v_lines)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_lines * current_rate
                        dur_str = f"{v_lines} câu"
                    else:
                        tot_words = sum(a["words"] for a in bd.values()) if bd else 0
                        v_pay = sum((bd.get(act, {}).get("words", 0)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else tot_words * current_rate
                        dur_str = f"{tot_words} từ"
                        
                    week_tot_dur += v_dur_min; week_tot_money += v_pay
                    
                    week_rows.append({
                        "Stt": idx, "Tiêu đề video": v_title, "Thời lượng": dur_str,
                        "Thành tiền": f"{v_pay:,.0f} VNĐ", "Diễn viên": v_actors
                    })

                total_studio_money += week_tot_money
                week_rows.append({
                    "Stt": "TỔNG", "Tiêu đề video": f"TỔNG CỘNG ({len(items)} video)",
                    "Thời lượng": f"{week_tot_dur} phút" if current_mode == "minute" else "-",
                    "Thành tiền": f"{week_tot_money:,.0f} VNĐ", "Diễn viên": "-"
                })

                df_week_display = pd.DataFrame(week_rows)
                st.dataframe(df_week_display, hide_index=True, use_container_width=True)
                excel_sheets_data[pw_name[:30]] = df_week_display

            with col_tr2:
                excel_payroll_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_payroll_buffer, engine='openpyxl') as writer:
                    for s_name, s_df in excel_sheets_data.items():
                        s_df.to_excel(writer, index=False, sheet_name=s_name.replace(":", "_").replace("/", "_"))
                excel_payroll_buffer.seek(0)

                st.download_button(
                    label="📊 XUẤT TẤT CẢ TUẦN RA EXCEL (.XLSX)", data=excel_payroll_buffer,
                    file_name="Theo_Doi_Video_Long_Tieng_MaiHan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
                )
        else: st.info("Chưa có dữ liệu video nào. Hãy chạy Re-Sync ở Tab 2 để tự động ghi nhận video mới!")

    with subtab_custom_rates:
        st.markdown("#### 💵 Bảng Điều Chỉnh Đơn Giá Cá Nhân (Theo Video & Diễn viên)")
        st.caption("Gõ đơn giá riêng cho từng người trong từng video (Ví dụ: Khánh video A 30k, video B 20k; Trúc video A 15k...):")
        
        if tracker_data:
            rate_editor_rows = []
            for v_idx, item in enumerate(tracker_data):
                pw = item.get("project_week", "Tuần 1"); v_title = item['video_title']
                v_dur = int(item.get("video_duration_min", 1)); v_acts = item.get("actors", "")
                custom_rates = item.get("custom_actor_rates", {})
                acting_actors = [a.strip().upper() for a in v_acts.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                
                for act in acting_actors:
                    act_rate = custom_rates.get(act, current_rate)
                    rate_editor_rows.append({
                        "v_idx": v_idx, "Tuần dự án": pw, "Tiêu đề video": v_title,
                        "Thời lượng (Phút)": v_dur, "Diễn viên": act, "Đơn giá cá nhân (VNĐ)": int(act_rate)
                    })

            if rate_editor_rows:
                df_custom_rates = pd.DataFrame(rate_editor_rows)
                edited_custom_rates_df = st.data_editor(
                    df_custom_rates[["Tuần dự án", "Tiêu đề video", "Thời lượng (Phút)", "Diễn viên", "Đơn giá cá nhân (VNĐ)"]],
                    column_config={
                        "Tuần dự án": st.column_config.TextColumn("Tuần", disabled=True),
                        "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video", disabled=True),
                        "Thời lượng (Phút)": st.column_config.NumberColumn("Độ dài (Phút)", disabled=True),
                        "Diễn viên": st.column_config.TextColumn("Diễn viên", disabled=True),
                        "Đơn giá cá nhân (VNĐ)": st.column_config.NumberColumn("Đơn giá riêng (Gõ để sửa)", format="%d", step=1000)
                    },
                    hide_index=True, use_container_width=True, key="custom_rates_editor_table_main"
                )

                if st.button("💾 LƯU ĐƠN GIÁ CÁ NHÂN", type="primary", use_container_width=True):
                    for idx_r, row in edited_custom_rates_df.iterrows():
                        v_i = df_custom_rates.iloc[idx_r]["v_idx"]; act_n = df_custom_rates.iloc[idx_r]["Diễn viên"]
                        new_r = float(row["Đơn giá cá nhân (VNĐ)"])
                        if "custom_actor_rates" not in tracker_data[v_i]: tracker_data[v_i]["custom_actor_rates"] = {}
                        tracker_data[v_i]["custom_actor_rates"][act_n] = new_r
                        
                    st.session_state['dubbing_tracker'] = tracker_data
                    save_json_db(TRACKER_DB_FILE, tracker_data)
                    st.success("✅ Đã lưu đơn giá riêng cho từng diễn viên thành công!"); time.sleep(1); st.rerun()
            else: st.info("Chưa có thông tin diễn viên lồng tiếng trong danh sách video.")
        else: st.info("Chưa có danh sách video để điều chỉnh đơn giá.")

    with subtab_payroll:
        st.markdown("#### Bảng Tổng kết & Báo cáo Lương Chi Tiết Từng Diễn Viên")
        if tracker_data:
            col_filter1, col_filter2 = st.columns([1, 1])
            all_project_weeks = sorted(list(set(item.get("project_week", "Tuần 1") for item in tracker_data)))
            with col_filter1: selected_week_filter = st.selectbox("🎯 Chọn Tuần Dự Án:", options=["TẤT CẢ CÁC TUẦN"] + all_project_weeks)

            if selected_week_filter == "TẤT CẢ CÁC TUẦN": active_tracker_data = tracker_data
            else: active_tracker_data = [item for item in tracker_data if item.get("project_week", "Tuần 1") == selected_week_filter]

            actor_weekly_map = {}
            for item in active_tracker_data:
                v_title = item['video_title']; v_dur = int(item.get("video_duration_min", 1))
                bd = item.get("actor_breakdown", {}); v_lines = item.get("total_lines", 0)
                custom_rates = item.get("custom_actor_rates", {})
                acting_actors = [a.strip().upper() for a in item.get("actors", "").split(",") if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]

                for act_name_clean in acting_actors:
                    if act_name_clean not in actor_weekly_map:
                        actor_weekly_map[act_name_clean] = {
                            "videos_count": 0, "total_video_mins": 0, "total_lines": 0, "total_words": 0, "video_rows": []
                        }
                    
                    act_rate = custom_rates.get(act_name_clean, current_rate)
                    act_lines = bd[act_name_clean]["lines"] if (bd and act_name_clean in bd) else v_lines
                    act_words = bd[act_name_clean]["words"] if (bd and act_name_clean in bd) else 0
                    
                    if current_mode == "minute": act_pay = v_dur * act_rate; dur_str = f"{v_dur} phút"
                    elif current_mode == "line": act_pay = act_lines * act_rate; dur_str = f"{act_lines} câu"
                    else: act_pay = act_words * act_rate; dur_str = f"{act_words} từ"

                    actor_weekly_map[act_name_clean]["videos_count"] += 1
                    actor_weekly_map[act_name_clean]["total_video_mins"] += v_dur
                    actor_weekly_map[act_name_clean]["total_lines"] += act_lines
                    actor_weekly_map[act_name_clean]["total_words"] += act_words
                    actor_weekly_map[act_name_clean]["video_rows"].append({
                        "Stt": len(actor_weekly_map[act_name_clean]["video_rows"]) + 1,
                        "Tiêu đề video": v_title, "Thời lượng": dur_str,
                        "Đơn giá": f"{act_rate:,.0f} VNĐ", "Thành tiền": f"{act_pay:,.0f} VNĐ", "Pay_Num": act_pay
                    })

            all_actor_names = sorted(list(actor_weekly_map.keys()))
            with col_filter2:
                selected_actor_view = st.selectbox("👤 Chọn Diễn viên để xem & xuất Báo cáo Lương cá nhân:", options=["TẤT CẢ DIỄN VIÊN"] + all_actor_names)

            st.markdown("---")

            if selected_actor_view != "TẤT CẢ DIỄN VIÊN":
                a_info = actor_weekly_map[selected_actor_view]; a_rows = a_info["video_rows"]
                a_tot_pay = sum(r["Pay_Num"] for r in a_rows)
                st.subheader(f"👤 PHIẾU BÁO CÁO THÙ LAO: {selected_actor_view}")
                st.caption(f"Dữ liệu thù lao lồng tiếng cho {selected_actor_view} ({selected_week_filter})")
                
                df_single_actor = pd.DataFrame(a_rows)[["Stt", "Tiêu đề video", "Thời lượng", "Đơn giá", "Thành tiền"]]
                st.dataframe(df_single_actor, hide_index=True, use_container_width=True)
                st.metric(f"💰 TỔNG THÙ LAO DỰ KIẾN TRẢ CHO {selected_actor_view}:", f"{a_tot_pay:,.0f} VNĐ")
                
                col_p_btn1, col_p_btn2 = st.columns(2)
                with col_p_btn1:
                    actor_docx_buf = generate_actor_salary_slip_docx(selected_actor_view, selected_week_filter, a_rows, a_tot_pay, current_mode)
                    st.download_button(
                        label=f"🖨️ IN / TẢI PHIẾU LƯƠNG WORD CỦA {selected_actor_view} (.DOCX)",
                        data=actor_docx_buf, file_name=f"PhieuLuong_{selected_actor_view}_{selected_week_filter}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True
                    )
                with col_p_btn2:
                    excel_single_buf = io.BytesIO()
                    with pd.ExcelWriter(excel_single_buf, engine='openpyxl') as writer: df_single_actor.to_excel(writer, index=False, sheet_name="Phieu Luong")
                    excel_single_buf.seek(0)
                    st.download_button(
                        label=f"📊 TẢI PHIẾU LƯƠNG CÁ NHÂN EXCEL (.XLSX)",
                        data=excel_single_buf, file_name=f"PhieuLuong_{selected_actor_view}_{selected_week_filter}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
                    )
            else:
                actor_payroll_rows = []; grand_actor_pay = 0; unique_video_titles = set(); grand_unique_mins = 0
                for item in active_tracker_data:
                    if item['video_title'] not in unique_video_titles:
                        unique_video_titles.add(item['video_title'])
                        grand_unique_mins += int(item.get("video_duration_min", 1))

                for idx, (act_name, info) in enumerate(sorted(actor_weekly_map.items()), 1):
                    if current_mode == "minute": unit_cnt = info["total_video_mins"]; dur_disp = f"{unit_cnt} phút"
                    elif current_mode == "line": unit_cnt = info["total_lines"]; dur_disp = f"{unit_cnt} câu"
                    else: unit_cnt = info["total_words"]; dur_disp = f"{unit_cnt} từ"

                    tot_p = sum(r["Pay_Num"] for r in info["video_rows"])
                    grand_actor_pay += tot_p

                    actor_payroll_rows.append({
                        "Stt": idx, "Diễn viên Lồng tiếng": act_name, "Số Video đã lồng": info["videos_count"],
                        "Tổng phút video": dur_disp, "Thành tiền Lương": f"{tot_p:,.0f} VNĐ",
                        "Danh sách Video tham gia": ", ".join([r["Tiêu đề video"] for r in info["video_rows"]])
                    })

                actor_payroll_rows.append({
                    "Stt": "TỔNG", "Diễn viên Lồng tiếng": f"TỔNG CỘNG ({len(actor_weekly_map)} Diễn viên)",
                    "Số Video đã lồng": "-", "Tổng phút video": f"{grand_unique_mins} phút" if current_mode == "minute" else "-",
                    "Thành tiền Lương": f"{grand_actor_pay:,.0f} VNĐ", "Danh sách Video tham gia": "-"
                })

                df_act_payroll = pd.DataFrame(actor_payroll_rows)
                st.metric(f"💰 TỔNG LƯƠNG CẦN CHI CHO DIỄN VIÊN ({selected_week_filter}):", f"{grand_actor_pay:,.0f} VNĐ")
                st.dataframe(
                    df_act_payroll[["Stt", "Diễn viên Lồng tiếng", "Số Video đã lồng", "Tổng phút video", "Thành tiền Lương", "Danh sách Video tham gia"]],
                    hide_index=True, use_container_width=True
                )

                excel_actor_buf = io.BytesIO()
                with pd.ExcelWriter(excel_actor_buf, engine='openpyxl') as writer: df_act_payroll.to_excel(writer, index=False, sheet_name="Luong Dien Vien")
                excel_actor_buf.seek(0)
                st.download_button(
                    label=f"📊 TẢI BÁO CÁO LƯƠNG TOÀN BỘ DIỄN VIÊN ({selected_week_filter}) EXCEL (.XLSX)",
                    data=excel_actor_buf, file_name=f"Bao_Cao_Luong_DienVien_{selected_week_filter}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True
                )
        else: st.info("Chưa có dữ liệu lồng tiếng để lập báo cáo lương.")

# ==========================================
# TAB 4: QUẢN LÝ DATABASE PHÂN VAI LỒNG TIẾNG
# ==========================================
with tab_cast_db:
    with st.container(border=True):
        st.subheader("🎭 BẢNG PHÂN VAI LỒNG TIẾNG (GLOBAL DATABASE)")
        st.markdown("Nơi thiết lập mặc định nhân vật Tiếng Anh nào sẽ do diễn viên lồng tiếng Việt nào đảm nhận cho Mai Han Team.")

        st.markdown("#### ➕ Thêm / Cập nhật Phân vai mới")
        c_c1, c_c2, c_c3 = st.columns([2, 2, 1.2])
        with c_c1: add_role_eng = st.text_input("Tên Nhân vật (Tiếng Anh):", placeholder="VD: Bri, Chase...", key=f"add_role_eng_{st.session_state['cast_input_key']}")
        with c_c2: add_actor_vn = st.text_input("Diễn viên Lồng tiếng (Tiếng Việt):", placeholder="VD: TRÚC, THIỆN...", key=f"add_actor_vn_{st.session_state['cast_input_key']}")
        with c_c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Phân Vai", use_container_width=True, type="primary", key="btn_add_cast"):
                if add_role_eng and add_actor_vn:
                    k = add_role_eng.upper().strip(); v = add_actor_vn.strip().upper()
                    st.session_state['custom_cast_mapping'][k] = v
                    save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                    st.session_state['cast_input_key'] += 1
                    st.success(f"✅ Đã gán thành công: `{add_role_eng}` ➔ `{add_actor_vn}`"); time.sleep(1); st.rerun()
                else: st.warning("Vui lòng nhập đầy đủ tên nhân vật và diễn viên!")

    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 📑 Danh sách Toàn bộ Bảng Phân Vai Đã Lưu")
        search_cast_query = st.text_input("🔍 Tìm kiếm Nhân vật hoặc Diễn viên lồng tiếng:", placeholder="Gõ tên nhân vật hoặc diễn viên...").strip().upper()
        all_cast_dict = st.session_state['custom_cast_mapping']

        if search_cast_query: filtered_cast = {k: v for k, v in all_cast_dict.items() if search_cast_query in k or search_cast_query in v.upper()}
        else: filtered_cast = all_cast_dict

        if filtered_cast:
            cast_db_table = []
            for eng_role, vn_actor in sorted(filtered_cast.items()):
                cast_db_table.append({"Nhân vật (Tiếng Anh)": eng_role, "Diễn viên Lồng tiếng (Tiếng Việt)": vn_actor, "Xóa khỏi Database": False})

            df_cast_db = pd.DataFrame(cast_db_table)
            st.caption(f"Đang hiển thị **{len(df_cast_db)}** vai lồng tiếng trong kho:")

            edited_cast_db_df = st.data_editor(
                df_cast_db,
                column_config={
                    "Nhân vật (Tiếng Anh)": st.column_config.TextColumn("Tên Nhân vật gốc (In hoa)", disabled=True),
                    "Diễn viên Lồng tiếng (Tiếng Việt)": st.column_config.TextColumn("Diễn viên lồng tiếng (Sửa trực tiếp)"),
                    "Xóa khỏi Database": st.column_config.CheckboxColumn("Xóa?")
                },
                disabled=["Nhân vật (Tiếng Anh)"], hide_index=True, use_container_width=True, key="global_cast_db_editor"
            )

            if st.button("💾 LƯU TOÀN BỘ CẬP NHẬT PHÂN VAI", type="primary", use_container_width=True, key="btn_save_global_cast"):
                new_cast_db = {}; deleted_cast_count = 0
                if search_cast_query:
                    for k, v in all_cast_dict.items():
                        if k not in filtered_cast: new_cast_db[k] = v

                for _, row in edited_cast_db_df.iterrows():
                    eng_k = str(row["Nhân vật (Tiếng Anh)"]).upper().strip()
                    act_v = str(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).strip().upper()
                    is_del = row["Xóa khỏi Database"]
                    if is_del: deleted_cast_count += 1
                    else:
                        if act_v: new_cast_db[eng_k] = act_v

                st.session_state['custom_cast_mapping'] = new_cast_db
                save_json_db(CAST_DB_FILE, new_cast_db)
                st.success(f"✅ Đã lưu cập nhật thành công! (Đã xóa {deleted_cast_count} vai)"); time.sleep(1); st.rerun()
        else: st.info("Không tìm thấy vai lồng tiếng nào khớp với từ khóa tìm kiếm.")

# ==========================================
# TAB 5: KHO DATABASE PHIÊN ÂM GIỌNG NAM (TỔNG HỢP)
# ==========================================
with tab_phonetic_db:
    with st.container(border=True):
        st.subheader("📚 Từ Điển Phiên Âm Giọng Nam (Global Database)")
        st.markdown("Nơi quản lý toàn bộ kho từ vựng Tiếng Anh và các bản phiên âm giọng Nam được lưu trữ lâu dài trên hệ thống.")
        
        st.markdown("#### 🔊 Nghe phát âm thử bất kỳ cụm từ/từ Tiếng Anh nào")
        col_test1, col_test2, col_test3 = st.columns([2.5, 1.5, 1.5])
        with col_test1: free_test_word = st.text_input("Nhập từ/cụm Tiếng Anh cần nghe thử:", placeholder="VD: Starbucks, Hamburger, McDonald's...", key="free_audio_text")
        with col_test2: free_accent = st.radio("Giọng phát âm:", options=["Giọng Mỹ (US)", "Giọng Anh (UK)"], horizontal=True, key="free_audio_accent")
        with col_test3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            free_listen_btn = st.button("🔊 Phát âm thanh", type="secondary", use_container_width=True, key="btn_free_listen")

        if free_listen_btn and free_test_word:
            tld = 'com' if "Mỹ" in free_accent else 'co.uk'
            test_fp = generate_english_audio(free_test_word, accent=tld)
            if test_fp: st.audio(test_fp, format="audio/mp3", autoplay=True)

        st.markdown("---")
        st.markdown("#### ➕ Bổ sung từ phiên âm mới vào Kho")
        c1, c2, c3 = st.columns([2, 2, 1.2])
        with c1: tab_add_eng = st.text_input("Từ Tiếng Anh gốc:", placeholder="VD: Burger", key=f"tab_add_eng_{st.session_state['pho_input_key']}")
        with c2: tab_add_pho = st.text_input("Phiên âm giọng Nam:", placeholder="VD: Bơ-gơ", key=f"tab_add_pho_{st.session_state['pho_input_key']}")
        with c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm vào Database", use_container_width=True, type="primary"):
                if tab_add_eng and tab_add_pho:
                    k = tab_add_eng.upper().strip(); v = tab_add_pho.strip()
                    st.session_state['custom_phonetics'][k] = v
                    save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                    st.session_state['pho_input_key'] += 1
                    st.success(f"✅ Đã thêm thành công: `{tab_add_eng}` ➔ `{tab_add_pho}`"); time.sleep(1); st.rerun()
                else: st.warning("Vui lòng điền đủ 2 ô!")

    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 📑 Danh sách toàn bộ Từ phiên âm đã lưu")
        search_query = st.text_input("🔍 Tìm kiếm từ Tiếng Anh hoặc Từ phiên âm:", placeholder="Gõ từ cần tìm ở đây...").strip().upper()
        all_phonetics_dict = st.session_state['custom_phonetics']
        
        if search_query: filtered_dict = {k: v for k, v in all_phonetics_dict.items() if search_query in k or search_query in v.upper()}
        else: filtered_dict = all_phonetics_dict

        if filtered_dict:
            db_table_data = []
            for eng_key, pho_val in sorted(filtered_dict.items()):
                db_table_data.append({"Từ Tiếng Anh": eng_key, "Phiên âm giọng Nam": pho_val, "Xóa khỏi Database": False})

            df_db = pd.DataFrame(db_table_data)
            st.caption(f"Đang hiển thị **{len(df_db)}** từ phiên âm trong hệ thống:")

            edited_db_df = st.data_editor(
                df_db,
                column_config={
                    "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh gốc (In hoa)", disabled=True),
                    "Phiên âm giọng Nam": st.column_config.TextColumn("Phiên âm giọng Nam (Sửa trực tiếp tại đây)"),
                    "Xóa khỏi Database": st.column_config.CheckboxColumn("Xóa?")
                },
                disabled=["Từ Tiếng Anh"], hide_index=True, use_container_width=True, key="global_phonetic_db_editor"
            )

            if st.button("💾 LƯU TOÀN BỘ CẬP NHẬT TRONG BẢNG", type="primary", use_container_width=True):
                new_db = {}; deleted_count = 0
                if search_query:
                    for k, v in all_phonetics_dict.items():
                        if k not in filtered_dict: new_db[k] = v

                for _, row in edited_db_df.iterrows():
                    eng_k = str(row["Từ Tiếng Anh"]).upper().strip(); pho_v = str(row["Phiên âm giọng Nam"]).strip()
                    is_delete = row["Xóa khỏi Database"]
                    if is_delete: deleted_count += 1
                    else:
                        if pho_v: new_db[eng_k] = pho_v

                st.session_state['custom_phonetics'] = new_db
                save_json_db(PHONETIC_DB_FILE, new_db)
                st.success(f"✅ Đã lưu cập nhật thành công! (Đã xóa {deleted_count} từ)"); time.sleep(1); st.rerun()
        else: st.info("Không tìm thấy từ phiên âm nào khớp với từ khóa tìm kiếm.")

# ==========================================
# TAB 6: BỘ CÔNG CỤ CHUYỂN ĐỔI (CONVERTER SUITE)
# ==========================================
with tab_tools:
    subtab_sub_conv, subtab_srt_excel, subtab_curr, subtab_dist, subtab_speed, subtab_mass_temp = st.tabs([
        "🎬 Kịch Bản Subtitle (SRT ⇄ DOCX)",
        "📊 SRT ➔ Excel (.xlsx)",
        "💵 Tiền Tệ (Currency)",
        "📏 Khoảng Cách (Distance)",
        "🚀 Vận Tốc (Speed)",
        "⚖️ Khối Lượng & Nhiệt Độ"
    ])

    # 1. BỘ CHUYỂN ĐỔI SUBTITLE KỊCH BẢN (SRT ⇄ DOCX)
    with subtab_sub_conv:
        st.markdown("#### 🎬 Bộ Công Cụ Chuyển Đổi Subtitle Chuyên Nghiệp")
        col_c1, col_c2 = st.columns(2)
        
        # Module 1: SRT -> DOCX
        with col_c1:
            with st.container(border=True):
                st.markdown("##### 📄 1. Chuyển SRT ➔ Word (.docx)")
                st.caption("Giữ nguyên cấu trúc dòng, định dạng font Times New Roman, 12pt:")
                
                srt_file = st.file_uploader("Tải file .srt của bạn vào đây:", type=['srt'], key="tool_srt_to_docx")
                if srt_file:
                    s_filename = srt_file.name
                    s_name_no_ext = os.path.splitext(s_filename)[0]
                    st.info(f"Đã nhận: **{s_filename}**")
                    
                    if st.button("✨ Chuyển SRT Sang Word", use_container_width=True, type="primary"):
                        try:
                            docx_buf = process_srt_to_docx(srt_file, s_name_no_ext)
                            st.success("✅ Chuyển đổi hoàn tất!")
                            st.download_button(
                                label="⬇️ Tải File Word (.docx)",
                                data=docx_buf,
                                file_name=f"CONVERTED_{s_name_no_ext}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        except Exception as e: st.error(f"Lỗi: {e}")

        # Module 2: DOCX -> SRT (Batch)
        with col_c2:
            with st.container(border=True):
                st.markdown("##### 📝 2. Chuyển Word (.docx) ➔ SRT (Batch hàng loạt)")
                st.caption("Tải 1 file hoặc hàng ngàn file Word kịch bản để tự động trích xuất SRT:")
                
                batch_docx_files = st.file_uploader(
                    "Tải 1 hoặc nhiều file .docx:",
                    type=['docx'],
                    accept_multiple_files=True,
                    key="tool_docx_to_srt_batch"
                )
                
                if batch_docx_files:
                    st.info(f"Đã chọn **{len(batch_docx_files)}** file Word.")
                    if st.button("✨ Chuyển Hàng Loạt Sang SRT", use_container_width=True, type="primary"):
                        try:
                            if len(batch_docx_files) == 1:
                                single_f = batch_docx_files[0]
                                s_name_no_ext = os.path.splitext(single_f.name)[0]
                                srt_bytes = process_docx_to_srt(single_f)
                                st.success("✅ Chuyển đổi hoàn tất!")
                                st.download_button(
                                    label=f"⬇️ Tải {s_name_no_ext}.srt",
                                    data=srt_bytes,
                                    file_name=f"{s_name_no_ext}.srt",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            else:
                                zip_buf = io.BytesIO()
                                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                    for doc_f in batch_docx_files:
                                        s_name_no_ext = os.path.splitext(doc_f.name)[0]
                                        srt_bytes = process_docx_to_srt(doc_f)
                                        zf.writestr(f"{s_name_no_ext}.srt", srt_bytes)
                                zip_buf.seek(0)
                                st.success(f"✅ Đã chuyển đổi thành công {len(batch_docx_files)} file!")
                                st.download_button(
                                    label="📦 Tải Trọn Bộ SRT (.ZIP)",
                                    data=zip_buf.getvalue(),
                                    file_name="Converted_SRT_Files.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                        except Exception as e: st.error(f"Lỗi: {e}")

    # 2. BỘ CHUYỂN ĐỔI SRT TO EXCEL WITH SPEAKER STYLING
    with subtab_srt_excel:
        st.markdown("#### 📊 Chuyển Đổi File Subtitle SRT ➔ Bảng Tính Excel (.xlsx)")
        st.caption("Tự động nhận diện nhân vật, tô màu phân biệt người nói và xuất file Excel có cấu trúc:")
        
        uploaded_srt_excel = st.file_uploader("Tải file .srt của bạn vào đây:", type=['srt'], key="tool_srt_to_excel")
        if uploaded_srt_excel is not None:
            try:
                try: srt_content_excel = uploaded_srt_excel.read().decode("utf-8")
                except UnicodeDecodeError: srt_content_excel = uploaded_srt_excel.read().decode("latin-1")
            except Exception:
                st.error("Lỗi mã hóa file. Vui lòng đảm bảo file SRT của bạn ở chuẩn mã hóa UTF-8.")
                srt_content_excel = None

            if srt_content_excel:
                with st.spinner('Đang phân tích dữ liệu SRT...'):
                    df_converted_excel = parse_srt_to_dataframe(srt_content_excel)
                
                if df_converted_excel.empty:
                    st.error("Không thể đọc được phụ đề nào từ file SRT này.")
                else:
                    st.markdown("##### 📊 Thống Kê Nhân Vật")
                    unique_spks = df_converted_excel['Speaker'].unique()
                    actual_spks = [s for s in unique_spks if s not in ["Unknown", ""]]
                    
                    st.success(f"**Tổng số Người nói được nhận dạng:** {len(actual_spks)} người.")
                    if actual_spks:
                        st.markdown(f"**Danh sách Người nói:** {', '.join(actual_spks)}")
                    else:
                        st.info("Không tìm thấy người nói rõ ràng (ngoại trừ các đoạn hội thoại không gắn tên).")

                    st.markdown("##### 👁️ Xem Trước Bảng Dữ Liệu Chuyển Đổi")
                    styled_excel_df = apply_excel_styles(df_converted_excel)
                    st.dataframe(styled_excel_df, use_container_width=True)

                    output_excel = io.BytesIO()
                    styled_excel_df.to_excel(output_excel, index=False, engine='openpyxl')
                    output_excel.seek(0)

                    orig_base_name = uploaded_srt_excel.name.rsplit('.', 1)[0]
                    excel_out_filename = f"{orig_base_name}.xlsx"
                    
                    st.download_button(
                        label=f"💾 TẢI FILE EXCEL (.XLSX): {excel_out_filename}",
                        data=output_excel.getvalue(),
                        file_name=excel_out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

    # 3. BỘ CHUYỂN ĐỔI TIỀN TỆ (CURRENCY)
    with subtab_curr:
        st.markdown("#### 💵 Quy Đổi Tiền Tệ Đa Ngoại Tệ")
        rates = {
            "VND": 1.0, "USD": 25400.0, "EUR": 27500.0, "GBP": 32000.0, "JPY": 165.0,
            "CNY": 3500.0, "KRW": 18.5, "AUD": 16800.0, "CAD": 18200.0, "SGD": 18900.0
        }
        c_col1, c_col2, c_col3 = st.columns([2, 1.5, 1.5])
        with c_col1: curr_amount = st.number_input("Số lượng tiền cần đổi:", value=100.0, min_value=0.0, step=10.0)
        with c_col2: from_curr = st.selectbox("Từ đồng tiền:", options=list(rates.keys()), index=1)
        with c_col3: to_curr = st.selectbox("Sang đồng tiền:", options=list(rates.keys()), index=0)
            
        amount_in_vnd = curr_amount * rates[from_curr]
        result_curr = amount_in_vnd / rates[to_curr]
        
        st.markdown("---")
        st.markdown(f"### 🎯 Kết Quả: **{curr_amount:,.2f} {from_curr}** = **{result_curr:,.2f} {to_curr}**")
        st.caption(f"Tỷ giá tham chiếu: 1 USD = {rates['USD']:,.0f} VND | 1 EUR = {rates['EUR']:,.0f} VND | 1 JPY = {rates['JPY']:,.1f} VND")

    # 4. BỘ CHUYỂN ĐỔI KHOẢNG CÁCH (DISTANCE)
    with subtab_dist:
        st.markdown("#### 📏 Quy Đổi Đơn Vị Khoảng Cách")
        dist_factors = {
            "Millimet (mm)": 0.001, "Centimet (cm)": 0.01, "Mét (m)": 1.0, "Kilômét (km)": 1000.0,
            "Inch (in)": 0.0254, "Foot (ft)": 0.3048, "Yard (yd)": 0.9144, "Dặm (Mile)": 1609.344
        }
        d_col1, d_col2, d_col3 = st.columns([2, 1.5, 1.5])
        with d_col1: dist_val = st.number_input("Giá trị khoảng cách:", value=1.0, min_value=0.0, step=1.0)
        with d_col2: from_dist = st.selectbox("Từ đơn vị:", options=list(dist_factors.keys()), index=3)
        with d_col3: to_dist = st.selectbox("Sang đơn vị:", options=list(dist_factors.keys()), index=2)
            
        meters = dist_val * dist_factors[from_dist]
        res_dist = meters / dist_factors[to_dist]
        
        st.markdown("---")
        st.markdown(f"### 🎯 Kết Quả: **{dist_val:,.4f} {from_dist}** = **{res_dist:,.4f} {to_dist}**")

    # 5. BỘ CHUYỂN ĐỔI VẬN TỐC (SPEED)
    with subtab_speed:
        st.markdown("#### 🚀 Quy Đổi Đơn Vị Vận Tốc")
        speed_factors = {
            "Mét/giây (m/s)": 1.0, "Kilômét/giờ (km/h)": 1 / 3.6,
            "Dặm/giờ (mph)": 0.44704, "Hải lý/giờ (Knot)": 0.514444
        }
        s_col1, s_col2, s_col3 = st.columns([2, 1.5, 1.5])
        with s_col1: speed_val = st.number_input("Giá trị vận tốc:", value=100.0, min_value=0.0, step=5.0)
        with s_col2: from_speed = st.selectbox("Từ đơn vị:", options=list(speed_factors.keys()), index=1)
        with s_col3: to_speed = st.selectbox("Sang đơn vị:", options=list(speed_factors.keys()), index=0)
            
        ms_val = speed_val * speed_factors[from_speed]
        res_speed = ms_val / speed_factors[to_speed]
        
        st.markdown("---")
        st.markdown(f"### 🎯 Kết Quả: **{speed_val:,.2f} {from_speed}** = **{res_speed:,.2f} {to_speed}**")

    # 6. KHỐI LƯỢNG & NHIỆT ĐỘ
    with subtab_mass_temp:
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            with st.container(border=True):
                st.markdown("##### ⚖️ Quy Đổi Khối Lượng")
                mass_factors = {
                    "Gram (g)": 0.001, "Kilôgram (kg)": 1.0, "Tấn": 1000.0,
                    "Ounce (oz)": 0.0283495, "Pound (lb)": 0.453592
                }
                m_val = st.number_input("Khối lượng:", value=1.0, min_value=0.0, key="m_val_in")
                m_from = st.selectbox("Từ:", options=list(mass_factors.keys()), index=1, key="m_from_sel")
                m_to = st.selectbox("Sang:", options=list(mass_factors.keys()), index=4, key="m_to_sel")
                
                kg_val = m_val * mass_factors[m_from]
                res_mass = kg_val / mass_factors[m_to]
                st.info(f"👉 **{m_val:,.2f} {m_from}** = **{res_mass:,.2f} {m_to}**")
                
        with m_col2:
            with st.container(border=True):
                st.markdown("##### 🌡️ Quy Đổi Nhiệt Độ")
                temp_val = st.number_input("Nhiệt độ:", value=37.0, key="temp_val_in")
                t_from = st.selectbox("Từ:", options=["Độ C (°C)", "Độ F (°F)", "Kelvin (K)"], index=0, key="t_from_sel")
                t_to = st.selectbox("Sang:", options=["Độ C (°C)", "Độ F (°F)", "Kelvin (K)"], index=1, key="t_to_sel")
                
                if "°C" in t_from: c_temp = temp_val
                elif "°F" in t_from: c_temp = (temp_val - 32) * 5 / 9
                else: c_temp = temp_val - 273.15
                
                if "°C" in t_to: res_temp = c_temp
                elif "°F" in t_to: res_temp = (c_temp * 9 / 5) + 32
                else: res_temp = c_temp + 273.15
                
                st.info(f"👉 **{temp_val:,.1f} {t_from}** = **{res_temp:,.1f} {t_to}**")

# --- FOOTER ---
st.markdown("""
<div class="saas-footer">
    ScriptPro Enterprise Edition • Designed for Mai Han Team
</div>
""", unsafe_allow_html=True)
