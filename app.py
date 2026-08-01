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

# --- CẤU HÌNH TRANG CHỦ STREAMLIT ---
st.set_page_config(page_title="Pro Script Editor", page_icon="🎬", layout="wide")

# --- HÀM ĐỌC / GHI DATABASE DỮ LIỆU CỤC BỘ ---
NON_SPEAKER_DB_FILE = "custom_non_speakers.json"
SPEAKER_DB_FILE = "custom_speakers.json"
PHONETIC_DB_FILE = "custom_phonetics.json"

# Database phiên âm giọng Nam mặc định
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
    # Gộp mặc định vào custom nếu chưa có
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
    """Quét tất cả từ Tiếng Anh trong LỜI THOẠI (Bỏ qua tên người nói)"""
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    eng_found = set()
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text or text.lower().startswith("srt conversion") or TIMECODE_REGEX.match(text):
            continue
        
        # Tách bỏ tên người nói
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

        # Tìm các từ Tiếng Anh
        for match in ENGLISH_WORD_REGEX.finditer(dialogue_content):
            word = match.group(0).strip()
            # Loại bỏ các từ không phải Tiếng Anh hoặc số đơn thuần
            if len(word) > 1 and not word.isdigit():
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
            elif i % 6 == 2: r, g, b = p, q, v
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
    """
    Xử lý chèn Phiên Âm Giọng Nam + Highlight từ Tiếng Anh trong ngoặc ()
    """
    if not current_text.strip(): return
    
    phonetic_db = st.session_state['custom_phonetics']
    
    if not enable_phonetic:
        # Nếu tắt chế độ phiên âm, xử lý HTML tag bình thường
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

    # Sắp xếp các cụm từ Tiếng Anh theo độ dài giảm dần để ưu tiên khớp từ dài trước
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
            
            # 1. Thêm đoạn văn bản tiếng Việt phía trước
            if start > last_end:
                paragraph.add_run(current_text[last_end:start])
                
            # 2. Lấy phiên âm tương ứng
            pho_text = phonetic_db.get(eng_word_original.upper(), eng_word_original)
            
            # 3. Thêm phiên âm giọng Nam
            run_pho = paragraph.add_run(f"{pho_text} ")
            
            # 4. Thêm từ Tiếng Anh gốc đặt trong ngoặc () và HIGHLIGHT VÀNG
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
        
        # TÊN NGƯỜI NÓI - KHÔNG BAO GIỜ PHIÊN ÂM
        run_speaker = new_paragraph.add_run(speaker_full)
        run_speaker.font.bold = True
        
        if enable_colors:
            run_speaker.font.color.rgb = get_speaker_color(speaker_name, speaker_color_map, used_colors)
        
        if len(speaker_full) > 10:
             new_paragraph.add_run('\t\t')
        else:
             new_paragraph.add_run('\t')

        # CHỈ PHIÊN ÂM LỜI THOẠI
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

# --- SIDEBAR (THANH ĐIỀU HƯỚNG) ---
st.sidebar.title("⚙️ Tùy chỉnh (Settings)")

if st.sidebar.button("🔄 Làm mới phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_file', 'new_filename', 'stats']:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.rerun()

st.sidebar.markdown("---")
enable_colors = st.sidebar.toggle("🌈 Bật tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm Tiếng Anh giọng Nam", value=True, help="Tự động chèn phiên âm giọng Nam trước từ Tiếng Anh (ngoặc đơn + tô màu vàng)")

# 1. Quản lý Người Nói Ưu Tiên (WHITELIST)
with st.sidebar.expander("🎭 Database Người nói (Whitelist)", expanded=False):
    manual_spk_input = st.text_area(
        "Nhập thủ công:", 
        height=80, 
        key=f"spk_manual_{st.session_state['spk_input_key']}"
    )
    upload_spk_file = st.file_uploader(
        "Tải file danh sách (.txt, .docx, .xlsx)", 
        type=['txt', 'docx', 'xlsx'], 
        key=f"spk_uploader_{st.session_state['spk_input_key']}"
    )
    
    if st.button("Lưu vào Database Người Nói", use_container_width=True):
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
            st.success(f"✅ Đã lưu {len(new_spks)} người nói vào Database!")
            time.sleep(1)
            st.rerun()

    if len(st.session_state['custom_speakers']) > 0:
        st.info(f"Database đã lưu: **{len(st.session_state['custom_speakers'])}** người nói.")

# 2. Quản lý Từ Nhiễu (BLACKLIST)
with st.sidebar.expander("🚫 Database Từ nhiễu (Non-speaker)", expanded=False):
    manual_input = st.text_area(
        "Nhập thủ công:", 
        height=80, 
        key=f"ns_manual_{st.session_state['ns_input_key']}"
    )
    upload_non_speaker = st.file_uploader(
        "Tải file danh sách (.txt, .docx, .xlsx)", 
        type=['txt', 'docx', 'xlsx'], 
        key=f"ns_uploader_{st.session_state['ns_input_key']}"
    )
    
    if st.button("Lưu vào Database Từ Nhiễu", use_container_width=True):
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
            st.success(f"✅ Đã lưu {len(new_phrases)} từ nhiễu vào Database!")
            time.sleep(1)
            st.rerun()

    if len(st.session_state['custom_non_speakers']) > 0:
        st.info(f"Database đã lưu: **{len(st.session_state['custom_non_speakers'])}** từ nhiễu.")

# --- GIAO DIỆN CHÍNH (UI) ---
st.title("🎬 Kịch Bản Pro - Word Script Editor")
st.markdown("Hệ thống tự động biên tập và làm đẹp kịch bản tiêu chuẩn quốc tế.")
st.markdown("---")

col1, col2 = st.columns([1.5, 1])

with col1:
    st.subheader("📁 1. Tải lên kịch bản")
    uploaded_file = st.file_uploader(
        "Kéo thả file Word (.docx) của bạn vào đây", 
        type=['docx'], 
        key=f"main_uploader_{st.session_state['uploader_key']}"
    )

    if uploaded_file is None:
        st.info("📌 **Vui lòng tải file kịch bản (.docx) ở trên để hiển thị phần Soát lỗi & Quản lý Phiên âm.**")
    else:
        original_filename = uploaded_file.name
        file_name_without_ext = os.path.splitext(original_filename)[0] 
        st.success(f"Đã nhận file: **{original_filename}**")

        speaker_regex = build_speaker_regex(st.session_state['custom_speakers'])
        candidates = scan_candidate_speakers(uploaded_file, speaker_regex)

        detected_speakers = []
        detected_non_speakers = []

        for name, count in candidates.items():
            if name.upper() in NON_SPEAKER_PHRASES:
                detected_non_speakers.append(f"{name} ({count} lần)")
            else:
                detected_speakers.append(f"{name} ({count} lần)")

        st.markdown("### 🔍 SOÁT LỖI NHẬN DIỆN NGƯỜI NÓI")
        
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

        # --- VÙNG LÀM VIỆC: QUẢN LÝ PHIÊN ÂM TIẾNG ANH (GIỌNG NAM) ---
        st.markdown("---")
        st.markdown("### 🗣️ VÙNG LÀM VIỆC: PHIÊN ÂM TIẾNG ANH (GIỌNG NAM)")
        st.caption("Quét và điều chỉnh phiên âm cho tất cả từ Tiếng Anh xuất hiện trong lời thoại kịch bản:")

        detected_eng_words = scan_english_words_in_dialogue(uploaded_file, speaker_regex)

        if detected_eng_words:
            # Tạo bảng dữ liệu 4 cột
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
                    "Đề xuất chỉnh sửa của bạn": st.column_config.TextColumn("Đề xuất phiên âm mới", help="Nhập cách phát âm bạn muốn ở đây"),
                    "Nạp vào Database": st.column_config.CheckboxColumn("Lưu Database?", default=True)
                },
                disabled=["Từ Tiếng Anh", "Phiên âm hiện tại"],
                hide_index=True,
                use_container_width=True,
                key="phonetic_editor_table"
            )

            if st.button("💾 Nạp các chỉnh sửa ở bảng trên vào Database Phiên Âm", type="secondary", use_container_width=True):
                updated_count = 0
                for _, row in edited_df.iterrows():
                    if row["Nạp vào Database"]:
                        eng_k = str(row["Từ Tiếng Anh"]).upper().strip()
                        pho_v = str(row["Đề xuất chỉnh sửa của bạn"]).strip()
                        if pho_v:
                            st.session_state['custom_phonetics'][eng_k] = pho_v
                            updated_count += 1
                
                save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                st.success(f"✅ Đã cập nhật thành công {updated_count} từ phiên âm vào Database!")
                time.sleep(1)
                st.rerun()
        else:
            st.info("Chưa phát hiện từ Tiếng Anh nào trong phần lời thoại của kịch bản này.")

        # --- HÀNG CUỐI: THÊM THỦ CÔNG TỪ PHIÊN ÂM MỚI ---
        st.markdown("#### ➕ Thêm thủ công từ phiên âm mới vào Database")
        add_col1, add_col2, add_col3 = st.columns([2, 2, 1.2])

        with add_col1:
            manual_eng = st.text_input("Từ Tiếng Anh:", placeholder="VD: Starbucks", key=f"manual_eng_{st.session_state['pho_input_key']}")
        with add_col2:
            manual_pho = st.text_input("Phiên âm giọng Nam:", placeholder="VD: Xì-ta-bắc", key=f"manual_pho_{st.session_state['pho_input_key']}")
        with add_col3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Bổ sung ngay", use_container_width=True):
                if manual_eng and manual_pho:
                    k = manual_eng.upper().strip()
                    v = manual_pho.strip()
                    st.session_state['custom_phonetics'][k] = v
                    save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                    st.session_state['pho_input_key'] += 1
                    st.success(f"✅ Đã thêm `{manual_eng}` $\rightarrow$ `{manual_pho}` vào Database!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("Vui lòng nhập đủ 2 ô!")

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
    st.subheader("📊 Bảng điều khiển (Dashboard)")
    st.markdown("Thống kê nhanh kịch bản của bạn")
    
    if 'stats' in st.session_state:
        stats = st.session_state['stats']
        st.metric(label="🎭 Tổng số Nhân vật", value=stats["total_speakers"])
        st.metric(label="💬 Tổng số Câu thoại", value=stats["total_lines"])
        
        top_name, top_count = stats["top_speaker"]
        st.info(f"👑 **Nhân vật nói nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
    else:
        st.info("Bảng thống kê sẽ xuất hiện sau khi bạn xử lý kịch bản.")
