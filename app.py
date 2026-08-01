import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
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

def load_json_db(filepath):
    """Đọc dữ liệu từ file Database JSON"""
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def save_json_db(filepath, data_set):
    """Ghi dữ liệu mới vào Database JSON"""
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(list(data_set), f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Không thể lưu vào Database: {e}")

# Khởi tạo Session State
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0
if 'spk_input_key' not in st.session_state:
    st.session_state['spk_input_key'] = 0
if 'ns_input_key' not in st.session_state:
    st.session_state['ns_input_key'] = 0

if 'custom_non_speakers' not in st.session_state:
    st.session_state['custom_non_speakers'] = load_json_db(NON_SPEAKER_DB_FILE)

if 'custom_speakers' not in st.session_state:
    st.session_state['custom_speakers'] = load_json_db(SPEAKER_DB_FILE)

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

def apply_html_formatting_to_run(paragraph, current_text):
    if not current_text.strip(): return
    matches = list(HTML_CONTENT_REGEX.finditer(current_text))
    last_end = 0
    for match in matches:
        tag_text = match.group(2)
        start, end = match.span()
        if start > last_end:
            paragraph.add_run(current_text[last_end:start])
        run_html = paragraph.add_run(tag_text)
        run_html.font.bold = True
        run_html.font.italic = True
        last_end = end
    if last_end < len(current_text):
        paragraph.add_run(current_text[last_end:])

def format_and_split_dialogue(document, text, enable_colors, speaker_color_map, used_colors, stats_counter, speaker_regex, preview_html_list):
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
        apply_html_formatting_to_run(new_paragraph, text)
        
        # Sửa CSS xem trước: Dùng padding-left + text-indent để KHÔNG bị cấn lề trái
        preview_html_list.append(
            f"<div style='padding-left: 100px; font-family: \"Times New Roman\"; font-size: 15px; line-height: 1.5; color: #111; margin-bottom: 2px; word-wrap: break-word;'>"
            f"{text}"
            f"</div>"
        )
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
            apply_html_formatting_to_run(continuation_paragraph, leading_content)
            
            preview_html_list.append(
                f"<div style='padding-left: 100px; font-family: \"Times New Roman\"; font-size: 15px; line-height: 1.5; color: #111; margin-bottom: 2px; word-wrap: break-word;'>"
                f"{leading_content}"
                f"</div>"
            )

        if speaker_name.upper() in NON_SPEAKER_PHRASES:
            content_block = text[start:]
            continuation_paragraph = document.add_paragraph()
            continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
            continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
            continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
            continuation_paragraph.add_run('\t')
            continuation_paragraph.paragraph_format.space_before = Pt(0)
            continuation_paragraph.paragraph_format.space_after = Pt(0)
            apply_html_formatting_to_run(continuation_paragraph, content_block)
            
            preview_html_list.append(
                f"<div style='padding-left: 100px; font-family: \"Times New Roman\"; font-size: 15px; line-height: 1.5; color: #111; margin-bottom: 2px; word-wrap: break-word;'>"
                f"{content_block}"
                f"</div>"
            )
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
        
        color_obj = get_speaker_color(speaker_name, speaker_color_map, used_colors)
        color_hex = f"rgb({color_obj[0]}, {color_obj[1]}, {color_obj[2]})" if enable_colors else "#000000"
        
        if enable_colors:
            run_speaker.font.color.rgb = color_obj
        
        if len(speaker_full) > 10:
             new_paragraph.add_run('\t\t')
        else:
             new_paragraph.add_run('\t')

        if content: apply_html_formatting_to_run(new_paragraph, content)
        new_paragraph.paragraph_format.space_before = Pt(0)
        new_paragraph.paragraph_format.space_after = Pt(0)
        last_processed_index = next_match_start
        
        # Căn chỉnh Hanging Indent chuẩn HTML không mất lề trái:
        preview_html_list.append(
            f"<div style='padding-left: 100px; text-indent: -100px; font-family: \"Times New Roman\"; font-size: 15px; line-height: 1.5; color: #111; margin-bottom: 2px; word-wrap: break-word;'>"
            f"<b style='color: {color_hex}; display: inline-block; min-width: 90px;'>{speaker_full}</b>&nbsp;&nbsp;{content}"
            f"</div>"
        )
    
    remaining_content = text[last_processed_index:].strip()
    if remaining_content:
        continuation_paragraph = document.add_paragraph()
        continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        continuation_paragraph.add_run('\t')
        continuation_paragraph.paragraph_format.space_before = Pt(0)
        continuation_paragraph.paragraph_format.space_after = Pt(0)
        apply_html_formatting_to_run(continuation_paragraph, remaining_content)
        
        preview_html_list.append(
            f"<div style='padding-left: 100px; font-family: \"Times New Roman\"; font-size: 15px; line-height: 1.5; color: #111; margin-bottom: 2px; word-wrap: break-word;'>"
            f"{remaining_content}"
            f"</div>"
        )

# --- MAIN PROCESSING ---
def process_docx(uploaded_file, file_name_without_ext, enable_colors):
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_200]
    random.shuffle(used_colors)
    stats_counter = Counter()
    preview_html_list = []
    
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
    
    preview_html_list.append(f"<h2 style='text-align: center; color: #000; font-family: \"Times New Roman\"; font-weight: bold; font-size: 22px; margin-bottom: 10px;'>{title_text}</h2>")
    
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
        
        preview_html_list.append(f"<div style='font-family: \"Times New Roman\"; color: #000; font-size: 15px; font-weight: bold; margin-bottom: 15px;'>{speaker_list_text}</div>")
    
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
            
            preview_html_list.append(f"<div style='font-family: \"Times New Roman\"; color: #222; font-weight: bold; font-size: 15px; margin-top: 12px; margin-bottom: 2px;'>{text}</div>")
        else:
            format_and_split_dialogue(document, text, enable_colors, speaker_color_map, used_colors, stats_counter, speaker_regex, preview_html_list)
            
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
            run.font.size = Pt(12)
        
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    stats = {
        "total_speakers": len(unique_speakers),
        "total_lines": sum(stats_counter.values()),
        "top_speaker": stats_counter.most_common(1)[0] if stats_counter else ("Không có", 0)
    }
    
    full_preview_html = "".join(preview_html_list)
    
    return modified_file, stats, full_preview_html

def clean_file_name_for_output(original_filename):
    name_without_ext = os.path.splitext(original_filename)[0]
    cleaned = re.sub(r'(CONVERTED_|FORMATTED_|\s*\(.*\)$|_edit$)', '', name_without_ext, flags=re.IGNORECASE).strip()
    return f"{cleaned}_edit.docx"

# --- SIDEBAR (THANH ĐIỀU HƯỚNG) ---
st.sidebar.title("⚙️ Tùy chỉnh (Settings)")

if st.sidebar.button("🔄 Làm mới phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_file', 'new_filename', 'stats', 'preview_html']:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.rerun()

st.sidebar.markdown("---")
enable_colors = st.sidebar.toggle("🌈 Bật tô màu nhân vật", value=True)

# 1. Quản lý Người Nói Ưu Tiên (WHITELIST)
with st.sidebar.expander("🎭 Database Người nói (Whitelist)", expanded=False):
    manual_spk_input = st.text_area(
        "Nhập thủ công (cách nhau bằng dấu phẩy hoặc xuống dòng):", 
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
        st.info("📌 **Vui lòng tải file kịch bản (.docx) ở trên để hiển thị phần Xem trước & Soát lỗi.**")
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

        st.markdown("### 🔍 SOÁT LỖI & XEM TRƯỚC TỰ ĐỘNG")
        st.markdown("Kiểm tra nhanh danh sách nhận diện trước khi bấm định dạng:")
        
        tab_spk, tab_non_spk = st.tabs(["🎭 Nhận diện là NGƯỜI NÓI", "🚫 Đang bị xem là TỪ NHIỄU"])
        
        with tab_spk:
            if detected_speakers:
                st.write(", ".join([f"`{s}`" for s in detected_speakers]))
                st.markdown("---")
                to_move_to_ns = st.multiselect(
                    "Phát hiện từ nào bị nhận diện sai? Chọn bên dưới để LƯU VÀO DATABASE TỪ NHIỄU:",
                    options=[name for name in candidates.keys() if name.upper() not in NON_SPEAKER_PHRASES],
                    key="select_to_ns"
                )
                if st.button("➡️ Đưa các từ chọn vào Database TỪ NHIỄU", type="secondary"):
                    if to_move_to_ns:
                        new_items = [item.upper() for item in to_move_to_ns]
                        st.session_state['custom_non_speakers'].update(new_items)
                        save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                        st.success(f"✅ Đã lưu thành công {len(new_items)} từ vào Database Từ Nhiễu!")
                        time.sleep(1)
                        st.rerun()
            else:
                st.info("Chưa tìm thấy cụm từ người nói nào.")

        with tab_non_spk:
            if detected_non_speakers:
                st.write(", ".join([f"`{s}`" for s in detected_non_speakers]))
                st.markdown("---")
                to_move_to_spk = st.multiselect(
                    "Từ nào thực ra là NGƯỜI NÓI? Chọn bên dưới để LƯU VÀO DATABASE NGƯỜI NÓI:",
                    options=[name for name in candidates.keys() if name.upper() in NON_SPEAKER_PHRASES],
                    key="select_to_spk"
                )
                if st.button("➡️ Đưa các từ chọn vào Database NGƯỜI NÓI", type="secondary"):
                    if to_move_to_spk:
                        st.session_state['custom_speakers'].update(to_move_to_spk)
                        save_json_db(SPEAKER_DB_FILE, st.session_state['custom_speakers'])
                        
                        for item in to_move_to_spk:
                            st.session_state['custom_non_speakers'].discard(item.upper())
                        save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                        
                        st.success(f"✅ Đã lưu thành công {len(to_move_to_spk)} tên vào Database Người Nói!")
                        time.sleep(1)
                        st.rerun()
            else:
                st.info("Không có cụm từ nào bị loại vào danh sách từ nhiễu.")

        st.markdown("---")
        if st.button("✨ 2. BẮT ĐẦU ĐỊNH DẠNG TỰ ĐỘNG", use_container_width=True, type="primary"):
            try:
                modified_file_io, stats, preview_html = process_docx(uploaded_file, file_name_without_ext, enable_colors)
                new_filename = clean_file_name_for_output(original_filename)
                
                st.session_state['processed_file'] = modified_file_io
                st.session_state['new_filename'] = new_filename
                st.session_state['stats'] = stats
                st.session_state['preview_html'] = preview_html
                
            except Exception as e:
                st.error(f"Đã có lỗi xảy ra: {e}")

        # KHU VỰC XEM TRƯỚC VÀ TẢI FILE SAU KHI ĐỊNH DẠNG
        if 'preview_html' in st.session_state:
            st.markdown("---")
            st.subheader("👁️ Xem trước kịch bản (Paper Preview)")
            st.caption("Khung xem trước mô phỏng trang giấy in Word. Bạn có thể kéo lăn chuột lên/xuống và sang trái/phải để kiểm tra:")
            
            # KHUNG XEM TRƯỚC MÔ PHỎNG TỜ GIẤY CÓ THANH CUỘN 2 CHIỀU
            paper_container_html = f"""
            <div style="
                background-color: #ffffff; 
                padding: 35px 40px; 
                border-radius: 8px; 
                border: 1px solid #ccc; 
                max-height: 520px; 
                overflow-y: auto; 
                overflow-x: auto; 
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                white-space: nowrap;
            ">
                <div style="display: inline-block; min-width: 650px;">
                    {st.session_state['preview_html']}
                </div>
            </div>
            """
            st.markdown(paper_container_html, unsafe_allow_html=True)

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
