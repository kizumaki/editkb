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
import pandas as pd # Thư viện để đọc Excel

# --- CẤU HÌNH TRANG CHỦ STREAMLIT ---
st.set_page_config(page_title="Pro Script Editor", page_icon="🎬", layout="wide")

# Khởi tạo Session State cho ứng dụng
if 'reset_key' not in st.session_state:
    st.session_state['reset_key'] = 0
if 'custom_non_speakers' not in st.session_state:
    st.session_state['custom_non_speakers'] = set()

# --- DANH SÁCH LỌC TỪ NHIỄU MẶC ĐỊNH ---
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

# Gom danh sách mặc định và danh sách tùy chỉnh
NON_SPEAKER_PHRASES = DEFAULT_NON_SPEAKER_PHRASES.union(st.session_state['custom_non_speakers'])

SPEAKER_REGEX_DELIMITER = re.compile(r"([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)

# --- HELPER FUNCTIONS ---
def extract_phrases_from_file(file_io, file_name):
    """Trích xuất từ ngữ từ các định dạng file khác nhau"""
    phrases = set()
    try:
        if file_name.endswith('.txt'):
            content = file_io.getvalue().decode("utf-8")
            phrases.update([line.strip().upper() for line in content.split('\n') if line.strip()])
        elif file_name.endswith('.docx'):
            doc = Document(io.BytesIO(file_io.getvalue()))
            for p in doc.paragraphs:
                if p.text.strip():
                    # Tách bằng dấu phẩy hoặc xuống dòng
                    parts = re.split(r'[,\n]', p.text)
                    phrases.update([part.strip().upper() for part in parts if part.strip()])
        elif file_name.endswith('.xlsx'):
            df = pd.read_excel(file_io, header=None)
            for col in df.columns:
                for item in df[col].dropna():
                    parts = re.split(r'[,\n]', str(item))
                    phrases.update([part.strip().upper() for part in parts if part.strip()])
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

def format_and_split_dialogue(document, text, enable_colors, speaker_color_map, used_colors, stats_counter):
    parts = SPEAKER_REGEX_DELIMITER.split(text)
    TAB_STOP_POSITION = Inches(1.0)
    
    if len(parts) == 1:
        new_paragraph = document.add_paragraph()
        new_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        new_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        new_paragraph.add_run('\t')
        new_paragraph.paragraph_format.space_after = Pt(0)
        new_paragraph.paragraph_format.space_before = Pt(0)
        apply_html_formatting_to_run(new_paragraph, text)
        return

    speaker_matches = list(SPEAKER_REGEX_DELIMITER.finditer(text))
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
            continuation_paragraph.paragraph_format.space_after = Pt(0)
            continuation_paragraph.paragraph_format.space_before = Pt(0)
            apply_html_formatting_to_run(continuation_paragraph, leading_content)

        if speaker_name.upper() in NON_SPEAKER_PHRASES:
            content_block = text[start:]
            continuation_paragraph = document.add_paragraph()
            continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
            continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
            continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
            continuation_paragraph.add_run('\t')
            apply_html_formatting_to_run(continuation_paragraph, content_block)
            continuation_paragraph.paragraph_format.space_after = Pt(0)
            continuation_paragraph.paragraph_format.space_before = Pt(0)
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

        if content: apply_html_formatting_to_run(new_paragraph, content)
        new_paragraph.paragraph_format.space_after = Pt(0)
        new_paragraph.paragraph_format.space_before = Pt(0)
        last_processed_index = next_match_start
    
    remaining_content = text[last_processed_index:].strip()
    if remaining_content:
        continuation_paragraph = document.add_paragraph()
        continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        continuation_paragraph.add_run('\t')
        continuation_paragraph.paragraph_format.space_after = Pt(0)
        continuation_paragraph.paragraph_format.space_before = Pt(0)
        apply_html_formatting_to_run(continuation_paragraph, remaining_content)

# --- MAIN PROCESSING ---
def process_docx(uploaded_file, file_name_without_ext, enable_colors):
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_200]
    random.shuffle(used_colors)
    stats_counter = Counter()
    
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
        for match in SPEAKER_REGEX_DELIMITER.finditer(text):
            speaker_name = match.group(1).strip()
            if speaker_name.upper() not in NON_SPEAKER_PHRASES and speaker_name not in unique_speakers:
                unique_speakers.append(speaker_name)
            
    if unique_speakers:
        speaker_list_text = "VAI: " + ", ".join(unique_speakers)
        p = document.add_paragraph(speaker_list_text)
        p.runs[0].font.name = 'Times New Roman'
        p.runs[0].font.size = Pt(12)
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
        else:
            format_and_split_dialogue(document, text, enable_colors, speaker_color_map, used_colors, stats_counter)
            
    progress_bar.progress(100)
    status_text.text("Định dạng hoàn tất!")
    time.sleep(0.5)
    progress_bar.empty()
    status_text.empty()
            
    for paragraph in document.paragraphs[start_index:]:
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
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
    
    return modified_file, stats

def clean_file_name_for_output(original_filename):
    name_without_ext = os.path.splitext(original_filename)[0]
    cleaned = re.sub(r'(CONVERTED_|FORMATTED_|\s*\(.*\)$|_edit$)', '', name_without_ext, flags=re.IGNORECASE).strip()
    return f"{cleaned}_edit.docx"

# --- SIDEBAR (THANH ĐIỀU HƯỚNG) ---
st.sidebar.title("⚙️ Tùy chỉnh (Settings)")

# Nút Làm Mới Phiên Làm Việc
if st.sidebar.button("🔄 Làm mới phiên làm việc", use_container_width=True, type="primary"):
    # Xóa các biến trạng thái cũ
    for key in ['processed_file', 'new_filename', 'stats']:
        if key in st.session_state:
            del st.session_state[key]
    st.session_state['reset_key'] += 1 # Thay đổi key uploader để reset
    st.rerun() # Tải lại toàn bộ trang

st.sidebar.markdown("---")
enable_colors = st.sidebar.toggle("🌈 Bật tô màu nhân vật", value=True)

# Giao diện thêm Từ Nhiễu (NON_SPEAKER_PHRASES)
with st.sidebar.expander("🚫 Quản lý Từ nhiễu (Non-speaker)", expanded=False):
    st.markdown("Thêm các cụm từ (Tiếng Anh/Việt) bị app nhận diện nhầm thành tên nhân vật. (VD: CẢNH 1, MÁY QUAY, CHÚ Ý...)")
    
    manual_input = st.text_area("Nhập thủ công (cách nhau bằng dấu phẩy hoặc xuống dòng):", height=100)
    
    upload_non_speaker = st.file_uploader(
        "Hoặc tải file lên (.txt, .docx, .xlsx)", 
        type=['txt', 'docx', 'xlsx'], 
        key=f"ns_uploader_{st.session_state['reset_key']}" # Tự làm mới khi nhấn Nút Reset
    )
    
    if st.button("Cập nhật vào hệ thống", use_container_width=True):
        new_phrases = set()
        
        # Xử lý Text Area
        if manual_input:
            parts = re.split(r'[,\n]', manual_input)
            new_phrases.update([p.strip().upper() for p in parts if p.strip()])
            
        # Xử lý File tải lên
        if upload_non_speaker:
            new_phrases.update(extract_phrases_from_file(upload_non_speaker, upload_non_speaker.name))
            
        if new_phrases:
            st.session_state['custom_non_speakers'].update(new_phrases)
            st.success(f"Đã cập nhật thêm {len(new_phrases)} cụm từ!")
            time.sleep(1)
            st.rerun()
        else:
            st.warning("Vui lòng nhập từ hoặc tải file lên!")

    # Hiển thị số lượng từ tùy chỉnh đã thêm
    if len(st.session_state['custom_non_speakers']) > 0:
        st.info(f"Đã tự thêm: **{len(st.session_state['custom_non_speakers'])}** từ nhiễu.")
        if st.button("🗑️ Xóa từ nhiễu tùy chỉnh"):
            st.session_state['custom_non_speakers'] = set()
            st.rerun()

# --- GIAO DIỆN CHÍNH (UI) ---
st.title("🎬 Kịch Bản Pro - Word Script Editor")
st.markdown("Hệ thống tự động biên tập và làm đẹp kịch bản tiêu chuẩn quốc tế.")
st.markdown("---")

col1, col2 = st.columns([1.5, 1])

with col1:
    st.subheader("📁 1. Tải lên kịch bản")
    
    # Gán key động để file uploader tự reset khi nhấn nút Làm mới
    uploaded_file = st.file_uploader(
        "Kéo thả file Word (.docx) của bạn vào đây", 
        type=['docx'], 
        key=f"main_uploader_{st.session_state['reset_key']}"
    )

    if uploaded_file is not None:
        original_filename = uploaded_file.name
        file_name_without_ext = os.path.splitext(original_filename)[0] 
        
        st.success(f"Đã nhận file: **{original_filename}**")
        
        if st.button("✨ 2. BẮT ĐẦU ĐỊNH DẠNG TỰ ĐỘNG", use_container_width=True):
            try:
                modified_file_io, stats = process_docx(uploaded_file, file_name_without_ext, enable_colors)
                new_filename = clean_file_name_for_output(original_filename)
                
                st.session_state['processed_file'] = modified_file_io
                st.session_state['new_filename'] = new_filename
                st.session_state['stats'] = stats
                
            except Exception as e:
                st.error(f"Đã có lỗi xảy ra: {e}")

        if 'processed_file' in st.session_state:
            st.download_button(
                label="⬇️ 3. TẢI FILE KỊCH BẢN ĐÃ CHUẨN HÓA",
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
