import streamlit as st
import io
import os
import pandas as pd

# Import các hàm dùng chung từ utils.py
from utils import (
    load_json_db, save_json_db, NON_SPEAKER_DB_FILE, SPEAKER_DB_FILE, 
    PHONETIC_DB_FILE, CAST_DB_FILE, TRACKER_DB_FILE, RATES_DB_FILE, 
    PRONOUN_REL_DB_FILE, SPEAKER_COLOR_DB_FILE, DEFAULT_CAST_MAPPING, 
    DEFAULT_FIXED_SPEAKER_COLORS, DEFAULT_SOUTH_VIETNAM_PHONETICS
)

# ==========================================
# 1. CẤU HÌNH TRANG CHỦ STREAMLIT
# ==========================================
st.set_page_config(
    page_title="ScriptPro Enterprise - Subtitle & Script Editor",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. KHỞI TẠO SESSION STATE
# ==========================================
if 'uploader_key' not in st.session_state: st.session_state['uploader_key'] = 0
if 'resync_uploader_key' not in st.session_state: st.session_state['resync_uploader_key'] = 0
if 'spk_input_key' not in st.session_state: st.session_state['spk_input_key'] = 0
if 'ns_input_key' not in st.session_state: st.session_state['ns_input_key'] = 0
if 'pho_input_key' not in st.session_state: st.session_state['pho_input_key'] = 0
if 'cast_input_key' not in st.session_state: st.session_state['cast_input_key'] = 0
if 'pronoun_input_key' not in st.session_state: st.session_state['pronoun_input_key'] = 0
if 'color_input_key' not in st.session_state: st.session_state['color_input_key'] = 0

if 'custom_non_speakers' not in st.session_state: st.session_state['custom_non_speakers'] = load_json_db(NON_SPEAKER_DB_FILE, set())
if 'custom_speakers' not in st.session_state: st.session_state['custom_speakers'] = load_json_db(SPEAKER_DB_FILE, set())

if 'custom_phonetics' not in st.session_state:
    loaded_pho = load_json_db(PHONETIC_DB_FILE, DEFAULT_SOUTH_VIETNAM_PHONETICS)
    st.session_state['custom_phonetics'] = {**DEFAULT_SOUTH_VIETNAM_PHONETICS, **loaded_pho}

if 'custom_cast_mapping' not in st.session_state:
    loaded_cast = load_json_db(CAST_DB_FILE, DEFAULT_CAST_MAPPING)
    st.session_state['custom_cast_mapping'] = {**DEFAULT_CAST_MAPPING, **loaded_cast}

if 'fixed_speaker_colors' not in st.session_state:
    st.session_state['fixed_speaker_colors'] = load_json_db(SPEAKER_COLOR_DB_FILE, DEFAULT_FIXED_SPEAKER_COLORS)

if 'dubbing_tracker' not in st.session_state: st.session_state['dubbing_tracker'] = load_json_db(TRACKER_DB_FILE, [])

if 'payroll_rates' not in st.session_state:
    st.session_state['payroll_rates'] = load_json_db(RATES_DB_FILE, {"mode": "minute", "unit_rate": 30000})

# ==========================================
# 3. SIDEBAR PANEL
# ==========================================
st.sidebar.markdown("### ⚡ Control Panel")

ui_theme_choice = st.sidebar.radio(
    "Lựa chọn Skin hiển thị:",
    options=["Mai Han Standard (Mặc định)", "Enterprise Pro (Tối ưu tương phản)"],
    index=0
)

if st.sidebar.button("🔄 Reset phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_docx', 'processed_ass', 'processed_srt', 'actor_zip', 'r_processed_docx', 'r_processed_ass', 'r_processed_srt', 'r_actor_zip']:
        if key in st.session_state: del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.session_state['resync_uploader_key'] += 1
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("#### 🎛️ Bật/Tắt Tính năng")
enable_colors = st.sidebar.toggle("🌈 Tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm giọng Nam", value=True)
enable_cast = st.sidebar.toggle("🎭 Phân vai lồng tiếng", value=True)

# ==========================================
# 4. HERO BANNER
# ==========================================
banner_html = f"""
<div style="background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); padding: 2rem; border-radius: 14px; color: white; margin-bottom: 1.5rem;">
    <div style="font-size: 0.8rem; font-weight: bold; background: #0284C7; display: inline-block; padding: 3px 10px; border-radius: 4px; margin-bottom: 8px;">{ui_theme_choice}</div>
    <div style="font-size: 2.2rem; font-weight: 800;">&#127916; ScriptPro Enterprise Studio</div>
    <div style="font-size: 1rem; color: #94A3B8; margin-top: 4px;">Hệ thống xử lý kịch bản lồng tiếng, chuẩn hóa định dạng Word, phân vai & báo cáo thù lao cá nhân thông minh.</div>
</div>
"""
st.markdown(banner_html, unsafe_allow_html=True)

# ==========================================
# 5. KHUNG MÀN HÌNH CHÍNH 9 TABS
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "🎬 Xử lý Kịch bản Gốc", 
    "🔄 Re-Sync Kịch Bản Biên Tập",
    "📋 Theo dõi & Báo cáo Lương",
    "🎭 Bảng Phân Vai Lồng Tiếng", 
    "📚 Kho Database Phiên Âm Giọng Nam",
    "🔀 Đối Chiếu 2 File Tiếng Anh & QC Dịch",
    "🔎 Soát Bất Nhất Thuật Ngữ & Xưng Hô",
    "🧹 Dọn Dẹp & Chuẩn Hóa Phụ Đề",
    "🧰 Bộ Công Cụ Chuyển Đổi"
])

with tab1:
    st.info("Đang kết nối Tab 1...")

# Các Tab tiếp theo sẽ được gọi từ các file tương ứng ở bước sau!
