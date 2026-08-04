import streamlit as st
import io
import os
import re
import time
import pandas as pd
from collections import Counter

from utils import (
    load_json_db, save_json_db, NON_SPEAKER_DB_FILE, SPEAKER_DB_FILE, 
    PHONETIC_DB_FILE, CAST_DB_FILE, TRACKER_DB_FILE, RATES_DB_FILE, 
    PRONOUN_REL_DB_FILE, SPEAKER_COLOR_DB_FILE, DEFAULT_CAST_MAPPING, 
    DEFAULT_FIXED_SPEAKER_COLORS, DEFAULT_SOUTH_VIETNAM_PHONETICS, extract_phrases_from_file,
    scan_candidate_speakers, scan_english_words_in_dialogue
)

from tab1_script import render_tab1
from tab2_resync import render_tab2
from tab3_payroll import render_tab3
from tab4_cast_color import render_tab4
from tab5_phonetic import render_tab5
from tab6_dual_align import render_tab6
from tab7_consistency import render_tab7
from tab8_cleaner import render_tab8
from tab9_tools import render_tab9

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
if 'bulk_uploader_key' not in st.session_state: st.session_state['bulk_uploader_key'] = 0
if 'spk_input_key' not in st.session_state: st.session_state['spk_input_key'] = 0
if 'ns_input_key' not in st.session_state: st.session_state['ns_input_key'] = 0
if 'pho_input_key' not in st.session_state: st.session_state['pho_input_key'] = 0
if 'cast_input_key' not in st.session_state: st.session_state['cast_input_key'] = 0
if 'pronoun_input_key' not in st.session_state: st.session_state['pronoun_input_key'] = 0
if 'color_input_key' not in st.session_state: st.session_state['color_input_key'] = 0
if 'textarea_clean_output' not in st.session_state: st.session_state['textarea_clean_output'] = ""

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

if 'custom_pronoun_rel' not in st.session_state:
    default_pronouns = {
        "TYLER|BILL": {"self": "tui", "target": "ông"},
        "CORY|EASTON": {"self": "tui", "target": "ông"},
        "COBY|COACH RAC": {"self": "tui", "target": "ông"}
    }
    st.session_state['custom_pronoun_rel'] = load_json_db(PRONOUN_REL_DB_FILE, default_pronouns)

if 'dubbing_tracker' not in st.session_state: st.session_state['dubbing_tracker'] = load_json_db(TRACKER_DB_FILE, [])

if 'payroll_rates' not in st.session_state:
    st.session_state['payroll_rates'] = load_json_db(RATES_DB_FILE, {"mode": "minute", "unit_rate": 30000})

# ==========================================
# 3. UNIFIED SIDEBAR (CONTROL PANEL)
# ==========================================
st.sidebar.markdown("### ⚡ Control Panel")

ui_theme_choice = st.sidebar.radio(
    "Lựa chọn Skin hiển thị:",
    options=["Mai Han Standard (Mặc định)", "Enterprise Pro (Tối ưu tương phản)"],
    index=0,
    help="Chế độ 'Mai Han Standard' giữ nguyên 100% giao diện truyền thống. Chế độ 'Enterprise Pro' mang lại phong cách Studio hiện đại, rõ nét và tương phản cao."
)

if st.sidebar.button("🔄 Reset phiên làm việc", use_container_width=True, type="primary"):
    for key in ['processed_docx', 'processed_ass', 'processed_srt', 'actor_zip', 'r_processed_docx', 'r_processed_ass', 'r_processed_srt', 'r_actor_zip']:
        if key in st.session_state: del st.session_state[key]
    st.session_state['uploader_key'] += 1
    st.session_state['resync_uploader_key'] += 1
    st.session_state['bulk_uploader_key'] += 1
    if 'bulk_spk_results' in st.session_state: del st.session_state['bulk_spk_results']
    if 'bulk_eng_results' in st.session_state: del st.session_state['bulk_eng_results']
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("#### 🎛️ Bật/Tắt Tính năng")
enable_colors = st.sidebar.toggle("🌈 Tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm giọng Nam", value=True, help="Tự động chèn phiên âm giọng Nam trước từ Tiếng Anh (ngoặc đơn + tô màu vàng)")
enable_cast = st.sidebar.toggle("🎭 Phân vai lồng tiếng", value=True, help="Hiển thị thông tin diễn viên lồng tiếng ở đầu trang và lần xuất hiện đầu tiên của nhân vật")

st.sidebar.markdown("---")
st.sidebar.markdown("#### 💾 Database Quản Lý Cụm Từ")

# KHỐI QUÉT KHO SRT/SCRIPT TỔNG HỢP (DÙNG BẢNG ẢO CHỐNG ĐƠ TRÌNH DUYỆT 100%)
with st.sidebar.expander("📦 Quét Kho SRT/Script Tổng Hợp", expanded=False):
    st.caption("Nạp hàng loạt file (.srt, .docx, .xlsx, .txt) để bóc tách Tên Vai & Từ Tiếng Anh cùng lúc.")
    
    if st.button("🗑️ Dọn dẹp danh sách file", key="btn_clear_bulk_scan", use_container_width=True):
        st.session_state['bulk_uploader_key'] += 1
        if 'bulk_spk_results' in st.session_state: del st.session_state['bulk_spk_results']
        if 'bulk_eng_results' in st.session_state: del st.session_state['bulk_eng_results']
        st.success("🧹 Đã làm sạch danh sách file!")
        time.sleep(0.5)
        st.rerun()

    bulk_files = st.file_uploader(
        "Kéo thả danh sách file vào đây:", 
        type=["srt", "docx", "txt", "xlsx"], 
        accept_multiple_files=True,
        key=f"bulk_srt_scanner_{st.session_state['bulk_uploader_key']}"
    )
    
    if bulk_files and st.button("🚀 Bóc tách Tổng hợp", key="btn_run_bulk_scan", use_container_width=True):
        all_candidate_speakers = Counter()
        all_english_words = set()
        
        custom_spks = st.session_state.get('custom_speakers', set())
        custom_non_spks = st.session_state.get('custom_non_speakers', set())
        
        with st.spinner(f"Đang quét {len(bulk_files)} file..."):
            for uploaded_file in bulk_files:
                spk_cand = scan_candidate_speakers(uploaded_file, custom_spks, custom_non_spks)
                all_candidate_speakers.update(spk_cand)
                
                eng_words = scan_english_words_in_dialogue(uploaded_file, custom_spks, custom_non_spks)
                all_english_words.update(eng_words)
        
        st.session_state['bulk_spk_results'] = all_candidate_speakers
        st.session_state['bulk_eng_results'] = sorted(list(all_english_words), key=lambda x: x.upper())
        st.success(f"✅ Đã quét xong {len(bulk_files)} file!")

    # 1. BẢNG TÊN VAI MỚI - BẢNG HIỂN THỊ SIÊU NHẸ (Virtual Scroll)
    if 'bulk_spk_results' in st.session_state and st.session_state['bulk_spk_results']:
        st.markdown("---")
        st.markdown("##### 👤 Tên Vai Mới Phát Hiện")
        new_spks = [s for s, c in st.session_state['bulk_spk_results'].items() if s not in st.session_state.get('custom_speakers', set())]
        
        if new_spks:
            st.caption(f"Tìm thấy **{len(new_spks)}** tên vai mới. Bỏ tích chọn `[ ]` những từ rác trực tiếp trong bảng:")
            df_spks = pd.DataFrame({"Lưu": [True] * len(new_spks), "Tên Vai": new_spks})
            edited_spk_df = st.data_editor(
                df_spks,
                column_config={
                    "Lưu": st.column_config.CheckboxColumn("Lưu?", default=True),
                    "Tên Vai": st.column_config.TextColumn("Tên Vai", disabled=True)
                },
                hide_index=True,
                height=220,
                use_container_width=True,
                key="data_editor_spks"
            )
            selected_spks = edited_spk_df[edited_spk_df["Lưu"] == True]["Tên Vai"].tolist()
            
            if st.button(f"➕ Thêm ({len(selected_spks)}) Vai đã chọn vào Whitelist", use_container_width=True):
                if selected_spks:
                    st.session_state['custom_speakers'].update(selected_spks)
                    save_json_db(SPEAKER_DB_FILE, st.session_state['custom_speakers'])
                    st.success(f"🎉 Đã lưu {len(selected_spks)} tên vai vào Whitelist!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("⚠️ Bạn chưa chọn tên vai nào!")
        else:
            st.info("Tất cả tên vai đều đã có trong Whitelist.")

    # 2. BẢNG TỪ TIẾNG ANH MỚI - BẢNG HIỂN THỊ SIÊU NHẸ (Virtual Scroll)
    if 'bulk_eng_results' in st.session_state and st.session_state['bulk_eng_results']:
        st.markdown("---")
        st.markdown("##### 🔤 Từ Tiếng Anh Mới Phát Hiện")
        existing_pho = st.session_state.get('custom_phonetics', {})
        new_words = [w for w in st.session_state['bulk_eng_results'] if w.upper() not in existing_pho]
        
        if new_words:
            st.caption(f"Tìm thấy **{len(new_words)}** từ Tiếng Anh mới. Bỏ tích chọn `[ ]` những từ rác trực tiếp trong bảng:")
            df_words = pd.DataFrame({"Lưu": [True] * len(new_words), "Từ Tiếng Anh": new_words})
            edited_eng_df = st.data_editor(
                df_words,
                column_config={
                    "Lưu": st.column_config.CheckboxColumn("Lưu?", default=True),
                    "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh", disabled=True)
                },
                hide_index=True,
                height=220,
                use_container_width=True,
                key="data_editor_eng"
            )
            selected_words = edited_eng_df[edited_eng_df["Lưu"] == True]["Từ Tiếng Anh"].tolist()
            
            if st.button(f"➕ Thêm ({len(selected_words)}) Từ đã chọn vào Kho Phiên Âm", use_container_width=True):
                if selected_words:
                    for w in selected_words:
                        st.session_state['custom_phonetics'][w.upper()] = w
                    save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                    st.success(f"🎉 Đã lưu {len(selected_words)} từ vào Kho Phiên Âm!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning("⚠️ Bạn chưa chọn từ nào!")
        else:
            st.info("Tất cả từ Tiếng Anh đều đã có trong Database Phiên Âm.")

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

# ==========================================
# 4. DYNAMIC CSS INJECTION THEO SKINS
# ==========================================
if "Enterprise Pro" in ui_theme_choice:
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; color: #0F172A; }
        .hero-container {
            background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
            padding: 2.2rem 2rem; border-radius: 14px; color: #FFFFFF; margin-bottom: 1.8rem;
            border-left: 6px solid #38BDF8; box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        }
        .hero-title { font-size: 2.3rem; font-weight: 800; margin: 0; color: #FFFFFF; }
        .hero-subtitle { font-size: 1.05rem; color: #94A3B8; margin-top: 0.4rem; }
        .badge-pro {
            background-color: #0284C7; color: #FFFFFF; padding: 4px 12px;
            border-radius: 6px; font-size: 0.75rem; font-weight: 700;
            text-transform: uppercase; display: inline-block; margin-bottom: 0.6rem;
        }
        .metric-card {
            background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px;
            padding: 1.25rem; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }
        .metric-label { font-size: 0.8rem; color: #475569; font-weight: 700; text-transform: uppercase; }
        .metric-value { font-size: 1.8rem; font-weight: 800; color: #0F172A; margin-top: 0.2rem; }
        .qc-card-warning {
            background-color: #FEF2F2; border-left: 5px solid #DC2626; color: #991B1B; padding: 12px 16px; border-radius: 8px; margin-bottom: 10px;
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
        .hero-title { font-size: 2.4rem; font-weight: 800; margin: 0; }
        .hero-subtitle { font-size: 1.05rem; opacity: 0.9; margin-top: 0.5rem; }
        .badge-pro {
            background-color: rgba(255, 255, 255, 0.2); backdrop-filter: blur(8px);
            padding: 4px 12px; border-radius: 9999px; font-size: 0.8rem; font-weight: 600;
            text-transform: uppercase; display: inline-block; margin-bottom: 0.8rem;
        }
        .metric-card {
            background: white; border: 1px solid #E2E8F0; border-radius: 12px; padding: 1.25rem;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        }
        .metric-label { font-size: 0.85rem; color: #64748B; font-weight: 500; text-transform: uppercase; }
        .metric-value { font-size: 1.8rem; font-weight: 700; color: #0F172A; margin-top: 0.25rem; }
        .qc-card-warning { background-color: #FEF2F2; border-left: 4px solid #EF4444; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 5. HERO BANNER
# ==========================================
st.markdown(f"""
<div class="hero-container">
    <div class="badge-pro">{ui_theme_choice}</div>
    <div class="hero-title">🎬 ScriptPro Enterprise Studio</div>
    <div class="hero-subtitle">Hệ thống xử lý kịch bản lồng tiếng, chuẩn hóa định dạng Word, phân vai & báo cáo thù lao cá nhân thông minh.</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 6. MÀN HÌNH CHÍNH TÁCH 9 TABS
# ==========================================
tab_script, tab_resync, tab_dub_tracker, tab_cast_db, tab_phonetic_db, tab_dual_align, tab_consistency, tab_cleaner, tab_tools = st.tabs([
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

with tab_script: render_tab1(enable_colors, enable_phonetic, enable_cast)
with tab_resync: render_tab2(enable_colors, enable_phonetic, enable_cast)
with tab_dub_tracker: render_tab3()
with tab_cast_db: render_tab4()
with tab_phonetic_db: render_tab5()
with tab_dual_align: render_tab6(enable_colors, enable_phonetic, enable_cast)
with tab_consistency: render_tab7()
with tab_cleaner: render_tab8()
with tab_tools: render_tab9()

st.markdown('<div class="saas-footer">Copyright © Mai Han Team. All Rights Reserved.</div>', unsafe_allow_html=True)
