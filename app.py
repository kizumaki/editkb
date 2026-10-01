import streamlit as st
import base64
import hmac
import os
import re
import time
from datetime import datetime, timedelta, timezone
import pandas as pd
from collections import Counter

from utils import (
    init_databases, sheets_enabled, save_json_db, clear_kept_files, build_backup_excel, NON_SPEAKER_DB_FILE, SPEAKER_DB_FILE,
    PHONETIC_DB_FILE, extract_phrases_from_file,
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
# 1a. LOGO MAI HAN TEAM
# ==========================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_BLUE = os.path.join(APP_DIR, "logo.png")         # logo màu gốc: dùng trên nền sáng
LOGO_WHITE = os.path.join(APP_DIR, "logo_white.png")  # logo trắng: dùng trên nền tối / dải băng xanh

def _is_dark_theme():
    try: return st.context.theme.type == "dark"
    except Exception: return False

@st.cache_data(show_spinner=False)
def _logo_base64(path):
    with open(path, "rb") as f: return base64.b64encode(f.read()).decode()

_logo_for_bg = LOGO_WHITE if _is_dark_theme() else LOGO_BLUE
st.logo(_logo_for_bg, size="large")

# ==========================================
# 1b. MẬT KHẨU VÀO APP (đặt APP_PASSWORD trong mục Secrets của Streamlit)
# ==========================================
def _get_app_password():
    try: return str(st.secrets.get("APP_PASSWORD", ""))
    except Exception: return ""

_app_password = _get_app_password()
if _app_password and not st.session_state.get("_authenticated"):
    _c_left, _c_mid, _c_right = st.columns([1, 1.2, 1])
    with _c_mid:
        st.image(_logo_for_bg, width=260)
        st.markdown("#### 🔒 ScriptPro Enterprise Studio")
        _login_box = st.container()
    with _login_box, st.form("login_form"):
        _entered = st.text_input("Nhập mật khẩu để sử dụng:", type="password")
        if st.form_submit_button("Vào", type="primary"):
            if hmac.compare_digest(_entered.encode("utf-8"), _app_password.encode("utf-8")):
                st.session_state["_authenticated"] = True
                st.rerun()
            else:
                st.error("Sai mật khẩu.")
    st.stop()

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

# Giữ giá trị các ô nhập quan trọng khi chuyển trang (Streamlit mặc định xoá ô của trang không mở)
for _k in ("resync_project_week", "dual_default_spk", "dual_hide_default_spk", "textarea_clean_output", "tab1_mode", "tab2_mode"):
    if _k in st.session_state: st.session_state[_k] = st.session_state[_k]

# Tải toàn bộ kho dữ liệu (Google Sheets khi chạy online, file JSON khi chạy trên máy) — 1 lần mỗi phiên
init_databases()

# ==========================================
# 3. UNIFIED SIDEBAR (CONTROL PANEL)
# ==========================================
st.sidebar.markdown("### Bảng điều khiển")

# Trạng thái kho dữ liệu
if st.session_state.get("_db_load_error"):
    st.sidebar.error("⚠️ Không tải được dữ liệu từ Google Sheets. Hãy tải lại trang (F5). Nếu vẫn lỗi, báo người quản lý.\n\n"
                     f"Chi tiết: {st.session_state['_db_load_error'][:300]}")
elif sheets_enabled():
    st.sidebar.success("☁️ Dữ liệu đang lưu trên Google Sheets")
else:
    st.sidebar.info("💻 Chạy trên máy: dữ liệu lưu thành file trong thư mục app")

if st.session_state.get("_db_save_error"):
    st.error(st.session_state.pop("_db_save_error"))

ui_theme_choice = st.sidebar.radio(
    "Kiểu dải băng đầu trang:",
    options=["Xanh Mai Han (mặc định)", "Tối, tương phản cao"],
    index=0,
    help="Chỉ đổi màu dải băng đầu trang. Không ảnh hưởng đến file xuất ra."
)

if st.sidebar.button("🔄 Bắt đầu phiên mới", use_container_width=True,
                     help="Xoá các file đã tải lên (ở mọi trang) và kết quả vừa xử lý để làm việc khác. Dữ liệu đã lưu (phiên âm, phân vai, lương...) KHÔNG bị xoá."):
    for key in ['processed_docx', 'processed_ass', 'processed_srt', 'actor_zip', 'stats',
                'r_processed_docx', 'r_processed_ass', 'r_processed_srt', 'r_actor_zip', 'resync_stats',
                'dual_vn_edits', 'dual_vn_edits_sig', 'manual_cleaned_orig_len', 'manual_cleaned_res_len',
                'batch_result_goc', 'batch_result_resync']:
        if key in st.session_state: del st.session_state[key]
    st.session_state['textarea_clean_output'] = ""
    clear_kept_files()  # xoá các file đã tải ở mọi trang
    st.session_state['uploader_key'] += 1
    st.session_state['resync_uploader_key'] += 1
    st.session_state['bulk_uploader_key'] += 1
    if 'bulk_spk_results' in st.session_state: del st.session_state['bulk_spk_results']
    if 'bulk_eng_results' in st.session_state: del st.session_state['bulk_eng_results']
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("#### Khi xuất kịch bản")
enable_colors = st.sidebar.toggle("🌈 Tô màu nhân vật", value=True)
enable_phonetic = st.sidebar.toggle("🗣️ Phiên âm giọng Nam", value=True, help="Tự động chèn phiên âm giọng Nam trước từ Tiếng Anh (ngoặc đơn + tô màu vàng)")
enable_cast = st.sidebar.toggle("🎭 Phân vai lồng tiếng", value=True, help="Hiển thị thông tin diễn viên lồng tiếng ở đầu trang và lần xuất hiện đầu tiên của nhân vật")

st.sidebar.markdown("---")
st.sidebar.markdown("#### Kho từ dùng chung")

# KHỐI QUÉT KHO SRT/SCRIPT TỔNG HỢP (THÊM NÚT CHỌN TẤT CẢ & BỎ CHỌN SIÊU TỐC)
with st.sidebar.expander("📦 Quét nhiều file tìm tên & từ mới", expanded=False):
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

    # 1. BẢNG TÊN VAI MỚI - BẢNG BẢO VỆ SIÊU NHẸ CÓ NÚT "CHỌN TẤT CẢ" VÀ "BỎ CHỌN"
    if 'bulk_spk_results' in st.session_state and st.session_state['bulk_spk_results']:
        st.markdown("---")
        st.markdown("##### 👤 Tên Vai Mới Phát Hiện")
        new_spks = [s for s, c in st.session_state['bulk_spk_results'].items() if s not in st.session_state.get('custom_speakers', set())]
        
        if new_spks:
            st.caption(f"Tìm thấy **{len(new_spks)}** tên vai mới:")
            
            # Khởi tạo trạng thái chọn tất cả / bỏ chọn
            if 'spk_select_all_state' not in st.session_state:
                st.session_state['spk_select_all_state'] = True
            if 'spk_editor_version' not in st.session_state:
                st.session_state['spk_editor_version'] = 0

            col_spk_1, col_spk_2 = st.columns(2)
            if col_spk_1.button("✅ Chọn tất cả", key="btn_spk_all", use_container_width=True):
                st.session_state['spk_select_all_state'] = True
                st.session_state['spk_editor_version'] += 1
                st.rerun()
            if col_spk_2.button("❌ Bỏ chọn", key="btn_spk_none", use_container_width=True):
                st.session_state['spk_select_all_state'] = False
                st.session_state['spk_editor_version'] += 1
                st.rerun()

            df_spks = pd.DataFrame({"Lưu": [st.session_state['spk_select_all_state']] * len(new_spks), "Tên Vai": new_spks})
            edited_spk_df = st.data_editor(
                df_spks,
                column_config={
                    "Lưu": st.column_config.CheckboxColumn("Lưu?", default=True),
                    "Tên Vai": st.column_config.TextColumn("Tên Vai", disabled=True)
                },
                hide_index=True,
                height=220,
                use_container_width=True,
                key=f"data_editor_spks_{st.session_state['spk_editor_version']}"
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

    # 2. BẢNG TỪ TIẾNG ANH MỚI - BẢNG BẢO VỆ SIÊU NHẸ CÓ NÚT "CHỌN TẤT CẢ" VÀ "BỎ CHỌN"
    if 'bulk_eng_results' in st.session_state and st.session_state['bulk_eng_results']:
        st.markdown("---")
        st.markdown("##### 🔤 Từ Tiếng Anh Mới Phát Hiện")
        existing_pho = st.session_state.get('custom_phonetics', {})
        new_words = [w for w in st.session_state['bulk_eng_results'] if w.upper() not in existing_pho]
        
        if new_words:
            st.caption(f"Tìm thấy **{len(new_words)}** từ Tiếng Anh mới:")
            
            # Khởi tạo trạng thái chọn tất cả / bỏ chọn
            if 'eng_select_all_state' not in st.session_state:
                st.session_state['eng_select_all_state'] = True
            if 'eng_editor_version' not in st.session_state:
                st.session_state['eng_editor_version'] = 0

            col_eng_1, col_eng_2 = st.columns(2)
            if col_eng_1.button("✅ Chọn tất cả", key="btn_eng_all", use_container_width=True):
                st.session_state['eng_select_all_state'] = True
                st.session_state['eng_editor_version'] += 1
                st.rerun()
            if col_eng_2.button("❌ Bỏ chọn", key="btn_eng_none", use_container_width=True):
                st.session_state['eng_select_all_state'] = False
                st.session_state['eng_editor_version'] += 1
                st.rerun()

            df_words = pd.DataFrame({"Lưu": [st.session_state['eng_select_all_state']] * len(new_words), "Từ Tiếng Anh": new_words})
            edited_eng_df = st.data_editor(
                df_words,
                column_config={
                    "Lưu": st.column_config.CheckboxColumn("Lưu?", default=True),
                    "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh", disabled=True)
                },
                hide_index=True,
                height=220,
                use_container_width=True,
                key=f"data_editor_eng_{st.session_state['eng_editor_version']}"
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

with st.sidebar.expander("🎭 Danh sách tên nhân vật", expanded=False):
    st.caption("Những tên này luôn được nhận là người nói.")
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

with st.sidebar.expander("🚫 Cụm từ KHÔNG phải tên nhân vật", expanded=False):
    st.caption("VD: \"Round 1:\", \"Update:\"... có dấu hai chấm nhưng không phải người nói.")
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

st.sidebar.markdown("---")
with st.sidebar.expander("💾 Sao lưu dữ liệu", expanded=False):
    st.caption("Tải về 1 file Excel chứa bản MỚI NHẤT của toàn bộ dữ liệu (phiên âm, phân vai, màu, xưng hô, lương...). "
               "Nên tải định kỳ, ví dụ cuối mỗi tuần, và cất ở nơi an toàn.")
    if st.button("📥 Chuẩn bị bản sao lưu", use_container_width=True, key="btn_prepare_backup"):
        try:
            with st.spinner("Đang đọc dữ liệu mới nhất..."):
                st.session_state["_backup_file"] = build_backup_excel().getvalue()
                st.session_state["_backup_time"] = datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d_%Hh%M")
        except Exception as e:
            st.error(f"Chưa tạo được bản sao lưu: {e}")
    if st.session_state.get("_backup_file"):
        st.download_button("⬇️ Tải file sao lưu (.xlsx)", data=st.session_state["_backup_file"],
                           file_name=f"SaoLuu_ScriptPro_{st.session_state['_backup_time']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary", key="dl_backup")

# ==========================================
# 4. DYNAMIC CSS INJECTION THEO SKINS
# ==========================================
if "Tối" in ui_theme_choice:
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
            background: linear-gradient(135deg, #183E80 0%, #2A63BA 100%);
            padding: 2.5rem 2rem; border-radius: 16px; color: white; margin-bottom: 2rem;
            box-shadow: 0 10px 25px -5px rgba(42, 99, 186, 0.35);
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
_vn_hour = datetime.now(timezone(timedelta(hours=7))).hour
_greeting = ("Chào buổi sáng" if 4 <= _vn_hour < 11 else "Chào buổi trưa" if _vn_hour < 13
             else "Chào buổi chiều" if _vn_hour < 18 else "Chào buổi tối")

st.markdown(f"""
<style>
    .hero-container {{ display: flex; align-items: center; gap: 1.75rem; flex-wrap: wrap;
                       padding: 1.4rem 1.75rem; margin-bottom: 1rem; }}
    .hero-logo {{ height: 76px; width: auto; flex-shrink: 0; }}
    .hero-text {{ flex: 1; min-width: 240px; }}
    .hero-greeting {{ font-size: 0.95rem; opacity: 0.85; margin-bottom: 0.15rem; }}
    .hero-title {{ font-size: 1.75rem; line-height: 1.2; }}
    .hero-subtitle {{ font-size: 0.98rem; margin-top: 0.3rem; }}
    .app-footer {{ color: #94A3B8; font-size: 0.8rem; text-align: center; margin-top: 3rem; }}
    @media (max-width: 640px) {{
        .hero-logo {{ height: 56px; }}
        .hero-title {{ font-size: 1.4rem; }}
        .hero-container {{ padding: 1.1rem 1.1rem; gap: 0.9rem; }}
    }}
</style>
<div class="hero-container">
    <img class="hero-logo" src="data:image/png;base64,{_logo_base64(LOGO_WHITE)}" alt="Mai Han Team">
    <div class="hero-text">
        <div class="hero-greeting">{_greeting}, team Mai Han 👋</div>
        <div class="hero-title">ScriptPro Studio</div>
        <div class="hero-subtitle">Xử lý kịch bản lồng tiếng, phân vai, phiên âm và báo cáo thù lao — gọn trong một chỗ.</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 6. MENU CHỨC NĂNG (4 NHÓM) — mỗi lần chỉ chạy trang đang mở cho nhẹ
# ==========================================
pages = {
    "Kịch bản": [
        st.Page(lambda: render_tab1(enable_colors, enable_phonetic, enable_cast), title="Xử lý kịch bản gốc",
                icon=":material/description:", default=True),  # trang mặc định luôn ở địa chỉ gốc "/"
        st.Page(lambda: render_tab2(enable_colors, enable_phonetic, enable_cast), title="Re-Sync bản đã biên tập",
                icon=":material/sync:", url_path="re-sync"),
    ],
    "Kiểm tra chất lượng": [
        st.Page(lambda: render_tab6(enable_colors, enable_phonetic, enable_cast), title="Đối chiếu 2 file tiếng Anh",
                icon=":material/compare_arrows:", url_path="doi-chieu"),
        st.Page(render_tab7, title="Soát xưng hô & thuật ngữ", icon=":material/fact_check:", url_path="soat-xung-ho"),
    ],
    "Quản lý": [
        st.Page(render_tab3, title="Theo dõi & báo cáo lương", icon=":material/payments:", url_path="luong"),
        st.Page(render_tab4, title="Phân vai & màu nhân vật", icon=":material/theater_comedy:", url_path="phan-vai"),
        st.Page(render_tab5, title="Kho phiên âm giọng Nam", icon=":material/record_voice_over:", url_path="phien-am"),
    ],
    "Công cụ": [
        st.Page(render_tab8, title="Dọn dẹp phụ đề", icon=":material/cleaning_services:", url_path="don-phu-de"),
        st.Page(render_tab9, title="Chuyển đổi định dạng", icon=":material/swap_horiz:", url_path="chuyen-doi"),
    ],
}
st.navigation(pages, position="top").run()

st.markdown('<div class="app-footer">© Mai Han Team</div>', unsafe_allow_html=True)
