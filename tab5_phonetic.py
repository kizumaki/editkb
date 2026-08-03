import streamlit as st
import io
import time
import pandas as pd
from gtts import gTTS
from utils import PHONETIC_DB_FILE, save_json_db

def generate_english_audio(text_to_speak, accent='com'):
    try:
        tts = gTTS(text=text_to_speak, lang='en', tld=accent)
        fp = io.BytesIO(); tts.write_to_fp(fp); fp.seek(0)
        return fp
    except Exception as e:
        st.error(f"Không thể tải âm thanh: {e}")
        return None

def render_tab5():
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
        with c1: tab_add_eng = st.text_input("Từ Tiếng Anh gốc:", placeholder="VD: Burger", key=f"tab_add_eng_{st.session_state.get('pho_input_key', 0)}")
        with c2: tab_add_pho = st.text_input("Phiên âm giọng Nam:", placeholder="VD: Bơ-gơ", key=f"tab_add_pho_{st.session_state.get('pho_input_key', 0)}")
        with c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm vào Database", use_container_width=True, type="primary"):
                if tab_add_eng and tab_add_pho:
                    k = tab_add_eng.upper().strip(); v = tab_add_pho.strip()
                    st.session_state['custom_phonetics'][k] = v
                    save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                    st.session_state['pho_input_key'] = st.session_state.get('pho_input_key', 0) + 1
                    st.success(f"✅ Đã thêm thành công: `{tab_add_eng}` ➔ `{tab_add_pho}`"); time.sleep(1); st.rerun()
                else: st.warning("Vui lòng điền đủ 2 ô!")

    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 📑 Danh sách toàn bộ Từ phiên âm đã lưu")
        search_query = st.text_input("🔍 Tìm kiếm từ Tiếng Anh hoặc Từ phiên âm:", placeholder="Gõ từ cần tìm ở đây...").strip().upper()
        all_phonetics_dict = st.session_state.get('custom_phonetics', {})
        
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
