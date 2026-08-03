import streamlit as st
import re
import pandas as pd
from collections import Counter
from utils import PRONOUN_REL_DB_FILE, save_json_db, ENGLISH_WORD_REGEX
from tab6_dual_align import parse_any_script_file_to_df

VN_SELF_PRONOUNS = ["tui", "tôi", "mình", "tao", "ta", "em", "anh", "chị", "cháu", "con", "tại hạ", "bản thân"]
VN_TARGET_PRONOUNS = ["ông", "bạn", "mày", "anh", "chị", "chú", "bác", "cậu", "bà", "cưng", "em", "ní", "mấy ní", "sư huynh", "huynh", "đệ"]

def render_tab7():
    st.subheader("🔎 SOÁT BẤT NHẤT THUẬT NGỮ & QUAN HỆ XƯNG HÔ NHÂN VẬT")
    st.markdown("Vùng làm việc phát hiện các câu thoại bị sượng xưng hô hoặc bất nhất bản dịch thuật ngữ/tên riêng.")

    subtab_pronoun, subtab_glossary = st.tabs([
        "👥 Quản Lý & Soát Lỗi Xưng Hô Nhân Vật", 
        "📚 Soát Bất Nhất Thuật Ngữ & Món Ăn/Tên Riêng"
    ])

    with subtab_pronoun:
        st.markdown("#### 1. Bảng Thiết Lập Quan Hệ Xưng Hô")
        col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns([2, 2, 1.5, 1.5, 1.2])
        with col_p1: rel_spk_a = st.text_input("Speaker A:", placeholder="VD: TYLER", key=f"p_spk_a_{st.session_state.get('pronoun_input_key', 0)}").strip().upper()
        with col_p2: rel_spk_b = st.text_input("Speaker B:", placeholder="VD: BILL", key=f"p_spk_b_{st.session_state.get('pronoun_input_key', 0)}").strip().upper()
        with col_p3: rel_self = st.text_input("Xưng (Self):", placeholder="VD: tui...", key=f"p_self_{st.session_state.get('pronoun_input_key', 0)}").strip().lower()
        with col_p4: rel_target = st.text_input("Gọi (Target):", placeholder="VD: ông...", key=f"p_target_{st.session_state.get('pronoun_input_key', 0)}").strip().lower()
        with col_p5:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Xưng Hô", type="primary", use_container_width=True, key="btn_add_pronoun"):
                if rel_spk_a and rel_spk_b and rel_self and rel_target:
                    key_pair = f"{rel_spk_a}|{rel_spk_b}"
                    st.session_state['custom_pronoun_rel'][key_pair] = {"self": rel_self, "target": rel_target}
                    save_json_db(PRONOUN_REL_DB_FILE, st.session_state['custom_pronoun_rel'])
                    st.session_state['pronoun_input_key'] = st.session_state.get('pronoun_input_key', 0) + 1
                    st.success("✅ Đã lưu quan hệ xưng hô!")

        st.markdown("---")
        uploaded_pronoun_script = st.file_uploader("Tải file Kịch bản Tiếng Việt (.srt/.docx) để soát xưng hô:", type=['srt', 'docx'], key="uploader_pronoun_qc")
        if uploaded_pronoun_script:
            c_spks_p = st.session_state.get('custom_speakers', set())
            c_non_spks_p = st.session_state.get('custom_non_speakers', set())
            df_p_script = parse_any_script_file_to_df(uploaded_pronoun_script.getvalue(), uploaded_pronoun_script.name, c_spks_p, c_non_spks_p)

            if not df_p_script.empty:
                st.success(f"✅ Đã quét **{len(df_p_script)}** câu thoại!")
                st.dataframe(df_p_script[['Start', 'End', 'Speaker', 'Dialogue']], use_container_width=True)

    with subtab_glossary:
        st.markdown("#### 📚 Soát Bất Nhất Thuật Ngữ & Bản Dịch")
        uploaded_glossary_script = st.file_uploader("Tải file Kịch bản (.srt/.docx) để kiểm tra thuật ngữ:", type=['srt', 'docx'], key="uploader_glossary_qc")
        if uploaded_glossary_script:
            c_spks_g = st.session_state.get('custom_speakers', set())
            c_non_spks_g = st.session_state.get('custom_non_speakers', set())
            df_g_script = parse_any_script_file_to_df(uploaded_glossary_script.getvalue(), uploaded_glossary_script.name, c_spks_g, c_non_spks_g)
            if not df_g_script.empty:
                st.success(f"✅ Phân tích thuật ngữ cho **{len(df_g_script)}** câu thoại!")
                st.dataframe(df_g_script[['Start', 'End', 'Speaker', 'Dialogue']], use_container_width=True)
