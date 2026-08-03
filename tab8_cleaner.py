import streamlit as st
import io
import os
import re
import zipfile
from docx import Document
from docx.shared import Pt
from utils import clean_and_normalize_text

def render_tab8():
    st.subheader("🧹 DỌN DẸP & CHUẨN HÓA PHỤ ĐỀ (TEXT NORMALIZER)")
    st.markdown("Tự động làm sạch kịch bản rác, bóc tách thẻ HTML/ASS rác, sửa lỗi gõ phím và khoảng trắng thừa.")

    subtab_paste_clean, subtab_file_clean = st.tabs([
        "✍️ Xử Lý & Dọn Dẹp Văn Bản Trực Tiếp", 
        "📁 Dọn Dẹp & Chuẩn Hóa File Hàng Loạt (.srt / .docx)"
    ])

    with subtab_paste_clean:
        col_opt1, col_opt2, col_opt3, col_opt4 = st.columns(4)
        with col_opt1: opt_strip_html = st.checkbox("🏷️ Xóa thẻ HTML", value=True)
        with col_opt2: opt_fix_punct = st.checkbox("✍️ Sửa dấu câu", value=True)
        with col_opt3: opt_cap_first = st.checkbox("🔤 Viết hoa đầu câu", value=True)
        with col_opt4: opt_remove_dash = st.checkbox("✂️ Xóa dấu - đầu câu", value=True)

        col_text_in, col_text_out = st.columns(2)
        with col_text_in:
            input_raw_text = st.text_area("Nội dung rác đầu vào:", value="<font color=\"red\">Cory :</font> Tránh xa  đồng đội tui ra !", height=200)
            btn_do_clean = st.button("🧹 THỰC THI DỌN DẸP", type="primary", use_container_width=True)

        with col_text_out:
            if btn_do_clean and input_raw_text:
                cleaned_res = clean_and_normalize_text(
                    input_raw_text, 
                    strip_all_tags=opt_strip_html, 
                    fix_punctuation=opt_fix_punct, 
                    normalize_spaces=True, 
                    capitalize_first=opt_cap_first, 
                    remove_leading_dash=opt_remove_dash
                )
                st.text_area("Kết quả làm sạch:", value=cleaned_res, height=200)
            else:
                st.text_area("Kết quả làm sạch:", value="", height=200)

    with subtab_file_clean:
        uploaded_clean_files = st.file_uploader("Tải file .srt/.docx cần giặt sạch:", type=['srt', 'docx'], accept_multiple_files=True)
        if uploaded_clean_files:
            st.info(f"Đã chọn **{len(uploaded_clean_files)}** file.")
