import streamlit as st
import io
import os
import zipfile
import pandas as pd
from docx import Document
from utils import parse_srt_to_dataframe, apply_excel_styles

def process_srt_to_docx_tool(uploaded_file):
    srt_content = uploaded_file.getvalue().decode('utf-8', errors='ignore')
    blocks = srt_content.strip().split('\n\n')
    doc = Document()
    for b in blocks:
        lines = b.split('\n')
        for l in lines:
            doc.add_paragraph(l)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

def render_tab9():
    st.subheader("🧰 BỘ CÔNG CỤ CHUYỂN ĐỔI (CONVERTER SUITE)")
    subtab_sub_conv, subtab_srt_excel, subtab_curr = st.tabs([
        "🎬 Kịch Bản Subtitle (SRT ⇄ DOCX)",
        "📊 SRT ➔ Excel (.xlsx)",
        "💵 Tiền Tệ (Currency)"
    ])

    with subtab_sub_conv:
        st.markdown("#### 🎬 Chuyển Đổi Phụ Đề Chuyên Nghiệp")
        batch_srt_files = st.file_uploader("Tải file .srt để chuyển sang .docx:", type=['srt'], accept_multiple_files=True, key="tool_srt_to_docx")
        if batch_srt_files:
            if st.button("✨ Chuyển SRT Sang Word", type="primary", use_container_width=True):
                single_f = batch_srt_files[0]
                docx_buf = process_srt_to_docx_tool(single_f)
                st.success("✅ Chuyển đổi thành công!")
                st.download_button("⬇️ Tải file Word (.docx)", data=docx_buf, file_name="Converted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    with subtab_srt_excel:
        uploaded_srt_excel = st.file_uploader("Tải file .srt để xuất Excel:", type=['srt'], key="tool_srt_excel")
        if uploaded_srt_excel:
            try: content = uploaded_srt_excel.read().decode("utf-8")
            except: content = uploaded_srt_excel.read().decode("latin-1")
            df_ex = parse_srt_to_dataframe(content)
            if not df_ex.empty:
                st.dataframe(apply_excel_styles(df_ex), use_container_width=True)
                out_excel = io.BytesIO()
                df_ex.to_excel(out_excel, index=False, engine='openpyxl')
                out_excel.seek(0)
                st.download_button("💾 TẢI FILE EXCEL (.XLSX)", data=out_excel.getvalue(), file_name="Converted_Subtitle.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

    with subtab_curr:
        st.markdown("#### 💵 Quy Đổi Tiền Tệ Đơn Giản")
        val = st.number_input("Số tiền USD:", value=100.0)
        st.success(f"👉 **{val} USD** = **{val * 25400:,.0f} VNĐ**")
