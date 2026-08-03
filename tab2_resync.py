import streamlit as st
import io
import os
from utils import clean_file_name_for_output
from tab1_script import process_tab1_docx

def render_tab2(enable_colors, enable_phonetic, enable_cast):
    st.subheader("🔄 RE-SYNC KỊCH BẢN ĐÃ BIÊN TẬP (CỠ PHÔNG 14PT)")
    st.markdown("Dành cho biên tập viên tải lại file kịch bản đã chỉnh sửa để re-sync chuẩn hóa toàn bộ phông chữ Times New Roman 14pt.")
    
    up_key = f"tab2_file_{st.session_state.get('resync_uploader_key', 0)}"
    uploaded_file = st.file_uploader("Tải file Kịch bản cần Re-sync (.docx):", type=['docx'], key=up_key)
    
    if uploaded_file is not None:
        file_name_without_ext = os.path.splitext(uploaded_file.name)[0]
        
        if st.button("🔄 2. BẮT ĐẦU RE-SYNC CỠ CHỮ 14PT", type="primary", use_container_width=True):
            try:
                with st.spinner("Đang Re-sync kịch bản..."):
                    out_docx = process_tab1_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic, enable_cast, font_size_pt=14)
                    st.session_state['r_processed_docx'] = out_docx
                    st.session_state['r_processed_name'] = file_name_without_ext
                st.success("✅ Re-sync hoàn tất mượt mà!")
            except Exception as e:
                st.error(f"Đã có lỗi xảy ra khi Re-sync: {e}")
                
        if 'r_processed_docx' in st.session_state:
            st.markdown("---")
            st.markdown("### 📥 Tải File Re-Sync")
            fn = clean_file_name_for_output(st.session_state.get('r_processed_name', 'KichBan'), tag="_resync", ext=".docx")
            st.download_button(
                label=f"⬇️ TẢI FILE WORD RE-SYNC 14PT ({fn})",
                data=st.session_state['r_processed_docx'],
                file_name=fn,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
