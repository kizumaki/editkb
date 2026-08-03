import streamlit as st
import io
import os
import re
import pandas as pd
from docx import Document
from utils import (
    parse_any_script_file_to_df, align_and_compare_english_scripts, 
    generate_qc_dual_excel, generate_aligned_docx_file
)

def render_tab6(enable_colors, enable_phonetic, enable_cast):
    st.subheader("🔀 ĐỐI CHIẾU 2 FILE TIẾNG ANH & SOÁT SỬA BẢN DỊCH VIỆT")
    st.markdown("So sánh file Tiếng Anh do Mai Han Team nghe với file Tiếng Anh Gốc do Khách gửi trễ.")

    col_spk_fb1, col_spk_fb2 = st.columns([1.8, 1.2])
    with col_spk_fb1:
        default_spk_input = st.text_input(
            "🎭 Tên người nói mặc định (dùng cho Báo cáo QC):", 
            placeholder="VD: Nick, Narrator, MC...", 
            value="", 
            help="Điền tên ở đây để Báo cáo QC Excel hiển thị 'Nick' rõ ràng thay vì 'Unknown:'."
        )

    with col_spk_fb2:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        hide_default_spk_export = st.checkbox(
            "🚫 Không in Tên mặc định vào Kịch bản xuất ra (Word/SRT)", 
            value=True,
            help="File Word và SRT xuất ra sẽ giữ nguyên văn bản gọn gàng không có chữ 'Nick:' ở từng câu."
        )

    col_dual1, col_dual2, col_dual3 = st.columns(3)
    with col_dual1:
        with st.container(border=True):
            st.markdown("##### 📄 1. File Tiếng Anh - Mai Han Team (.srt/docx)")
            uploaded_mh_eng = st.file_uploader("Tải file tiếng Anh - Mai Han Team (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_mh_eng")

    with col_dual2:
        with st.container(border=True):
            st.markdown("##### 📄 2. File Tiếng Anh của Khách (.srt/docx)")
            uploaded_off_eng = st.file_uploader("Tải file tiếng Anh của Khách (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_off_eng")

    with col_dual3:
        with st.container(border=True):
            st.markdown("##### 📄 3. File Tiếng Việt Hiện Tại (Vietnamese Script)")
            uploaded_vn_script = st.file_uploader("Tải file Tiếng Việt (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_vn_script")

    if uploaded_mh_eng is not None and uploaded_off_eng is not None:
        c_spks = st.session_state.get('custom_speakers', set())
        c_non_spks = st.session_state.get('custom_non_speakers', set())
        fallback_spk_name = default_spk_input.strip() if default_spk_input.strip() else "Unknown"

        df_mh_eng = parse_any_script_file_to_df(uploaded_mh_eng.getvalue(), uploaded_mh_eng.name, c_spks, c_non_spks, fallback_spk_name)
        df_off_eng = parse_any_script_file_to_df(uploaded_off_eng.getvalue(), uploaded_off_eng.name, c_spks, c_non_spks, fallback_spk_name)
        df_vn_script = parse_any_script_file_to_df(uploaded_vn_script.getvalue(), uploaded_vn_script.name, c_spks, c_non_spks, fallback_spk_name) if uploaded_vn_script else None

        if not df_mh_eng.empty and not df_off_eng.empty:
            df_aligned = align_and_compare_english_scripts(df_mh_eng, df_off_eng, df_vn_script, fallback_spk_name)

            tot_rows = len(df_aligned)
            missing_cnt = sum(1 for st_v in df_aligned['Trạng thái QC'] if '🔴' in str(st_v))
            word_diff_cnt = sum(1 for st_v in df_aligned['Trạng thái QC'] if '🟡' in str(st_v))
            mismatch_cnt = sum(1 for st_v in df_aligned['Trạng thái QC'] if '🔵' in str(st_v))

            st.markdown("---")
            st.markdown("#### 📊 Báo Cáo QC So Sánh Nội Dung Tiếng Anh")

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                st.markdown(f'<div class="metric-card"><div class="metric-label">💬 Tổng số câu thoại</div><div class="metric-value">{tot_rows}</div></div>', unsafe_allow_html=True)
            with col_m2:
                st.markdown(f'<div class="metric-card"><div class="metric-label">🔴 Bỏ sót câu thoại</div><div class="metric-value" style="color:#DC2626;">{missing_cnt}</div></div>', unsafe_allow_html=True)
            with col_m3:
                st.markdown(f'<div class="metric-card"><div class="metric-label">🟡 Khác từ vựng (AI nghe nhầm)</div><div class="metric-value" style="color:#D97706;">{word_diff_cnt}</div></div>', unsafe_allow_html=True)
            with col_m4:
                st.markdown(f'<div class="metric-card"><div class="metric-label">🔵 Lệch người nói</div><div class="metric-value" style="color:#2563EB;">{mismatch_cnt}</div></div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### 👁️ Workspace Bảng Đối Chiếu Tiếng Anh & Chỉnh Sửa Dịch Tiếng Việt")
            qc_filter = st.selectbox("🔍 Lọc danh sách câu theo Trạng thái QC:", options=["TẤT CẢ CÁC CÂU", "🟡 Chỉ xem câu KHÁC TỪ VỰNG (Cần sửa dịch)", "🔴 Chỉ xem câu BỎ SÓT THOẠI", "🔵 Chỉ xem câu LỆCH VAI"])

            if "🔴" in qc_filter: df_display = df_aligned[df_aligned['Trạng thái QC'].str.contains('🔴', na=False)]
            elif "🟡" in qc_filter: df_display = df_aligned[df_aligned['Trạng thái QC'].str.contains('🟡', na=False)]
            elif "🔵" in qc_filter: df_display = df_aligned[df_aligned['Trạng thái QC'].str.contains('🔵', na=False)]
            else: df_display = df_aligned

            edited_aligned_df = st.data_editor(
                df_display[['Stt', 'Timecode', 'Tiếng Anh Mai Han (AI/Heard)', 'Tiếng Anh Khách (Official)', 'Dịch Tiếng Việt (Cần Sửa)', 'Trạng thái QC', 'Ghi chú QC']],
                column_config={
                    "Stt": st.column_config.NumberColumn("Stt", disabled=True, width="small"),
                    "Timecode": st.column_config.TextColumn("Mốc Timecode Mai Han", disabled=True),
                    "Tiếng Anh Mai Han (AI/Heard)": st.column_config.TextColumn("Bản Anh Mai Han nghe", disabled=True),
                    "Tiếng Anh Khách (Official)": st.column_config.TextColumn("Bản Anh Khách gửi chuẩn", disabled=True),
                    "Dịch Tiếng Việt (Cần Sửa)": st.column_config.TextColumn("Kịch bản Tiếng Việt (Sửa trực tiếp)"),
                    "Trạng thái QC": st.column_config.TextColumn("Trạng thái", disabled=True),
                    "Ghi chú QC": st.column_config.TextColumn("Ghi chú từ vựng khác biệt", disabled=True)
                },
                hide_index=True, use_container_width=True, key="dual_english_editor_table"
            )

            st.markdown("---")
            st.markdown("#### ⬇️ Xuất File Kịch Bản Tiếng Việt Hoàn Chỉnh")
            col_ex1, col_ex2, col_ex3 = st.columns(3)
            base_out_name = os.path.splitext(uploaded_mh_eng.name)[0]

            with col_ex1:
                qc_excel_buf = generate_qc_dual_excel(df_aligned)
                st.download_button(
                    label="📊 TẢI BÁO CÁO ĐỐI CHIẾU EXCEL (.XLSX)", data=qc_excel_buf,
                    file_name=f"{base_out_name}_DoiChieu_English_QC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True
                )

            with col_ex2:
                aligned_docx_buf = generate_aligned_docx_file(
                    df_aligned, base_out_name, enable_colors, enable_phonetic, enable_cast,
                    hide_default_spk=hide_default_spk_export, fallback_spk_name=fallback_spk_name
                )
                st.download_button(
                    label="📄 TẢI WORD VIỆT HOÀN CHỈNH (.DOCX)", data=aligned_docx_buf,
                    file_name=f"{base_out_name}_VI_Final.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary", use_container_width=True
                )

            with col_ex3:
                srt_out_lines = []
                for idx_s, r_s in df_aligned.iterrows():
                    spk_val = r_s['Speaker_MH']
                    is_explicit = r_s.get('Is_Explicit_MH', True)
                    
                    should_show_spk = True
                    if hide_default_spk_export and (not is_explicit or spk_val.upper() == fallback_spk_name.upper() or spk_val.upper() == "UNKNOWN"):
                        should_show_spk = False

                    if should_show_spk and spk_val: spk_text = f"{spk_val}: {r_s['Dialogue_VN']}"
                    else: spk_text = r_s['Dialogue_VN']

                    srt_out_lines.append(f"{idx_s+1}\n{r_s['Timecode']}\n{spk_text}\n")
                srt_out_bytes = "\n".join(srt_out_lines).encode('utf-8-sig')

                st.download_button(
                    label="📝 TẢI SUBTITLE SRT VIỆT (.SRT)", data=srt_out_bytes,
                    file_name=f"{base_out_name}_VI_Final.srt",
                    mime="text/plain", type="primary", use_container_width=True
                )
