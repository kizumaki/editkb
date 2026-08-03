import streamlit as st
import io
import os
import zipfile
from collections import Counter
import pandas as pd
from docx import Document
from utils import (
    process_srt_to_docx, process_docx_to_srt, parse_srt_to_dataframe, 
    apply_excel_styles, find_all_speaker_tags, save_json_db, NON_SPEAKER_PHRASES, 
    NON_SPEAKER_DB_FILE, SPEAKER_DB_FILE, TIMECODE_REGEX, generate_reaper_region_csv, 
    generate_pro_tools_csv, generate_cmx3600_edl
)

def render_tab9():
    subtab_sub_conv, subtab_srt_excel, subtab_daw_markers, subtab_curr, subtab_dist, subtab_speed, subtab_mass_temp = st.tabs([
        "🎬 Kịch Bản Subtitle (SRT ⇄ DOCX)",
        "📊 SRT ➔ Excel (.xlsx)",
        "🎛️ DAW Marker Timeline",
        "💵 Tiền Tệ (Currency)",
        "📏 Khoảng Cách (Distance)",
        "🚀 Vận Tốc (Speed)",
        "⚖️ Khối Lượng & Nhiệt Độ"
    ])

    with subtab_sub_conv:
        st.markdown("#### 🎬 Bộ Công Cụ Chuyển Đổi Subtitle Chuyên Nghiệp")
        col_c1, col_c2 = st.columns(2)
        
        with col_c1:
            with st.container(border=True):
                st.markdown("##### 📄 1. Chuyển SRT ➔ Word (.docx)")
                batch_srt_files = st.file_uploader("Tải 1 hoặc nhiều file .srt:", type=['srt'], accept_multiple_files=True, key="tool_srt_to_docx_batch")
                
                if batch_srt_files:
                    st.info(f"Đã chọn **{len(batch_srt_files)}** file SRT.")
                    if st.button("✨ Chuyển SRT Sang Word", use_container_width=True, type="primary"):
                        try:
                            if len(batch_srt_files) == 1:
                                single_f = batch_srt_files[0]
                                s_name_no_ext = os.path.splitext(single_f.name)[0]
                                docx_buf = process_srt_to_docx(single_f, s_name_no_ext)
                                st.success("✅ Chuyển đổi hoàn tất!")
                                st.download_button(
                                    label=f"⬇️ Tải {s_name_no_ext}.docx", data=docx_buf,
                                    file_name=f"{s_name_no_ext}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True
                                )
                            else:
                                zip_buf = io.BytesIO()
                                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                    for srt_f in batch_srt_files:
                                        s_name_no_ext = os.path.splitext(srt_f.name)[0]
                                        docx_buf = process_srt_to_docx(srt_f, s_name_no_ext)
                                        zf.writestr(f"{s_name_no_ext}.docx", docx_buf.getvalue())
                                zip_buf.seek(0)
                                st.success(f"✅ Đã chuyển đổi {len(batch_srt_files)} file!")
                                st.download_button(label="📦 Tải Trọn Bộ Word (.ZIP)", data=zip_buf.getvalue(), file_name="Word_Files.zip", mime="application/zip", use_container_width=True)
                        except Exception as e: st.error(f"Lỗi: {e}")

        with col_c2:
            with st.container(border=True):
                st.markdown("##### 📝 2. Chuyển Word (.docx) ➔ SRT (Batch hàng loạt)")
                batch_docx_files = st.file_uploader("Tải 1 hoặc nhiều file .docx:", type=['docx'], accept_multiple_files=True, key="tool_docx_to_srt_batch")
                
                if batch_docx_files:
                    st.info(f"Đã chọn **{len(batch_docx_files)}** file Word.")
                    if st.button("✨ Chuyển Hàng Loạt Sang SRT", use_container_width=True, type="primary"):
                        try:
                            if len(batch_docx_files) == 1:
                                single_f = batch_docx_files[0]
                                s_name_no_ext = os.path.splitext(single_f.name)[0]
                                srt_bytes = process_docx_to_srt(single_f)
                                st.success("✅ Chuyển đổi hoàn tất!")
                                st.download_button(
                                    label=f"⬇️ Tải {s_name_no_ext}.srt", data=srt_bytes,
                                    file_name=f"{s_name_no_ext}.srt", mime="text/plain", use_container_width=True
                                )
                            else:
                                zip_buf = io.BytesIO()
                                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                    for doc_f in batch_docx_files:
                                        s_name_no_ext = os.path.splitext(doc_f.name)[0]
                                        srt_bytes = process_docx_to_srt(doc_f)
                                        zf.writestr(f"{s_name_no_ext}.srt", srt_bytes)
                                zip_buf.seek(0)
                                st.success(f"✅ Đã chuyển đổi {len(batch_docx_files)} file!")
                                st.download_button(label="📦 Tải Trọn Bộ SRT (.ZIP)", data=zip_buf.getvalue(), file_name="SRT_Files.zip", mime="application/zip", use_container_width=True)
                        except Exception as e: st.error(f"Lỗi: {e}")

    with subtab_srt_excel:
        st.markdown("#### 📊 Chuyển Đổi File Subtitle SRT ➔ Bảng Tính Excel (.xlsx)")
        uploaded_srt_excel = st.file_uploader("Tải file .srt của bạn vào đây:", type=['srt'], key="tool_srt_to_excel")
        if uploaded_srt_excel is not None:
            try:
                try: srt_content_excel = uploaded_srt_excel.read().decode("utf-8")
                except UnicodeDecodeError: srt_content_excel = uploaded_srt_excel.read().decode("latin-1")
            except Exception:
                st.error("Lỗi mã hóa file.")
                srt_content_excel = None

            if srt_content_excel:
                custom_spks_ex = st.session_state.get('custom_speakers', set())
                custom_non_spks_ex = st.session_state.get('custom_non_speakers', set())

                df_converted_excel = parse_srt_to_dataframe(srt_content_excel, custom_spks_ex, custom_non_spks_ex)
                if not df_converted_excel.empty:
                    styled_excel_df = apply_excel_styles(df_converted_excel)
                    st.dataframe(styled_excel_df, use_container_width=True)

                    output_excel = io.BytesIO()
                    styled_excel_df.to_excel(output_excel, index=False, engine='openpyxl')
                    output_excel.seek(0)

                    orig_base_name = uploaded_srt_excel.name.rsplit('.', 1)[0]
                    st.download_button(
                        label=f"💾 TẢI FILE EXCEL (.XLSX): {orig_base_name}.xlsx", data=output_excel.getvalue(),
                        file_name=f"{orig_base_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True
                    )

    with subtab_daw_markers:
        st.markdown("#### 🎛️ Tự Động Tạo File Marker Timeline Cho Phần Mềm Thu Âm DAW")
        uploaded_marker_file = st.file_uploader("Tải file Kịch bản (.srt/.docx) của bạn vào đây:", type=['srt', 'docx'], key="tool_marker_uploader")
        
        if uploaded_marker_file is not None:
            m_filename = uploaded_marker_file.name
            m_base_name = os.path.splitext(m_filename)[0]
            custom_spks_m = st.session_state.get('custom_speakers', set())
            custom_non_spks_m = st.session_state.get('custom_non_speakers', set())
            
            if m_filename.endswith('.srt'):
                try: m_srt_text = uploaded_marker_file.read().decode("utf-8")
                except UnicodeDecodeError: m_srt_text = uploaded_marker_file.read().decode("latin-1")
                df_markers = parse_srt_to_dataframe(m_srt_text, custom_spks_m, custom_non_spks_m)
            else:
                s_bytes = process_docx_to_srt(uploaded_marker_file)
                m_srt_text = s_bytes.decode("utf-8", errors="ignore")
                df_markers = parse_srt_to_dataframe(m_srt_text, custom_spks_m, custom_non_spks_m)

            if not df_markers.empty:
                st.success(f"✅ Đã trích xuất **{len(df_markers)}** câu thoại!")
                col_m1, col_m2, col_m3 = st.columns(3)
                
                with col_m1:
                    reaper_csv_str = generate_reaper_region_csv(df_markers)
                    st.download_button("🎧 REAPER (Region CSV)", data=reaper_csv_str.encode('utf-8-sig'), file_name=f"{m_base_name}_Reaper.csv", mime="text/csv", type="primary", use_container_width=True)

                with col_m2:
                    pt_csv_str = generate_pro_tools_csv(df_markers)
                    st.download_button("🎛️ PRO TOOLS (Track Markers)", data=pt_csv_str.encode('utf-8-sig'), file_name=f"{m_base_name}_ProTools.csv", mime="text/csv", type="primary", use_container_width=True)

                with col_m3:
                    edl_str = generate_cmx3600_edl(df_markers)
                    st.download_button("🎬 CMX 3600 EDL (Premiere/Resolve)", data=edl_str.encode('utf-8-sig'), file_name=f"{m_base_name}.edl", mime="text/plain", type="primary", use_container_width=True)

    with subtab_curr:
        st.markdown("#### 💵 Quy Đổi Tiền Tệ Đa Ngoại Tệ")
        rates = {"VND": 1.0, "USD": 25400.0, "EUR": 27500.0, "GBP": 32000.0, "JPY": 165.0, "CNY": 3500.0, "KRW": 18.5, "AUD": 16800.0, "CAD": 18200.0, "SGD": 18900.0}
        c_col1, c_col2, c_col3 = st.columns([2, 1.5, 1.5])
        with c_col1: curr_amount = st.number_input("Số lượng tiền cần đổi:", value=100.0, min_value=0.0, step=10.0)
        with c_col2: from_curr = st.selectbox("Từ đồng tiền:", options=list(rates.keys()), index=1)
        with c_col3: to_curr = st.selectbox("Sang đồng tiền:", options=list(rates.keys()), index=0)
            
        amount_in_vnd = curr_amount * rates[from_curr]
        result_curr = amount_in_vnd / rates[to_curr]
        st.markdown(f"### 🎯 Kết Quả: **{curr_amount:,.2f} {from_curr}** = **{result_curr:,.2f} {to_curr}**")

    with subtab_dist:
        st.markdown("#### 📏 Quy Đổi Đơn Vị Khoảng Cách")
        dist_factors = {"Millimet (mm)": 0.001, "Centimet (cm)": 0.01, "Mét (m)": 1.0, "Kilômét (km)": 1000.0, "Inch (in)": 0.0254, "Foot (ft)": 0.3048, "Yard (yd)": 0.9144, "Dặm (Mile)": 1609.344}
        d_col1, d_col2, d_col3 = st.columns([2, 1.5, 1.5])
        with d_col1: dist_val = st.number_input("Giá trị khoảng cách:", value=1.0, min_value=0.0, step=1.0)
        with d_col2: from_dist = st.selectbox("Từ đơn vị:", options=list(dist_factors.keys()), index=3)
        with d_col3: to_dist = st.selectbox("Sang đơn vị:", options=list(dist_factors.keys()), index=2)
            
        res_dist = (dist_val * dist_factors[from_dist]) / dist_factors[to_dist]
        st.markdown(f"### 🎯 Kết Quả: **{dist_val:,.4f} {from_dist}** = **{res_dist:,.4f} {to_dist}**")

    with subtab_speed:
        st.markdown("#### 🚀 Quy Đổi Đơn Vị Vận Tốc")
        speed_factors = {"Mét/giây (m/s)": 1.0, "Kilômét/giờ (km/h)": 1 / 3.6, "Dặm/giờ (mph)": 0.44704, "Hải lý/giờ (Knot)": 0.514444}
        s_col1, s_col2, s_col3 = st.columns([2, 1.5, 1.5])
        with s_col1: speed_val = st.number_input("Giá trị vận tốc:", value=100.0, min_value=0.0, step=5.0)
        with s_col2: from_speed = st.selectbox("Từ đơn vị:", options=list(speed_factors.keys()), index=1)
        with s_col3: to_speed = st.selectbox("Sang đơn vị:", options=list(speed_factors.keys()), index=0)
            
        res_speed = (speed_val * speed_factors[from_speed]) / speed_factors[to_speed]
        st.markdown(f"### 🎯 Kết Quả: **{speed_val:,.2f} {from_speed}** = **{res_speed:,.2f} {to_speed}**")

    with subtab_mass_temp:
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            with st.container(border=True):
                st.markdown("##### ⚖️ Quy Đổi Khối Lượng")
                mass_factors = {"Gram (g)": 0.001, "Kilôgram (kg)": 1.0, "Tấn": 1000.0, "Ounce (oz)": 0.0283495, "Pound (lb)": 0.453592}
                m_val = st.number_input("Khối lượng:", value=1.0, min_value=0.0, key="m_val_in")
                m_from = st.selectbox("Từ:", options=list(mass_factors.keys()), index=1, key="m_from_sel")
                m_to = st.selectbox("Sang:", options=list(mass_factors.keys()), index=4, key="m_to_sel")
                res_mass = (m_val * mass_factors[m_from]) / mass_factors[m_to]
                st.info(f"👉 **{m_val:,.2f} {m_from}** = **{res_mass:,.2f} {m_to}**")
                
        with m_col2:
            with st.container(border=True):
                st.markdown("##### 🌡️ Quy Đổi Nhiệt Độ")
                temp_val = st.number_input("Nhiệt độ:", value=37.0, key="temp_val_in")
                t_from = st.selectbox("Từ:", options=["Độ C (°C)", "Độ F (°F)", "Kelvin (K)"], index=0, key="t_from_sel")
                t_to = st.selectbox("Sang:", options=["Độ C (°C)", "Độ F (°F)", "Kelvin (K)"], index=1, key="t_to_sel")
                
                if "°C" in t_from: c_temp = temp_val
                elif "°F" in t_from: c_temp = (temp_val - 32) * 5 / 9
                else: c_temp = temp_val - 273.15
                
                if "°C" in t_to: res_temp = c_temp
                elif "°F" in t_to: res_temp = (c_temp * 9 / 5) + 32
                else: res_temp = c_temp + 273.15
                
                st.info(f"👉 **{temp_val:,.1f} {t_from}** = **{res_temp:,.1f} {t_to}**")
