import streamlit as st
import io
import os
import re
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from utils import timecode_to_sec, calculate_time_overlap, parse_srt_to_dataframe

def parse_any_script_file_to_df(file_bytes, filename, custom_speakers=None, non_speakers=None, default_speaker="Unknown"):
    if filename.lower().endswith('.srt'):
        try: content_str = file_bytes.decode('utf-8')
        except UnicodeDecodeError: content_str = file_bytes.decode('latin-1')
        return parse_srt_to_dataframe(content_str, custom_speakers, non_speakers, default_speaker)
    elif filename.lower().endswith('.docx'):
        doc = Document(io.BytesIO(file_bytes))
        paragraphs_text = [p.text.strip() for p in doc.paragraphs if p.text.strip() != ""]
        timecode_pattern = re.compile(r'\d{2}:\d{2}:\d{2}[,.]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[,.]\d{3}')
        
        srt_lines = []
        i = 0
        while i < len(paragraphs_text):
            line = paragraphs_text[i]
            if line.isdigit() and i + 1 < len(paragraphs_text) and timecode_pattern.search(paragraphs_text[i+1]):
                srt_lines.append(line)
                srt_lines.append(paragraphs_text[i+1])
                i += 2
                while i < len(paragraphs_text):
                    if paragraphs_text[i].isdigit() and i + 1 < len(paragraphs_text) and timecode_pattern.search(paragraphs_text[i+1]): break
                    else: srt_lines.append(paragraphs_text[i]); i += 1
            else: i += 1
        content_str = "\n".join(srt_lines)
        return parse_srt_to_dataframe(content_str, custom_speakers, non_speakers, default_speaker)
    return pd.DataFrame(columns=['Start', 'End', 'Speaker', 'Dialogue', 'Is_Explicit'])

def align_and_compare_english_scripts(df_mh_eng, df_off_eng, df_vn=None, default_speaker="Unknown"):
    aligned_rows = []
    fallback_spk = default_speaker.strip() if default_speaker and default_speaker.strip() else "Unknown"
    
    for idx_mh, row_mh in df_mh_eng.iterrows():
        s_mh = timecode_to_sec(row_mh['Start'])
        e_mh = timecode_to_sec(row_mh['End'])
        is_explicit_mh = row_mh.get('Is_Explicit', False)
        
        best_off_matches = []
        for idx_off, row_off in df_off_eng.iterrows():
            s_off = timecode_to_sec(row_off['Start'])
            e_off = timecode_to_sec(row_off['End'])
            overlap = calculate_time_overlap(s_mh, e_mh, s_off, e_off)
            if overlap > 0.1:
                best_off_matches.append((idx_off, overlap, row_off))
                
        if best_off_matches:
            off_dialogues = [str(m[2]['Dialogue']) for m in best_off_matches if pd.notna(m[2]['Dialogue'])]
            off_text_combined = " ".join(off_dialogues)
            off_spk = best_off_matches[0][2]['Speaker']
        else:
            off_text_combined = ""
            off_spk = ""

        vn_text_combined = ""
        vn_spk = ""
        if df_vn is not None and not df_vn.empty:
            best_vn_matches = []
            for idx_vn, row_vn in df_vn.iterrows():
                s_vn = timecode_to_sec(row_vn['Start'])
                e_vn = timecode_to_sec(row_vn['End'])
                overlap_vn = calculate_time_overlap(s_mh, e_mh, s_vn, e_vn)
                if overlap_vn > 0.1:
                    best_vn_matches.append((overlap_vn, row_vn))
            if best_vn_matches:
                vn_text_combined = " ".join([str(m[1]['Dialogue']) for m in best_vn_matches if pd.notna(m[1]['Dialogue'])])
                vn_spk = best_vn_matches[0][1]['Speaker']
        else:
            vn_spk = str(row_mh['Speaker']) if pd.notna(row_mh['Speaker']) else ""

        qc_status = "🟢 Khớp chuẩn"
        qc_details = "Nội dung Tiếng Anh khớp chuẩn nghĩa"
        
        mh_diag_clean = str(row_mh['Dialogue']).strip() if pd.notna(row_mh['Dialogue']) else ""
        off_diag_clean = off_text_combined.strip()
        
        if not off_diag_clean:
            qc_status = "🔴 Thiếu câu gốc Khách"
            qc_details = "File Mai Han có câu này nhưng file Khách không thấy có"
        elif not mh_diag_clean:
            qc_status = "🔴 Thiếu câu Mai Han"
            qc_details = "File Khách có câu này nhưng Mai Han nghe bị bỏ sót"

        spk_display_mh = str(row_mh['Speaker']) if pd.notna(row_mh['Speaker']) and str(row_mh['Speaker']) != "Unknown" else fallback_spk
        spk_display_off = off_spk if off_spk and off_spk != "Unknown" else fallback_spk
        spk_display_vn = vn_spk if vn_spk and vn_spk != "Unknown" else fallback_spk

        aligned_rows.append({
            "Stt": idx_mh + 1,
            "Timecode": f"{row_mh['Start']} --> {row_mh['End']}",
            "Start": row_mh['Start'],
            "End": row_mh['End'],
            "Tiếng Anh Mai Han (AI/Heard)": f"{spk_display_mh}: {row_mh['Dialogue']}",
            "Tiếng Anh Khách (Official)": f"{spk_display_off}: {off_text_combined}" if spk_display_off else off_text_combined,
            "Dịch Tiếng Việt (Cần Sửa)": f"{spk_display_vn}: {vn_text_combined}" if spk_display_vn else vn_text_combined,
            "Speaker_MH": spk_display_mh,
            "Is_Explicit_MH": is_explicit_mh,
            "Speaker_Off": spk_display_off,
            "Speaker_VN": spk_display_vn,
            "Dialogue_MH": row_mh['Dialogue'],
            "Dialogue_Off": off_text_combined,
            "Dialogue_VN": vn_text_combined,
            "Trạng thái QC": qc_status,
            "Ghi chú QC": qc_details
        })

    return pd.DataFrame(aligned_rows)

def render_tab6(enable_colors, enable_phonetic, enable_cast):
    st.subheader("🔀 ĐỐI CHIẾU 2 FILE TIẾNG ANH & SOÁT SỬA BẢN DỊCH VIỆT")
    st.markdown("So sánh file Tiếng Anh do Mai Han Team nghe với file Tiếng Anh Gốc do Khách gửi trễ.")

    col_spk_fb1, col_spk_fb2 = st.columns([1.8, 1.2])
    with col_spk_fb1:
        default_spk_input = st.text_input("🎭 Tên người nói mặc định (dùng cho Báo cáo QC):", placeholder="VD: Nick, MC...", value="")

    col_dual1, col_dual2, col_dual3 = st.columns(3)
    with col_dual1:
        uploaded_mh_eng = st.file_uploader("1. File tiếng Anh Mai Han (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_mh_eng")
    with col_dual2:
        uploaded_off_eng = st.file_uploader("2. File tiếng Anh Khách (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_off_eng")
    with col_dual3:
        uploaded_vn_script = st.file_uploader("3. File Tiếng Việt (.srt/docx):", type=['srt', 'docx'], key="uploader_dual_vn_script")

    if uploaded_mh_eng is not None and uploaded_off_eng is not None:
        c_spks = st.session_state.get('custom_speakers', set())
        c_non_spks = st.session_state.get('custom_non_speakers', set())
        fallback_spk_name = default_spk_input.strip() if default_spk_input.strip() else "Unknown"

        df_mh_eng = parse_any_script_file_to_df(uploaded_mh_eng.getvalue(), uploaded_mh_eng.name, c_spks, c_non_spks, fallback_spk_name)
        df_off_eng = parse_any_script_file_to_df(uploaded_off_eng.getvalue(), uploaded_off_eng.name, c_spks, c_non_spks, fallback_spk_name)
        df_vn_script = parse_any_script_file_to_df(uploaded_vn_script.getvalue(), uploaded_vn_script.name, c_spks, c_non_spks, fallback_spk_name) if uploaded_vn_script else None

        if not df_mh_eng.empty and not df_off_eng.empty:
            df_aligned = align_and_compare_english_scripts(df_mh_eng, df_off_eng, df_vn_script, fallback_spk_name)
            st.success(f"✅ Đã đối chiếu thành công **{len(df_aligned)}** câu thoại!")
            st.dataframe(df_aligned[['Stt', 'Timecode', 'Tiếng Anh Mai Han (AI/Heard)', 'Tiếng Anh Khách (Official)', 'Dịch Tiếng Việt (Cần Sửa)', 'Trạng thái QC']], use_container_width=True)
