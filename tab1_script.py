import streamlit as st
import os
import time
import pandas as pd
from utils import (
    scan_candidate_speakers, scan_english_words_in_dialogue, 
    generate_english_audio, process_docx, clean_file_name_for_output, 
    generate_actor_docx, save_json_db, CAST_DB_FILE, PHONETIC_DB_FILE, 
    SPEAKER_DB_FILE, NON_SPEAKER_DB_FILE, DEFAULT_NON_SPEAKER_PHRASES
)

def render_tab1(enable_colors, enable_phonetic, enable_cast):
    col1, col2 = st.columns([1.6, 1])

    with col1:
        with st.container(border=True):
            st.markdown("### 📁 Tải lên file Kịch bản Word gốc (.docx)")
            uploaded_file = st.file_uploader(
                "Kéo thả file .docx gốc của bạn vào đây", 
                type=['docx'], 
                key=f"main_uploader_{st.session_state['uploader_key']}"
            )

        if uploaded_file is None:
            st.info("📌 **Vui lòng tải file kịch bản (.docx) ở trên để hiển thị công cụ biên tập.**")
        else:
            original_filename = uploaded_file.name
            file_name_without_ext = os.path.splitext(original_filename)[0] 
            st.success(f"📄 Đã nhận file thành công: **{original_filename}**")

            custom_spks = st.session_state.get('custom_speakers', set())
            custom_non_spks = st.session_state.get('custom_non_speakers', set())
            non_spk_phrases = DEFAULT_NON_SPEAKER_PHRASES.union(custom_non_spks)

            candidates = scan_candidate_speakers(uploaded_file, custom_spks, custom_non_spks)

            detected_speakers_names = [name for name in candidates.keys() if name.upper() not in non_spk_phrases]
            detected_non_speakers_names = [name for name in candidates.keys() if name.upper() in non_spk_phrases]

            detected_speakers = [f"{name} ({candidates[name]} lần)" for name in detected_speakers_names]
            detected_non_speakers = [f"{name} ({candidates[name]} lần)" for name in detected_non_speakers_names]

            with st.container(border=True):
                st.markdown("### 🎭 Phân Vai Lồng Tiếng Cho Kịch Bản Hiện Tại")
                st.caption("Xem và gán người lồng tiếng Việt cho từng nhân vật trong file kịch bản này:")

                if detected_speakers_names:
                    cast_table_data = []
                    for spk_name in detected_speakers_names:
                        current_actor = st.session_state['custom_cast_mapping'].get(spk_name.upper(), "")
                        cast_table_data.append({
                            "Nhân vật (Tiếng Anh)": spk_name,
                            "Diễn viên Lồng tiếng (Tiếng Việt)": current_actor,
                            "Nạp vào Database": True
                        })

                    df_cast = pd.DataFrame(cast_table_data)

                    edited_cast_df = st.data_editor(
                        df_cast,
                        column_config={
                            "Nhân vật (Tiếng Anh)": st.column_config.TextColumn("Nhân vật (Kịch bản gốc)", disabled=True),
                            "Diễn viên Lồng tiếng (Tiếng Việt)": st.column_config.TextColumn("Diễn viên lồng tiếng (Sửa trực tiếp)"),
                            "Nạp vào Database": st.column_config.CheckboxColumn("Lưu Database?", default=True)
                        },
                        disabled=["Nhân vật (Tiếng Anh)"],
                        hide_index=True,
                        use_container_width=True,
                        key="script_cast_editor_table"
                    )

                    if st.button("💾 Lưu Bảng Phân Vai Kịch Bản Này Vào Database", type="secondary", use_container_width=True):
                        updated_cast_count = 0
                        for _, row in edited_cast_df.iterrows():
                            if row["Nạp vào Database"]:
                                spk_k = str(row["Nhân vật (Tiếng Anh)"]).upper().strip()
                                act_v = str(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).strip().upper()
                                if act_v:
                                    st.session_state['custom_cast_mapping'][spk_k] = act_v
                                    updated_cast_count += 1
                        save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                        st.success(f"✅ Đã lưu phân vai cho {updated_cast_count} nhân vật vào Database!")
                        time.sleep(1); st.rerun()

            with st.container(border=True):
                st.markdown("### 🔍 Soát Lỗi Nhận Diện Tên Người Nói")
                tab_spk, tab_non_spk = st.tabs(["🎭 Nhận diện là NGƯỜI NÓI", "🚫 Đang bị xem là TỪ NHIỄU"])
                
                with tab_spk:
                    if detected_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_speakers]))
                        to_move_to_ns = st.multiselect(
                            "Phát hiện từ nào bị nhận diện sai? Chọn để LƯU VÀO DATABASE TỪ NHIỄU:",
                            options=[name for name in candidates.keys() if name.upper() not in non_spk_phrases],
                            key="select_to_ns"
                        )
                        if st.button("➡️ Đưa vào Database TỪ NHIỄU", type="secondary"):
                            if to_move_to_ns:
                                new_items = [item.upper() for item in to_move_to_ns]
                                st.session_state['custom_non_speakers'].update(new_items)
                                save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                                st.success(f"✅ Đã lưu {len(new_items)} từ vào Database Từ Nhiễu!")
                                time.sleep(1); st.rerun()
                    else: st.info("Chưa tìm thấy cụm từ người nói nào.")

                with tab_non_spk:
                    if detected_non_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_non_speakers]))
                        to_move_to_spk = st.multiselect(
                            "Từ nào thực ra là NGƯỜI NÓI? Chọn để LƯU VÀO DATABASE NGƯỜI NÓI:",
                            options=[name for name in candidates.keys() if name.upper() in non_spk_phrases],
                            key="select_to_spk"
                        )
                        if st.button("➡️ Đưa vào Database NGƯỜI NÓI", type="secondary"):
                            if to_move_to_spk:
                                st.session_state['custom_speakers'].update(to_move_to_spk)
                                save_json_db(SPEAKER_DB_FILE, st.session_state['custom_speakers'])
                                for item in to_move_to_spk: st.session_state['custom_non_speakers'].discard(item.upper())
                                save_json_db(NON_SPEAKER_DB_FILE, st.session_state['custom_non_speakers'])
                                st.success(f"✅ Đã lưu {len(to_move_to_spk)} tên vào Database Người Nói!")
                                time.sleep(1); st.rerun()
                    else: st.info("Không có cụm từ nào bị loại vào danh sách từ nhiễu.")

            with st.container(border=True):
                st.markdown("### 🗣️ Từ Tiếng Anh Xuất Hiện Trong Kịch Bản")
                st.caption("Quét và điều chỉnh phiên âm riêng cho kịch bản này (Đã qua bộ lọc thông minh):")

                detected_eng_words = scan_english_words_in_dialogue(uploaded_file, custom_spks, custom_non_spks)

                if detected_eng_words:
                    st.markdown("#### 🔊 Trình nghe phát âm chuẩn giọng bản xứ (Google US/UK)")
                    col_listen1, col_listen2, col_listen3 = st.columns([2.5, 1.5, 1.5])
                    
                    with col_listen1:
                        word_to_listen = st.selectbox("Chọn từ cần nghe phát âm:", options=detected_eng_words, key="script_listen_select")
                    with col_listen2:
                        accent_choice = st.radio("Giọng phát âm:", options=["Giọng Mỹ (US)", "Giọng Anh (UK)"], horizontal=True, key="script_accent_radio")
                    with col_listen3:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        listen_btn = st.button("🔊 Nghe phát âm", type="secondary", use_container_width=True)

                    if listen_btn and word_to_listen:
                        tld = 'com' if "Mỹ" in accent_choice else 'co.uk'
                        audio_fp = generate_english_audio(word_to_listen, accent=tld)
                        if audio_fp: st.audio(audio_fp, format="audio/mp3", autoplay=True)

                    st.markdown("---")
                    
                    table_data = []
                    for word in detected_eng_words:
                        current_pho = st.session_state['custom_phonetics'].get(word.upper(), word)
                        table_data.append({
                            "Từ Tiếng Anh": word,
                            "Phiên âm hiện tại": current_pho,
                            "Đề xuất chỉnh sửa của bạn": current_pho,
                            "Nạp vào Database": True
                        })

                    df_eng = pd.DataFrame(table_data)

                    edited_df = st.data_editor(
                        df_eng,
                        column_config={
                            "Từ Tiếng Anh": st.column_config.TextColumn("Từ Tiếng Anh gốc", disabled=True),
                            "Phiên âm hiện tại": st.column_config.TextColumn("Phiên âm gán hiện tại", disabled=True),
                            "Đề xuất chỉnh sửa của bạn": st.column_config.TextColumn("Đề xuất phiên âm mới"),
                            "Nạp vào Database": st.column_config.CheckboxColumn("Lưu Database?", default=True)
                        },
                        disabled=["Từ Tiếng Anh", "Phiên âm hiện tại"],
                        hide_index=True,
                        use_container_width=True,
                        key="phonetic_script_table"
                    )

                    if st.button("💾 Nạp chỉnh sửa kịch bản này vào Database Phiên Âm", type="secondary", use_container_width=True):
                        updated_count = 0
                        for _, row in edited_df.iterrows():
                            if row["Nạp vào Database"]:
                                eng_k = str(row["Từ Tiếng Anh"]).upper().strip()
                                pho_v = str(row["Đề xuất chỉnh sửa của bạn"]).strip()
                                if pho_v:
                                    st.session_state['custom_phonetics'][eng_k] = pho_v
                                    updated_count += 1
                        
                        save_json_db(PHONETIC_DB_FILE, st.session_state['custom_phonetics'])
                        st.success(f"✅ Đã cập nhật {updated_count} từ phiên âm vào Database!")
                        time.sleep(1); st.rerun()
                else: st.info("Không phát hiện từ Tiếng Anh / Tên riêng nước ngoài nào trong phần lời thoại kịch bản này.")

            st.markdown("---")
            if st.button("✨ 2. BẮT ĐẦU ĐỊNH DẠNG TỰ ĐỘNG", use_container_width=True, type="primary"):
                try:
                    modified_docx, ass_f, srt_f, act_zip, stats = process_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic, enable_cast, is_resync=False, font_size_pt=12)
                    
                    st.session_state['processed_docx'] = modified_docx
                    st.session_state['processed_ass'] = ass_f
                    st.session_state['processed_srt'] = srt_f
                    st.session_state['actor_zip'] = act_zip
                    st.session_state['docx_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".docx")
                    st.session_state['ass_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".ass")
                    st.session_state['srt_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".srt")
                    st.session_state['zip_name'] = clean_file_name_for_output(original_filename, tag="_KichBan_TachVai", ext=".zip")
                    st.session_state['stats'] = stats
                    
                except Exception as e: st.error(f"Đã có lỗi xảy ra: {e}")

            if 'processed_docx' in st.session_state:
                st.markdown("---")
                qc_warns = st.session_state['stats'].get("qc_warnings", [])
                if qc_warns:
                    with st.expander("🔍 BÁO CÁO CẢNH BÁO CHẤT LƯỢNG (QC & CPS CHECKER)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(qc_warns) > 10: st.info(f"...và thêm {len(qc_warns)-10} cảnh báo khác.")
                
                st.markdown("### ⬇️ 3. TẢI VỀ CÁC FILE ĐÃ XỬ LÝ HOÀN HẢO")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                with col_dl1:
                    st.download_button(
                        label="📄 FILE WORD KỊCH BẢN (.DOCX)",
                        data=st.session_state['processed_docx'],
                        file_name=st.session_state['docx_name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True
                    )
                with col_dl2:
                    st.download_button(
                        label="🎬 PHỤ ĐỀ CAO CẤP (.ASS)",
                        data=st.session_state['processed_ass'],
                        file_name=st.session_state['ass_name'],
                        mime="text/plain", use_container_width=True
                    )
                with col_dl3:
                    st.download_button(
                        label="📝 PHỤ ĐỀ CHUẨN (.SRT)",
                        data=st.session_state['processed_srt'],
                        file_name=st.session_state['srt_name'],
                        mime="text/plain", use_container_width=True
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ KỊCH BẢN TÁCH VAI RIÊNG CHO PHÒNG THU LỒNG TIẾNG")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                act_map = st.session_state['stats'].get("actor_dialogue_map", {})
                if act_map:
                    col_act1, col_act2 = st.columns([2, 1])
                    with col_act1:
                        selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(act_map.keys()))
                        if selected_actor:
                            act_buf = generate_actor_docx(st.session_state['stats']['video_title'], selected_actor, act_map[selected_actor], font_size_pt=12)
                            st.download_button(
                                label=f"⬇️ TẢI FILE WORD RIÊNG CHO {selected_actor} (.DOCX)",
                                data=act_buf,
                                file_name=f"KichBan_{selected_actor}_{st.session_state['stats']['video_title']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    with col_act2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 TẢI TRỌN BỘ KỊCH BẢN TÁCH VAI (.ZIP)",
                            data=st.session_state['actor_zip'],
                            file_name=st.session_state['zip_name'],
                            mime="application/zip", type="secondary", use_container_width=True
                        )
                st.balloons()

    with col2:
        st.markdown("### 📊 SaaS Analytics")
        if 'stats' in st.session_state:
            stats = st.session_state['stats']
            st.markdown(f"""
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">🎭 Tổng số Nhân vật</div>
                <div class="metric-value">{stats["total_speakers"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">💬 Tổng số Câu thoại</div>
                <div class="metric-value">{stats["total_lines"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">⏱️ Độ dài Video</div>
                <div class="metric-value">{stats["video_duration_min"]} phút</div>
            </div>
            """, unsafe_allow_html=True)
            top_name, top_count = stats["top_speaker"]
            st.info(f"👑 **Nhân vật thoại nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
        else: st.info("Bảng phân tích dữ liệu kịch bản sẽ xuất hiện tại đây sau khi bạn xử lý file.")
