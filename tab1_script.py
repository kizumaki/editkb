import streamlit as st
import os
import time
import pandas as pd
from utils import (
    kept_file_uploader, data_signature, warn_if_stale,
    scan_candidate_speakers, scan_english_words_in_dialogue, 
    generate_english_audio, process_docx, clean_file_name_for_output, 
    generate_actor_docx, save_json_db, CAST_DB_FILE, PHONETIC_DB_FILE, 
    DEFAULT_NON_SPEAKER_PHRASES, clean_cell, add_names, render_names_result
)
from learn_core import norm_eng_key, pho_missing
from batch_tools import render_batch_processing

def render_tab1(enable_colors, enable_phonetic, enable_cast):
    work_mode = st.radio("Cách làm:", ["Một file (soát kỹ từng bước)", "Nhiều file cùng lúc"],
                         horizontal=True, key="tab1_mode", label_visibility="collapsed")
    if work_mode.startswith("Nhiều"):
        render_batch_processing(False, enable_colors, enable_phonetic, enable_cast)
        return

    col1, col2 = st.columns([1.6, 1])

    with col1:
        with st.container(border=True):
            st.markdown("### 📁 Tải lên file Kịch bản Word gốc (.docx)")
            uploaded_file = kept_file_uploader(
                "Kéo thả file .docx gốc của bạn vào đây", 
                type=['docx'], 
                key="main_uploader"
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

                    if st.button("💾 Lưu phân vai của kịch bản này", type="secondary", use_container_width=True):
                        updated_cast_count = 0
                        for _, row in edited_cast_df.iterrows():
                            if row["Nạp vào Database"]:
                                spk_k = clean_cell(row["Nhân vật (Tiếng Anh)"]).upper()
                                act_v = clean_cell(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).upper()
                                if act_v:
                                    st.session_state['custom_cast_mapping'][spk_k] = act_v
                                    updated_cast_count += 1
                        save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                        st.success(f"✅ Đã lưu phân vai cho {updated_cast_count} nhân vật vào Database!")
                        time.sleep(1); st.rerun()

            with st.container(border=True):
                st.markdown("### 🔍 Kiểm tra tên người nói")
                tab_spk, tab_non_spk = st.tabs(["🎭 Đang nhận là người nói", "🚫 Đang bị bỏ qua"])
                
                with tab_spk:
                    if detected_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_speakers]))
                        to_move_to_ns = st.multiselect(
                            "Phát hiện từ nào bị nhận diện sai? Chọn để LƯU VÀO DATABASE TỪ NHIỄU:",
                            options=[name for name in candidates.keys() if name.upper() not in non_spk_phrases],
                            key="select_to_ns"
                        )
                        render_names_result("ns", slot="tab1")
                        if st.button("➡️ Chuyển sang \"không phải tên nhân vật\"", type="secondary"):
                            if to_move_to_ns:
                                add_names("ns", to_move_to_ns, move=True, slot="tab1"); st.rerun()  # chuyển hẳn: bỏ khỏi danh sách tên
                    else: st.info("Chưa tìm thấy cụm từ người nói nào.")

                with tab_non_spk:
                    if detected_non_speakers:
                        st.write(", ".join([f"`{s}`" for s in detected_non_speakers]))
                        to_move_to_spk = st.multiselect(
                            "Từ nào thực ra là NGƯỜI NÓI? Chọn để LƯU VÀO DATABASE NGƯỜI NÓI:",
                            options=[name for name in candidates.keys() if name.upper() in non_spk_phrases],
                            key="select_to_spk"
                        )
                        render_names_result("spk", slot="tab1")
                        if st.button("➡️ Chuyển sang danh sách tên nhân vật", type="secondary"):
                            if to_move_to_spk:
                                add_names("spk", to_move_to_spk, move=True, slot="tab1"); st.rerun()
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
                        current_pho = st.session_state['custom_phonetics'].get(norm_eng_key(word), "")
                        if pho_missing(word, current_pho): current_pho = ""  # "Pizza -> Pizza" = chưa có phiên âm
                        table_data.append({
                            "Từ Tiếng Anh": word,
                            "Phiên âm hiện tại": current_pho or "(chưa có)",
                            "Đề xuất chỉnh sửa của bạn": current_pho,
                            "Nạp vào Database": False
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

                    if st.button("💾 Lưu phiên âm vào kho", type="secondary", use_container_width=True):
                        updated_count = 0
                        kho = st.session_state['custom_phonetics']
                        for _, row in edited_df.iterrows():
                            eng_k = norm_eng_key(clean_cell(row["Từ Tiếng Anh"]))
                            pho_v = clean_cell(row["Đề xuất chỉnh sửa của bạn"])
                            # Chỉ lưu dòng được tích, có phiên âm THẬT (không phải chép lại chữ tiếng Anh) và khác kho
                            if row["Nạp vào Database"] and pho_v and not pho_missing(eng_k, pho_v) and pho_v != kho.get(eng_k):
                                kho[eng_k] = pho_v
                                updated_count += 1
                        if updated_count: save_json_db(PHONETIC_DB_FILE, kho)
                        st.success(f"✅ Đã cập nhật {updated_count} từ phiên âm vào Database!")
                        time.sleep(1); st.rerun()
                else: st.info("Không phát hiện từ Tiếng Anh / Tên riêng nước ngoài nào trong phần lời thoại kịch bản này.")

            st.markdown("---")
            if st.button("✨ Bắt đầu định dạng tự động", use_container_width=True, type="primary"):
                try:
                    modified_docx, ass_f, srt_f, act_zip, stats = process_docx(uploaded_file, file_name_without_ext, enable_colors, enable_phonetic, enable_cast, is_resync=False, font_size_pt=14)
                    
                    st.session_state['processed_docx'] = modified_docx
                    st.session_state['processed_ass'] = ass_f
                    st.session_state['processed_srt'] = srt_f
                    st.session_state['actor_zip'] = act_zip
                    st.session_state['docx_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".docx")
                    st.session_state['ass_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".ass")
                    st.session_state['srt_name'] = clean_file_name_for_output(original_filename, tag="_edit", ext=".srt")
                    st.session_state['zip_name'] = clean_file_name_for_output(original_filename, tag="_KichBan_TachVai", ext=".zip")
                    st.session_state['stats'] = stats
                    st.session_state['_celebrate_tab1'] = True  # chỉ chúc mừng 1 lần ngay sau khi xử lý xong
                    st.session_state['processed_sig'] = data_signature(enable_colors, enable_phonetic, enable_cast)

                except Exception as e: st.error(f"Đã có lỗi xảy ra: {e}")

            if 'processed_docx' in st.session_state:
                st.markdown("---")
                warn_if_stale('processed_sig', enable_colors, enable_phonetic, enable_cast, button_label="✨ Bắt đầu định dạng tự động")
                qc_warns = st.session_state['stats'].get("qc_warnings", [])
                if qc_warns:
                    with st.expander("🔍 Cảnh báo chất lượng (tốc độ đọc, phân vai)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(qc_warns) > 10: st.info(f"...và thêm {len(qc_warns)-10} cảnh báo khác.")
                
                st.markdown("### ⬇️ Tải về file đã xử lý")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                with col_dl1:
                    st.download_button(
                        label="📄 Kịch bản Word (.docx)",
                        data=st.session_state['processed_docx'],
                        file_name=st.session_state['docx_name'],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True
                    )
                with col_dl2:
                    st.download_button(
                        label="🎬 Phụ đề có màu (.ass)",
                        data=st.session_state['processed_ass'],
                        file_name=st.session_state['ass_name'],
                        mime="text/plain", use_container_width=True
                    )
                with col_dl3:
                    st.download_button(
                        label="📝 Phụ đề thường (.srt)",
                        data=st.session_state['processed_srt'],
                        file_name=st.session_state['srt_name'],
                        mime="text/plain", use_container_width=True
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ Kịch bản tách riêng cho từng diễn viên")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                act_map = st.session_state['stats'].get("actor_dialogue_map", {})
                if act_map:
                    col_act1, col_act2 = st.columns([2, 1])
                    with col_act1:
                        selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(act_map.keys()))
                        if selected_actor:
                            act_buf = generate_actor_docx(st.session_state['stats']['video_title'], selected_actor, act_map[selected_actor], font_size_pt=14)
                            st.download_button(
                                label=f"⬇️ Tải file của {selected_actor} (.docx)",
                                data=act_buf,
                                file_name=f"KichBan_{selected_actor}_{st.session_state['stats']['video_title']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    with col_act2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 Tải trọn bộ (.zip)",
                            data=st.session_state['actor_zip'],
                            file_name=st.session_state['zip_name'],
                            mime="application/zip", type="secondary", use_container_width=True
                        )
                if st.session_state.pop('_celebrate_tab1', False): st.balloons()

    with col2:
        st.markdown("### 📊 Thống kê kịch bản")
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
