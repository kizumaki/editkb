import streamlit as st
import os
import time
import pandas as pd
from utils import (
    process_docx, clean_file_name_for_output, generate_actor_docx, 
    save_json_db, TRACKER_DB_FILE
)

def render_tab2(enable_colors, enable_phonetic, enable_cast):
    col_r1, col_r2 = st.columns([1.6, 1])
    
    with col_r1:
        with st.container(border=True):
            st.markdown("### 🔄 Tải lên file Kịch bản ĐÃ BIÊN TẬP THỦ CÔNG (.docx)")
            st.caption("Dành riêng cho file kịch bản đã được team biên tập chỉnh sửa lời thoại. Tự động phục hồi màu sắc, phân vai và **xuất phông chữ 14pt cực nét** cho phòng thu.")
            
            resync_file = st.file_uploader(
                "Kéo thả file .docx đã biên tập vào đây", 
                type=['docx'], 
                key=f"resync_uploader_{st.session_state['resync_uploader_key']}"
            )
            project_week_input = st.text_input("📌 Gán Tuần Dự Án cho video này:", value="Tuần 1", help="VD: Tuần 1, Tuần 2, Tuần 1 - Đợt Phim A...")
            
        if resync_file is not None:
            r_filename = resync_file.name
            r_name_no_ext = os.path.splitext(r_filename)[0]
            st.success(f"📄 Đã nhận file kịch bản biên tập: **{r_filename}**")
            
            st.markdown("---")
            if st.button("✨ 2. BẮT ĐẦU RE-SYNC & CHUẨN HÓA LẠI ĐỊNH DẠNG (14PT)", use_container_width=True, type="primary", key="btn_resync_start"):
                try:
                    r_docx, r_ass, r_srt, r_zip, r_stats = process_docx(resync_file, r_name_no_ext, enable_colors, enable_phonetic, enable_cast, is_resync=True, font_size_pt=14)
                    
                    st.session_state['r_processed_docx'] = r_docx
                    st.session_state['r_processed_ass'] = r_ass
                    st.session_state['r_processed_srt'] = r_srt
                    st.session_state['r_actor_zip'] = r_zip
                    st.session_state['r_docx_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".docx")
                    st.session_state['r_ass_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".ass")
                    st.session_state['r_srt_name'] = clean_file_name_for_output(r_filename, tag="_final", ext=".srt")
                    st.session_state['r_zip_name'] = clean_file_name_for_output(r_filename, tag="_KichBan_TachVai_Final", ext=".zip")
                    st.session_state['resync_stats'] = r_stats
                    
                    video_title = r_stats.get("video_title", r_name_no_ext)
                    actors_list = r_stats.get("actors_list", [])
                    actors_str = ", ".join(actors_list) if actors_list else "Chưa có thông tin"
                    actor_breakdown = r_stats.get("actor_stats_breakdown", {})
                    total_lines = r_stats.get("total_lines", 0)
                    video_dur_min = r_stats.get("video_duration_min", 1)
                    
                    today_str = time.strftime("%d/%m/%Y")
                    assigned_week = project_week_input.strip() if project_week_input else "Tuần 1"
                    
                    tracker_list = st.session_state['dubbing_tracker']
                    existing_entry = next((item for item in tracker_list if item['video_title'].upper() == video_title.upper()), None)
                    
                    curr_def_rate = st.session_state['payroll_rates'].get("unit_rate", 30000)
                    custom_actor_rates = {a.upper(): curr_def_rate for a in actors_list}
                    
                    entry_data = {
                        "video_title": video_title, "actors": actors_str, "actor_breakdown": actor_breakdown,
                        "total_lines": total_lines, "video_duration_min": video_dur_min, "date": today_str,
                        "project_week": assigned_week, "custom_actor_rates": custom_actor_rates
                    }
                    
                    if existing_entry:
                        if "custom_actor_rates" in existing_entry: entry_data["custom_actor_rates"] = existing_entry["custom_actor_rates"]
                        existing_entry.update(entry_data)
                    else: tracker_list.append(entry_data)
                    
                    st.session_state['dubbing_tracker'] = tracker_list
                    save_json_db(TRACKER_DB_FILE, tracker_list)
                    
                except Exception as e: st.error(f"Lỗi xảy ra khi Re-Sync: {e}")
                    
            if 'r_processed_docx' in st.session_state:
                st.markdown("---")
                
                # BÁO CÁO CẢNH BÁO CHẤT LƯỢNG (QC)
                r_qc_warns = st.session_state['resync_stats'].get("qc_warnings", [])
                if r_qc_warns:
                    with st.expander("🔍 BÁO CÁO CẢNH BÁO CHẤT LƯỢNG (QC & CPS CHECKER)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in r_qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(r_qc_warns) > 10: st.info(f"...và thêm {len(r_qc_warns)-10} cảnh báo khác.")
                
                # BÁO CÁO SO SÁNH ĐỘ TOÀN VẸN (DIFF CHECK: EDIT VS FINAL)
                integrity = st.session_state['resync_stats'].get("integrity_report", {})
                if integrity:
                    with st.expander("🛡️ BÁO CÁO SO SÁNH ĐỘ TOÀN VẸN (FILE EDIT VS FINAL)", expanded=True):
                        st.caption("Kiểm tra đối soát tự động giữa kịch bản biên tập nạp vào và kịch bản Final tạo ra:")
                        ic1, ic2, ic3, ic4 = st.columns(4)
                        ic1.metric("Mốc TC (Edit)", integrity.get("tc_in_cnt", 0))
                        ic2.metric("Mốc TC (Final)", integrity.get("tc_out_cnt", 0))
                        ic3.metric("Dòng thoại (Edit)", integrity.get("line_in_cnt", 0))
                        ic4.metric("Dòng thoại (Final)", integrity.get("line_out_cnt", 0))
                        
                        diff_issues = integrity.get("diff_issues", [])
                        if not diff_issues:
                            st.success("🎉 **BẢO CHỨNG 100% KHỚP CHUẨN!** Tất cả timecode và nội dung thoại giữa file Edit và Final trùng khớp hoàn toàn.")
                        else:
                            st.warning(f"⚠️ Phát hiện **{len(diff_issues)}** điểm sai lệch cần lưu ý:")
                            st.dataframe(pd.DataFrame(diff_issues), use_container_width=True)

                st.markdown("### ⬇️ 3. TẢI VỀ CÁC FILE CHUẨN HOÀN HẢO (PHÔNG 14PT)")
                col_rdl1, col_rdl2, col_rdl3 = st.columns(3)
                with col_rdl1:
                    st.download_button(
                        label="📄 FILE WORD KỊCH BẢN (14PT)", data=st.session_state['r_processed_docx'],
                        file_name=st.session_state['r_docx_name'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True, key="btn_resync_dl_docx"
                    )
                with col_rdl2:
                    st.download_button(
                        label="🎬 PHỤ ĐỀ CAO CẤP (.ASS)", data=st.session_state['r_processed_ass'],
                        file_name=st.session_state['r_ass_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_ass"
                    )
                with col_rdl3:
                    st.download_button(
                        label="📝 PHỤ ĐỀ CHUẨN (.SRT)", data=st.session_state['r_processed_srt'],
                        file_name=st.session_state['r_srt_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_srt"
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ KỊCH BẢN TÁCH VAI RIÊNG CHO PHÒNG THU LỒNG TIẾNG (14PT)")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                r_act_map = st.session_state['resync_stats'].get("actor_dialogue_map", {})
                if r_act_map:
                    col_ract1, col_ract2 = st.columns([2, 1])
                    with col_ract1:
                        r_selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(r_act_map.keys()), key="resync_select_actor")
                        if r_selected_actor:
                            r_act_buf = generate_actor_docx(st.session_state['resync_stats']['video_title'], r_selected_actor, r_act_map[r_selected_actor], font_size_pt=14)
                            st.download_button(
                                label=f"⬇️ TẢI FILE WORD RIÊNG CHO {r_selected_actor} (14PT)", data=r_act_buf,
                                file_name=f"KichBan_{r_selected_actor}_{st.session_state['resync_stats']['video_title']}_Final.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="btn_dl_single_actor_resync"
                            )
                    with col_ract2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 TẢI TRỌN BỘ KỊCH BẢN TÁCH VAI (.ZIP)", data=st.session_state['r_actor_zip'],
                            file_name=st.session_state['r_zip_name'], mime="application/zip", type="secondary", use_container_width=True, key="btn_dl_zip_actor_resync"
                        )
                st.balloons()
        else: st.info("📌 **Hãy tải file kịch bản đã qua chỉnh sửa thủ công để hệ thống phục hồi lại định dạng chuẩn.**")

    with col_r2:
        st.markdown("### 📊 Re-Sync Analytics")
        if 'resync_stats' in st.session_state:
            r_stats = st.session_state['resync_stats']
            st.markdown(f"""
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">🎭 Tổng số Nhân vật</div>
                <div class="metric-value">{r_stats["total_speakers"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">💬 Tổng số Câu thoại</div>
                <div class="metric-value">{r_stats["total_lines"]}</div>
            </div>
            <div class="metric-card" style="margin-bottom: 12px;">
                <div class="metric-label">⏱️ Độ dài Video</div>
                <div class="metric-value">{r_stats["video_duration_min"]} phút</div>
            </div>
            """, unsafe_allow_html=True)
            top_name, top_count = r_stats["top_speaker"]
            st.info(f"👑 **Nhân vật thoại nhiều nhất:** \n\n**{top_name}** với {top_count} câu thoại.")
        else: st.info("Thống kê file Re-Sync sẽ xuất hiện tại đây sau khi hoàn tất.")
