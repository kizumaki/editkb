import streamlit as st
import os
import time
import pandas as pd
from utils import (
    kept_file_uploader,
    process_docx, clean_file_name_for_output, generate_actor_docx, 
    save_json_db, TRACKER_DB_FILE, record_video_in_tracker, project_picker
)
from batch_tools import render_batch_processing

def render_tab2(enable_colors, enable_phonetic, enable_cast):
    work_mode = st.radio("Cách làm:", ["Một file (soát kỹ từng bước)", "Nhiều file cùng lúc"],
                         horizontal=True, key="tab2_mode", label_visibility="collapsed")
    if work_mode.startswith("Nhiều"):
        render_batch_processing(True, enable_colors, enable_phonetic, enable_cast)
        return

    col_r1, col_r2 = st.columns([1.6, 1])
    
    with col_r1:
        with st.container(border=True):
            st.markdown("### 🔄 Tải lên kịch bản đã biên tập (.docx)")
            st.caption("Dành cho kịch bản team đã sửa lời thoại bằng tay. App sẽ tô lại màu, phân vai và xuất **cỡ chữ 14** cho phòng thu.")
            
            resync_file = kept_file_uploader(
                "Kéo thả file .docx đã biên tập vào đây", 
                type=['docx'], 
                key="resync_uploader"
            )
            col_pj1, col_pj2 = st.columns(2)
            with col_pj1: project_id_input = project_picker("📁 Video này thuộc dự án:", key="resync_project_id")
            with col_pj2: project_week_input = st.text_input("📌 Ghi chú tuần/đợt (không bắt buộc):", key="resync_project_week",
                                                             help="Chỉ để ghi chú, VD: Tuần 1. Lương được tính theo dự án và ngày ghi nhận video.")
            
        if resync_file is not None:
            r_filename = resync_file.name
            r_name_no_ext = os.path.splitext(r_filename)[0]
            st.success(f"📄 Đã nhận file kịch bản biên tập: **{r_filename}**")
            
            st.markdown("---")
            if st.button("✨ Bắt đầu Re-Sync (cỡ chữ 14)", use_container_width=True, type="primary", key="btn_resync_start"):
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
                    st.session_state['_celebrate_tab2'] = True  # chỉ chúc mừng 1 lần ngay sau khi xử lý xong
                    
                    record_video_in_tracker(r_stats, r_name_no_ext, project_week_input, project_id_input)
                    save_json_db(TRACKER_DB_FILE, st.session_state['dubbing_tracker'])
                    
                except Exception as e: st.error(f"Lỗi xảy ra khi Re-Sync: {e}")
                    
            if 'r_processed_docx' in st.session_state:
                st.markdown("---")
                r_qc_warns = st.session_state['resync_stats'].get("qc_warnings", [])
                if r_qc_warns:
                    with st.expander("🔍 Cảnh báo chất lượng (tốc độ đọc, phân vai)", expanded=True):
                        st.caption("Danh sách cảnh báo về tốc độ đọc thoại hoặc gán phân vai để BTV rà soát:")
                        for w in r_qc_warns[:10]: st.markdown(f"<div class='qc-card-warning'>{w}</div>", unsafe_allow_html=True)
                        if len(r_qc_warns) > 10: st.info(f"...và thêm {len(r_qc_warns)-10} cảnh báo khác.")
                
                # BÁO CÁO SO SÁNH ĐỘ TOÀN VẸN (DIFF CHECK)
                integrity = st.session_state['resync_stats'].get("integrity_report", {})
                if integrity:
                    with st.expander("🛡️ So sánh file biên tập với file Final", expanded=True):
                        st.caption("Kiểm tra đối soát tự động giữa kịch bản biên tập nạp vào và kịch bản Final tạo ra:")
                        ic1, ic2, ic3, ic4 = st.columns(4)
                        ic1.metric("Mốc TC (Edit)", integrity.get("tc_in_cnt", 0))
                        ic2.metric("Mốc TC (Final)", integrity.get("tc_out_cnt", 0))
                        ic3.metric("Dòng thoại (Edit)", integrity.get("line_in_cnt", 0))
                        ic4.metric("Dòng thoại (Final)", integrity.get("line_out_cnt", 0))
                        
                        diff_issues = integrity.get("diff_issues", [])
                        if not diff_issues:
                            st.success("✅ Không thấy sai lệch: số mốc timecode khớp, không có câu thoại nào bị hụt chữ đáng kể.")
                        else:
                            st.warning(f"⚠️ Phát hiện **{len(diff_issues)}** điểm sai lệch cần lưu ý:")
                            st.dataframe(pd.DataFrame(diff_issues), use_container_width=True)

                st.markdown("### ⬇️ Tải về file đã chuẩn hóa (cỡ chữ 14)")
                col_rdl1, col_rdl2, col_rdl3 = st.columns(3)
                with col_rdl1:
                    st.download_button(
                        label="📄 Kịch bản Word (cỡ 14)", data=st.session_state['r_processed_docx'],
                        file_name=st.session_state['r_docx_name'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True, key="btn_resync_dl_docx"
                    )
                with col_rdl2:
                    st.download_button(
                        label="🎬 Phụ đề có màu (.ass)", data=st.session_state['r_processed_ass'],
                        file_name=st.session_state['r_ass_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_ass"
                    )
                with col_rdl3:
                    st.download_button(
                        label="📝 Phụ đề thường (.srt)", data=st.session_state['r_processed_srt'],
                        file_name=st.session_state['r_srt_name'], mime="text/plain", use_container_width=True, key="btn_resync_dl_srt"
                    )
                    
                st.markdown("---")
                st.markdown("#### 🎙️ Kịch bản tách riêng cho từng diễn viên (cỡ 14)")
                st.caption("Mỗi diễn viên chỉ nhận đúng câu thoại của mình, giúp thu âm nhanh và không xao nhãng:")
                
                r_act_map = st.session_state['resync_stats'].get("actor_dialogue_map", {})
                if r_act_map:
                    col_ract1, col_ract2 = st.columns([2, 1])
                    with col_ract1:
                        r_selected_actor = st.selectbox("Chọn Diễn viên lồng tiếng để tải file riêng:", options=list(r_act_map.keys()), key="resync_select_actor")
                        if r_selected_actor:
                            r_act_buf = generate_actor_docx(st.session_state['resync_stats']['video_title'], r_selected_actor, r_act_map[r_selected_actor], font_size_pt=14)
                            st.download_button(
                                label=f"⬇️ Tải file của {r_selected_actor} (.docx)", data=r_act_buf,
                                file_name=f"KichBan_{r_selected_actor}_{st.session_state['resync_stats']['video_title']}_Final.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="btn_dl_single_actor_resync"
                            )
                    with col_ract2:
                        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                        st.download_button(
                            label="📦 Tải trọn bộ (.zip)", data=st.session_state['r_actor_zip'],
                            file_name=st.session_state['r_zip_name'], mime="application/zip", type="secondary", use_container_width=True, key="btn_dl_zip_actor_resync"
                        )
                if st.session_state.pop('_celebrate_tab2', False): st.balloons()
        else: st.info("📌 **Hãy tải file kịch bản đã qua chỉnh sửa thủ công để hệ thống phục hồi lại định dạng chuẩn.**")

    with col_r2:
        st.markdown("### 📊 Thống kê bản Re-Sync")
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
