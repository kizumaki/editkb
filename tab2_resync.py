import streamlit as st
import io
import time
import pandas as pd
from utils import process_docx, clean_file_name_for_output, save_json_db, TRACKER_DB_FILE

def render_tab2(enable_colors, enable_phonetic, enable_cast):
    st.markdown("### 🔄 Re-Sync Kịch Bản Đã Biên Tập Thủ Công")
    st.caption("Dành riêng cho file kịch bản đã được team biên tập chỉnh sửa lời thoại. Tự động phục hồi màu sắc, phân vai và xuất phông chữ 14pt cực nét cho phòng thu.")
    
    r_uploaded_file = st.file_uploader(
        "Kéo thả file .docx đã biên tập vào đây", 
        type=["docx"], 
        key=f"resync_uploader_{st.session_state['resync_uploader_key']}"
    )
    
    selected_week = st.text_input("📌 Gán Tuần Dự Án cho video này:", value="Tuần 1", help="Dùng để gom nhóm xuất báo cáo lương lồng tiếng ở Tab 3")

    if r_uploaded_file is not None:
        st.success(f"📄 Đã nhận file kịch bản biên tập: **{r_uploaded_file.name}**")
        
        if st.button("🚀 2. BẮT ĐẦU RE-SYNC & CHUẨN HÓA LẠI ĐỊNH DẠNG (14PT)", type="primary", use_container_width=True):
            with st.spinner("Đang Re-Sync & chuẩn hóa kịch bản..."):
                r_docx_f, r_ass_f, r_srt_f, r_zip_f, r_stats = process_docx(
                    uploaded_file=r_uploaded_file,
                    file_name_without_ext=r_uploaded_file.name,
                    enable_colors=enable_colors,
                    enable_phonetic=enable_phonetic,
                    enable_cast=enable_cast,
                    is_resync=True
                )
                
                st.session_state['r_processed_docx'] = r_docx_f
                st.session_state['r_processed_ass'] = r_ass_f
                st.session_state['r_processed_srt'] = r_srt_f
                st.session_state['r_actor_zip'] = r_zip_f
                st.session_state['r_stats'] = r_stats
                st.session_state['r_filename'] = r_uploaded_file.name
                
                # Ghi nhận vào Tracker Lương
                if r_stats and r_stats.get('actor_stats_breakdown'):
                    v_title = r_stats['video_title']
                    v_dur = r_stats['video_duration_min']
                    for act_name, act_info in r_stats['actor_stats_breakdown'].items():
                        new_record = {
                            "video_title": v_title,
                            "actor_name": act_name,
                            "week_name": selected_week,
                            "duration_min": v_dur,
                            "line_count": act_info["lines"],
                            "word_count": act_info["words"],
                            "timestamp": time.strftime("%Y-%m-%d %H:%M")
                        }
                        existing = [r for r in st.session_state.get('dubbing_tracker', []) if r.get('video_title') == v_title and r.get('actor_name') == act_name and r.get('week_name') == selected_week]
                        if not existing:
                            if 'dubbing_tracker' not in st.session_state:
                                st.session_state['dubbing_tracker'] = []
                            st.session_state['dubbing_tracker'].append(new_record)
                    save_json_db(TRACKER_DB_FILE, st.session_state['dubbing_tracker'])

    # HIỂN THỊ KẾT QUẢ VÀ BÁO CÁO SO SÁNH ĐỘ TOÀN VẸN
    if 'r_processed_docx' in st.session_state:
        st.markdown("---")
        st.markdown("### 📥 Tải Xuất File Re-Sync Final")
        
        orig_name = st.session_state.get('r_filename', 'script.docx')
        out_docx_name = clean_file_name_for_output(orig_name, tag="_resync", ext=".docx")
        out_ass_name = clean_file_name_for_output(orig_name, tag="_resync", ext=".ass")
        out_srt_name = clean_file_name_for_output(orig_name, tag="_resync", ext=".srt")
        out_zip_name = clean_file_name_for_output(orig_name, tag="_KichBan_DienVien", ext=".zip")

        col1, col2, col3, col4 = st.columns(4)
        col1.download_button("📄 Tải Docx Final (14pt)", st.session_state['r_processed_docx'], file_name=out_docx_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        col2.download_button("🎬 Tải Phụ Đề ASS", st.session_state['r_processed_ass'], file_name=out_ass_name, mime="text/plain", use_container_width=True)
        col3.download_button("💬 Tải Phụ Đề SRT", st.session_state['r_processed_srt'], file_name=out_srt_name, mime="text/plain", use_container_width=True)
        col4.download_button("📦 Gói Kịch Bản Diễn Viên", st.session_state['r_actor_zip'], file_name=out_zip_name, mime="application/zip", use_container_width=True)

        # 🔍 KHỐI BÁO CÁO SO SÁNH ĐỘ TOÀN VẸN (DIFF CHECK)
        r_stats = st.session_state.get('r_stats', {})
        integrity = r_stats.get('integrity_report', {})
        
        st.markdown("---")
        st.markdown("### 🔍 Báo Cáo So Sánh Độ Toàn Vẹn (File Edit vs File Final)")
        
        if integrity:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Mốc Timecode (Edit)", integrity.get("tc_in_cnt", 0))
            m2.metric("Mốc Timecode (Final)", integrity.get("tc_out_cnt", 0))
            m3.metric("Số dòng thoại (Edit)", integrity.get("line_in_cnt", 0))
            m4.metric("Số dòng thoại (Final)", integrity.get("line_out_cnt", 0))
            
            diff_issues = integrity.get("diff_issues", [])
            if not diff_issues:
                st.success("🎉 KHỚP CHUẨN 100%! Tất cả timecode và nội dung thoại trùng khớp hoàn toàn. Không phát hiện mất chữ, sụt giảm ký tự hay lệch timecode nào giữa file Biên tập và Final.")
            else:
                st.warning(f"⚠️ Phát hiện **{len(diff_issues)}** điểm khác biệt cần lưu ý giữa File Edit và File Final:")
                df_diff = pd.DataFrame(diff_issues)
                st.dataframe(df_diff, use_container_width=True)
