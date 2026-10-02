import streamlit as st
import io
import os
import zipfile
import pandas as pd
from utils import (
    kept_file_uploader, process_docx, clean_file_name_for_output,
    record_video_in_tracker, save_json_db, TRACKER_DB_FILE, project_picker
)

def render_batch_processing(is_resync, enable_colors, enable_phonetic, enable_cast):
    """Xử lý nhiều kịch bản .docx một lượt, gom kết quả vào 1 file .zip."""
    mode = "resync" if is_resync else "goc"
    tag = "_final" if is_resync else "_edit"
    font_size = 14  # kịch bản gốc & Re-Sync cùng định dạng
    result_key = f"batch_result_{mode}"

    with st.container(border=True):
        st.markdown("### 📚 Xử lý nhiều kịch bản cùng lúc")
        st.caption("Kéo thả nhiều file .docx một lượt. Mỗi file được xử lý y hệt chế độ một file, kết quả gom vào 1 file .zip "
                   "(mỗi kịch bản 1 thư mục: Word, phụ đề .ass/.srt và kịch bản tách vai).")
        files = kept_file_uploader("Kéo thả các file .docx vào đây", type=["docx"],
                                   accept_multiple_files=True, key=f"batch_{mode}_uploader")
        if is_resync:
            col_pj1, col_pj2 = st.columns(2)
            with col_pj1: project_picker("📁 TẤT CẢ video trong lượt này thuộc dự án:", key="resync_project_id")
            with col_pj2: st.text_input("📌 Ghi chú tuần/đợt (không bắt buộc):", key="resync_project_week")
        else:
            st.caption("💡 Chế độ này không có bước soát tên nhân vật/phiên âm cho từng file. "
                       "Nếu kịch bản có nhân vật hoặc từ tiếng Anh mới, nên xử lý riêng file đó ở chế độ một file trước.")

    if not files:
        st.info("📌 Tải lên các file kịch bản (.docx) ở trên để bắt đầu.")
        return

    st.success(f"Đã nhận **{len(files)}** file.")
    if st.button(f"✨ Xử lý {len(files)} kịch bản", type="primary", use_container_width=True, key=f"btn_batch_{mode}"):
        summary, zip_buf, ok_count = [], io.BytesIO(), 0
        progress = st.progress(0, text="Đang bắt đầu...")
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, f in enumerate(files, 1):
                progress.progress((i - 1) / len(files), text=f"Đang xử lý {i}/{len(files)}: {f.name}")
                name_no_ext = os.path.splitext(f.name)[0]
                folder = clean_file_name_for_output(f.name, tag="", ext="")
                try:
                    docx_f, ass_f, srt_f, act_zip, stats = process_docx(
                        f, name_no_ext, enable_colors, enable_phonetic, enable_cast,
                        is_resync=is_resync, font_size_pt=font_size)
                    zf.writestr(f"{folder}/{clean_file_name_for_output(f.name, tag=tag, ext='.docx')}", docx_f.getvalue())
                    zf.writestr(f"{folder}/{clean_file_name_for_output(f.name, tag=tag, ext='.ass')}", ass_f.getvalue())
                    zf.writestr(f"{folder}/{clean_file_name_for_output(f.name, tag=tag, ext='.srt')}", srt_f.getvalue())
                    with zipfile.ZipFile(act_zip) as az:
                        for n in az.namelist():
                            zf.writestr(f"{folder}/Tach_vai/{n}", az.read(n))
                    if is_resync:
                        record_video_in_tracker(stats, name_no_ext, st.session_state.get("resync_project_week"),
                                                st.session_state.get("resync_project_id"))
                    warns = stats.get("qc_warnings", [])
                    integ = stats.get("integrity_report", {}).get("diff_issues", [])
                    summary.append({
                        "File": f.name, "Kết quả": "✅ Xong",
                        "Nhân vật": stats.get("total_speakers", 0), "Câu thoại": stats.get("total_lines", 0),
                        "Thời lượng (phút)": stats.get("video_duration_min", 0),
                        "Cảnh báo cần xem": len(warns) + len(integ),
                        "Ghi chú": next((w for w in warns if "chưa gán diễn viên" in w), "")[:150],
                    })
                    ok_count += 1
                except Exception as e:
                    note = ("File hỏng hoặc không phải file Word thật (.docx). Thử mở bằng Word rồi lưu lại."
                            if isinstance(e, (zipfile.BadZipFile, KeyError)) else f"Lỗi khi xử lý: {e}"[:150])
                    summary.append({"File": f.name, "Kết quả": "❌ Lỗi", "Nhân vật": 0, "Câu thoại": 0,
                                    "Thời lượng (phút)": 0, "Cảnh báo cần xem": 0, "Ghi chú": note})
        progress.progress(1.0, text="Hoàn tất!")
        if is_resync and ok_count:
            save_json_db(TRACKER_DB_FILE, st.session_state["dubbing_tracker"])  # lưu bảng lương 1 lần cho cả lượt
        zip_buf.seek(0)
        st.session_state[result_key] = {"zip": zip_buf.getvalue(), "summary": summary, "ok": ok_count, "total": len(files)}

    res = st.session_state.get(result_key)
    if res:
        st.markdown("---")
        st.markdown(f"### ⬇️ Kết quả: {res['ok']}/{res['total']} file xử lý thành công")
        if res["ok"] < res["total"]:
            st.warning("Có file bị lỗi, xem cột **Ghi chú** bên dưới. Các file còn lại vẫn có trong file .zip.")
        st.dataframe(pd.DataFrame(res["summary"]), hide_index=True, use_container_width=True)
        if is_resync and res["ok"]:
            _pname = next((p.get("name") for p in st.session_state.get("projects", [])
                           if p.get("project_id") == st.session_state.get("resync_project_id")), "")
            st.caption(f"📋 Đã ghi {res['ok']} video vào bảng lương (dự án: **{_pname}**).")
        st.download_button("📦 Tải toàn bộ kết quả (.zip)", data=res["zip"],
                           file_name=f"KetQua_{'ReSync' if is_resync else 'KichBanGoc'}_{res['total']}_file.zip",
                           mime="application/zip", type="primary", use_container_width=True, key=f"dl_batch_{mode}")
