import streamlit as st
import io
import time
import pandas as pd
from utils import TRACKER_DB_FILE, RATES_DB_FILE, save_json_db, generate_actor_salary_slip_docx

def render_tab3():
    st.subheader("📋 BÁO CÁO BẢNG TÍNH LƯƠNG LỒNG TIẾNG THEO TUẦN DỰ ÁN")
    
    with st.expander("⚙️ CẤU HÌNH ĐƠN GIÁ MẶC ĐỊNH SẢN XUẤT", expanded=False):
        c_rate1, c_rate2 = st.columns([2, 3])
        curr_mode = st.session_state['payroll_rates'].get("mode", "minute")
        mode_idx = 0 if curr_mode == "minute" else (1 if curr_mode == "line" else 2)
        
        with c_rate1:
            rate_mode_choice = st.radio(
                "Cách tính thù lao:", options=["Theo Phút video (phút)", "Theo Câu thoại (câu)", "Theo Số từ (từ)"], 
                index=mode_idx, key="radio_payroll_mode"
            )
        with c_rate2:
            default_unit_rate = st.number_input(
                "Đơn giá mặc định khi thêm video mới (VNĐ):", value=int(st.session_state['payroll_rates'].get("unit_rate", 30000)), 
                step=5000, key="num_payroll_unit_rate"
            )
            unit_label = "phút" if "Phút" in rate_mode_choice else ("câu" if "Câu" in rate_mode_choice else "từ")
            st.caption(f"👉 **Đơn giá mặc định áp dụng:** `{default_unit_rate:,.0f}` VNĐ / {unit_label}")
            
        new_mode = "minute" if "Phút" in rate_mode_choice else ("line" if "Câu" in rate_mode_choice else "word")
            
        if new_mode != st.session_state['payroll_rates'].get("mode") or default_unit_rate != st.session_state['payroll_rates'].get("unit_rate"):
            st.session_state['payroll_rates'] = {"mode": new_mode, "unit_rate": default_unit_rate}
            save_json_db(RATES_DB_FILE, st.session_state['payroll_rates'])

    current_mode = st.session_state['payroll_rates'].get("mode", "minute")
    current_rate = st.session_state['payroll_rates'].get("unit_rate", 30000)

    subtab_video, subtab_custom_rates, subtab_payroll = st.tabs([
        "📹 Bảng Theo Dõi Video Theo Tuần Dự Án", "💵 Chỉnh Đơn Giá Cá Nhân (Video x Diễn viên)", "💰 Báo Cáo Lương Chi Tiết Từng Diễn Viên"
    ])

    tracker_data = st.session_state['dubbing_tracker']

    with subtab_video:
        st.markdown("#### Bảng quản lý danh sách Video và Thù lao")
        if tracker_data:
            formatted_editor_list = []
            for idx, item in enumerate(tracker_data, 1):
                p_week = item.get("project_week", "Tuần 1")
                v_dur_min = int(item.get("video_duration_min", 1))
                v_lines = item.get("total_lines", 0)
                v_acts = item.get("actors", "")
                bd = item.get("actor_breakdown", {})
                custom_rates = item.get("custom_actor_rates", {})
                raw_acts = [a.strip().upper() for a in v_acts.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                
                if current_mode == "minute":
                    v_pay = sum(v_dur_min * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_dur_min * current_rate
                    dur_str = f"{v_dur_min} phút"
                elif current_mode == "line":
                    v_pay = sum((bd.get(act, {}).get("lines", v_lines)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else v_lines * current_rate
                    dur_str = f"{v_lines} câu"
                else:
                    tot_words = sum(a["words"] for a in bd.values()) if bd else 0
                    v_pay = sum((bd.get(act, {}).get("words", 0)) * custom_rates.get(act, current_rate) for act in raw_acts) if raw_acts else tot_words * current_rate
                    dur_str = f"{tot_words} từ"
                    
                formatted_editor_list.append({
                    "Stt": idx, "Tuần dự án": p_week, "Tiêu đề video": item['video_title'],
                    "Thời lượng": dur_str, "Thành tiền Video": f"{v_pay:,.0f} VNĐ", "Diễn viên": v_acts, "Xóa": False
                })

            df_editor_raw = pd.DataFrame(formatted_editor_list)
            edited_tracker_df = st.data_editor(
                df_editor_raw,
                column_config={
                    "Stt": st.column_config.NumberColumn("Stt", disabled=True, width="small"),
                    "Tuần dự án": st.column_config.TextColumn("Tuần dự án"),
                    "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video"),
                    "Thời lượng": st.column_config.TextColumn("Thời lượng", disabled=True),
                    "Thành tiền Video": st.column_config.TextColumn("Thành tiền", disabled=True),
                    "Diễn viên": st.column_config.TextColumn("Diễn viên"),
                    "Xóa": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="dubbing_tracker_editor_main"
            )

            if st.button("💾 LƯU THAY ĐỔI TRÊN BẢNG VIDEO", type="primary", use_container_width=True):
                new_tracker = []
                for idx_r, row in edited_tracker_df.iterrows():
                    if not row["Xóa"]:
                        orig_item = tracker_data[idx_r]
                        orig_item['project_week'] = str(row["Tuần dự án"]).strip()
                        orig_item['video_title'] = str(row["Tiêu đề video"]).strip()
                        new_acts_raw = [a.strip().upper() for a in str(row["Diễn viên"]).split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                        orig_item['actors'] = ", ".join(new_acts_raw) if new_acts_raw else "CHƯA CÓ THÔNG TIN"
                        new_tracker.append(orig_item)
                        
                st.session_state['dubbing_tracker'] = new_tracker
                save_json_db(TRACKER_DB_FILE, new_tracker)
                st.success("✅ Đã lưu thay đổi!"); time.sleep(1); st.rerun()
        else:
            st.info("Chưa có dữ liệu video nào.")

    with subtab_custom_rates:
        st.markdown("#### 💵 Chỉnh Đơn Giá Cá Nhân (Theo Video & Diễn viên)")
        if tracker_data:
            rate_editor_rows = []
            for v_idx, item in enumerate(tracker_data):
                pw = item.get("project_week", "Tuần 1"); v_title = item['video_title']
                v_dur = int(item.get("video_duration_min", 1)); v_acts = item.get("actors", "")
                custom_rates = item.get("custom_actor_rates", {})
                acting_actors = [a.strip().upper() for a in v_acts.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                
                for act in acting_actors:
                    act_rate = custom_rates.get(act, current_rate)
                    rate_editor_rows.append({
                        "v_idx": v_idx, "Tuần dự án": pw, "Tiêu đề video": v_title,
                        "Thời lượng (Phút)": v_dur, "Diễn viên": act, "Đơn giá cá nhân (VNĐ)": int(act_rate)
                    })

            if rate_editor_rows:
                df_custom_rates = pd.DataFrame(rate_editor_rows)
                edited_custom_rates_df = st.data_editor(
                    df_custom_rates[["Tuần dự án", "Tiêu đề video", "Thời lượng (Phút)", "Diễn viên", "Đơn giá cá nhân (VNĐ)"]],
                    column_config={
                        "Tuần dự án": st.column_config.TextColumn("Tuần", disabled=True),
                        "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video", disabled=True),
                        "Thời lượng (Phút)": st.column_config.NumberColumn("Độ dài (Phút)", disabled=True),
                        "Diễn viên": st.column_config.TextColumn("Diễn viên", disabled=True),
                        "Đơn giá cá nhân (VNĐ)": st.column_config.NumberColumn("Đơn giá riêng", format="%d", step=1000)
                    },
                    hide_index=True, use_container_width=True, key="custom_rates_editor_table_main"
                )

                if st.button("💾 LƯU ĐƠN GIÁ CÁ NHÂN", type="primary", use_container_width=True):
                    for idx_r, row in edited_custom_rates_df.iterrows():
                        v_i = df_custom_rates.iloc[idx_r]["v_idx"]; act_n = df_custom_rates.iloc[idx_r]["Diễn viên"]
                        new_r = float(row["Đơn giá cá nhân (VNĐ)"])
                        if "custom_actor_rates" not in tracker_data[v_i]: tracker_data[v_i]["custom_actor_rates"] = {}
                        tracker_data[v_i]["custom_actor_rates"][act_n] = new_r
                        
                    st.session_state['dubbing_tracker'] = tracker_data
                    save_json_db(TRACKER_DB_FILE, tracker_data)
                    st.success("✅ Đã lưu đơn giá riêng thành công!"); time.sleep(1); st.rerun()

    with subtab_payroll:
        st.markdown("#### 💰 Báo cáo Lương Chi Tiết Từng Diễn Viên")
        if tracker_data:
            st.info("Tính năng báo cáo thù lao chi tiết theo tuần sẵn sàng!")
        else:
            st.info("Chưa có dữ liệu để lập báo cáo lương.")
