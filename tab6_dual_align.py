import streamlit as st
import io
import time
import pandas as pd
from utils import (
    TRACKER_DB_FILE, RATES_DB_FILE, save_json_db, 
    generate_actor_salary_slip_docx
)

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
        st.markdown("#### Bảng quản lý danh sách Video và Thù lao (Định dạng chuẩn 6 cột)")
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
                    "Tuần dự án": st.column_config.TextColumn("Tuần dự án (Sửa trực tiếp)"),
                    "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video (Sửa trực tiếp)"),
                    "Thời lượng": st.column_config.TextColumn("Thời lượng", disabled=True),
                    "Thành tiền Video": st.column_config.TextColumn("Thành tiền (Tự động tính theo giá riêng)", disabled=True),
                    "Diễn viên": st.column_config.TextColumn("Diễn viên (Sửa trực tiếp)"),
                    "Xóa": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="dubbing_tracker_editor_main"
            )

            col_tr1, col_tr2 = st.columns([1, 1])
            with col_tr1:
                if st.button("💾 LƯU THAY ĐỔI TRÊN BẢNG VIDEO", type="primary", use_container_width=True):
                    new_tracker = []; deleted_cnt = 0
                    for idx_r, row in edited_tracker_df.iterrows():
                        if row["Xóa"]: deleted_cnt += 1
                        else:
                            orig_item = tracker_data[idx_r]
                            orig_item['project_week'] = str(row["Tuần dự án"]).strip()
                            orig_item['video_title'] = str(row["Tiêu đề video"]).strip()
                            
                            new_acts_raw = [a.strip().upper() for a in str(row["Diễn viên"]).split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                            orig_item['actors'] = ", ".join(new_acts_raw) if new_acts_raw else "CHƯA CÓ THÔNG TIN"
                            
                            old_bd = orig_item.get("actor_breakdown", {}); new_bd = {}
                            c_rates = orig_item.get("custom_actor_rates", {})
                            for act in new_acts_raw:
                                if act in old_bd: new_bd[act] = old_bd[act]
                                else: new_bd[act] = {"lines": 0, "words": 0}
                                if act not in c_rates: c_rates[act] = current_rate
                            
                            orig_item['actor_breakdown'] = new_bd
                            orig_item['custom_actor_rates'] = c_rates
                            new_tracker.append(orig_item)
                            
                    st.session_state['dubbing_tracker'] = new_tracker
                    save_json_db(TRACKER_DB_FILE, new_tracker)
                    st.success(f"✅ Đã lưu thay đổi & đồng bộ dữ liệu!"); time.sleep(1); st.rerun()

            st.markdown("---")
            st.markdown("#### 📑 Xem theo Bảng chia Tuần Dự Án (Có Hàng Tổng)")

            project_weeks_map = {}
            for item in tracker_data:
                pw = item.get("project_week", "Tuần 1")
                if pw not in project_weeks_map: project_weeks_map[pw] = []
                project_weeks_map[pw].append(item)

            total_studio_money = 0; excel_sheets_data = {}

            for pw_name, items in sorted(project_weeks_map.items()):
                st.markdown(f"##### 📅 {pw_name.upper()}")
                week_rows = []; week_tot_dur = 0; week_tot_money = 0
                
                for idx, item in enumerate(items, 1):
                    v_title = item['video_title']; v_dur_min = int(item.get("video_duration_min", 1))
                    v_lines = item.get("total_lines", 0); v_actors = item.get("actors", "")
                    bd = item.get("actor_breakdown", {}); custom_rates = item.get("custom_actor_rates", {})
                    
                    raw_acts = [a.strip().upper() for a in v_actors.split(',') if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]
                    
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
                        
                    week_tot_dur += v_dur_min; week_tot_money += v_pay
                    
                    week_rows.append({
                        "Stt": idx, "Tiêu đề video": v_title, "Thời lượng": dur_str,
                        "Thành tiền": f"{v_pay:,.0f} VNĐ", "Diễn viên": v_actors
                    })

                total_studio_money += week_tot_money
                week_rows.append({
                    "Stt": "TỔNG", "Tiêu đề video": f"TỔNG CỘNG ({len(items)} video)",
                    "Thời lượng": f"{week_tot_dur} phút" if current_mode == "minute" else "-",
                    "Thành tiền": f"{week_tot_money:,.0f} VNĐ", "Diễn viên": "-"
                })

                df_week_display = pd.DataFrame(week_rows)
                st.dataframe(df_week_display, hide_index=True, use_container_width=True)
                excel_sheets_data[pw_name[:30]] = df_week_display

            with col_tr2:
                excel_payroll_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_payroll_buffer, engine='openpyxl') as writer:
                    for s_name, s_df in excel_sheets_data.items():
                        s_df.to_excel(writer, index=False, sheet_name=s_name.replace(":", "_").replace("/", "_"))
                excel_payroll_buffer.seek(0)

                st.download_button(
                    label="📊 XUẤT TẤT CẢ TUẦN RA EXCEL (.XLSX)", data=excel_payroll_buffer,
                    file_name="Theo_Doi_Video_Long_Tieng_MaiHan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
                )
        else: st.info("Chưa có dữ liệu video nào. Hãy chạy Re-Sync ở Tab 2 để tự động ghi nhận video mới!")

    with subtab_custom_rates:
        st.markdown("#### 💵 Bảng Điều Chỉnh Đơn Giá Cá Nhân (Theo Video & Diễn viên)")
        st.caption("Gõ đơn giá riêng cho từng người trong từng video (Ví dụ: Khánh video A 30k, video B 20k; Trúc video A 15k...):")
        
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
                        "Đơn giá cá nhân (VNĐ)": st.column_config.NumberColumn("Đơn giá riêng (Gõ để sửa)", format="%d", step=1000)
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
                    st.success("✅ Đã lưu đơn giá riêng cho từng diễn viên thành công!"); time.sleep(1); st.rerun()
            else: st.info("Chưa có thông tin diễn viên lồng tiếng trong danh sách video.")
        else: st.info("Chưa có danh sách video để điều chỉnh đơn giá.")

    with subtab_payroll:
        st.markdown("#### Bảng Tổng kết & Báo cáo Lương Chi Tiết Từng Diễn Viên")
        if tracker_data:
            col_filter1, col_filter2 = st.columns([1, 1])
            all_project_weeks = sorted(list(set(item.get("project_week", "Tuần 1") for item in tracker_data)))
            with col_filter1: selected_week_filter = st.selectbox("🎯 Chọn Tuần Dự Án:", options=["TẤT CẢ CÁC TUẦN"] + all_project_weeks)

            if selected_week_filter == "TẤT CẢ CÁC TUẦN": active_tracker_data = tracker_data
            else: active_tracker_data = [item for item in tracker_data if item.get("project_week", "Tuần 1") == selected_week_filter]

            actor_weekly_map = {}
            for item in active_tracker_data:
                v_title = item['video_title']; v_dur = int(item.get("video_duration_min", 1))
                bd = item.get("actor_breakdown", {}); v_lines = item.get("total_lines", 0)
                custom_rates = item.get("custom_actor_rates", {})
                acting_actors = [a.strip().upper() for a in item.get("actors", "").split(",") if a.strip() and a.strip() != "CHƯA CÓ THÔNG TIN"]

                for act_name_clean in acting_actors:
                    if act_name_clean not in actor_weekly_map:
                        actor_weekly_map[act_name_clean] = {
                            "videos_count": 0, "total_video_mins": 0, "total_lines": 0, "total_words": 0, "video_rows": []
                        }
                    
                    act_rate = custom_rates.get(act_name_clean, current_rate)
                    act_lines = bd[act_name_clean]["lines"] if (bd and act_name_clean in bd) else v_lines
                    act_words = bd[act_name_clean]["words"] if (bd and act_name_clean in bd) else 0
                    
                    if current_mode == "minute": act_pay = v_dur * act_rate; dur_str = f"{v_dur} phút"
                    elif current_mode == "line": act_pay = act_lines * act_rate; dur_str = f"{act_lines} câu"
                    else: act_pay = act_words * act_rate; dur_str = f"{act_words} từ"

                    actor_weekly_map[act_name_clean]["videos_count"] += 1
                    actor_weekly_map[act_name_clean]["total_video_mins"] += v_dur
                    actor_weekly_map[act_name_clean]["total_lines"] += act_lines
                    actor_weekly_map[act_name_clean]["total_words"] += act_words
                    actor_weekly_map[act_name_clean]["video_rows"].append({
                        "Stt": len(actor_weekly_map[act_name_clean]["video_rows"]) + 1,
                        "Tiêu đề video": v_title, "Thời lượng": dur_str,
                        "Đơn giá": f"{act_rate:,.0f} VNĐ", "Thành tiền": f"{act_pay:,.0f} VNĐ", "Pay_Num": act_pay
                    })

            all_actor_names = sorted(list(actor_weekly_map.keys()))
            with col_filter2:
                selected_actor_view = st.selectbox("👤 Chọn Diễn viên để xem & xuất Báo cáo Lương cá nhân:", options=["TẤT CẢ DIỄN VIÊN"] + all_actor_names)

            st.markdown("---")

            if selected_actor_view != "TẤT CẢ DIỄN VIÊN":
                a_info = actor_weekly_map[selected_actor_view]; a_rows = a_info["video_rows"]
                a_tot_pay = sum(r["Pay_Num"] for r in a_rows)
                st.subheader(f"👤 PHIẾU BÁO CÁO THÙ LAO: {selected_actor_view}")
                st.caption(f"Dữ liệu thù lao lồng tiếng cho {selected_actor_view} ({selected_week_filter})")
                
                df_single_actor = pd.DataFrame(a_rows)[["Stt", "Tiêu đề video", "Thời lượng", "Đơn giá", "Thành tiền"]]
                st.dataframe(df_single_actor, hide_index=True, use_container_width=True)
                st.metric(f"💰 TỔNG THÙ LAO DỰ KIẾN TRẢ CHO {selected_actor_view}:", f"{a_tot_pay:,.0f} VNĐ")
                
                col_p_btn1, col_p_btn2 = st.columns(2)
                with col_p_btn1:
                    actor_docx_buf = generate_actor_salary_slip_docx(selected_actor_view, selected_week_filter, a_rows, a_tot_pay, current_mode)
                    st.download_button(
                        label=f"🖨️ IN / TẢI PHIẾU LƯƠNG WORD CỦA {selected_actor_view} (.DOCX)",
                        data=actor_docx_buf, file_name=f"PhieuLuong_{selected_actor_view}_{selected_week_filter}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", use_container_width=True
                    )
                with col_p_btn2:
                    excel_single_buf = io.BytesIO()
                    with pd.ExcelWriter(excel_single_buf, engine='openpyxl') as writer: df_single_actor.to_excel(writer, index=False, sheet_name="Phieu Luong")
                    excel_single_buf.seek(0)
                    st.download_button(
                        label=f"📊 TẢI PHIẾU LƯƠNG CÁ NHÂN EXCEL (.XLSX)",
                        data=excel_single_buf, file_name=f"PhieuLuong_{selected_actor_view}_{selected_week_filter}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
                    )
            else:
                actor_payroll_rows = []; grand_actor_pay = 0; unique_video_titles = set(); grand_unique_mins = 0
                for item in active_tracker_data:
                    if item['video_title'] not in unique_video_titles:
                        unique_video_titles.add(item['video_title'])
                        grand_unique_mins += int(item.get("video_duration_min", 1))

                for idx, (act_name, info) in enumerate(sorted(actor_weekly_map.items()), 1):
                    if current_mode == "minute": unit_cnt = info["total_video_mins"]; dur_disp = f"{unit_cnt} phút"
                    elif current_mode == "line": unit_cnt = info["total_lines"]; dur_disp = f"{unit_cnt} câu"
                    else: unit_cnt = info["total_words"]; dur_disp = f"{unit_cnt} từ"

                    tot_p = sum(r["Pay_Num"] for r in info["video_rows"])
                    grand_actor_pay += tot_p

                    actor_payroll_rows.append({
                        "Stt": idx, "Diễn viên Lồng tiếng": act_name, "Số Video đã lồng": info["videos_count"],
                        "Tổng phút video": dur_disp, "Thành tiền Lương": f"{tot_p:,.0f} VNĐ",
                        "Danh sách Video tham gia": ", ".join([r["Tiêu đề video"] for r in info["video_rows"]])
                    })

                actor_payroll_rows.append({
                    "Stt": "TỔNG", "Diễn viên Lồng tiếng": f"TỔNG CỘNG ({len(actor_weekly_map)} Diễn viên)",
                    "Số Video đã lồng": "-", "Tổng phút video": f"{grand_unique_mins} phút" if current_mode == "minute" else "-",
                    "Thành tiền Lương": f"{grand_actor_pay:,.0f} VNĐ", "Danh sách Video tham gia": "-"
                })

                df_act_payroll = pd.DataFrame(actor_payroll_rows)
                st.metric(f"💰 TỔNG LƯƠNG CẦN CHI CHO DIỄN VIÊN ({selected_week_filter}):", f"{grand_actor_pay:,.0f} VNĐ")
                st.dataframe(
                    df_act_payroll[["Stt", "Diễn viên Lồng tiếng", "Số Video đã lồng", "Tổng phút video", "Thành tiền Lương", "Danh sách Video tham gia"]],
                    hide_index=True, use_container_width=True
                )

                excel_actor_buf = io.BytesIO()
                with pd.ExcelWriter(excel_actor_buf, engine='openpyxl') as writer: df_act_payroll.to_excel(writer, index=False, sheet_name="Luong Dien Vien")
                excel_actor_buf.seek(0)
                st.download_button(
                    label=f"📊 TẢI BÁO CÁO LƯƠNG TOÀN BỘ DIỄN VIÊN ({selected_week_filter}) EXCEL (.XLSX)",
                    data=excel_actor_buf, file_name=f"Bao_Cao_Luong_DienVien_{selected_week_filter}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True
                )
        else: st.info("Chưa có dữ liệu lồng tiếng để lập báo cáo lương.")
