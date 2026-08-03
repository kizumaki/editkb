import streamlit as st
import time
import pandas as pd
from utils import CAST_DB_FILE, SPEAKER_COLOR_DB_FILE, save_json_db, hex_to_rgb

def render_tab4():
    subtab_cast_map, subtab_color_map = st.tabs([
        "🎭 Bảng Phân Vai Diễn Viên (Global Database)",
        "🎨 Bảng Màu & Highlight Nhân Vật Cố Định"
    ])

    with subtab_cast_map:
        st.subheader("🎭 BẢNG PHÂN VAI LỒNG TIẾNG (GLOBAL DATABASE)")
        st.markdown("Nơi thiết lập mặc định nhân vật Tiếng Anh nào sẽ do diễn viên lồng tiếng Việt nào đảm nhận cho Mai Han Team.")

        st.markdown("#### ➕ Thêm / Cập nhật Phân vai mới")
        c_c1, c_c2, c_c3 = st.columns([2, 2, 1.2])
        with c_c1: add_role_eng = st.text_input("Tên Nhân vật (Tiếng Anh):", placeholder="VD: Bri, Chase...", key=f"add_role_eng_{st.session_state.get('cast_input_key', 0)}")
        with c_c2: add_actor_vn = st.text_input("Diễn viên Lồng tiếng (Tiếng Việt):", placeholder="VD: TRÚC, THIỆN...", key=f"add_actor_vn_{st.session_state.get('cast_input_key', 0)}")
        with c_c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Phân Vai", use_container_width=True, type="primary", key="btn_add_cast"):
                if add_role_eng and add_actor_vn:
                    k = add_role_eng.upper().strip(); v = add_actor_vn.strip().upper()
                    st.session_state['custom_cast_mapping'][k] = v
                    save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                    st.session_state['cast_input_key'] = st.session_state.get('cast_input_key', 0) + 1
                    st.success(f"✅ Đã gán thành công: `{add_role_eng}` ➔ `{add_actor_vn}`"); time.sleep(1); st.rerun()
                else: st.warning("Vui lòng nhập đầy đủ tên nhân vật và diễn viên!")

        st.markdown("---")
        search_cast_query = st.text_input("🔍 Tìm kiếm Nhân vật hoặc Diễn viên lồng tiếng:", placeholder="Gõ tên nhân vật hoặc diễn viên...").strip().upper()
        all_cast_dict = st.session_state.get('custom_cast_mapping', {})

        if search_cast_query: filtered_cast = {k: v for k, v in all_cast_dict.items() if search_cast_query in k or search_cast_query in v.upper()}
        else: filtered_cast = all_cast_dict

        if filtered_cast:
            cast_db_table = []
            for eng_role, vn_actor in sorted(filtered_cast.items()):
                cast_db_table.append({"Nhân vật (Tiếng Anh)": eng_role, "Diễn viên Lồng tiếng (Tiếng Việt)": vn_actor, "Xóa khỏi Database": False})

            df_cast_db = pd.DataFrame(cast_db_table)
            st.caption(f"Đang hiển thị **{len(df_cast_db)}** vai lồng tiếng trong kho:")

            edited_cast_db_df = st.data_editor(
                df_cast_db,
                column_config={
                    "Nhân vật (Tiếng Anh)": st.column_config.TextColumn("Tên Nhân vật gốc (In hoa)", disabled=True),
                    "Diễn viên Lồng tiếng (Tiếng Việt)": st.column_config.TextColumn("Diễn viên lồng tiếng (Sửa trực tiếp)"),
                    "Xóa khỏi Database": st.column_config.CheckboxColumn("Xóa?")
                },
                disabled=["Nhân vật (Tiếng Anh)"], hide_index=True, use_container_width=True, key="global_cast_db_editor"
            )

            if st.button("💾 LƯU TOÀN BỘ CẬP NHẬT PHÂN VAI", type="primary", use_container_width=True, key="btn_save_global_cast"):
                new_cast_db = {}; deleted_cast_count = 0
                if search_cast_query:
                    for k, v in all_cast_dict.items():
                        if k not in filtered_cast: new_cast_db[k] = v

                for _, row in edited_cast_db_df.iterrows():
                    eng_k = str(row["Nhân vật (Tiếng Anh)"]).upper().strip()
                    act_v = str(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).strip().upper()
                    is_del = row["Xóa khỏi Database"]
                    if is_del: deleted_cast_count += 1
                    else:
                        if act_v: new_cast_db[eng_k] = act_v

                st.session_state['custom_cast_mapping'] = new_cast_db
                save_json_db(CAST_DB_FILE, new_cast_db)
                st.success(f"✅ Đã lưu cập nhật thành công! (Đã xóa {deleted_cast_count} vai)"); time.sleep(1); st.rerun()
        else: st.info("Không tìm thấy vai lồng tiếng nào khớp với từ khóa tìm kiếm.")

    with subtab_color_map:
        st.subheader("🎨 BẢNG MÀU CHỮ & HIGHLIGHT CỐ ĐỊNH")
        st.markdown("Nơi cấu hình màu chữ và màu highlight cố định cho các nhân vật đặc biệt (VD: ALL màu đỏ nền vàng).")
        fixed_color_dict = st.session_state.get('fixed_speaker_colors', {})

        col_col1, col_col2, col_col3, col_col4 = st.columns([1.8, 1.5, 1.5, 1.2])
        with col_col1:
            new_color_spk = st.text_input("Tên Nhân vật Cố định:", placeholder="VD: ALL, CORY...", key=f"add_color_spk_{st.session_state.get('color_input_key', 0)}").strip().upper()
        with col_col2:
            enable_tc = st.checkbox("Tô Màu Chữ", value=True, key=f"chk_tc_{st.session_state.get('color_input_key', 0)}")
            new_text_hex = st.color_picker("Chọn Màu Chữ:", "#FF0000", key=f"add_tc_picker_{st.session_state.get('color_input_key', 0)}") if enable_tc else None
        with col_col3:
            enable_hc = st.checkbox("Tô Highlight Nền", value=False, key=f"chk_hc_{st.session_state.get('color_input_key', 0)}")
            new_hl_hex = st.color_picker("Chọn Màu Highlight:", "#FFFF00", key=f"add_hc_picker_{st.session_state.get('color_input_key', 0)}") if enable_hc else None
        with col_col4:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Quy Tắc Màu", use_container_width=True, type="primary", key="btn_add_fixed_color"):
                if new_color_spk:
                    tc_tuple = hex_to_rgb(new_text_hex) if enable_tc else None
                    hc_tuple = hex_to_rgb(new_hl_hex) if enable_hc else None
                    st.session_state['fixed_speaker_colors'][new_color_spk] = {"text_color": tc_tuple, "highlight_color": hc_tuple}
                    save_json_db(SPEAKER_COLOR_DB_FILE, st.session_state['fixed_speaker_colors'])
                    st.session_state['color_input_key'] = st.session_state.get('color_input_key', 0) + 1
                    st.success(f"✅ Đã lưu màu cho `{new_color_spk}`!"); time.sleep(1); st.rerun()

        if fixed_color_dict:
            color_rows = []
            for spk, cfg in sorted(fixed_color_dict.items()):
                tc = cfg.get("text_color") if isinstance(cfg, dict) else cfg
                hc = cfg.get("highlight_color") if isinstance(cfg, dict) else None
                color_rows.append({"Nhân vật": spk, "Màu chữ (RGB)": str(tc), "Màu Highlight": str(hc), "Xóa": False})

            df_colors = pd.DataFrame(color_rows)
            edited_colors_df = st.data_editor(
                df_colors,
                column_config={
                    "Nhân vật": st.column_config.TextColumn("Nhân vật", disabled=True),
                    "Màu chữ (RGB)": st.column_config.TextColumn("Màu chữ", disabled=True),
                    "Màu Highlight": st.column_config.TextColumn("Màu Highlight", disabled=True),
                    "Xóa": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="fixed_colors_editor_table"
            )

            if st.button("💾 LƯU BẢNG MÀU CỐ ĐỊNH", type="primary", use_container_width=True, key="btn_save_colors"):
                new_colors = {}
                for _, row in edited_colors_df.iterrows():
                    if not row["Xóa"]:
                        spk_k = str(row["Nhân vật"]).upper().strip()
                        new_colors[spk_k] = fixed_color_dict.get(spk_k, {})
                st.session_state['fixed_speaker_colors'] = new_colors
                save_json_db(SPEAKER_COLOR_DB_FILE, new_colors)
                st.success("✅ Đã cập nhật Bảng Màu Cố Định!"); time.sleep(1); st.rerun()
