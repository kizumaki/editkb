import streamlit as st
import time
import pandas as pd
from utils import CAST_DB_FILE, SPEAKER_COLOR_DB_FILE, save_json_db, hex_to_rgb

def render_tab4():
    subtab_cast_map, subtab_color_map = st.tabs([
        "🎭 Bảng Phân Vai Diễn Viên (Global Database)",
        "🎨 Bảng Màu & Highlight Nhân Vật Cố Định"
    ])

    # SUBTAB 1: PHÂN VAI DIỄN VIÊN
    with subtab_cast_map:
        st.subheader("🎭 BẢNG PHÂN VAI LỒNG TIẾNG")
        c_c1, c_c2, c_c3 = st.columns([2, 2, 1.2])
        with c_c1: add_role_eng = st.text_input("Tên Nhân vật (Tiếng Anh):", placeholder="VD: Bri...", key=f"add_role_eng_{st.session_state['cast_input_key']}")
        with c_c2: add_actor_vn = st.text_input("Diễn viên Lồng tiếng (Tiếng Việt):", placeholder="VD: TRÚC...", key=f"add_actor_vn_{st.session_state['cast_input_key']}")
        with c_c3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Phân Vai", use_container_width=True, type="primary", key="btn_add_cast"):
                if add_role_eng and add_actor_vn:
                    k = add_role_eng.upper().strip(); v = add_actor_vn.strip().upper()
                    st.session_state['custom_cast_mapping'][k] = v
                    save_json_db(CAST_DB_FILE, st.session_state['custom_cast_mapping'])
                    st.session_state['cast_input_key'] += 1
                    st.success(f"✅ Đã gán thành công: `{add_role_eng}` ➔ `{add_actor_vn}`"); time.sleep(1); st.rerun()

        all_cast_dict = st.session_state['custom_cast_mapping']
        if all_cast_dict:
            cast_db_table = [{"Nhân vật (Tiếng Anh)": k, "Diễn viên Lồng tiếng (Tiếng Việt)": v, "Xóa khỏi Database": False} for k, v in sorted(all_cast_dict.items())]
            df_cast_db = pd.DataFrame(cast_db_table)
            edited_cast_db_df = st.data_editor(
                df_cast_db,
                column_config={
                    "Nhân vật (Tiếng Anh)": st.column_config.TextColumn("Tên Nhân vật", disabled=True),
                    "Diễn viên Lồng tiếng (Tiếng Việt)": st.column_config.TextColumn("Diễn viên lồng tiếng"),
                    "Xóa khỏi Database": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="global_cast_db_editor"
            )

            if st.button("💾 LƯU BẢNG PHÂN VAI", type="primary", use_container_width=True, key="btn_save_global_cast"):
                new_cast_db = {}
                for _, row in edited_cast_db_df.iterrows():
                    if not row["Xóa khỏi Database"]:
                        eng_k = str(row["Nhân vật (Tiếng Anh)"]).upper().strip()
                        act_v = str(row["Diễn viên Lồng tiếng (Tiếng Việt)"]).strip().upper()
                        if act_v: new_cast_db[eng_k] = act_v

                st.session_state['custom_cast_mapping'] = new_cast_db
                save_json_db(CAST_DB_FILE, new_cast_db)
                st.success("✅ Đã lưu cập nhật Phân Vai!"); time.sleep(1); st.rerun()

    # SUBTAB 2: BẢNG MÀU CỐ ĐỊNH
    with subtab_color_map:
        st.subheader("🎨 BẢNG MÀU CHỮ & HIGHLIGHT CỐ ĐỊNH")
        fixed_color_dict = st.session_state['fixed_speaker_colors']
        
        col_col1, col_col2, col_col3, col_col4 = st.columns([1.8, 1.5, 1.5, 1.2])
        with col_col1:
            new_color_spk = st.text_input("Tên Nhân vật Cố định:", placeholder="VD: ALL, CORY...", key=f"add_color_spk_{st.session_state['color_input_key']}").strip().upper()
        with col_col2:
            enable_tc = st.checkbox("Tô Màu Chữ", value=True, key=f"chk_tc_{st.session_state['color_input_key']}")
            new_text_hex = st.color_picker("Chọn Màu Chữ:", "#FF0000", key=f"add_tc_picker_{st.session_state['color_input_key']}") if enable_tc else None
        with col_col3:
            enable_hc = st.checkbox("Tô Highlight Nền", value=False, key=f"chk_hc_{st.session_state['color_input_key']}")
            new_hl_hex = st.color_picker("Chọn Màu Highlight:", "#FFFF00", key=f"add_hc_picker_{st.session_state['color_input_key']}") if enable_hc else None
        with col_col4:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Quy Tắc Màu", use_container_width=True, type="primary", key="btn_add_fixed_color"):
                if new_color_spk:
                    tc_tuple = hex_to_rgb(new_text_hex) if enable_tc else None
                    hc_tuple = hex_to_rgb(new_hl_hex) if enable_hc else None
                    st.session_state['fixed_speaker_colors'][new_color_spk] = {"text_color": tc_tuple, "highlight_color": hc_tuple}
                    save_json_db(SPEAKER_COLOR_DB_FILE, st.session_state['fixed_speaker_colors'])
                    st.session_state['color_input_key'] += 1
                    st.success(f"✅ Đã lưu màu cho `{new_color_spk}`!"); time.sleep(1); st.rerun()
