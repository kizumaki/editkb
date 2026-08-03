import streamlit as st
import re
import pandas as pd
from collections import Counter
from utils import PRONOUN_REL_DB_FILE, save_json_db, ENGLISH_WORD_REGEX, parse_any_script_file_to_df

VN_SELF_PRONOUNS = ["tui", "tôi", "mình", "tao", "ta", "em", "anh", "chị", "cháu", "con", "tại hạ", "bản thân"]
VN_TARGET_PRONOUNS = ["ông", "bạn", "mày", "anh", "chị", "chú", "bác", "cậu", "bà", "cưng", "em", "ní", "mấy ní", "sư huynh", "huynh", "đệ"]

def render_tab7():
    st.subheader("🔎 SOÁT BẤT NHẤT THUẬT NGỮ & QUAN HỆ XƯNG HÔ NHÂN VẬT")
    st.markdown("Vùng làm việc phát hiện các câu thoại bị sượng xưng hô hoặc bất nhất bản dịch thuật ngữ/tên món ăn.")

    subtab_pronoun, subtab_glossary = st.tabs([
        "👥 Quản Lý & Soát Lỗi Xưng Hô Nhân Vật", 
        "📚 Soát Bất Nhất Thuật Ngữ & Món Ăn/Tên Riêng"
    ])

    with subtab_pronoun:
        st.markdown("#### 1. Bảng Thiết Lập Quan Hệ Xưng Hô (Lưu Database)")
        col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns([2, 2, 1.5, 1.5, 1.2])
        with col_p1: rel_spk_a = st.text_input("Người Nói (Speaker A):", placeholder="VD: TYLER", key=f"p_spk_a_{st.session_state.get('pronoun_input_key', 0)}").strip().upper()
        with col_p2: rel_spk_b = st.text_input("Người Nghe (Speaker B):", placeholder="VD: BILL", key=f"p_spk_b_{st.session_state.get('pronoun_input_key', 0)}").strip().upper()
        with col_p3: rel_self = st.text_input("Xưng (Self):", placeholder="VD: tui...", key=f"p_self_{st.session_state.get('pronoun_input_key', 0)}").strip().lower()
        with col_p4: rel_target = st.text_input("Gọi (Target):", placeholder="VD: ông...", key=f"p_target_{st.session_state.get('pronoun_input_key', 0)}").strip().lower()
        with col_p5:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ Thêm Xưng Hô", type="primary", use_container_width=True, key="btn_add_pronoun"):
                if rel_spk_a and rel_spk_b and rel_self and rel_target:
                    key_pair = f"{rel_spk_a}|{rel_spk_b}"
                    st.session_state['custom_pronoun_rel'][key_pair] = {"self": rel_self, "target": rel_target}
                    save_json_db(PRONOUN_REL_DB_FILE, st.session_state['custom_pronoun_rel'])
                    st.session_state['pronoun_input_key'] = st.session_state.get('pronoun_input_key', 0) + 1
                    st.success("✅ Đã lưu quan hệ xưng hô!")

        rel_db_dict = st.session_state.get('custom_pronoun_rel', {})
        if rel_db_dict:
            rel_data_table = []
            for pair_key, val in sorted(rel_db_dict.items()):
                spk_a, spk_b = pair_key.split("|") if "|" in pair_key else (pair_key, "ALL")
                rel_data_table.append({
                    "Người Nói": spk_a, "Người Nghe": spk_b,
                    "Xưng (Self)": val.get("self", "tui"), "Gọi (Target)": val.get("target", "ông"),
                    "Xóa": False
                })

            df_rel = pd.DataFrame(rel_data_table)
            edited_rel_df = st.data_editor(
                df_rel,
                column_config={
                    "Người Nói": st.column_config.TextColumn("Người Nói", disabled=True),
                    "Người Nghe": st.column_config.TextColumn("Người Nghe", disabled=True),
                    "Xưng (Self)": st.column_config.TextColumn("Đại từ Xưng (Self)"),
                    "Gọi (Target)": st.column_config.TextColumn("Đại từ Gọi (Target)"),
                    "Xóa": st.column_config.CheckboxColumn("Xóa?")
                },
                hide_index=True, use_container_width=True, key="pronoun_rel_editor_table"
            )

            if st.button("💾 LƯU BẢNG QUY TẮC XƯNG HÔ VÀO DATABASE", type="secondary", use_container_width=True):
                new_rel_db = {}
                for _, row in edited_rel_df.iterrows():
                    if not row["Xóa"]:
                        pk = f"{str(row['Người Nói']).upper()}|{str(row['Người Nghe']).upper()}"
                        new_rel_db[pk] = {"self": str(row["Xưng (Self)"]).lower(), "target": str(row["Gọi (Target)"]).lower()}
                st.session_state['custom_pronoun_rel'] = new_rel_db
                save_json_db(PRONOUN_REL_DB_FILE, new_rel_db)
                st.success("✅ Đã lưu cập nhật Bảng Xưng Hô!"); time.sleep(1); st.rerun()

        st.markdown("---")
        st.markdown("#### 2. Công Cụ Soát Lỗi Xưng Hô Tự Động Trong Kịch Bản")
        uploaded_pronoun_script = st.file_uploader("Tải file Kịch bản Tiếng Việt (.srt/.docx) để kiểm tra xưng hô:", type=['srt', 'docx'], key="uploader_pronoun_qc")

        if uploaded_pronoun_script is not None:
            c_spks_p = st.session_state.get('custom_speakers', set())
            c_non_spks_p = st.session_state.get('custom_non_speakers', set())
            df_p_script = parse_any_script_file_to_df(uploaded_pronoun_script.getvalue(), uploaded_pronoun_script.name, c_spks_p, c_non_spks_p)

            if not df_p_script.empty:
                pronoun_audit_rows = []
                speaker_pronoun_stats = {}

                for idx, r in df_p_script.iterrows():
                    spk = str(r['Speaker']).strip()
                    diag = str(r['Dialogue']).strip()
                    words = [w.lower() for w in re.findall(r'\b\w+\b', diag)]

                    found_self = [w for w in words if w in VN_SELF_PRONOUNS]
                    found_target = [w for w in words if w in VN_TARGET_PRONOUNS]

                    if spk not in speaker_pronoun_stats:
                        speaker_pronoun_stats[spk] = {"self": Counter(), "target": Counter()}

                    for s_w in found_self: speaker_pronoun_stats[spk]["self"][s_w] += 1
                    for t_w in found_target: speaker_pronoun_stats[spk]["target"][t_w] += 1

                for idx, r in df_p_script.iterrows():
                    spk = str(r['Speaker']).strip()
                    diag = str(r['Dialogue']).strip()
                    words = [w.lower() for w in re.findall(r'\b\w+\b', diag)]

                    found_self = [w for w in words if w in VN_SELF_PRONOUNS]
                    found_target = [w for w in words if w in VN_TARGET_PRONOUNS]

                    top_self = speaker_pronoun_stats[spk]["self"].most_common(1)[0][0] if speaker_pronoun_stats[spk]["self"] else ""
                    top_target = speaker_pronoun_stats[spk]["target"].most_common(1)[0][0] if speaker_pronoun_stats[spk]["target"] else ""

                    is_unusual = False
                    warn_msg = []

                    if found_self and top_self and any(w != top_self for w in found_self):
                        is_unusual = True
                        warn_msg.append(f"Xưng '{', '.join(found_self)}' (Đa số xưng '{top_self}')")

                    if found_target and top_target and any(w != top_target for w in found_target):
                        is_unusual = True
                        warn_msg.append(f"Gọi '{', '.join(found_target)}' (Đa số gọi '{top_target}')")

                    status_str = "🟡 Nghi vấn lệch xưng hô" if is_unusual else "🟢 Ok"
                    pronoun_audit_rows.append({
                        "Stt": idx + 1, "Timecode": f"{r['Start']} --> {r['End']}",
                        "Nhân vật": spk, "Câu thoại Tiếng Việt": diag,
                        "Trạng thái": status_str, "Chi tiết QC": "; ".join(warn_msg) if warn_msg else "Xưng hô khớp với tần suất chính"
                    })

                df_p_audit = pd.DataFrame(pronoun_audit_rows)
                unusual_cnt = sum(1 for st_v in df_p_audit['Trạng thái'] if '🟡' in str(st_v))

                col_pm1, col_pm2 = st.columns(2)
                with col_pm1:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">💬 Tổng số câu thoại đã quét</div><div class="metric-value">{len(df_p_audit)}</div></div>', unsafe_allow_html=True)
                with col_pm2:
                    st.markdown(f'<div class="metric-card"><div class="metric-label">🟡 Câu nghi vấn lệch xưng hô</div><div class="metric-value" style="color:#D97706;">{unusual_cnt}</div></div>', unsafe_allow_html=True)

                st.markdown("##### 👁️ Bảng Báo Cáo Soát Lỗi Xưng Hô Chi Tiết")
                p_filter = st.checkbox("🟡 Chỉ hiển thị các câu nghi vấn lệch xưng hô", value=True)
                df_p_disp = df_p_audit[df_p_audit['Trạng thái'].str.contains('🟡', na=False)] if p_filter else df_p_audit
                st.dataframe(df_p_disp[['Stt', 'Timecode', 'Nhân vật', 'Câu thoại Tiếng Việt', 'Trạng thái', 'Chi tiết QC']], hide_index=True, use_container_width=True)

    with subtab_glossary:
        st.markdown("#### 📚 Soát Bất Nhất Thuật Ngữ & Bản Dịch Tiếng Việt")
        uploaded_glossary_script = st.file_uploader("Tải file Kịch bản (.srt/.docx) để kiểm tra thuật ngữ:", type=['srt', 'docx'], key="uploader_glossary_qc")

        if uploaded_glossary_script is not None:
            c_spks_g = st.session_state.get('custom_speakers', set())
            c_non_spks_g = st.session_state.get('custom_non_speakers', set())
            df_g_script = parse_any_script_file_to_df(uploaded_glossary_script.getvalue(), uploaded_glossary_script.name, c_spks_g, c_non_spks_g)

            if not df_g_script.empty:
                all_dialogues_str = " ".join(df_g_script['Dialogue'].dropna().tolist())
                eng_matches = ENGLISH_WORD_REGEX.findall(all_dialogues_str)
                eng_counts = Counter([w for w in eng_matches if is_candidate_english_word(w)])

                glossary_audit_rows = []
                for word, cnt in eng_counts.most_common(20):
                    matching_lines = df_g_script[df_g_script['Dialogue'].str.contains(r'\b' + re.escape(word) + r'\b', case=False, na=False)]
                    sample_texts = matching_lines['Dialogue'].head(3).tolist()

                    glossary_audit_rows.append({
                        "Thuật ngữ / Tên riêng": word,
                        "Số lần lặp lại": cnt,
                        "Các câu thoại chứa từ này": " | ".join(sample_texts)
                    })

                if glossary_audit_rows:
                    df_g_audit = pd.DataFrame(glossary_audit_rows)
                    st.markdown("##### 📑 Bảng Thống Kê Thuật Ngữ Lặp Lại Trong Kịch Bản")
                    st.dataframe(df_g_audit, hide_index=True, use_container_width=True)
