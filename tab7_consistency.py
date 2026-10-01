import streamlit as st
import re
import time
import pandas as pd
from collections import Counter
from utils import kept_file_uploader, PRONOUN_REL_DB_FILE, save_json_db, ENGLISH_WORD_REGEX, parse_any_script_file_to_df, is_candidate_english_word, clean_cell

from pronoun_qc import (analyze as analyze_pronouns, SELF_TERMS as PRONOUN_SELF_TERMS, TARGET_TERMS as PRONOUN_TARGET_TERMS,
                        THIRD_TERMS as PRONOUN_THIRD_TERMS, KINSHIP_TERMS as PRONOUN_KIN_TERMS)

def render_tab7():
    st.subheader("Soát xưng hô & thuật ngữ")
    st.markdown("Vùng làm việc phát hiện các câu thoại bị sượng xưng hô hoặc bất nhất bản dịch thuật ngữ/tên món ăn.")

    subtab_pronoun, subtab_glossary = st.tabs([
        "👥 Xưng hô nhân vật", 
        "📚 Thuật ngữ & tên riêng"
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
                    st.success("✅ Đã lưu quan hệ xưng hô!"); time.sleep(1); st.rerun()

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

            if st.button("💾 Lưu bảng xưng hô", type="secondary", use_container_width=True):
                new_rel_db = {}
                for _, row in edited_rel_df.iterrows():
                    if not row["Xóa"]:
                        pk = f"{str(row['Người Nói']).upper()}|{str(row['Người Nghe']).upper()}"
                        new_rel_db[pk] = {"self": clean_cell(row["Xưng (Self)"]).lower(), "target": clean_cell(row["Gọi (Target)"]).lower()}
                st.session_state['custom_pronoun_rel'] = new_rel_db
                save_json_db(PRONOUN_REL_DB_FILE, new_rel_db)
                st.success("✅ Đã lưu cập nhật Bảng Xưng Hô!"); time.sleep(1); st.rerun()

        st.markdown("---")
        st.markdown("#### 2. Công Cụ Soát Lỗi Xưng Hô Tự Động Trong Kịch Bản")
        uploaded_pronoun_script = kept_file_uploader("Tải file Kịch bản Tiếng Việt (.srt/.docx) để kiểm tra xưng hô:", type=['srt', 'docx'], key="uploader_pronoun_qc")

        if uploaded_pronoun_script is not None:
            c_spks_p = st.session_state.get('custom_speakers', set())
            c_non_spks_p = st.session_state.get('custom_non_speakers', set())
            df_p_script = parse_any_script_file_to_df(uploaded_pronoun_script.getvalue(), uploaded_pronoun_script.name, c_spks_p, c_non_spks_p)

            if not df_p_script.empty:
                lines = [{"stt": i + 1, "timecode": f"{r['Start']} --> {r['End']}",
                          "speaker": str(r['Speaker']).strip(), "text": str(r['Dialogue']).strip()}
                         for i, (_, r) in enumerate(df_p_script.iterrows())]
                results, summary = analyze_pronouns(lines, st.session_state.get('custom_pronoun_rel', {}))
                n_red = sum(1 for x in results if x["status"].startswith("🔴"))
                n_yellow = sum(1 for x in results if x["status"].startswith("🟡"))

                col_pm1, col_pm2, col_pm3 = st.columns(3)
                col_pm1.markdown(f'<div class="metric-card"><div class="metric-label">💬 Câu thoại đã quét</div><div class="metric-value">{len(results)}</div></div>', unsafe_allow_html=True)
                col_pm2.markdown(f'<div class="metric-card"><div class="metric-label">🔴 Trái quy tắc đã lưu</div><div class="metric-value" style="color:#DC2626;">{n_red}</div></div>', unsafe_allow_html=True)
                col_pm3.markdown(f'<div class="metric-card"><div class="metric-label">🟡 Lệch thói quen nhân vật</div><div class="metric-value" style="color:#D97706;">{n_yellow}</div></div>', unsafe_allow_html=True)

                st.markdown("##### 👥 Mỗi nhân vật đang xưng / gọi thế nào")
                st.caption("Số trong ngoặc = số lần dùng. Từ thân tộc (anh/chị/em/ông/bà/con...) vừa để xưng vừa để gọi, "
                           "nên app chỉ thống kê, không tự báo lỗi — người soát nhìn bảng này để phát hiện chỗ bất thường.")
                st.dataframe(pd.DataFrame(summary), hide_index=True, use_container_width=True)

                st.markdown("##### 👁️ Chi tiết từng câu")
                p_filter = st.checkbox("Chỉ hiện câu có vấn đề (🔴 / 🟡)", value=True, key="pronoun_only_issues")
                rows = [{"Stt": x["stt"], "Timecode": x["timecode"], "Nhân vật": x["speaker"], "Câu thoại": x["text"],
                         "Xưng hô tìm thấy": x["found"], "Trạng thái": x["status"], "Chi tiết": x["detail"]}
                        for x in results if not p_filter or not x["status"].startswith("🟢")]
                if rows: st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
                else: st.success("Không phát hiện câu nào lệch xưng hô.")
                with st.expander("ℹ️ App nhận diện những cách xưng hô nào?"):
                    st.markdown(f"**Xưng:** {', '.join(PRONOUN_SELF_TERMS)}\n\n**Gọi:** {', '.join(PRONOUN_TARGET_TERMS)}\n\n"
                                f"**Ngôi thứ ba:** {', '.join(PRONOUN_THIRD_TERMS)}\n\n**Thân tộc (chỉ thống kê):** {', '.join(PRONOUN_KIN_TERMS)}\n\n"
                                "Thiếu từ nào thường gặp, báo người quản lý app để bổ sung.")

    with subtab_glossary:
        st.markdown("#### 📚 Soát Bất Nhất Thuật Ngữ & Bản Dịch Tiếng Việt")
        uploaded_glossary_script = kept_file_uploader("Tải file Kịch bản (.srt/.docx) để kiểm tra thuật ngữ:", type=['srt', 'docx'], key="uploader_glossary_qc")

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
