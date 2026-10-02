import streamlit as st
import re
import time
import pandas as pd
from collections import Counter
import io
from utils import kept_file_uploader, PRONOUN_REL_DB_FILE, save_json_db, ENGLISH_WORD_REGEX, parse_any_script_file_to_df, is_candidate_english_word, clean_cell, log_event
from learn_core import merge_pronoun_rels, norm_pair_key

from pronoun_qc import (analyze as analyze_pronouns, SELF_TERMS as PRONOUN_SELF_TERMS, TARGET_TERMS as PRONOUN_TARGET_TERMS,
                        THIRD_TERMS as PRONOUN_THIRD_TERMS, KINSHIP_TERMS as PRONOUN_KIN_TERMS)

KEEP, USE_NEW = "Giữ cũ", "Dùng mới"
_KNOWN_TERMS = PRONOUN_SELF_TERMS + PRONOUN_TARGET_TERMS + PRONOUN_THIRD_TERMS + PRONOUN_KIN_TERMS


def _rel_template():
    buf = io.BytesIO()
    pd.DataFrame({"Người nói": ["TYLER"], "Người nghe": ["BILL"], "Xưng": ["tui"], "Gọi": ["ông"]}).to_excel(buf, index=False)
    return buf.getvalue()


def _apply_rels(items, source):
    """Mọi lần thêm quan hệ xưng hô đi qua đây: mới -> thêm; trùng -> bỏ qua; khác -> hỏi; từ lạ -> cảnh báo."""
    kho = st.session_state['custom_pronoun_rel']
    res = merge_pronoun_rels(kho, items, _KNOWN_TERMS)
    if res["added"] and save_json_db(PRONOUN_REL_DB_FILE, kho):
        log_event("Thêm quan hệ xưng hô", f"{source}: " + ", ".join(k for k, _ in res["added"])[:400])
    st.session_state["_rel_result"] = res
    st.session_state['pronoun_input_key'] = st.session_state.get('pronoun_input_key', 0) + 1
    st.rerun()


def _render_rel_result():
    res = st.session_state.get("_rel_result")
    if not res: return
    if res.get("msg"): st.success(res["msg"])
    if res["added"]: st.success(f"✅ Đã thêm {len(res['added'])} cặp: " + ", ".join(k.replace("|", " → ") for k, _ in res["added"][:8]))
    if res["same"]: st.info(f"ℹ️ {len(res['same'])} cặp đã có sẵn, giống hệt — bỏ qua.")
    if res["odd"]: st.warning("⚠️ Từ xưng/gọi **lạ** (không có trong bộ từ xưng hô, có thể gõ sai): " +
                              "; ".join(f"«{w}» ({k.replace('|', ' → ')})" for k, w in res["odd"][:8]))
    if res["conflicts"]:
        st.warning(f"⚠️ {len(res['conflicts'])} cặp ĐÃ CÓ nhưng cách xưng hô KHÁC. App **chưa ghi đè** — chọn cho từng cặp:")
        df = pd.DataFrame([{"Cặp": k.replace("|", " → "), "Kho đang có": f"{o.get('self')} / {o.get('target')}",
                            "Mới nhập": f"{n['self']} / {n['target']}", "Chọn": KEEP} for k, o, n in res["conflicts"]])
        ed = st.data_editor(df, hide_index=True, use_container_width=True, key=f"rel_conf_{st.session_state.get('pronoun_input_key', 0)}",
                            column_config={c: st.column_config.TextColumn(disabled=True) for c in ("Cặp", "Kho đang có", "Mới nhập")} |
                                          {"Chọn": st.column_config.SelectboxColumn(options=[KEEP, USE_NEW], required=True)})
        if st.button("✔️ Áp dụng lựa chọn", type="primary", key="btn_rel_conf"):
            kho = st.session_state['custom_pronoun_rel']
            changed = [(k, o, n) for (k, o, n), (_, r) in zip(res["conflicts"], ed.iterrows()) if r["Chọn"] == USE_NEW]
            for k, _, n in changed: kho[k] = n
            if changed and save_json_db(PRONOUN_REL_DB_FILE, kho):
                log_event("Sửa quan hệ xưng hô (chọn dùng mới)", "; ".join(f"{k}: {o} → {n}" for k, o, n in changed)[:500])
            st.session_state["_rel_result"] = {"added": [], "same": [], "odd": [], "conflicts": [],
                                               "msg": f"✅ Đã cập nhật {len(changed)} cặp, giữ nguyên {len(res['conflicts']) - len(changed)} cặp."}
            st.rerun()
    if st.button("Đóng thông báo", key="btn_rel_close"): st.session_state.pop("_rel_result", None); st.rerun()


def render_tab7():
    st.subheader("Soát xưng hô & thuật ngữ")
    st.markdown("Vùng làm việc phát hiện các câu thoại bị sượng xưng hô hoặc bất nhất bản dịch thuật ngữ/tên món ăn.")

    subtab_pronoun, subtab_glossary = st.tabs([
        "👥 Xưng hô nhân vật", 
        "📚 Thuật ngữ & tên riêng"
    ])

    with subtab_pronoun:
        st.markdown("#### 1. Bảng Thiết Lập Quan Hệ Xưng Hô (Lưu Database)")
        _render_rel_result()
        ver = st.session_state.get('pronoun_input_key', 0)
        t_hand, t_file = st.tabs(["✍️ Gõ tay", "📄 Tải file"])
        with t_hand:
            col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns([2, 2, 1.5, 1.5, 1.2])
            with col_p1: rel_spk_a = st.text_input("Người Nói (Speaker A):", placeholder="VD: TYLER", key=f"p_spk_a_{ver}")
            with col_p2: rel_spk_b = st.text_input("Người Nghe (Speaker B):", placeholder="VD: BILL", key=f"p_spk_b_{ver}")
            with col_p3: rel_self = st.text_input("Xưng (Self):", placeholder="VD: tui...", key=f"p_self_{ver}")
            with col_p4: rel_target = st.text_input("Gọi (Target):", placeholder="VD: ông...", key=f"p_target_{ver}")
            with col_p5:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                if st.button("➕ Thêm Xưng Hô", type="primary", use_container_width=True, key="btn_add_pronoun"):
                    if rel_spk_a.strip() and rel_spk_b.strip() and (rel_self.strip() or rel_target.strip()):
                        _apply_rels([(rel_spk_a, rel_spk_b, rel_self, rel_target)], "gõ tay")
                    else: st.warning("Điền Người nói, Người nghe và ít nhất 1 trong 2 ô Xưng/Gọi.")
            k_now = norm_pair_key(rel_spk_a, rel_spk_b)
            cur = st.session_state['custom_pronoun_rel'].get(k_now)
            if isinstance(cur, dict): st.caption(f"ℹ️ Kho đang có cho cặp này: xưng **{cur.get('self')}**, gọi **{cur.get('target')}**")
        with t_file:
            st.caption("Excel/CSV 4 cột: **Người nói, Người nghe, Xưng, Gọi**.")
            st.download_button("⬇️ Tải file mẫu (Excel)", data=_rel_template(), file_name="Mau_Xung_Ho.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_rel_tpl")
            up = st.file_uploader("Chọn file", type=["xlsx", "xls", "csv"], key=f"rel_upload_{ver}")
            if up and st.button("📥 Nạp file", type="primary", key="btn_rel_upload"):
                try:
                    df_up = (pd.read_csv(up, dtype=str) if up.name.lower().endswith(".csv") else pd.read_excel(up, dtype=str)).dropna(how="all")
                    if df_up.shape[1] < 4: st.error("File cần đủ 4 cột: Người nói, Người nghe, Xưng, Gọi.")
                    else: _apply_rels([tuple(clean_cell(x) for x in r[:4]) for r in df_up.itertuples(index=False)], f"file {up.name}")
                except Exception as e: st.error(f"Không đọc được file: {e}")

        rel_db_dict = st.session_state.get('custom_pronoun_rel', {})
        if rel_db_dict:
            rel_data_table = []
            for pair_key, val in sorted(rel_db_dict.items()):
                spk_a, _, spk_b = pair_key.partition("|")
                spk_b = spk_b or "ALL"
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
