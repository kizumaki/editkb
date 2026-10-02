import streamlit as st
import io
import re
import pandas as pd
from utils import decode_text, PHONETIC_DB_FILE, save_json_db, generate_english_audio, clean_cell, log_event
import learn_core as L

KEEP, USE_NEW = "Giữ cũ", "Dùng mới"


def _parse_upload(f):
    """File Excel/CSV (cột 1 = tiếng Anh, cột 2 = phiên âm) hoặc .txt (mỗi dòng "Phiên âm (English)" hoặc "English = phiên âm")."""
    name = f.name.lower()
    if name.endswith((".xlsx", ".xls", ".csv")):
        df = pd.read_csv(io.BytesIO(f.getvalue()), dtype=str) if name.endswith(".csv") else pd.read_excel(io.BytesIO(f.getvalue()), dtype=str)
        df = df.dropna(how="all")
        if df.shape[1] < 2: return [], ["File cần ít nhất 2 cột: Tiếng Anh, Phiên âm."]
        return [(clean_cell(a), clean_cell(b)) for a, b in zip(df.iloc[:, 0], df.iloc[:, 1]) if clean_cell(a) and clean_cell(b)], []
    items, bad = [], []
    for line in decode_text(f.getvalue()).splitlines():
        t = line.strip()
        if not t: continue
        m = re.match(r"^(.+?)\s*\(([^()]+)\)\s*$", t)              # Phiên âm (English)
        if m: items.append((m.group(2).strip(), m.group(1).strip())); continue
        parts = re.split(r"\s*(?:=|\t|;|\||->|→)\s*", t, maxsplit=1)  # English = phiên âm
        if len(parts) == 2 and parts[0] and parts[1]: items.append((parts[0], parts[1]))
        else: bad.append(t)
    return items, bad


def _template():
    buf = io.BytesIO()
    pd.DataFrame({"Tiếng Anh": ["Minecraft", "block"], "Phiên âm": ["Mai-ráp", "lốc"]}).to_excel(buf, index=False)
    return buf.getvalue()


def _apply(items, proper_mode, source):
    """proper_mode: True/False = viết hoa tên riêng / viết thường; None = giữ như người nhập gõ."""
    kho = st.session_state["custom_phonetics"]
    items = [(e, L.normalize_pho(p, proper_mode)) for e, p in items if e and p]
    res = L.merge_phonetics(kho, items)
    if res["added"]:
        if save_json_db(PHONETIC_DB_FILE, kho): log_event("Thêm phiên âm", f"{source}: {len(res['added'])} từ")
    st.session_state["_pho_result"] = res
    st.session_state["pho_input_key"] = st.session_state.get("pho_input_key", 0) + 1
    st.rerun()


def _show_result():
    res = st.session_state.get("_pho_result")
    if not res: return
    if res.get("msg"): st.success(res["msg"])
    if res["added"]: st.success(f"✅ Đã thêm **{len(res['added'])}** từ mới: " + ", ".join(f"{p} ({e})" for e, p in res["added"][:8]) + (" ..." if len(res["added"]) > 8 else ""))
    if res["same"]: st.info(f"ℹ️ **{len(res['same'])}** từ đã có sẵn, giống hệt — bỏ qua.")
    if res["near"]:
        st.warning("⚠️ Có từ **gần giống** từ đã có, kiểm tra xem có bị trùng không: " +
                   "; ".join(f"«{a}» gần giống «{b}» ({p})" for a, b, p in res["near"][:6]))
    if res["conflicts"]:
        st.warning(f"⚠️ **{len(res['conflicts'])}** từ ĐÃ CÓ trong kho nhưng phiên âm KHÁC. App **chưa ghi đè** — hãy chọn cho từng từ:")
        df = pd.DataFrame([{"Tiếng Anh": e, "Kho đang có": old, "Mới nhập": new, "Chọn": KEEP} for e, old, new in res["conflicts"]])
        ed = st.data_editor(df, hide_index=True, use_container_width=True, key=f"pho_conf_{st.session_state.get('pho_input_key', 0)}",
                            column_config={"Tiếng Anh": st.column_config.TextColumn(disabled=True),
                                           "Kho đang có": st.column_config.TextColumn(disabled=True),
                                           "Mới nhập": st.column_config.TextColumn(disabled=True),
                                           "Chọn": st.column_config.SelectboxColumn(options=[KEEP, USE_NEW], required=True)})
        if st.button("✔️ Áp dụng lựa chọn", type="primary", key="btn_pho_conf"):
            kho = st.session_state["custom_phonetics"]
            changed = [(r["Tiếng Anh"], r["Kho đang có"], r["Mới nhập"]) for _, r in ed.iterrows() if r["Chọn"] == USE_NEW]
            for e, _, new in changed: kho[e] = new
            if changed and save_json_db(PHONETIC_DB_FILE, kho):
                log_event("Sửa phiên âm (chọn dùng mới)", "; ".join(f"{e}: {o} → {n}" for e, o, n in changed)[:500])
            st.session_state["_pho_result"] = {"added": [], "same": [], "near": [], "conflicts": [],
                                               "msg": f"✅ Đã cập nhật {len(changed)} từ theo bản mới, giữ nguyên {len(ed) - len(changed)} từ."}
            st.rerun()
    if st.button("Đóng thông báo", key="btn_pho_close"): st.session_state.pop("_pho_result", None); st.rerun()


def _guess_proper(eng, cur, caps, names, eng_case=None):
    """Đoán tên riêng cho từ ĐÃ CÓ trong kho. Trả về (True/False, lý do, chắc chắn?)."""
    if eng in names: return True, "Là tên nhân vật trong danh sách", True
    words = eng.split()
    if caps and words:
        c = caps.get(words[0], [0, 0])
        if sum(c) >= 3:
            r = c[0] / sum(c)
            if r >= 0.7: return True, f"Phụ đề gốc viết hoa {c[0]}/{sum(c)} lần", True
            if r <= 0.3: return False, f"Phụ đề gốc viết thường {c[1]}/{sum(c)} lần", True
    up, low = (eng_case or {}).get(eng, [0, 0])
    if up + low >= 2 and up / (up + low) >= 0.8: return True, f"Kịch bản team viết hoa chữ tiếng Anh {up}/{up + low} lần", True
    if up + low >= 2 and low / (up + low) >= 0.8: return False, f"Kịch bản team viết thường chữ tiếng Anh {low}/{up + low} lần", True
    return cur[:1].isupper(), "Ít gặp hoặc lúc hoa lúc thường — giữ như đang viết, bạn tự chọn", False


def _render_case_tool():
    with st.expander("🔠 Chuẩn hoá chữ hoa/thường cho toàn kho (theo quy chuẩn: chỉ viết hoa tên riêng)", expanded=False):
        kho = st.session_state["custom_phonetics"]
        R = st.session_state.get("learn_result")
        caps = (R or {}).get("caps") or {}
        names = {L.norm_eng_key(s) for s in st.session_state.get("custom_speakers", set())} | set(st.session_state.get("custom_cast_mapping", {}))
        if not caps:
            st.info("💡 Để máy đoán **chính xác hơn**: vào trang **Học từ kịch bản cũ**, nạp file .zip và bấm **Phân tích** "
                    "(không cần lưu gì), rồi quay lại đây. Khi đó máy xem thêm từ nào được viết hoa giữa câu trong phụ đề gốc tiếng Anh.")
        rows = []
        for eng, cur in sorted(kho.items()):
            if L.pho_missing(eng, cur): continue
            proper, why, sure = _guess_proper(eng, cur, caps, names, (R or {}).get("pho_eng_case"))
            after = L.normalize_pho(cur, proper)
            rows.append({"Áp dụng": sure and after != cur, "Tiếng Anh": eng, "Đang viết": cur, "Tên riêng": proper,
                         "Sau chuẩn hoá": after, "Máy đoán dựa vào": why})
        if not rows: st.caption("Kho chưa có từ nào."); return
        only_change = st.toggle("Chỉ xem những từ sẽ thay đổi", value=True, key="case_only_change")
        df = pd.DataFrame(rows)
        view = df[df["Sau chuẩn hoá"] != df["Đang viết"]] if only_change else df
        st.caption(f"**{len(view)}** từ. Máy đã tích sẵn những từ đoán **chắc chắn**. Sửa cột **Tên riêng?** nếu máy đoán sai "
                   "(cột \"Sau chuẩn hoá\" cập nhật sau khi bấm áp dụng). Tắt nút gạt phía trên để xem cả những từ đang viết thường mà thực ra là tên riêng (VD: en-ni → En-ni).")
        ed = st.data_editor(view, hide_index=True, use_container_width=True, height=min(520, 38 + 35 * max(len(view), 1)),
                            key=f"case_tool_{only_change}_{st.session_state.get('pho_input_key', 0)}",
                            column_config={"Áp dụng": st.column_config.CheckboxColumn("Áp dụng?", width="small"),
                                           "Tiếng Anh": st.column_config.TextColumn(disabled=True),
                                           "Đang viết": st.column_config.TextColumn(disabled=True),
                                           "Tên riêng": st.column_config.CheckboxColumn("Tên riêng?", help="Tích = viết hoa chữ đầu mỗi từ; bỏ tích = viết thường"),
                                           "Sau chuẩn hoá": st.column_config.TextColumn(disabled=True),
                                           "Máy đoán dựa vào": st.column_config.TextColumn(disabled=True)})
        sel = ed[ed["Áp dụng"] | (ed["Tên riêng"] != view["Tên riêng"])]  # đổi ô "Tên riêng?" là tự áp dụng dòng đó
        if st.button(f"✔️ Áp dụng cho {len(sel)} từ", type="primary", key="btn_case_apply", disabled=sel.empty):
            changed = []
            for _, r in sel.iterrows():
                new = L.normalize_pho(r["Đang viết"], bool(r["Tên riêng"]))
                if new != kho.get(r["Tiếng Anh"]): changed.append((r["Tiếng Anh"], kho.get(r["Tiếng Anh"]), new)); kho[r["Tiếng Anh"]] = new
            if changed and save_json_db(PHONETIC_DB_FILE, kho):
                log_event("Chuẩn hoá hoa/thường kho phiên âm", f"{len(changed)} từ: " + "; ".join(f"{e}: {o}→{n}" for e, o, n in changed)[:450])
                st.session_state["_pho_table_msg"] = ("success", f"✅ Đã chuẩn hoá {len(changed)} từ.")
            elif not changed: st.session_state["_pho_table_msg"] = ("info", "Không có từ nào cần đổi.")
            st.session_state["pho_input_key"] = st.session_state.get("pho_input_key", 0) + 1
            st.rerun()


def render_tab5():
    st.subheader("Kho phiên âm giọng Nam")
    st.markdown("Kho dùng chung của cả team (lưu trên Google Sheets). Quy chuẩn: **Phiên âm (English)**, "
                "chỉ **viết hoa tên riêng** (tên người, địa danh, thương hiệu, tên game): *Chim-mi (Jimmy)*, *Mai-ráp (Minecraft)*; "
                "từ thường viết thường: *lốc (block)*. Kho là **chuẩn duy nhất** — Re-Sync sẽ sửa phiên âm trong kịch bản theo kho.")

    with st.container(border=True):
        st.markdown("#### 🔊 Nghe phát âm thử")
        col_test1, col_test2, col_test3 = st.columns([2.5, 1.5, 1.5])
        with col_test1: free_test_word = st.text_input("Từ/cụm tiếng Anh cần nghe thử:", placeholder="VD: Starbucks, McDonald's...", key="free_audio_text")
        with col_test2: free_accent = st.radio("Giọng phát âm:", options=["Giọng Mỹ (US)", "Giọng Anh (UK)"], horizontal=True, key="free_audio_accent")
        with col_test3:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            free_listen_btn = st.button("🔊 Phát âm thanh", use_container_width=True, key="btn_free_listen")
        if free_listen_btn and free_test_word:
            test_fp = generate_english_audio(free_test_word, accent='com' if "Mỹ" in free_accent else 'co.uk')
            if test_fp: st.audio(test_fp, format="audio/mp3", autoplay=True)

    with st.container(border=True):
        st.markdown("#### ➕ Thêm từ mới")
        _show_result()
        ver = st.session_state.get("pho_input_key", 0)
        t_hand, t_file = st.tabs(["✍️ Gõ tay", "📄 Tải file"])
        with t_hand:
            c1, c2, c3 = st.columns([2, 2, 1.6])
            eng = c1.text_input("Từ tiếng Anh gốc:", placeholder="VD: Jimmy", key=f"tab_add_eng_{ver}")
            pho = c2.text_input("Phiên âm giọng Nam:", placeholder="VD: Chim-mi", key=f"tab_add_pho_{ver}")
            kind = c3.radio("Loại từ:", ["Từ thường", "Tên riêng"], horizontal=True, key=f"tab_add_kind_{ver}",
                            help="Tên riêng (người, địa danh, thương hiệu, game) → viết hoa chữ đầu. Từ thường → viết thường.")
            k = L.norm_eng_key(eng)
            kho = st.session_state["custom_phonetics"]
            if k and k in kho and not L.pho_missing(k, kho[k]): st.caption(f"ℹ️ Kho đang có: **{kho[k]} ({eng.strip()})**")
            elif k and L.near_keys(k, kho): st.caption("⚠️ Gần giống từ đã có: " + ", ".join(f"{n} → {kho[n]}" for n in L.near_keys(k, kho)))
            if st.button("➕ Thêm vào kho", type="primary", key="btn_add_pho"):
                if eng.strip() and pho.strip(): _apply([(eng, pho)], kind == "Tên riêng", "gõ tay")
                else: st.warning("Vui lòng điền đủ 2 ô!")
        with t_file:
            st.caption("Excel/CSV: **cột 1 = tiếng Anh, cột 2 = phiên âm**. Hoặc file .txt, mỗi dòng `Chim-mi (Jimmy)` hoặc `Jimmy = Chim-mi`. "
                       "Viết hoa/thường được **giữ đúng như trong file** (chỉ sửa lỗi gõ).")
            st.download_button("⬇️ Tải file mẫu (Excel)", data=_template(), file_name="Mau_Phien_Am.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_pho_tpl")
            up = st.file_uploader("Chọn file", type=["xlsx", "xls", "csv", "txt"], key=f"pho_upload_{ver}")
            if up and st.button("📥 Nạp file vào kho", type="primary", key="btn_pho_upload"):
                try:
                    items, bad = _parse_upload(up)
                except Exception as e:
                    items, bad = [], [f"Không đọc được file: {e}"]
                if bad: st.warning(f"{len(bad)} dòng không hiểu được, đã bỏ qua: " + "; ".join(bad[:5]))
                if items: _apply(items, None, f"file {up.name}")
                elif not bad: st.warning("File không có dòng nào hợp lệ.")

    _render_case_tool()

    with st.container(border=True):
        st.markdown("#### 📑 Toàn bộ kho phiên âm")
        if st.session_state.get("_pho_table_msg"):
            kind, msg = st.session_state.pop("_pho_table_msg"); getattr(st, kind)(msg)
        all_pho = st.session_state.get('custom_phonetics', {})
        missing = {k for k, v in all_pho.items() if L.pho_missing(k, v)}
        c1, c2 = st.columns([3, 1.4])
        q = c1.text_input("🔍 Tìm từ tiếng Anh hoặc phiên âm:", placeholder="Gõ từ cần tìm...").strip().upper()
        only_missing = c2.toggle(f"Chỉ xem từ chưa có phiên âm ({len(missing)})", key="pho_only_missing",
                                 help="Từ được thêm bằng nút quét ở thanh bên nhưng chưa ai điền phiên âm — sẽ KHÔNG được chèn vào kịch bản.")
        shown = {k: v for k, v in all_pho.items() if (not q or q in k or q in v.upper()) and (not only_missing or k in missing)}
        if not shown:
            st.info("Không có từ nào khớp."); return
        df = pd.DataFrame([{"Tiếng Anh": k, "Phiên âm": "" if k in missing else v, "Xóa": False} for k, v in sorted(shown.items())])
        st.caption(f"Đang hiển thị **{len(df)}** / {len(all_pho)} từ. Sửa trực tiếp trong bảng rồi bấm Lưu.")
        ed = st.data_editor(df, hide_index=True, use_container_width=True, key=f"global_phonetic_db_editor_{q}_{only_missing}_{ver}",
                            column_config={"Tiếng Anh": st.column_config.TextColumn("Tiếng Anh gốc", disabled=True),
                                           "Phiên âm": st.column_config.TextColumn("Phiên âm (sửa trực tiếp)"),
                                           "Xóa": st.column_config.CheckboxColumn("Xóa?", width="small")})
        if st.button("💾 Lưu thay đổi trong bảng", type="primary", use_container_width=True, key="btn_pho_table_save"):
            changes, deleted = [], []
            for _, r in ed.iterrows():
                k, v = r["Tiếng Anh"], clean_cell(r["Phiên âm"])
                if r["Xóa"]: all_pho.pop(k, None); deleted.append(k)
                elif v and v != all_pho.get(k): changes.append(f"{k}: {all_pho.get(k)} → {v}"); all_pho[k] = v
            typed = {r["Tiếng Anh"]: clean_cell(r["Phiên âm"]) for _, r in ed.iterrows() if not r["Xóa"] and clean_cell(r["Phiên âm"])}
            if not (changes or deleted):
                st.session_state["_pho_table_msg"] = ("info", "Không có gì thay đổi.")
            elif save_json_db(PHONETIC_DB_FILE, all_pho):
                log_event("Sửa kho phiên âm", (f"Sửa {len(changes)}: " + "; ".join(changes) + f" | Xoá: {', '.join(deleted)}")[:500])
                fixed = [f"«{t}» → «{all_pho.get(L.norm_eng_key(k))}»" for k, t in typed.items()
                         if all_pho.get(L.norm_eng_key(k)) not in (None, t) and k in {c.split(':')[0] for c in changes}]
                msg = f"✅ Đã lưu lên kho: sửa {len(changes)} từ, xoá {len(deleted)} từ."
                if fixed: msg += " Đã tự chuẩn hoá theo quy chuẩn (chỉ hoa chữ đầu, bỏ dấu cách thừa): " + "; ".join(fixed[:8])
                st.session_state["_pho_table_msg"] = ("success", msg)
            else:
                st.session_state["_pho_table_msg"] = ("error", "❌ CHƯA lưu được lên Google Sheets — thay đổi chưa được ghi. Thử bấm lưu lại sau ít phút.")
            st.session_state["pho_input_key"] = ver + 1
            st.rerun()
