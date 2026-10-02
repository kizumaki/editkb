"""Trang "Học từ kịch bản cũ": nạp kịch bản team đã biên tập -> gợi ý cho các kho. Người duyệt tích chọn rồi mới lưu."""
import io
import os
import re
import zipfile
import streamlit as st
import pandas as pd
from docx import Document
from utils import (
    kept_file_uploader, save_json_db, log_event, clean_cell, VN_SYLLABLES,
    PHONETIC_DB_FILE, SPEAKER_DB_FILE, NON_SPEAKER_DB_FILE, CAST_DB_FILE, PRONOUN_REL_DB_FILE
)
import learn_core as L
import pronoun_qc


class _ZipEntry:
    def __init__(self, name, data): self.name, self._d = name, data
    def getvalue(self): return self._d

def _expand_zips(files):
    """File .zip (VD: cả thư mục kịch bản mẫu nén lại) -> tách thành từng file .docx/.srt bên trong."""
    out = []
    for f in files:
        if not f.name.lower().endswith(".zip"): out.append(f); continue
        with zipfile.ZipFile(io.BytesIO(f.getvalue())) as z:
            for info in z.infolist():
                n = info.filename
                if info.is_dir() or "__MACOSX" in n or os.path.basename(n).startswith(("~$", ".")): continue
                if n.lower().endswith((".docx", ".srt")): out.append(_ZipEntry(os.path.basename(n), z.read(info)))
    return out

def _read_files(files):
    docx, srt, skipped = [], [], []
    files = _expand_zips(files)
    bar = st.progress(0.0, text="Đang đọc file...")
    for i, f in enumerate(files, 1):
        bar.progress(i / len(files), text=f"Đang đọc {i}/{len(files)}: {f.name}")
        name = f.name.lower()
        try:
            if name.endswith(".docx"):
                docx.append((f.name, [p.text for p in Document(io.BytesIO(f.getvalue())).paragraphs]))
            elif name.endswith(".srt"):
                text = f.getvalue().decode("utf-8-sig", errors="ignore")
                body = " ".join(l for l in text.splitlines() if l.strip() and not L.TIMECODE.match(l) and not l.strip().isdigit())
                if len(L.VN_MARK.findall(body)) > 0.02 * max(len(body), 1): skipped.append(f"{f.name} (phụ đề tiếng Việt — chỉ dùng phụ đề gốc tiếng Anh)")
                else: srt.append((f.name, text))
        except Exception as e:
            skipped.append(f"{f.name} (không đọc được: {e})"[:150])
    bar.empty()
    return docx, srt, skipped


def render_tab10():
    st.subheader("Học từ kịch bản đã biên tập")
    st.caption("Nạp các file **Word team đã biên tập** (và nếu có, **phụ đề gốc tiếng Anh .srt**). App rút ra gợi ý cho kho phiên âm, "
               "tên nhân vật, cụm không phải tên nhân vật, phân vai, xưng hô. **Chỉ những dòng được tích chọn mới được lưu.** "
               "Nội dung kịch bản không được lưu lại ở đâu cả, chỉ lưu các mục bạn duyệt.")
    st.caption("💡 Cách nhanh nhất: trên máy, bấm đúp file **`D:\\editkb\\GOM_KICH_BAN_MAU.bat`** → nó tạo 1 file "
               "**`KichBan_Mau_MaiHan.zip`** chứa toàn bộ kịch bản mẫu ở `D:\\HOC` → kéo thả đúng 1 file đó vào đây. "
               "Nạp lại bao nhiêu lần cũng được: mục nào đã có trong kho sẽ không bị đề xuất lại.")
    files = kept_file_uploader("Kéo thả file .zip, hoặc nhiều file .docx / .srt vào đây", type=["docx", "srt", "zip"],
                               accept_multiple_files=True, key="learn_uploader")
    if files and st.button(f"🔍 Phân tích {len(files)} file", type="primary", key="btn_learn_run"):
        docx, srt, skipped = _read_files(files)
        with st.spinner("Đang học từ kịch bản..."):
            st.session_state["learn_result"] = L.learn(docx, srt, VN_SYLLABLES, pronoun_qc.find_terms)
            st.session_state["learn_meta"] = (len(docx), len(srt), skipped)
        for k in [k for k in st.session_state if str(k).startswith("learn_ed_")]: del st.session_state[k]
    R = st.session_state.get("learn_result")
    if not R:
        st.info("📌 Tải lên các file rồi bấm **Phân tích**.")
        return
    n_docx, n_srt, skipped = st.session_state["learn_meta"]
    st.success(f"Đã học từ **{n_docx}** kịch bản Word và **{n_srt}** phụ đề gốc.")
    if skipped:
        with st.expander(f"⚠️ {len(skipped)} file không dùng"): st.write("\n".join(f"- {s}" for s in skipped))

    tabs = st.tabs(["🔤 Phiên âm", "🎭 Tên nhân vật", "🚫 Không phải tên", "🎙️ Phân vai", "💬 Xưng hô"])
    with tabs[0]: _phonetics(R)
    with tabs[1]: _speakers(R)
    with tabs[2]: _non_speakers(R)
    with tabs[3]: _cast(R)
    with tabs[4]: _pronouns(R)


def _editor(df, key, config, note):
    st.caption(note)
    c1, c2, _ = st.columns([1, 1, 3])
    ver = st.session_state.get(f"{key}_v", 0)
    if c1.button("✅ Chọn tất cả", key=f"{key}_all"): st.session_state[f"{key}_force"] = True; st.session_state[f"{key}_v"] = ver + 1; st.rerun()
    if c2.button("❌ Bỏ chọn hết", key=f"{key}_none"): st.session_state[f"{key}_force"] = False; st.session_state[f"{key}_v"] = ver + 1; st.rerun()
    force = st.session_state.get(f"{key}_force")
    if force is not None: df["Lưu"] = force
    return st.data_editor(df, column_config={"Lưu": st.column_config.CheckboxColumn("Lưu?", width="small"), **config},
                          hide_index=True, use_container_width=True, height=min(560, 38 + 35 * len(df)), key=f"learn_ed_{key}_{ver}")


def _done(key, n, what):
    st.session_state.pop(f"{key}_force", None)
    st.session_state[f"{key}_v"] = st.session_state.get(f"{key}_v", 0) + 1
    log_event("Học từ kịch bản cũ", f"Lưu {n} {what}")
    st.session_state["_learn_msg"] = f"✅ Đã lưu {n} {what}."
    st.rerun()


def _msg():
    m = st.session_state.pop("_learn_msg", None)
    if m: st.success(m)


# ---------- 1. phiên âm ----------
def _phonetics(R):
    _msg()
    kho = st.session_state["custom_phonetics"]
    frequent_speakers = {s.upper() for s, n in R["spk"].items() if n >= 3}
    rows = []
    for eng, variants in sorted(R["pho"].items(), key=lambda x: -sum(x[1].values())):
        proper = L.is_proper(eng, R["caps"], R["pho_eng_case"], variants)
        best = L.normalize_pho(variants.most_common(1)[0][0], proper)
        cur = kho.get(eng)
        if cur and L.normalize_pho(cur).lower() == best.lower(): continue  # đã có, giống hệt
        total = sum(variants.values())
        others = ", ".join(f"{v} ({n})" for v, n in variants.most_common(4)[1:])
        rows.append({"Lưu": bool(not cur and total >= 2 and eng not in frequent_speakers),
                     "Tiếng Anh": eng, "Phiên âm": best, "Tên riêng": proper, "Số lần": total,
                     "Cách viết khác": others, "Kho đang có": cur or "", "Ví dụ ở file": R["pho_src"].get(eng, "")})
    if not rows: st.success("Kho phiên âm đã có đủ mọi từ học được."); return
    df = _editor(pd.DataFrame(rows), "pho", {
        "Tiếng Anh": st.column_config.TextColumn(disabled=True),
        "Phiên âm": st.column_config.TextColumn("Phiên âm (sửa được)"),
        "Tên riêng": st.column_config.CheckboxColumn("Tên riêng?", help="Tích = tên người/địa danh/thương hiệu -> viết hoa chữ đầu (Chim-mi). Bỏ tích = danh từ thường -> viết thường (lốc)."),
        "Số lần": st.column_config.NumberColumn(disabled=True, width="small"),
        "Cách viết khác": st.column_config.TextColumn("Team từng viết khác", disabled=True),
        "Kho đang có": st.column_config.TextColumn(disabled=True),
        "Ví dụ ở file": st.column_config.TextColumn(disabled=True),
    }, f"**{len(rows)}** từ. Đã tự chọn sẵn từ **mới** xuất hiện từ 2 lần. Từ **đã có trong kho nhưng khác** thì không chọn sẵn — "
       "tích vào nếu muốn thay. Tên nhân vật chính (Preston, Bri...) không chọn sẵn để tránh chèn phiên âm mỗi lần nhắc tên.")
    sel = df[df["Lưu"]]
    if st.button(f"💾 Lưu {len(sel)} phiên âm vào kho", type="primary", key="btn_save_pho", disabled=sel.empty):
        for _, r in sel.iterrows():
            pho = clean_cell(r["Phiên âm"])
            if pho: kho[r["Tiếng Anh"]] = L.normalize_pho(pho, bool(r["Tên riêng"]))
        if save_json_db(PHONETIC_DB_FILE, kho): _done("pho", len(sel), "phiên âm")


# ---------- 2. tên nhân vật ----------
def _speakers(R):
    _msg()
    have = {s.upper() for s in st.session_state["custom_speakers"]}
    rows, seen = [], set()
    for src, counter in (("Kịch bản tiếng Việt", R["spk"]), ("Phụ đề gốc tiếng Anh", R["srt_names"])):
        for name, n in counter.most_common():
            u = name.strip().upper()
            if u in have or u in seen or not u: continue
            seen.add(u)
            rows.append({"Lưu": bool(n >= 2 and not L.has_vn(name) and not re.search(r"\d", name)),
                         "Tên nhân vật": name.strip(), "Số lần": n, "Học từ": src})
    if not rows: st.success("Danh sách tên nhân vật đã đủ."); return
    df = _editor(pd.DataFrame(rows), "spk", {
        "Tên nhân vật": st.column_config.TextColumn(disabled=True),
        "Số lần": st.column_config.NumberColumn(disabled=True, width="small"),
        "Học từ": st.column_config.TextColumn(disabled=True),
    }, f"**{len(rows)}** tên chưa có trong danh sách. Tên tiếng Việt (\"Giọng nam\", \"Nữ 1\"...) không chọn sẵn vì kịch bản gốc tiếng Anh không dùng.")
    sel = df[df["Lưu"]]
    if st.button(f"💾 Lưu {len(sel)} tên nhân vật", type="primary", key="btn_save_spk", disabled=sel.empty):
        st.session_state["custom_speakers"].update(sel["Tên nhân vật"].tolist())
        if save_json_db(SPEAKER_DB_FILE, st.session_state["custom_speakers"]): _done("spk", len(sel), "tên nhân vật")


# ---------- 3. không phải tên nhân vật ----------
def _non_speakers(R):
    _msg()
    have = {s.upper() for s in st.session_state["custom_non_speakers"]}
    spk = {s.upper() for s in st.session_state["custom_speakers"]}
    rows = [{"Lưu": n >= 2, "Cụm từ": p, "Số lần": n, "Ví dụ": R["ns_src"].get(p, ("", ""))[1]}
            for p, n in R["ns"].most_common() if p not in have and p not in spk]
    if not rows:
        st.success("Không có cụm mới." if R["ns"] or st.session_state["learn_meta"][1] else
                   "Cần nạp **phụ đề gốc tiếng Anh (.srt)** để học mục này."); return
    df = _editor(pd.DataFrame(rows), "ns", {
        "Cụm từ": st.column_config.TextColumn(disabled=True),
        "Số lần": st.column_config.NumberColumn(disabled=True, width="small"),
        "Ví dụ": st.column_config.TextColumn("Câu ví dụ trong phụ đề gốc", disabled=True),
    }, f"**{len(rows)}** cụm có dấu hai chấm trong phụ đề gốc nhưng KHÔNG phải tên người nói (VD: \"I was like:\"). Chọn sẵn cụm gặp từ 2 lần.")
    sel = df[df["Lưu"]]
    if st.button(f"💾 Lưu {len(sel)} cụm", type="primary", key="btn_save_ns", disabled=sel.empty):
        st.session_state["custom_non_speakers"].update(sel["Cụm từ"].str.upper().tolist())
        if save_json_db(NON_SPEAKER_DB_FILE, st.session_state["custom_non_speakers"]): _done("ns", len(sel), "cụm không phải tên nhân vật")


# ---------- 4. phân vai ----------
def _cast(R):
    _msg()
    kho = st.session_state["custom_cast_mapping"]
    rows = []
    for ch, cnt in sorted(R["cast"].items(), key=lambda x: -sum(x[1].values())):
        actor, n = cnt.most_common(1)[0]
        total = sum(cnt.values())
        cur = kho.get(ch, "")
        if cur.upper() == actor: continue
        clear = n / total >= 0.7 and total >= 2
        rows.append({"Lưu": bool(clear and not cur and not L.has_vn(ch) and not re.search(r"\d", ch)),
                     "Nhân vật": ch, "Diễn viên": actor, "Số lần": total,
                     "Diễn viên khác từng lồng": ", ".join(f"{a} ({k})" for a, k in cnt.most_common()[1:]), "Kho đang có": cur})
    if not rows: st.success("Bảng phân vai đã khớp."); return
    df = _editor(pd.DataFrame(rows), "cast", {
        "Nhân vật": st.column_config.TextColumn(disabled=True),
        "Diễn viên": st.column_config.TextColumn("Diễn viên (sửa được)"),
        "Số lần": st.column_config.NumberColumn(disabled=True, width="small"),
        "Diễn viên khác từng lồng": st.column_config.TextColumn(disabled=True),
        "Kho đang có": st.column_config.TextColumn(disabled=True),
    }, f"**{len(rows)}** nhân vật. Chỉ chọn sẵn khi **1 diễn viên chiếm đa số rõ ràng** và kho chưa có. "
       "Nhân vật đã có trong kho mà kịch bản cũ ghi khác (có thể do đổi diễn viên) thì không chọn sẵn — kho hiện tại được giữ nguyên.")
    sel = df[df["Lưu"]]
    if st.button(f"💾 Lưu {len(sel)} phân vai", type="primary", key="btn_save_cast", disabled=sel.empty):
        for _, r in sel.iterrows():
            a = clean_cell(r["Diễn viên"]).upper()
            if a: kho[r["Nhân vật"]] = a
        if save_json_db(CAST_DB_FILE, kho): _done("cast", len(sel), "phân vai")


# ---------- 5. xưng hô ----------
def _pronouns(R):
    _msg()
    kho = st.session_state["custom_pronoun_rel"]
    rows = []
    for key in sorted(set(R["pair_self"]) | set(R["pair_target"])):
        if key in kho: continue
        s, t = L.dominant(R["pair_self"].get(key, {})), L.dominant(R["pair_target"].get(key, {}))
        if not (s and t): continue
        a, b = key.split("|", 1)
        rows.append({"Lưu": not (L.has_vn(a) or L.has_vn(b) or re.search(r"\d", key)), "Người nói": a, "Nói với": b,
                     "Xưng": s, "Gọi": t, "Số câu": sum(R["pair_self"][key].values()) + sum(R["pair_target"][key].values())})
    if not rows: st.info("Chưa đủ dữ liệu để rút ra cặp xưng hô rõ ràng (cần nhiều câu qua lại giữa 2 nhân vật)."); return
    df = _editor(pd.DataFrame(rows).sort_values("Số câu", ascending=False), "pr", {
        "Người nói": st.column_config.TextColumn(disabled=True), "Nói với": st.column_config.TextColumn(disabled=True),
        "Xưng": st.column_config.TextColumn("Xưng (sửa được)"), "Gọi": st.column_config.TextColumn("Gọi (sửa được)"),
        "Số câu": st.column_config.NumberColumn(disabled=True, width="small"),
    }, f"**{len(rows)}** cặp có cách xưng & gọi chiếm ưu thế rõ (từ 60% trở lên). Dùng cho trang Soát xưng hô.")
    sel = df[df["Lưu"]]
    if st.button(f"💾 Lưu {len(sel)} cặp xưng hô", type="primary", key="btn_save_pr", disabled=sel.empty):
        for _, r in sel.iterrows():
            kho[f"{r['Người nói']}|{r['Nói với']}"] = {"self": clean_cell(r["Xưng"]), "target": clean_cell(r["Gọi"])}
        if save_json_db(PRONOUN_REL_DB_FILE, kho): _done("pr", len(sel), "cặp xưng hô")
