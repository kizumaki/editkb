import streamlit as st
import io
from datetime import datetime, timedelta, timezone
import pandas as pd
from utils import (
    TRACKER_DB_FILE, PROJECTS_DB_FILE, PAYROLL_LOCKS_DB_FILE, save_json_db, reload_json_db,
    ensure_payroll_ready, generate_actor_salary_slip_docx, clean_cell, log_event
)
from auth import is_admin, current_user
import payroll_core as pc

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _now_vn(): return datetime.now(timezone(timedelta(hours=7)))

def _who():
    u = current_user() or {}
    return u.get("display_name") or u.get("username", "")

def _flash(kind, msg): st.session_state.setdefault("_pay_flash", []).append((kind, msg))

def _show_flash():
    # Luôn tạo 1 ô cố định (kể cả khi không có thông báo): nếu không, bảng thẻ bên dưới bị đổi vị trí
    # và Streamlit tự nhảy về thẻ đầu tiên ở lần tải lại sau.
    box = st.container()
    for kind, msg in st.session_state.pop("_pay_flash", []): getattr(box, kind)(msg)

def _bump(key): st.session_state[key] = st.session_state.get(key, 0) + 1

def _to_date(v):
    if v is None: return None
    try:
        if pd.isna(v): return None
    except (TypeError, ValueError): pass
    if isinstance(v, pd.Timestamp): return v.date()
    return pc.parse_date(v)

def _to_num(v):
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (TypeError, ValueError): return None

def _excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items(): df.to_excel(writer, index=False, sheet_name=name[:31])
    return buf.getvalue()

def _project_select(label, key, projects, allow_all=False):
    pmap = pc.project_map(projects)
    ids = [str(p["project_id"]) for p in projects]
    opts = (["__all__"] if allow_all else []) + ids
    if st.session_state.get(key) not in opts: st.session_state.pop(key, None)
    return st.selectbox(label, opts, key=key,
                        format_func=lambda i: "Tất cả dự án" if i == "__all__" else
                        pc.project_name(pmap, i) + ("" if str(pmap[i].get("active", "1")) != "0" else " (ngừng dùng)"))

def _slip_rows(rows):
    return [{"Stt": n, "Tiêu đề video": r["video_title"], "Thời lượng": pc.units_text(r["units"], r["mode"]),
             "Đơn giá": pc.money(r["rate"]), "Thành tiền": pc.money(r["pay"])} for n, r in enumerate(rows, 1)]


def render_tab3():
    ensure_payroll_ready()
    st.subheader("Theo dõi video & báo cáo lương")
    _show_flash()
    tracker = st.session_state["dubbing_tracker"]
    projects = st.session_state["projects"]
    locks = st.session_state["payroll_locks"]
    t_progress, t_projects, t_video_rates, t_pay = st.tabs([
        "📋 Tiến độ video", "📁 Dự án & đơn giá", "💵 Giá riêng từng video", "💰 Lương & chốt sổ"])
    with t_progress: _render_progress(tracker, projects, locks)
    with t_projects: _render_projects(tracker, projects)
    with t_video_rates: _render_video_rates(tracker, projects, locks)
    with t_pay: _render_payroll(tracker, projects, locks)


# ==========================================
# 1. TIẾN ĐỘ VIDEO
# ==========================================
def _render_progress(tracker, projects, locks):
    pmap = pc.project_map(projects)
    lp = pc.locked_pairs(locks)
    c1, c2, c3 = st.columns([1.3, 1, 1.3])
    with c1: f_proj = _project_select("Dự án", "trk_project", projects, allow_all=True)
    with c2: f_stage = st.selectbox("Giai đoạn", ["__all__"] + pc.STAGES, key="trk_stage",
                                    format_func=lambda s: "Tất cả giai đoạn" if s == "__all__" else s)
    with c3: f_text = st.text_input("Tìm theo tên video", key="trk_search", placeholder="Gõ vài chữ trong tên video...")

    in_proj = [i for i, it in enumerate(tracker) if f_proj == "__all__" or str(it.get("project_id")) == f_proj]
    m_cols = st.columns(len(pc.STAGES))
    for s, col in zip(pc.STAGES, m_cols):
        col.metric(s, sum(1 for i in in_proj if tracker[i].get("stage") == s))

    with st.expander("➕ Thêm video để theo dõi từ khâu dịch (trước khi Re-Sync)", expanded=False):
        st.caption("Đặt **tên video trùng tên file kịch bản** (bỏ đuôi .docx). Khi Re-Sync file đó, app tự cập nhật vào đúng dòng này.")
        active = [p for p in projects if str(p.get("active", "1")) != "0"] or projects
        with st.form("form_add_video", clear_on_submit=True):
            a1, a2 = st.columns(2)
            new_title = a1.text_input("Tên video")
            new_pid = a2.selectbox("Dự án", [str(p["project_id"]) for p in active], format_func=lambda i: pc.project_name(pmap, i))
            a3, a4, a5, a6 = st.columns(4)
            new_stage = a3.selectbox("Giai đoạn", pc.STAGES)
            new_assignee = a4.text_input("Người phụ trách")
            new_date = a5.date_input("Ngày ghi nhận", value=_now_vn().date(), format="DD/MM/YYYY")
            new_deadline = a6.date_input("Hạn giao", value=None, format="DD/MM/YYYY", help="Để trống nếu chưa biết")
            if st.form_submit_button("➕ Thêm video", type="primary"):
                t = new_title.strip()
                if not t: st.error("Hãy gõ tên video.")
                elif any(str(it.get("video_title", "")).strip().upper() == t.upper() for it in tracker):
                    st.error(f"Đã có video tên «{t}» trong danh sách.")
                else:
                    tracker.append({"video_title": t, "project_id": new_pid, "stage": new_stage, "assignee": new_assignee.strip(),
                                    "date": pc.fmt_date(new_date), "actors": pc.NO_ACTOR, "actor_breakdown": {},
                                    "total_lines": 0, "video_duration_min": 0, "custom_actor_rates": {}, "project_week": "",
                                    "deadline": pc.fmt_date(new_deadline) if new_deadline else "",
                                    "stage_updated": pc.fmt_date(_now_vn().date())})
                    if save_json_db(TRACKER_DB_FILE, tracker):
                        _flash("success", f"✅ Đã thêm video «{t}».")
                        _bump("trk_ver"); st.rerun()

    idxs = [i for i in in_proj
            if (f_stage == "__all__" or tracker[i].get("stage") == f_stage)
            and f_text.strip().lower() in str(tracker[i].get("video_title", "")).lower()]
    if not idxs:
        st.info("Chưa có video nào khớp bộ lọc. Video tự vào đây mỗi lần Re-Sync (menu Kịch bản → Re-Sync bản đã biên tập), "
                "hoặc thêm tay ở ô ➕ phía trên.")
        return

    rows = []
    for i in idxs:
        it = tracker[i]
        rows.append({
            "Chốt": "🔒" if pc.video_is_locked(it, lp) else "",
            "Tiêu đề video": it.get("video_title", ""),
            "Dự án": pc.project_name(pmap, it.get("project_id")),
            "Ngày": pc.parse_date(it.get("date")),
            "Hạn giao": pc.parse_date(it.get("deadline")),
            "Giai đoạn": it.get("stage") if it.get("stage") in pc.STAGES else None,
            "Người phụ trách": it.get("assignee", ""),
            "Diễn viên": it.get("actors", ""),
            "Khối lượng": f"{int(pc._num(it.get('video_duration_min')))} phút · {int(pc._num(it.get('total_lines')))} câu",
            "Ghi chú": it.get("project_week", ""),
            "Xóa": False,
        })
    df = pd.DataFrame(rows, index=idxs)
    edited = st.data_editor(
        df,
        column_config={
            "Chốt": st.column_config.TextColumn("🔒", disabled=True, width="small", help="🔒 = video đã chốt lương (một phần hoặc toàn bộ)"),
            "Tiêu đề video": st.column_config.TextColumn("Tiêu đề video", width="large"),
            "Dự án": st.column_config.SelectboxColumn("Dự án", options=[p.get("name", "") for p in projects], required=True),
            "Ngày": st.column_config.DateColumn("Ngày", format="DD/MM/YYYY", help="Ngày ghi nhận — dùng để chia đợt lương"),
            "Hạn giao": st.column_config.DateColumn("Hạn giao", format="DD/MM/YYYY", help="Hạn giao — trang Tổng quan dùng để cảnh báo trễ hạn"),
            "Giai đoạn": st.column_config.SelectboxColumn("Giai đoạn", options=pc.STAGES),
            "Người phụ trách": st.column_config.TextColumn("Người phụ trách"),
            "Diễn viên": st.column_config.TextColumn("Diễn viên", help="Cách nhau bằng dấu phẩy"),
            "Khối lượng": st.column_config.TextColumn("Khối lượng", disabled=True),
            "Ghi chú": st.column_config.TextColumn("Ghi chú tuần/đợt"),
            "Xóa": st.column_config.CheckboxColumn("Xóa?", width="small"),
        },
        hide_index=True, use_container_width=True,
        key=f"trk_editor_{f_proj}_{f_stage}_{f_text}_{st.session_state.get('trk_ver', 0)}",
    )

    b1, b2 = st.columns(2)
    if b1.button("💾 Lưu thay đổi", type="primary", use_container_width=True, key="btn_trk_save"):
        _save_progress(tracker, projects, lp, edited)
    b2.download_button("📊 Tải danh sách đang xem ra Excel", data=_excel({"Tien do": df.drop(columns=["Xóa"])}),
                       file_name="Tien_Do_Video_MaiHan.xlsx", mime=XLSX_MIME, use_container_width=True, key="dl_trk")


def _save_progress(tracker, projects, lp, edited):
    name_to_id = {p.get("name", ""): str(p["project_id"]) for p in projects}
    warns, deleted = [], []
    for i, row in edited.iterrows():
        it = tracker[i]
        title = str(it.get("video_title", ""))
        locked_actors = [k[1] for k in lp if k[0] == title.strip().upper()]
        if row["Xóa"]:
            if locked_actors: warns.append(f"«{title}»: đã chốt lương nên không xoá được.")
            else: deleted.append(i)
            continue
        new_title = clean_cell(row["Tiêu đề video"]) or title
        new_pid = name_to_id.get(clean_cell(row["Dự án"]), str(it.get("project_id")))
        new_date = _to_date(row["Ngày"])
        new_actors = pc.actors_of({"actors": clean_cell(row["Diễn viên"])})
        if new_title.upper() != title.upper() and any(
                str(x.get("video_title", "")).strip().upper() == new_title.upper() for x in tracker if x is not it):
            warns.append(f"«{new_title}»: trùng tên video khác, giữ tên cũ."); new_title = title
        if locked_actors:
            old_date = pc.parse_date(it.get("date"))
            if new_title.upper() != title.upper(): warns.append(f"«{title}»: đã chốt lương, không đổi tên được."); new_title = title
            if new_pid != str(it.get("project_id")): warns.append(f"«{title}»: đã chốt lương, không đổi dự án được."); new_pid = str(it.get("project_id"))
            if new_date != old_date: warns.append(f"«{title}»: đã chốt lương, không đổi ngày được."); new_date = old_date
            missing = [a for a in locked_actors if a not in new_actors]
            if missing:
                warns.append(f"«{title}»: {', '.join(missing)} đã chốt lương, không bỏ khỏi video được.")
                new_actors += missing
        it["video_title"] = new_title
        it["project_id"] = new_pid
        it["date"] = pc.fmt_date(new_date) if new_date else ""
        new_stage = clean_cell(row["Giai đoạn"])
        if new_stage != it.get("stage", ""): it["stage_updated"] = pc.fmt_date(_now_vn().date())  # để biết video "đứng yên" bao lâu
        it["stage"] = new_stage
        dl = _to_date(row["Hạn giao"])
        it["deadline"] = pc.fmt_date(dl) if dl else ""
        it["assignee"] = clean_cell(row["Người phụ trách"])
        it["project_week"] = clean_cell(row["Ghi chú"])
        it["actors"] = ", ".join(new_actors) if new_actors else pc.NO_ACTOR
        old_bd = it.get("actor_breakdown") or {}
        it["actor_breakdown"] = {a: old_bd.get(a, {"lines": 0, "words": 0}) for a in new_actors}
    removed = [tracker[i].get("video_title", "") for i in deleted]
    for i in sorted(deleted, reverse=True): tracker.pop(i)
    if save_json_db(TRACKER_DB_FILE, tracker):
        if removed: log_event("Xoá video khỏi bảng lương", ", ".join(removed))
        _flash("success", "✅ Đã lưu thay đổi.")
        for w in warns: _flash("warning", "⚠️ " + w)
        _bump("trk_ver"); st.rerun()


# ==========================================
# 2. DỰ ÁN & ĐƠN GIÁ
# ==========================================
def _render_projects(tracker, projects):
    st.markdown("#### Danh sách dự án")
    st.caption("Mỗi dự án có **cách tính** và **đơn giá mặc định** riêng. Đổi 2 thứ này sẽ áp dụng cho mọi video "
               "**chưa chốt sổ** của dự án; phần đã chốt giữ nguyên số tiền.")
    labels = list(pc.MODE_LABELS.values())
    label_to_mode = {v: k for k, v in pc.MODE_LABELS.items()}
    df = pd.DataFrame([{
        "Tên dự án": p.get("name", ""),
        "Cách tính": pc.MODE_LABELS.get(p.get("mode"), labels[0]),
        "Đơn giá mặc định": int(pc._num(p.get("unit_rate"))),
        "Đang dùng": str(p.get("active", "1")) != "0",
        "Số video": sum(1 for it in tracker if str(it.get("project_id")) == str(p["project_id"])),
    } for p in projects])
    edited = st.data_editor(
        df,
        column_config={
            "Tên dự án": st.column_config.TextColumn("Tên dự án", required=True, width="large"),
            "Cách tính": st.column_config.SelectboxColumn("Cách tính", options=labels, required=True),
            "Đơn giá mặc định": st.column_config.NumberColumn("Đơn giá mặc định (VNĐ)", min_value=0, step=1000, format="%,d", required=True),
            "Đang dùng": st.column_config.CheckboxColumn("Đang dùng", help="Bỏ chọn để ẩn dự án khỏi ô chọn ở trang Re-Sync (dữ liệu vẫn giữ)"),
            "Số video": st.column_config.NumberColumn("Số video", disabled=True),
        },
        hide_index=True, use_container_width=True, key=f"proj_editor_{st.session_state.get('proj_ver', 0)}",
    )
    if st.button("💾 Lưu danh sách dự án", type="primary", key="btn_proj_save"):
        names = [clean_cell(n) for n in edited["Tên dự án"]]
        if any(not n for n in names): st.error("Tên dự án không được để trống.")
        elif len({n.upper() for n in names}) != len(names): st.error("Có 2 dự án trùng tên. Hãy đặt tên khác nhau.")
        else:
            changes = []
            for p, (_, row), name in zip(projects, edited.iterrows(), names):
                new = {"name": name, "mode": label_to_mode.get(row["Cách tính"], p.get("mode", "minute")),
                       "unit_rate": _to_num(row["Đơn giá mặc định"]) or 0, "active": "1" if row["Đang dùng"] else "0"}
                diff = [f"{k}: {p.get(k)} → {v}" for k, v in new.items() if str(p.get(k)) != str(v)]
                if diff: changes.append(f"{p.get('name')} ({'; '.join(diff)})")
                p.update(new)
            if save_json_db(PROJECTS_DB_FILE, projects):
                if changes: log_event("Sửa dự án", " | ".join(changes))
                _flash("success", "✅ Đã lưu danh sách dự án."); _bump("proj_ver"); st.rerun()

    with st.expander("➕ Tạo dự án mới", expanded=False):
        with st.form("form_add_project", clear_on_submit=True):
            n1, n2, n3 = st.columns([2, 1.3, 1.3])
            p_name = n1.text_input("Tên dự án", placeholder="VD: Preston mùa 3")
            p_mode = n2.selectbox("Cách tính", labels)
            p_rate = n3.number_input("Đơn giá mặc định (VNĐ)", min_value=0, value=30000, step=1000)
            if st.form_submit_button("➕ Tạo dự án", type="primary"):
                name = p_name.strip()
                if not name: st.error("Hãy gõ tên dự án.")
                elif any(str(p.get("name", "")).strip().upper() == name.upper() for p in projects): st.error("Đã có dự án trùng tên.")
                else:
                    projects.append({"project_id": pc.new_id("DA"), "name": name, "mode": label_to_mode[p_mode],
                                     "unit_rate": p_rate, "actor_rates": {}, "active": "1", "created": pc.fmt_date(_now_vn().date())})
                    if save_json_db(PROJECTS_DB_FILE, projects):
                        log_event("Tạo dự án", f"{name} · {p_mode} · {p_rate}")
                        _flash("success", f"✅ Đã tạo dự án «{name}»."); _bump("proj_ver"); st.rerun()

    st.markdown("---")
    st.markdown("#### Đơn giá riêng của diễn viên trong dự án")
    st.caption("Ví dụ: dự án A trả Khánh 35,000/phút trong khi giá mặc định là 30,000. **Để trống = dùng giá mặc định của dự án.** "
               "Muốn thêm diễn viên chưa có trong danh sách: bấm dòng trống cuối bảng.")
    pid = _project_select("Chọn dự án", "rate_project", projects)
    p = pc.project_map(projects)[pid]
    ar = p.get("actor_rates") or {}
    seen = set(ar)
    for it in tracker:
        if str(it.get("project_id")) == pid: seen.update(pc.actors_of(it))
    df_r = pd.DataFrame([{"Diễn viên": a, "Đơn giá riêng": _to_num(ar.get(a))} for a in sorted(seen)],
                        columns=["Diễn viên", "Đơn giá riêng"])
    df_r["Đơn giá riêng"] = df_r["Đơn giá riêng"].astype("float")
    st.caption(f"Dự án **{p.get('name')}**: {pc.MODE_LABELS.get(p.get('mode'), '')}, giá mặc định "
               f"**{pc.money(pc._num(p.get('unit_rate')))} / {pc.MODES.get(p.get('mode'), '')}**.")
    edited_r = st.data_editor(
        df_r, num_rows="dynamic",
        column_config={
            "Diễn viên": st.column_config.TextColumn("Diễn viên"),
            "Đơn giá riêng": st.column_config.NumberColumn("Đơn giá riêng (VNĐ) — trống = giá mặc định", min_value=0, step=1000, format="%,d"),
        },
        hide_index=True, use_container_width=True, key=f"actor_rate_editor_{pid}_{st.session_state.get('proj_ver', 0)}",
    )
    if st.button("💾 Lưu đơn giá diễn viên", type="primary", key="btn_actor_rate_save"):
        new_ar = {}
        for _, row in edited_r.iterrows():
            a, r = clean_cell(row["Diễn viên"]).upper(), _to_num(row["Đơn giá riêng"])
            if a and r is not None: new_ar[a] = int(r) if float(r).is_integer() else r
        if new_ar != ar:
            old = dict(ar); p["actor_rates"] = new_ar
            if save_json_db(PROJECTS_DB_FILE, projects):
                log_event("Sửa đơn giá diễn viên theo dự án", f"{p.get('name')}: {old} → {new_ar}")
                _flash("success", "✅ Đã lưu đơn giá diễn viên."); _bump("proj_ver"); st.rerun()
        else: st.info("Không có gì thay đổi.")


# ==========================================
# 3. GIÁ RIÊNG TỪNG VIDEO
# ==========================================
def _render_video_rates(tracker, projects, locks):
    st.caption("Dùng khi 1 video đặc biệt có giá khác (VD: video khó). Ưu tiên: **giá riêng video** → **giá diễn viên × dự án** → "
               "**giá mặc định dự án**. Gõ số mới vào cột đơn giá để đặt giá riêng; tích **Dùng giá chung** để bỏ giá riêng.")
    pid = _project_select("Chọn dự án", "vr_project", projects)
    project = pc.project_map(projects)[pid]
    lp = pc.locked_pairs(locks)
    rows, metas = [], []
    for i, it in enumerate(tracker):
        if str(it.get("project_id")) != pid: continue
        for a in pc.actors_of(it):
            c = pc.compute_pay(it, a, project)
            locked = pc.pair_key(it.get("video_title", ""), a) in lp
            rows.append({"Chốt": "🔒" if locked else "", "Video": it.get("video_title", ""), "Diễn viên": a,
                         "Khối lượng": pc.units_text(c["units"], c["mode"]), "Đơn giá": c["rate"],
                         "Dùng giá chung": c["source"] != "Riêng video này", "Nguồn giá": c["source"],
                         "Thành tiền": pc.money(c["pay"])})
            metas.append((i, a, c["rate"], locked))
    if not rows:
        st.info("Dự án này chưa có video nào có diễn viên."); return
    edited = st.data_editor(
        pd.DataFrame(rows),
        column_config={
            "Chốt": st.column_config.TextColumn("🔒", disabled=True, width="small"),
            "Video": st.column_config.TextColumn("Video", disabled=True, width="medium"),
            "Diễn viên": st.column_config.TextColumn("Diễn viên", disabled=True),
            "Khối lượng": st.column_config.TextColumn("Khối lượng", disabled=True),
            "Đơn giá": st.column_config.NumberColumn("Đơn giá áp dụng (VNĐ)", min_value=0, step=1000, format="%,d"),
            "Dùng giá chung": st.column_config.CheckboxColumn("Dùng giá chung"),
            "Nguồn giá": st.column_config.TextColumn("Giá đang lấy từ", disabled=True),
            "Thành tiền": st.column_config.TextColumn("Thành tiền (chưa lưu)", disabled=True),
        },
        hide_index=True, use_container_width=True, key=f"vr_editor_{pid}_{st.session_state.get('vr_ver', 0)}",
    )
    if st.button("💾 Lưu giá riêng", type="primary", key="btn_vr_save"):
        warns, changed = [], []
        for (i, a, orig, locked), (_, row) in zip(metas, edited.iterrows()):
            it = tracker[i]
            cr = it.setdefault("custom_actor_rates", {})
            new_rate = _to_num(row["Đơn giá"])
            use_common = bool(row["Dùng giá chung"])
            had_custom = a in cr
            if new_rate is not None and abs(new_rate - orig) > 0.5: action = ("set", new_rate)
            elif use_common and had_custom: action = ("del", None)
            elif not use_common and not had_custom: action = ("set", orig)
            else: continue
            if locked:
                warns.append(f"«{it.get('video_title')}» – {a}: đã chốt lương, không đổi giá được."); continue
            if action[0] == "set": cr[a] = action[1]
            else: cr.pop(a, None)
            changed.append(f"{it.get('video_title')} – {a}: {'giá chung' if action[0] == 'del' else action[1]}")
        if changed:
            if save_json_db(TRACKER_DB_FILE, tracker):
                log_event("Sửa giá riêng video", " | ".join(changed)[:500])
                _flash("success", f"✅ Đã lưu {len(changed)} thay đổi giá.")
        for w in warns: _flash("warning", "⚠️ " + w)
        if not changed and not warns: _flash("info", "Không có gì thay đổi.")
        _bump("vr_ver"); st.rerun()


# ==========================================
# 4. LƯƠNG & CHỐT SỔ
# ==========================================
def _render_payroll(tracker, projects, locks):
    today = _now_vn().date()
    if "pay_from" not in st.session_state: st.session_state["pay_from"] = today.replace(day=1)
    if "pay_to" not in st.session_state: st.session_state["pay_to"] = today
    c1, c2, c3 = st.columns([1.6, 1, 1])
    with c1: pid = _project_select("Dự án", "pay_project", projects)
    with c2: d_from = st.date_input("Từ ngày", key="pay_from", format="DD/MM/YYYY")
    with c3: d_to = st.date_input("Đến ngày", key="pay_to", format="DD/MM/YYYY")
    if d_from > d_to:
        st.error("«Từ ngày» đang sau «Đến ngày». Hãy chọn lại."); return
    project = pc.project_map(projects)[pid]
    mode = project.get("mode", "minute")
    period = f"{project.get('name')} · {pc.fmt_date(d_from)} – {pc.fmt_date(d_to)}"
    st.caption(f"Cách tính của dự án: **{pc.MODE_LABELS.get(mode, '')}**, giá mặc định **{pc.money(pc._num(project.get('unit_rate')))}"
               f" / {pc.MODES.get(mode, '')}**. Video được tính theo **ngày ghi nhận** (cột Ngày ở thẻ Tiến độ video).")

    undated = pc.undated_videos(tracker, pid)
    if undated:
        names = ", ".join(f"«{u.get('video_title')}»" for u in undated[:5]) + (" ..." if len(undated) > 5 else "")
        st.warning(f"⚠️ {len(undated)} video chưa có ngày nên **không được tính vào đợt nào**: {names}. "
                   "Hãy điền ngày ở thẻ 📋 Tiến độ video.")

    rows = pc.period_rows(tracker, projects, locks, pid, d_from, d_to)
    if not rows:
        st.info("Không có video nào có diễn viên trong dự án và khoảng ngày này.")
    else:
        actors = sorted({r["actor"] for r in rows})
        summary = []
        for a in actors:
            ar = [r for r in rows if r["actor"] == a]
            done = sum(r["pay"] for r in ar if r["locked"]); todo = sum(r["pay"] for r in ar if not r["locked"])
            summary.append({"Diễn viên": a, "Số video": len(ar), "Khối lượng": pc.units_text(sum(r["units"] for r in ar), mode),
                            "Đã chốt": pc.money(done), "Chưa chốt": pc.money(todo), "Tổng cộng": pc.money(done + todo)})
        g_done = sum(r["pay"] for r in rows if r["locked"]); g_todo = sum(r["pay"] for r in rows if not r["locked"])
        m1, m2, m3 = st.columns(3)
        m1.metric("💰 Tổng lương đợt này", pc.money(g_done + g_todo))
        m2.metric("🔒 Đã chốt sổ", pc.money(g_done))
        m3.metric("⏳ Chưa chốt", pc.money(g_todo))
        df_sum = pd.DataFrame(summary)
        st.dataframe(df_sum, hide_index=True, use_container_width=True)

        detail = pd.DataFrame([{"Diễn viên": r["actor"], "Video": r["video_title"], "Ngày": pc.fmt_date(r["date"]),
                                "Khối lượng": pc.units_text(r["units"], r["mode"]), "Đơn giá": pc.money(r["rate"]),
                                "Thành tiền": pc.money(r["pay"]), "Giá lấy từ": r["source"],
                                "Trạng thái": "🔒 Đã chốt" if r["locked"] else "Chưa chốt"} for r in rows])
        st.download_button("📊 Tải báo cáo lương đợt này ra Excel (tổng hợp + chi tiết)",
                           data=_excel({"Tong hop": df_sum, "Chi tiet": detail}),
                           file_name=f"Bao_Cao_Luong_{pc.fmt_date(d_from).replace('/', '-')}_{pc.fmt_date(d_to).replace('/', '-')}.xlsx",
                           mime=XLSX_MIME, use_container_width=True, key="dl_pay_report")

        st.markdown("#### 👤 Chi tiết & phiếu lương từng diễn viên")
        a_sel = st.selectbox("Chọn diễn viên", actors, key="pay_actor")
        a_rows = [r for r in rows if r["actor"] == a_sel]
        st.dataframe(detail[detail["Diễn viên"] == a_sel].drop(columns=["Diễn viên"]), hide_index=True, use_container_width=True)
        a_total = sum(r["pay"] for r in a_rows)
        st.metric(f"Tổng thù lao của {a_sel}", pc.money(a_total))
        s1, s2 = st.columns(2)
        s1.download_button(f"🖨️ Phiếu lương Word của {a_sel}",
                           data=generate_actor_salary_slip_docx(a_sel, period, _slip_rows(a_rows), a_total, mode),
                           file_name=f"PhieuLuong_{a_sel}_{pc.fmt_date(d_from).replace('/', '-')}.docx", mime=DOCX_MIME,
                           type="primary", use_container_width=True, key="dl_slip_docx")
        s2.download_button("📊 Phiếu lương Excel", data=_excel({"Phieu luong": pd.DataFrame(_slip_rows(a_rows))}),
                           file_name=f"PhieuLuong_{a_sel}_{pc.fmt_date(d_from).replace('/', '-')}.xlsx", mime=XLSX_MIME,
                           use_container_width=True, key="dl_slip_xlsx")

        _render_lock_box(tracker, projects, project, rows, d_from, d_to)

    _render_lock_history(locks, project, mode)


def _render_lock_box(tracker, projects, project, rows, d_from, d_to):
    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 🔒 Chốt sổ")
        live = [r for r in rows if not r["locked"]]
        if not live:
            st.success("✅ Toàn bộ lương trong đợt này đã được chốt sổ."); return
        st.caption("Chốt sổ = **lưu cứng số tiền** của từng video. Sau khi chốt, đổi đơn giá hay sửa video cũng **không làm đổi** "
                   "số tiền đã chốt. Chỉ **Quản trị** mở khoá được. Mỗi lần chốt đều được ghi vào nhật ký.")
        live_actors = sorted({r["actor"] for r in live})
        ver = st.session_state.get("lock_ver", 0)
        chosen = st.multiselect("Chốt cho diễn viên", live_actors, default=live_actors, key=f"lock_actors_{ver}")
        sel = [r for r in live if r["actor"] in chosen]
        st.markdown(f"Sẽ chốt **{len(sel)}** dòng (video × diễn viên), tổng **{pc.money(sum(r['pay'] for r in sel))}**.")
        ok = st.checkbox("Tôi đã kiểm tra kỹ số liệu ở trên.", key=f"lock_confirm_{ver}")
        if st.button("🔒 Chốt sổ", type="primary", disabled=not (ok and sel), key="btn_lock"):
            # Đọc lại bản mới nhất trước khi chốt, tránh chốt trùng với người khác vừa chốt
            if not all(reload_json_db(f) for f in (PAYROLL_LOCKS_DB_FILE, TRACKER_DB_FILE, PROJECTS_DB_FILE)): return
            locks = st.session_state["payroll_locks"]
            fresh_project = pc.project_map(st.session_state["projects"]).get(str(project["project_id"]), project)
            fresh = pc.period_rows(st.session_state["dubbing_tracker"], st.session_state["projects"], locks,
                                   project["project_id"], d_from, d_to)
            now_text = _now_vn().strftime("%d/%m/%Y %H:%M")
            new_locks = [l for l in (pc.make_lock(fresh_project, a, fresh, d_from, d_to, _who(), now_text) for a in chosen) if l]
            if not new_locks:
                _flash("info", "Không còn gì để chốt (có thể người khác vừa chốt xong)."); st.rerun()
            locks.extend(new_locks)
            if save_json_db(PAYROLL_LOCKS_DB_FILE, locks):
                for l in new_locks:
                    log_event("Chốt sổ lương", f"{l['lock_id']} · {l['project_name']} · {l['actor']} · {l['from_date']}–{l['to_date']} · "
                                               f"{len(l['rows'])} video · {pc.money(l['total'])}")
                _flash("success", f"🔒 Đã chốt sổ cho {len(new_locks)} diễn viên, tổng {pc.money(sum(l['total'] for l in new_locks))}.")
                _bump("lock_ver"); st.rerun()


def _render_lock_history(locks, project, mode):
    mine = sorted([l for l in locks if str(l.get("project_id")) == str(project["project_id"])],
                  key=lambda l: str(l.get("lock_id", "")), reverse=True)
    with st.expander(f"📚 Các đợt đã chốt của dự án này ({len(mine)})", expanded=False):
        if not mine:
            st.caption("Chưa có đợt chốt nào."); return
        st.dataframe(pd.DataFrame([{
            "Mã": l.get("lock_id"), "Diễn viên": l.get("actor"), "Từ ngày": l.get("from_date"), "Đến ngày": l.get("to_date"),
            "Số video": len(l.get("rows") or []), "Số tiền": pc.money(pc._num(l.get("total"))),
            "Người chốt": l.get("locked_by"), "Lúc chốt": l.get("locked_at"),
            "Trạng thái": "🔒 Đang chốt" if l.get("status", pc.LOCKED) == pc.LOCKED else
                          f"🔓 Đã mở khoá ({l.get('unlocked_by')}, {l.get('unlocked_at')}): {l.get('note', '')}",
        } for l in mine]), hide_index=True, use_container_width=True)

        by_id = {l["lock_id"]: l for l in mine}
        fmt = lambda i: f"{by_id[i].get('actor')} · {by_id[i].get('from_date')}–{by_id[i].get('to_date')} · {pc.money(pc._num(by_id[i].get('total')))}"
        pick = st.selectbox("Tải lại phiếu lương của đợt đã chốt", list(by_id), format_func=fmt, key="lock_pick")
        l = by_id[pick]
        l_rows = [{"video_title": r.get("video_title"), "units": int(pc._num(r.get("units"))), "mode": l.get("mode", mode),
                   "rate": pc._num(r.get("rate")), "pay": pc._num(r.get("pay"))} for r in l.get("rows") or []]
        st.download_button("🖨️ Tải phiếu lương Word của đợt này",
                           data=generate_actor_salary_slip_docx(l.get("actor"), f"{l.get('project_name')} · {l.get('from_date')} – {l.get('to_date')} (đã chốt)",
                                                                _slip_rows(l_rows), pc._num(l.get("total")), l.get("mode", mode)),
                           file_name=f"PhieuLuong_{l.get('actor')}_{pick}.docx", mime=DOCX_MIME, key="dl_lock_slip")

        st.markdown("---")
        active = [i for i in by_id if by_id[i].get("status", pc.LOCKED) == pc.LOCKED]
        if not is_admin():
            st.caption("🔓 Chỉ tài khoản **Quản trị** mở khoá được đợt đã chốt."); return
        if not active: return
        st.markdown("##### 🔓 Mở khoá (chỉ Quản trị)")
        st.caption("Mở khoá xong, các video của đợt đó quay về «chưa chốt» và tính lại theo đơn giá hiện tại.")
        ver = st.session_state.get("unlock_ver", 0)
        u_pick = st.selectbox("Đợt cần mở khoá", active, format_func=fmt, key=f"unlock_pick_{ver}")
        reason = st.text_input("Lý do mở khoá (bắt buộc)", key=f"unlock_reason_{ver}")
        if st.button("🔓 Mở khoá đợt này", disabled=not reason.strip(), key="btn_unlock"):
            if not reload_json_db(PAYROLL_LOCKS_DB_FILE): return
            target = next((x for x in st.session_state["payroll_locks"] if x.get("lock_id") == u_pick), None)
            if not target or target.get("status", pc.LOCKED) != pc.LOCKED:
                _flash("info", "Đợt này đã được mở khoá trước đó."); _bump("unlock_ver"); st.rerun()
            target.update({"status": pc.UNLOCKED, "unlocked_by": _who(),
                           "unlocked_at": _now_vn().strftime("%d/%m/%Y %H:%M"), "note": reason.strip()})
            if save_json_db(PAYROLL_LOCKS_DB_FILE, st.session_state["payroll_locks"]):
                log_event("Mở khoá chốt sổ lương", f"{u_pick} · {target.get('actor')} · {pc.money(pc._num(target.get('total')))} · lý do: {reason.strip()}")
                _flash("success", "🔓 Đã mở khoá."); _bump("unlock_ver"); st.rerun()
