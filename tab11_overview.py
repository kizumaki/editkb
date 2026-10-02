"""Trang Tổng quan cho người quản lý: việc gì trễ, việc gì sắp đến hạn, video nào đứng yên, ai đang ôm nhiều việc."""
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
from utils import ensure_payroll_ready
import payroll_core as pc


def _today(): return datetime.now(timezone(timedelta(hours=7))).date()


def _stage_since(it):
    return pc.parse_date(it.get("stage_updated")) or pc.parse_date(it.get("date"))


def _df(rows):
    df = pd.DataFrame(rows)
    for c in ("Còn (ngày)", "Đứng yên (ngày)"):
        if c in df: df[c] = df[c].map(lambda v: "—" if v is None or pd.isna(v) else str(int(v)))  # "—" thay vì chữ "None"
    return df


def render_tab11():
    ensure_payroll_ready()
    st.subheader("Tổng quan công việc")
    tracker = st.session_state["dubbing_tracker"]
    projects = st.session_state["projects"]
    pmap = pc.project_map(projects)
    today = _today()

    c1, c2, c3 = st.columns([1.6, 1, 1])
    active_ids = [str(p["project_id"]) for p in projects if str(p.get("active", "1")) != "0"]
    with c1:
        f_proj = st.selectbox("Dự án", ["__all__"] + active_ids, key="ov_project",
                              format_func=lambda i: "Tất cả dự án đang chạy" if i == "__all__" else
                              pc.project_name(pmap, i))
    with c2: soon_days = st.number_input("Cảnh báo trước hạn (ngày)", 1, 30, 3, key="ov_soon")
    with c3: stuck_days = st.number_input("Coi là đứng yên sau (ngày)", 1, 60, 5, key="ov_stuck")

    items = [it for it in tracker if (str(it.get("project_id")) in active_ids if f_proj == "__all__" else str(it.get("project_id")) == f_proj)]
    open_items = [it for it in items if it.get("stage") != "Xong"]

    def row(it):
        dl = pc.parse_date(it.get("deadline"))
        since = _stage_since(it)
        p = pmap.get(str(it.get("project_id")), {})
        return {"Video": it.get("video_title", ""), "Dự án": p.get("name", ""),
                "Giai đoạn": it.get("stage", "") or "(chưa ghi)", "Người phụ trách": it.get("assignee", "") or "—",
                "Hạn giao": pc.fmt_date(dl) if dl else "—",
                "Còn (ngày)": (dl - today).days if dl else None,
                "Đứng yên (ngày)": (today - since).days if since else None}

    overdue = [it for it in open_items if (d := pc.parse_date(it.get("deadline"))) and d < today]
    soon = [it for it in open_items if (d := pc.parse_date(it.get("deadline"))) and today <= d <= today + timedelta(days=soon_days)]
    no_deadline = [it for it in open_items if not pc.parse_date(it.get("deadline"))]
    stuck = [it for it in open_items if (s := _stage_since(it)) and (today - s).days >= stuck_days]

    m = st.columns(5)
    m[0].metric("🎬 Đang làm", len(open_items))
    m[1].metric("🔴 Trễ hạn", len(overdue))
    m[2].metric(f"🟠 Đến hạn trong {soon_days} ngày", len(soon))
    m[3].metric(f"⏸️ Đứng yên ≥ {stuck_days} ngày", len(stuck))
    m[4].metric("❔ Chưa có hạn giao", len(no_deadline))

    if not items:
        st.info("Chưa có video nào. Video tự vào danh sách khi Re-Sync, hoặc thêm tay ở trang **Theo dõi & báo cáo lương → Tiến độ video**.")
        return

    st.markdown("#### 🚨 Cần xử lý ngay")
    urgent = sorted(overdue + [x for x in soon if x not in overdue],
                    key=lambda it: pc.parse_date(it.get("deadline")))
    if urgent:
        df = _df([row(it) for it in urgent])
        st.dataframe(df.style.apply(lambda r: ["background-color: rgba(239,68,68,0.15)" if str(r["Còn (ngày)"]).startswith("-")
                                               else "background-color: rgba(245,158,11,0.15)"] * len(r), axis=1),
                     hide_index=True, use_container_width=True,
                     column_config={"Còn (ngày)": st.column_config.TextColumn(help="Số âm = đã trễ bấy nhiêu ngày")})
    else:
        st.success(f"✅ Không có video nào trễ hạn hoặc đến hạn trong {soon_days} ngày tới.")

    if stuck:
        st.markdown(f"#### ⏸️ Video đứng yên ở một giai đoạn từ {stuck_days} ngày trở lên")
        st.dataframe(_df([row(it) for it in sorted(stuck, key=_stage_since)]), hide_index=True, use_container_width=True)

    left, right = st.columns(2)
    with left:
        st.markdown("#### 📊 Số video theo giai đoạn")
        by_proj = {}
        for it in items:
            name = pmap.get(str(it.get("project_id")), {}).get("name", "?")
            by_proj.setdefault(name, {s: 0 for s in pc.STAGES})
            if it.get("stage") in pc.STAGES: by_proj[name][it["stage"]] += 1
        st.dataframe(pd.DataFrame(by_proj).T.rename_axis("Dự án").reset_index(), hide_index=True, use_container_width=True)
    with right:
        st.markdown("#### 👥 Việc đang ôm theo người phụ trách")
        load = {}
        for it in open_items:
            who = it.get("assignee", "") or "(chưa giao)"
            d = load.setdefault(who, {"Người phụ trách": who, "Đang làm": 0, "Trễ hạn": 0, "Sắp đến hạn": 0})
            d["Đang làm"] += 1
            if it in overdue: d["Trễ hạn"] += 1
            if it in soon: d["Sắp đến hạn"] += 1
        if load:
            st.dataframe(pd.DataFrame(sorted(load.values(), key=lambda d: (-d["Trễ hạn"], -d["Đang làm"]))), hide_index=True, use_container_width=True)
        else:
            st.caption("Không có việc đang làm.")

    if no_deadline:
        with st.expander(f"❔ {len(no_deadline)} video đang làm nhưng CHƯA có hạn giao"):
            st.caption("Điền hạn giao ở trang **Theo dõi & báo cáo lương → Tiến độ video** (cột *Hạn giao*) để được cảnh báo đúng lúc.")
            st.dataframe(_df([row(it) for it in no_deadline]).drop(columns=["Hạn giao", "Còn (ngày)"]),
                         hide_index=True, use_container_width=True)
