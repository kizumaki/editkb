"""Lõi tính lương theo dự án (không phụ thuộc giao diện — để kiểm thử được riêng).

Thứ tự ưu tiên đơn giá cho 1 diễn viên trong 1 video:
  1. Đơn giá riêng của video đó (custom_actor_rates)
  2. Đơn giá riêng của diễn viên trong dự án (actor_rates của dự án)
  3. Đơn giá mặc định của dự án
Cặp (video, diễn viên) đã chốt sổ thì dùng số liệu lưu cứng trong đợt chốt, không tính lại.
"""
import secrets
from datetime import date, datetime, timedelta

STAGES = ["Dịch", "Biên tập", "Thu âm", "Xong"]
MODES = {"minute": "phút", "line": "câu", "word": "từ"}
MODE_LABELS = {"minute": "Theo phút video", "line": "Theo câu thoại", "word": "Theo số từ"}
DEFAULT_PROJECT_ID = "mac_dinh"
DEFAULT_PROJECT_NAME = "Dự án mặc định"
NO_ACTOR = "CHƯA CÓ THÔNG TIN"
LOCKED, UNLOCKED = "chot", "mo_khoa"


# ---------- ngày tháng ----------
def parse_date(v):
    """Đọc ngày từ 'dd/mm/yyyy', 'yyyy-mm-dd', số ngày kiểu Google Sheets, hoặc date. Không đọc được -> None."""
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < v < 80000: return date(1899, 12, 30) + timedelta(days=int(v))
        return None
    s = str(v or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    return None

def fmt_date(d):
    return d.strftime("%d/%m/%Y") if d else ""


# ---------- video, diễn viên, khối lượng ----------
def actors_of(item):
    out = []
    for a in str(item.get("actors", "")).split(","):
        a = a.strip().upper()
        if a and a != NO_ACTOR and a not in out: out.append(a)
    return out

def _num(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0

def actor_units(item, actor, mode):
    """Khối lượng tính tiền của 1 diễn viên trong 1 video (phút / câu / từ)."""
    bd = item.get("actor_breakdown") or {}
    if mode == "minute":
        return int(_num(item.get("video_duration_min", 0)))
    stats = bd.get(actor)
    if isinstance(stats, dict):
        return int(_num(stats.get("lines" if mode == "line" else "words", 0)))
    # Dữ liệu rất cũ không có chi tiết từng người: theo câu thì lấy tổng câu như bản cũ; theo từ thì không đoán.
    if mode == "line" and not bd:
        return int(_num(item.get("total_lines", 0)))
    return 0

def units_text(units, mode):
    return f"{units:,} {MODES.get(mode, '')}"

def money(v):
    return f"{v:,.0f} VNĐ"  # team quen dùng dấu phẩy ngăn hàng nghìn: 30,000 VNĐ


# ---------- dự án & đơn giá ----------
def project_map(projects):
    return {str(p.get("project_id")): p for p in projects}

def project_name(projects_by_id, pid):
    p = projects_by_id.get(str(pid))
    return p.get("name", pid) if p else f"(dự án đã xoá: {pid})"

def inherited_rate(project, actor):
    """Đơn giá khi video KHÔNG có giá riêng: giá diễn viên × dự án, nếu không có thì giá mặc định dự án."""
    ar = (project or {}).get("actor_rates") or {}
    if actor in ar and ar[actor] not in (None, ""):
        return _num(ar[actor]), "Diễn viên × dự án"
    return _num((project or {}).get("unit_rate", 0)), "Mặc định dự án"

def effective_rate(item, actor, project):
    cr = item.get("custom_actor_rates") or {}
    if actor in cr and cr[actor] not in (None, ""):
        return _num(cr[actor]), "Riêng video này"
    return inherited_rate(project, actor)

def compute_pay(item, actor, project):
    mode = (project or {}).get("mode", "minute")
    units = actor_units(item, actor, mode)
    rate, source = effective_rate(item, actor, project)
    return {"units": units, "mode": mode, "rate": rate, "source": source, "pay": round(units * rate)}


# ---------- chốt sổ ----------
def pair_key(video_title, actor):
    return (str(video_title).strip().upper(), str(actor).strip().upper())

def active_locks(locks):
    return [l for l in locks if l.get("status", LOCKED) == LOCKED]

def locked_pairs(locks):
    """{(VIDEO, DIỄN VIÊN): đợt chốt} — đợt chốt sớm nhất thắng nếu (hiếm) bị trùng."""
    out = {}
    for l in sorted(active_locks(locks), key=lambda x: str(x.get("lock_id", ""))):  # mã chứa ngày giờ yyyymmddHHMMSS
        for r in l.get("rows") or []:
            out.setdefault(pair_key(r.get("video_title", ""), l.get("actor", "")), l)
    return out

def video_is_locked(item, lp):
    t = str(item.get("video_title", "")).strip().upper()
    return any(k[0] == t for k in lp)

def new_id(prefix):
    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S')}-{secrets.token_hex(2)}"


def period_rows(tracker, projects, locks, project_id, d_from, d_to):
    """Toàn bộ dòng lương (video × diễn viên) của 1 dự án trong khoảng ngày.
    Mỗi dòng: actor, video_title, date, units, mode, rate, source, pay, locked (bool), lock_id."""
    pmap = project_map(projects)
    project = pmap.get(str(project_id))
    lp = locked_pairs(locks)
    rows, seen_locked = [], set()
    # 1. Phần đã chốt: lấy số liệu lưu cứng
    for l in active_locks(locks):
        if str(l.get("project_id")) != str(project_id): continue
        for r in l.get("rows") or []:
            d = parse_date(r.get("date"))
            k = pair_key(r.get("video_title", ""), l.get("actor", ""))
            if d is None or not (d_from <= d <= d_to) or k in seen_locked or lp.get(k) is not l: continue
            seen_locked.add(k)
            rows.append({"actor": l.get("actor", ""), "video_title": r.get("video_title", ""), "date": d,
                         "units": int(_num(r.get("units"))), "mode": l.get("mode", r.get("mode", "minute")),
                         "rate": _num(r.get("rate")), "source": "Đã chốt", "pay": round(_num(r.get("pay"))),
                         "locked": True, "lock_id": l.get("lock_id", "")})
    # 2. Phần chưa chốt: tính theo đơn giá hiện tại
    for item in tracker:
        if str(item.get("project_id")) != str(project_id): continue
        d = parse_date(item.get("date"))
        if d is None or not (d_from <= d <= d_to): continue
        for a in actors_of(item):
            if pair_key(item.get("video_title", ""), a) in lp: continue
            c = compute_pay(item, a, project)
            rows.append({"actor": a, "video_title": item.get("video_title", ""), "date": d, **c,
                         "locked": False, "lock_id": ""})
    rows.sort(key=lambda r: (r["actor"], r["date"], str(r["video_title"]).upper()))
    return rows

def undated_videos(tracker, project_id):
    return [i for i in tracker if str(i.get("project_id")) == str(project_id) and parse_date(i.get("date")) is None]

def make_lock(project, actor, rows, d_from, d_to, username, now_text):
    """Tạo 1 đợt chốt sổ từ các dòng CHƯA chốt của 1 diễn viên."""
    live = [r for r in rows if r["actor"] == actor and not r["locked"]]
    if not live: return None
    return {
        "lock_id": new_id("CS"), "project_id": project.get("project_id"), "project_name": project.get("name", ""),
        "actor": actor, "from_date": fmt_date(d_from), "to_date": fmt_date(d_to),
        "mode": project.get("mode", "minute"), "total": sum(r["pay"] for r in live),
        "rows": [{"video_title": r["video_title"], "date": fmt_date(r["date"]), "units": r["units"],
                  "rate": r["rate"], "pay": r["pay"]} for r in live],
        "locked_by": username, "locked_at": now_text, "status": LOCKED, "unlocked_by": "", "unlocked_at": "", "note": "",
    }


# ---------- chuyển dữ liệu cũ ----------
def migrate(projects, tracker, old_rates, today_text):
    """Đưa dữ liệu kiểu cũ (1 cách tính chung cho mọi video) sang kiểu theo dự án. Chạy nhiều lần vẫn an toàn.
    Trả về (projects_changed, tracker_changed)."""
    p_changed = t_changed = False
    old_mode = str((old_rates or {}).get("mode", "minute")).lower()
    if old_mode not in MODES: old_mode = "minute"
    old_rate = _num((old_rates or {}).get("unit_rate", 30000)) or 30000
    ids = {str(p.get("project_id")) for p in projects}
    needs_default = not projects or any(str(i.get("project_id", "")).strip() not in ids for i in tracker)
    if needs_default and DEFAULT_PROJECT_ID not in ids:
        projects.append({"project_id": DEFAULT_PROJECT_ID, "name": DEFAULT_PROJECT_NAME, "mode": old_mode,
                         "unit_rate": old_rate, "actor_rates": {}, "active": "1", "created": today_text})
        ids.add(DEFAULT_PROJECT_ID); p_changed = True
    for item in tracker:
        if str(item.get("project_id", "")).strip() in ids: continue
        item["project_id"] = DEFAULT_PROJECT_ID
        # Bản cũ tự điền đơn giá mặc định vào "giá riêng" của mọi video -> bỏ những giá trùng giá mặc định cũ,
        # chỉ giữ những giá thật sự được sửa tay.
        cr = item.get("custom_actor_rates") or {}
        item["custom_actor_rates"] = {a: r for a, r in cr.items() if abs(_num(r) - old_rate) > 0.5}
        item.setdefault("stage", "Thu âm")
        item.setdefault("assignee", "")
        t_changed = True  # video không có ngày: KHÔNG tự gán hôm nay (sẽ lọt vào đợt lương hiện tại) — trang lương sẽ cảnh báo
    return p_changed, t_changed
