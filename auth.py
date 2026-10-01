import streamlit as st
import hashlib
import hmac
import re
import secrets
import time
import pandas as pd
from utils import ACCOUNTS_DB_FILE, save_json_db, sheets_enabled, log_event, read_audit_log

# ==========================================
# QUYỀN: mỗi trang chức năng + 2 quyền phụ ở thanh bên
# ==========================================
PAGES = {
    "goc": "Xử lý kịch bản gốc",
    "resync": "Re-Sync bản đã biên tập",
    "doi_chieu": "Đối chiếu 2 file tiếng Anh",
    "soat_xung_ho": "Soát xưng hô & thuật ngữ",
    "luong": "Theo dõi & báo cáo lương",
    "phan_vai": "Phân vai & màu nhân vật",
    "phien_am": "Kho phiên âm giọng Nam",
    "don_phu_de": "Dọn dẹp phụ đề",
    "chuyen_doi": "Chuyển đổi định dạng",
}
EXTRA_PERMS = {
    "kho_tu": "Sửa kho từ dùng chung (thanh bên)",
    "sao_luu": "Tải bản sao lưu dữ liệu",
}
ALL_PERMS = {**PAGES, **EXTRA_PERMS}

ROLES = {
    "quan_tri": ("Quản trị", list(ALL_PERMS)),
    "ke_toan": ("Kế toán / Quản lý dự án", ["luong", "sao_luu"]),
    "bien_tap": ("Biên tập viên", ["goc", "resync", "doi_chieu", "soat_xung_ho", "phan_vai", "phien_am",
                                   "don_phu_de", "chuyen_doi", "kho_tu"]),
    "ctv_qc": ("Cộng tác viên QC", ["doi_chieu"]),
}

# ==========================================
# MẬT KHẨU: lưu dạng mã hoá 1 chiều (PBKDF2), không ai đọc ngược ra được
# ==========================================
def _hash_password(password, salt_hex=None):
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return salt.hex(), digest.hex()

def _check_password(password, salt_hex, hash_hex):
    if not salt_hex or not hash_hex: return False
    return hmac.compare_digest(_hash_password(password, salt_hex)[1], hash_hex)

def _app_secret():
    try: return str(st.secrets.get("APP_PASSWORD", ""))
    except Exception: return ""

def _accounts():
    return st.session_state.setdefault("accounts", [])

def _find(username):
    u = (username or "").strip().lower()
    return next((a for a in _accounts() if str(a.get("username", "")).lower() == u), None)

def _validate_new_password(pw, pw2):
    if len(pw) < 8: return "Mật khẩu phải có ít nhất 8 ký tự."
    if pw != pw2: return "Hai lần nhập mật khẩu không khớp."
    return ""

def _set_password(account, pw):
    account["salt"], account["password_hash"] = _hash_password(pw)

# ==========================================
# NGƯỜI ĐANG ĐĂNG NHẬP
# ==========================================
def current_user():
    return st.session_state.get("_user")

def is_admin():
    u = current_user()
    return bool(u) and u.get("role") == "quan_tri"

def has_perm(perm):
    u = current_user()
    if not u: return False
    if u.get("role") == "quan_tri": return True
    return perm in (u.get("pages") or [])

def _start_session(account):
    st.session_state["_user"] = {
        "username": account["username"], "display_name": account.get("display_name") or account["username"],
        "role": account.get("role", "ctv_qc"), "pages": list(account.get("pages") or []),
    }
    log_event("Đăng nhập")

# ==========================================
# MÀN HÌNH ĐĂNG NHẬP (gọi ở đầu app; nếu chưa đăng nhập thì dừng app tại đây)
# ==========================================
def require_login(logo_path):
    if current_user(): return
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.image(logo_path, width=240)
        if not _accounts():
            _first_setup()
        else:
            _login_form()
    st.stop()

def _first_setup():
    st.markdown("#### 🛠️ Thiết lập lần đầu: tạo tài khoản Quản trị")
    secret = _app_secret()
    if sheets_enabled() and not secret:
        st.error("Cần đặt `APP_PASSWORD` trong mục Secrets của Streamlit trước, để chỉ chủ app mới tạo được tài khoản Quản trị.")
        return
    with st.form("first_admin"):
        if secret:
            st.caption("Nhập mật khẩu chủ app (APP_PASSWORD trong Secrets) để xác nhận bạn là chủ app.")
            owner_pw = st.text_input("Mật khẩu chủ app", type="password")
        username = st.text_input("Tên đăng nhập cho bạn (viết liền, không dấu, VD: trung)")
        display = st.text_input("Tên hiển thị (VD: Anh Trung)")
        pw = st.text_input("Mật khẩu mới (ít nhất 8 ký tự)", type="password")
        pw2 = st.text_input("Nhập lại mật khẩu mới", type="password")
        if st.form_submit_button("Tạo tài khoản Quản trị", type="primary"):
            if secret and not hmac.compare_digest(owner_pw.encode(), secret.encode()):
                time.sleep(1); st.error("Mật khẩu chủ app không đúng."); return
            err = _validate_username(username) or _validate_new_password(pw, pw2)
            if err: st.error(err); return
            acc = {"username": username.strip().lower(), "display_name": display.strip() or username.strip(),
                   "role": "quan_tri", "pages": ROLES["quan_tri"][1], "active": "1"}
            _set_password(acc, pw)
            _accounts().append(acc)
            if save_json_db(ACCOUNTS_DB_FILE, _accounts()):
                _start_session(acc); st.rerun()

def _login_form():
    st.markdown("#### 🔒 Đăng nhập ScriptPro Studio")
    with st.form("login_form"):
        username = st.text_input("Tên đăng nhập")
        pw = st.text_input("Mật khẩu", type="password")
        if st.form_submit_button("Đăng nhập", type="primary", use_container_width=True):
            acc = _find(username)
            if acc and str(acc.get("active", "1")) != "0" and _check_password(pw, acc.get("salt"), acc.get("password_hash")):
                _start_session(acc); st.rerun()
            else:
                time.sleep(1.5)  # làm chậm việc dò mật khẩu
                st.error("Sai tên đăng nhập hoặc mật khẩu, hoặc tài khoản đã bị khoá.")
    with st.expander("Quên mật khẩu?"):
        st.caption("Nhân viên: nhờ Quản trị đặt lại mật khẩu. "
                   "Quản trị quên mật khẩu: dùng mật khẩu chủ app (APP_PASSWORD trong Secrets) để đặt lại.")
        secret = _app_secret()
        if secret:
            with st.form("owner_reset"):
                owner_pw = st.text_input("Mật khẩu chủ app", type="password")
                username = st.text_input("Tên đăng nhập cần đặt lại")
                pw = st.text_input("Mật khẩu mới (ít nhất 8 ký tự)", type="password")
                pw2 = st.text_input("Nhập lại mật khẩu mới", type="password")
                if st.form_submit_button("Đặt lại mật khẩu"):
                    if not hmac.compare_digest(owner_pw.encode(), secret.encode()):
                        time.sleep(1.5); st.error("Mật khẩu chủ app không đúng."); return
                    acc = _find(username)
                    err = ("Không có tài khoản này." if not acc else "") or _validate_new_password(pw, pw2)
                    if err: st.error(err); return
                    _set_password(acc, pw); acc["active"] = "1"
                    if save_json_db(ACCOUNTS_DB_FILE, _accounts()):
                        log_event("Chủ app đặt lại mật khẩu", acc["username"])
                        st.success("Đã đặt lại. Bạn có thể đăng nhập bằng mật khẩu mới.")

def _validate_username(username):
    u = (username or "").strip().lower()
    if not re.fullmatch(r"[a-z0-9._-]{3,30}", u):
        return "Tên đăng nhập: 3–30 ký tự, chỉ gồm chữ thường không dấu, số, dấu chấm, gạch dưới, gạch ngang."
    if _find(u): return "Tên đăng nhập này đã có người dùng."
    return ""

# ==========================================
# Ô TÀI KHOẢN Ở THANH BÊN
# ==========================================
def render_user_box():
    u = current_user()
    role_name = ROLES.get(u["role"], ("?",))[0]
    st.sidebar.markdown(f"👤 **{u['display_name']}**  \n<span style='font-size:0.85rem;opacity:0.75'>{role_name}</span>",
                        unsafe_allow_html=True)
    c1, c2 = st.sidebar.columns(2)
    if c2.button("Đăng xuất", use_container_width=True, key="btn_logout"):
        log_event("Đăng xuất")
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()
    if c1.button("Đổi mật khẩu", use_container_width=True, key="btn_show_pw"):
        st.session_state["_show_pw_form"] = not st.session_state.get("_show_pw_form", False)
    if st.session_state.get("_show_pw_form"):
        with st.sidebar.form("change_pw"):
            old = st.text_input("Mật khẩu hiện tại", type="password")
            pw = st.text_input("Mật khẩu mới (ít nhất 8 ký tự)", type="password")
            pw2 = st.text_input("Nhập lại mật khẩu mới", type="password")
            if st.form_submit_button("Lưu mật khẩu mới", type="primary"):
                acc = _find(u["username"])
                if not acc or not _check_password(old, acc.get("salt"), acc.get("password_hash")):
                    time.sleep(1); st.error("Mật khẩu hiện tại không đúng."); return
                err = _validate_new_password(pw, pw2)
                if err: st.error(err); return
                _set_password(acc, pw)
                if save_json_db(ACCOUNTS_DB_FILE, _accounts()):
                    log_event("Đổi mật khẩu")
                    st.session_state["_show_pw_form"] = False
                    st.success("Đã đổi mật khẩu.")

# ==========================================
# TRANG QUẢN TRỊ: TÀI KHOẢN & QUYỀN
# ==========================================
def _perm_label(p): return ALL_PERMS.get(p, p)

def render_accounts_page():
    if not is_admin():
        st.error("Chỉ Quản trị mới vào được trang này."); return
    st.subheader("Tài khoản & quyền")
    st.caption("Thay đổi quyền có hiệu lực từ **lần đăng nhập sau** của người đó.")

    rows = [{"Tên đăng nhập": a["username"], "Tên hiển thị": a.get("display_name", ""),
             "Vai trò": ROLES.get(a.get("role"), ("?",))[0],
             "Được dùng": "Tất cả" if a.get("role") == "quan_tri" else ", ".join(_perm_label(p) for p in a.get("pages") or []),
             "Trạng thái": "🟢 Hoạt động" if str(a.get("active", "1")) != "0" else "🔴 Đã khoá"} for a in _accounts()]
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    t_add, t_edit, t_log = st.tabs(["➕ Thêm tài khoản", "✏️ Sửa / khoá / đặt lại mật khẩu", "📜 Nhật ký hoạt động"])

    with t_add:
        role = st.selectbox("Vai trò", list(ROLES), format_func=lambda r: ROLES[r][0], key="add_role")
        with st.form("add_account", clear_on_submit=True):
            c1, c2 = st.columns(2)
            username = c1.text_input("Tên đăng nhập (viết liền, không dấu)")
            display = c2.text_input("Tên hiển thị")
            pages = ROLES[role][1]
            if role != "quan_tri":
                pages = st.multiselect("Được dùng (đã chọn sẵn theo vai trò, có thể thêm/bớt)", list(ALL_PERMS),
                                       default=ROLES[role][1], format_func=_perm_label)
            pw = st.text_input("Mật khẩu ban đầu (ít nhất 8 ký tự, gửi riêng cho người đó)", type="password")
            pw2 = st.text_input("Nhập lại mật khẩu", type="password")
            if st.form_submit_button("Tạo tài khoản", type="primary"):
                err = _validate_username(username) or _validate_new_password(pw, pw2)
                if err: st.error(err)
                else:
                    acc = {"username": username.strip().lower(), "display_name": display.strip() or username.strip(),
                           "role": role, "pages": list(pages), "active": "1"}
                    _set_password(acc, pw)
                    _accounts().append(acc)
                    if save_json_db(ACCOUNTS_DB_FILE, _accounts()):
                        log_event("Tạo tài khoản", f"{acc['username']} ({ROLES[role][0]})")
                        st.success(f"Đã tạo tài khoản **{acc['username']}**."); time.sleep(1); st.rerun()

    with t_edit:
        if not _accounts(): st.info("Chưa có tài khoản."); return
        uname = st.selectbox("Chọn tài khoản", [a["username"] for a in _accounts()], key="edit_user",
                             format_func=lambda n: f"{n} — {(_find(n) or {}).get('display_name', '')}")
        acc = _find(uname)
        is_self = uname == current_user()["username"]
        admins_active = [a for a in _accounts() if a.get("role") == "quan_tri" and str(a.get("active", "1")) != "0"]
        role = st.selectbox("Vai trò", list(ROLES), index=list(ROLES).index(acc.get("role", "ctv_qc")),
                            format_func=lambda r: ROLES[r][0], key=f"edit_role_{uname}")
        with st.form(f"edit_account_{uname}"):
            display = st.text_input("Tên hiển thị", value=acc.get("display_name", ""))
            pages = ROLES["quan_tri"][1]
            if role != "quan_tri":
                current = acc.get("pages") if role == acc.get("role") else ROLES[role][1]
                pages = st.multiselect("Được dùng", list(ALL_PERMS), default=[p for p in current if p in ALL_PERMS],
                                       format_func=_perm_label)
            active = st.checkbox("Cho phép đăng nhập (bỏ tích = khoá tài khoản)", value=str(acc.get("active", "1")) != "0")
            new_pw = st.text_input("Đặt lại mật khẩu (để trống nếu không đổi)", type="password")
            if st.form_submit_button("Lưu thay đổi", type="primary"):
                losing_admin = acc.get("role") == "quan_tri" and (role != "quan_tri" or not active)
                if losing_admin and len(admins_active) <= 1:
                    st.error("Không thể: đây là Quản trị cuối cùng. Hãy tạo thêm 1 Quản trị khác trước."); return
                if is_self and losing_admin:
                    st.error("Bạn không thể tự gỡ quyền Quản trị hoặc tự khoá chính mình."); return
                if new_pw and len(new_pw) < 8:
                    st.error("Mật khẩu mới phải có ít nhất 8 ký tự."); return
                acc.update({"display_name": display.strip() or uname, "role": role, "pages": list(pages),
                            "active": "1" if active else "0"})
                if new_pw: _set_password(acc, new_pw)
                if save_json_db(ACCOUNTS_DB_FILE, _accounts()):
                    log_event("Sửa tài khoản", f"{uname}: {ROLES[role][0]}, {'hoạt động' if active else 'khoá'}"
                                               + (", đặt lại mật khẩu" if new_pw else ""))
                    st.success("Đã lưu."); time.sleep(1); st.rerun()

    with t_log:
        if st.button("📜 Xem 300 hoạt động gần nhất", key="btn_load_log"):
            try:
                logs = read_audit_log(300)
                st.dataframe(pd.DataFrame(logs, columns=["Thời gian", "Người dùng", "Hành động", "Chi tiết"]),
                             hide_index=True, use_container_width=True)
            except Exception as e:
                st.error(f"Chưa đọc được nhật ký: {e}")
