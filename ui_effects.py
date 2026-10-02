"""Hiệu ứng giao diện nhẹ cho ScriptPro (chỉ dùng CSS, không làm chậm app).

- Tắt được bằng công tắc "Hiệu ứng chuyển động" ở thanh bên.
- Máy nào bật chế độ "giảm chuyển động" của hệ điều hành thì tự tắt hiệu ứng.
- Nếu Streamlit cập nhật làm đổi cấu trúc trang, hiệu ứng chỉ mất đi, app vẫn chạy bình thường.
"""
import streamlit as st

BRAND = "42, 99, 186"  # màu xanh logo Mai Han (RGB)

# Phần luôn bật: chỉ là đổ bóng / bo góc / màu khi rê chuột, không có chuyển động
_BASE_CSS = f"""
html {{ scroll-behavior: smooth; }}
.metric-card {{ border-radius: 14px; border-top: 3px solid rgba({BRAND}, 0.85) !important; }}
[data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"] {{ border-radius: 14px; }}
[data-testid="stFileUploaderDropzone"] {{ border: 1.5px dashed rgba({BRAND}, 0.35); border-radius: 12px; }}
[data-testid="stFileUploaderDropzone"]:hover {{ border-color: rgba({BRAND}, 0.9); background: rgba({BRAND}, 0.06); }}
[data-testid="stExpander"] details {{ border-radius: 12px; }}

/* ---------- Thanh bên: gọn, hiện đại ---------- */
section[data-testid="stSidebar"] {{
    background-image: linear-gradient(180deg, rgba({BRAND}, 0.09) 0%, rgba({BRAND}, 0.02) 35%, transparent 70%); }}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {{ gap: 0.55rem; }}
/* Thẻ người dùng */
.mh-user {{ display: flex; align-items: center; gap: 12px; padding: 12px 14px; border-radius: 16px; margin-bottom: 0.5rem !important;
            background: rgba({BRAND}, 0.08); border: 1px solid rgba({BRAND}, 0.18); }}
.mh-avatar {{ width: 42px; height: 42px; border-radius: 50%; flex-shrink: 0; display: flex; align-items: center; justify-content: center;
              background: linear-gradient(135deg, #183E80, #2A63BA); color: #fff; font-weight: 700; font-size: 0.95rem;
              box-shadow: 0 6px 14px -6px rgba({BRAND}, 0.8); }}
.mh-user-text {{ min-width: 0; }}
section[data-testid="stSidebar"] .stMarkdownContainer:has(.mh-user),
section[data-testid="stSidebar"] .stMarkdownContainer:has(.mh-side-gap) {{ margin-bottom: 0 !important; }}
section[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.mh-user) {{ padding-bottom: 0.15rem; }}
.mh-user-name {{ font-weight: 700; font-size: 1rem; line-height: 1.25; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.mh-user-role {{ display: inline-block; margin-top: 4px; padding: 1px 9px; border-radius: 999px; font-size: 0.72rem; font-weight: 600;
                 background: rgba({BRAND}, 0.16); }}
/* Nút ở thanh bên: thấp, chữ nhỏ, không xuống dòng */
section[data-testid="stSidebar"] .stButton > button,
section[data-testid="stSidebar"] .stDownloadButton > button,
section[data-testid="stSidebar"] [data-testid="stPopover"] button {{
    min-height: 36px; padding: 4px 10px; border-radius: 10px; }}
section[data-testid="stSidebar"] .stButton > button p,
section[data-testid="stSidebar"] .stDownloadButton > button p,
section[data-testid="stSidebar"] [data-testid="stPopover"] button p {{ font-size: 0.86rem; white-space: nowrap; }}
/* Trạng thái dữ liệu: 1 dòng nhỏ có chấm màu */
.mh-status {{ display: flex; align-items: center; gap: 8px; font-size: 0.78rem; opacity: 0.8; padding: 0 4px; }}
.mh-dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; background: #22C55E; box-shadow: 0 0 0 3px rgba(34, 197, 94, 0.2); }}
.mh-status.local .mh-dot {{ background: #F59E0B; box-shadow: 0 0 0 3px rgba(245, 158, 11, 0.2); }}
/* Khoảng trống nhỏ giữa các nhóm (không chữ, không vạch kẻ) */
.mh-side-gap {{ height: 0.35rem; }}
/* Ô mở rộng ở thanh bên */
section[data-testid="stSidebar"] [data-testid="stExpander"] details {{
    border-color: rgba({BRAND}, 0.18); background: rgba({BRAND}, 0.03); }}
section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {{ font-size: 0.88rem; }}
"""

# Phần chuyển động: tắt được
_MOTION_CSS = f"""
@keyframes mhFadeUp {{ from {{ opacity: 0; transform: translateY(12px); }} to {{ opacity: 1; transform: none; }} }}
@keyframes mhFlow {{ 0% {{ background-position: 0% 50%; }} 50% {{ background-position: 100% 50%; }} 100% {{ background-position: 0% 50%; }} }}
@keyframes mhShine {{ 0% {{ transform: translateX(-150%) skewX(-20deg); }} 55%, 100% {{ transform: translateX(450%) skewX(-20deg); }} }}
@keyframes mhPulse {{ 0%, 100% {{ transform: scale(0.92); opacity: 0.75; }} 50% {{ transform: scale(1.06); opacity: 1; }} }}

/* Khối nội dung MỚI xuất hiện (mở / chuyển trang, kết quả vừa xử lý) hiện ra mờ dần.
   Khối đã có sẵn thì Streamlit giữ nguyên nên không bị nhấp nháy khi bấm nút. */
[data-testid="stMainBlockContainer"] [data-testid="stElementContainer"] {{ animation: mhFadeUp 0.35s cubic-bezier(.2,.7,.2,1) both; }}

/* Dải băng đầu trang: màu xanh chảy chậm + vệt sáng lướt qua */
.hero-container {{ position: relative; overflow: hidden; background-size: 220% 220% !important;
                   animation: mhFlow 16s ease-in-out infinite; }}
.hero-container::after {{ content: ""; position: absolute; inset: 0 auto 0 0; width: 22%;
                          background: linear-gradient(90deg, transparent, rgba(255,255,255,0.13), transparent);
                          animation: mhShine 8s ease-in-out 1.5s infinite; pointer-events: none; }}
.hero-logo {{ transition: transform 0.35s cubic-bezier(.2,.7,.2,1); }}
.hero-container:hover .hero-logo {{ transform: scale(1.04) rotate(-1deg); }}

/* Thẻ thống kê nhấc nhẹ lên khi rê chuột */
.metric-card {{ transition: transform 0.25s cubic-bezier(.2,.7,.2,1), box-shadow 0.25s; }}
.metric-card:hover {{ transform: translateY(-3px); box-shadow: 0 14px 28px -14px rgba({BRAND}, 0.45); }}

/* Khung viền: đổ bóng mềm khi rê chuột (không nhấc lên, để khỏi chao đảo khi đang làm việc) */
[data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"] {{ transition: box-shadow 0.3s, border-color 0.3s; }}
[data-testid="stLayoutWrapper"] > [data-testid="stVerticalBlock"]:hover {{ box-shadow: 0 10px 30px -18px rgba({BRAND}, 0.5); }}

/* Nút bấm: nhấc nhẹ + đổ bóng khi rê chuột, lún xuống khi bấm */
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {{
    transition: transform 0.15s ease, box-shadow 0.2s ease, filter 0.2s ease; }}
.stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {{
    transform: translateY(-1px); box-shadow: 0 8px 18px -10px rgba({BRAND}, 0.7); }}
.stButton > button:active, .stDownloadButton > button:active, [data-testid="stFormSubmitButton"] > button:active {{
    transform: translateY(0) scale(0.98); }}

/* Biểu tượng chờ: logo Mai Han "thở" thay cho vòng xoay mặc định */
[data-testid="stSpinner"] i, [data-testid="stSpinner"] svg {{ display: none !important; }}
[data-testid="stSpinner"] > div::before {{ content: ""; display: inline-block; width: 56px; height: 26px;
    margin-right: 10px; vertical-align: middle; background: var(--mh-logo) center / contain no-repeat;
    animation: mhPulse 1.1s ease-in-out infinite; }}

/* Tôn trọng cài đặt "giảm chuyển động" của máy người dùng */
@media (prefers-reduced-motion: reduce) {{
    *, *::before, *::after {{ animation: none !important; transition: none !important; }}
}}
"""

def inject_effects(enabled, spinner_logo_b64):
    css = _BASE_CSS
    if enabled:
        css += f":root {{ --mh-logo: url('data:image/png;base64,{spinner_logo_b64}'); }}\n" + _MOTION_CSS
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
