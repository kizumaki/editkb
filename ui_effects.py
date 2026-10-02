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

# ---------- Thanh menu trên cùng: kính mờ + mục dạng viên bo tròn, nhóm đang mở nền xanh ----------
_NAV_CSS = f"""
header[data-testid="stHeader"] {{
    background: var(--mh-nav-bg) !important; backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid rgba({BRAND}, 0.14); box-shadow: 0 8px 28px -22px rgba({BRAND}, 0.9); }}
[data-testid="stTopNavSection"] {{
    padding: 6px 14px !important; border-radius: 999px; gap: 4px;
    transition: background-color 0.2s ease, color 0.2s ease, box-shadow 0.2s ease; }}
[data-testid="stTopNavSection"] p {{ font-weight: 600; font-size: 0.92rem; letter-spacing: 0.01em; }}
[data-testid="stTopNavSection"]:hover {{ background: rgba({BRAND}, 0.10); color: var(--mh-nav-accent); }}
[aria-expanded="true"] > [data-testid="stTopNavSection"] {{ background: rgba({BRAND}, 0.14); color: var(--mh-nav-accent); }}
/* Nhóm chứa trang đang mở: viên nền xanh đậm, chữ trắng */
[data-testid="stToolbar"] div[style*="order"]:has(a[aria-current="page"]) [data-testid="stTopNavSection"] {{
    background: linear-gradient(135deg, #183E80, #2A63BA); color: #fff !important;
    box-shadow: 0 6px 16px -8px rgba({BRAND}, 0.9); }}
[data-testid="stToolbar"] div[style*="order"]:has(a[aria-current="page"]) [data-testid="stTopNavSection"] p,
[data-testid="stToolbar"] div[style*="order"]:has(a[aria-current="page"]) [data-testid="stTopNavSection"] svg {{ color: #fff !important; }}
/* Danh sách thả xuống */
[data-baseweb="popover"]:has([data-testid="stTopNavPopover"]) > div {{
    border-radius: 14px !important; overflow: hidden; border: 1px solid rgba({BRAND}, 0.14);
    box-shadow: 0 22px 44px -18px rgba(15, 23, 42, 0.45) !important; }}
[data-testid="stTopNavPopover"] {{ padding: 6px !important; }}
a[data-testid="stTopNavDropdownLink"] {{ border-radius: 10px; padding: 8px 12px !important; margin: 1px 0;
    transition: background-color 0.15s ease, transform 0.15s ease; }}
a[data-testid="stTopNavDropdownLink"]:hover {{ background: rgba({BRAND}, 0.09) !important; }}
a[data-testid="stTopNavDropdownLink"][aria-current="page"] {{ background: rgba({BRAND}, 0.13) !important;
    box-shadow: inset 3px 0 0 #2A63BA; }}
a[data-testid="stTopNavDropdownLink"][aria-current="page"] p {{ font-weight: 700; color: var(--mh-nav-accent); }}
"""
_NAV_MOTION_CSS = """
@keyframes mhDrop { from { opacity: 0; transform: translateY(-6px) scale(0.98); } to { opacity: 1; transform: none; } }
[data-baseweb="popover"] [data-testid="stTopNavPopover"] { animation: mhDrop 0.18s cubic-bezier(.2,.7,.2,1) both; transform-origin: top left; }
a[data-testid="stTopNavDropdownLink"]:hover { transform: translateX(3px); }
"""

# ---------- Linh vật: bé chibi lồng tiếng (tự vẽ, không dùng hình của ai) ----------
# Đi qua lại trong khoảng trống giữa thanh menu và dải băng, thỉnh thoảng nhảy / vẫy micro, chớp mắt;
# rê chuột vào thì nhảy cẫng lên. Chỉ hiện khi bật "Hiệu ứng chuyển động"; ẩn trên màn hình nhỏ.
_MASCOT_SVG = """
<svg viewBox="0 0 100 124" width="74" height="92" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
  <defs>
    <radialGradient id="mhSkin" cx="42%" cy="38%" r="65%"><stop offset="0" stop-color="#FFEBDD"/><stop offset="1" stop-color="#F4BE9C"/></radialGradient>
    <linearGradient id="mhShirt" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#3B78D6"/><stop offset="1" stop-color="#183E80"/></linearGradient>
    <radialGradient id="mhHair" cx="40%" cy="25%" r="80%"><stop offset="0" stop-color="#5A3B2E"/><stop offset="1" stop-color="#2C1A14"/></radialGradient>
    <radialGradient id="mhCup" cx="35%" cy="30%" r="75%"><stop offset="0" stop-color="#6FA0EC"/><stop offset="1" stop-color="#1F4C9A"/></radialGradient>
  </defs>
  <g class="mh-leg mh-leg-l"><rect x="38" y="96" width="9" height="20" rx="4.5" fill="#22325A"/><ellipse cx="42" cy="117" rx="7" ry="4" fill="#14203D"/></g>
  <g class="mh-leg mh-leg-r"><rect x="53" y="96" width="9" height="20" rx="4.5" fill="#22325A"/><ellipse cx="58" cy="117" rx="7" ry="4" fill="#14203D"/></g>
  <path d="M30 100 Q30 76 50 76 Q70 76 70 100 Z" fill="url(#mhShirt)"/>
  <path d="M50 96 C44 91 42 87.5 45 85.5 C47 84 49 85 50 86.6 C51 85 53 84 55 85.5 C58 87.5 56 91 50 96 Z" fill="#FF8FA3"/>
  <g class="mh-arm"><ellipse cx="70" cy="84" rx="5" ry="8" fill="url(#mhShirt)" transform="rotate(-35 70 84)"/>
    <rect x="72" y="66" width="6" height="16" rx="3" fill="#3A3F4B"/><circle cx="75" cy="64" r="6" fill="#565C6B"/>
    <circle cx="73.5" cy="62.5" r="2" fill="#9AA1B2"/></g>
  <ellipse cx="31" cy="86" rx="5" ry="8" fill="url(#mhShirt)" transform="rotate(25 31 86)"/>
  <circle cx="50" cy="46" r="31" fill="url(#mhSkin)"/>
  <path d="M19 44 Q18 14 50 13 Q82 14 81 44 Q74 30 62 28 Q60 36 50 33 Q42 36 37 29 Q26 32 19 44 Z" fill="url(#mhHair)"/>
  <path d="M17 46 Q14 10 50 9 Q86 10 83 46" fill="none" stroke="#2A63BA" stroke-width="5.5" stroke-linecap="round"/>
  <ellipse cx="18" cy="50" rx="7" ry="10" fill="url(#mhCup)"/><ellipse cx="82" cy="50" rx="7" ry="10" fill="url(#mhCup)"/>
  <g class="mh-eyes">
    <ellipse cx="38" cy="51" rx="5" ry="6.5" fill="#262838"/><ellipse cx="62" cy="51" rx="5" ry="6.5" fill="#262838"/>
    <circle cx="36.5" cy="48.5" r="2" fill="#fff"/><circle cx="60.5" cy="48.5" r="2" fill="#fff"/>
  </g>
  <ellipse cx="30" cy="61" rx="5" ry="3" fill="#FF8FA3" opacity=".55"/><ellipse cx="70" cy="61" rx="5" ry="3" fill="#FF8FA3" opacity=".55"/>
  <path d="M45 63 Q50 68 55 63" fill="none" stroke="#7A3B2E" stroke-width="2.2" stroke-linecap="round"/>
</svg>
"""

_MASCOT_CSS = """
.mh-lane { position: relative; height: 0; pointer-events: none; }
.mh-lane-inner { position: absolute; left: 45%; right: 8%; bottom: 6px; height: 100px; }
.mh-walker { position: absolute; bottom: 0; left: 0; width: 74px; animation: mhWalk 26s linear infinite; pointer-events: auto; }
.mh-flip { animation: mhFlip 26s steps(1) infinite; transform-origin: center; }
.mh-hop { animation: mhAutoHop 7.3s ease-in-out infinite; }
.mh-walker:hover .mh-hop { animation: mhHop 0.6s cubic-bezier(.3,1.6,.5,1) 1; }
.mh-bob { animation: mhBob 0.45s ease-in-out infinite alternate; }
.mh-shadow { position: absolute; left: 15px; bottom: -2px; width: 44px; height: 7px; border-radius: 50%;
             background: rgba(15,23,42,.18); animation: mhShadow 7.3s ease-in-out infinite; }
.mh-leg-l { animation: mhLeg 0.45s ease-in-out infinite alternate; transform-origin: 42px 96px; }
.mh-leg-r { animation: mhLeg 0.45s ease-in-out infinite alternate-reverse; transform-origin: 58px 96px; }
.mh-arm { animation: mhWave 26s ease-in-out infinite; transform-origin: 66px 84px; }
.mh-eyes { animation: mhBlink 4.2s infinite; transform-origin: 50px 51px; }
.mh-bubble { position: absolute; bottom: 94px; left: 50%; transform: translateX(-50%); white-space: nowrap;
             background: #fff; color: #183E80; font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 10px;
             box-shadow: 0 4px 12px -4px rgba(24,62,128,.45); opacity: 0; animation: mhBubble 26s ease-in-out infinite; }
@keyframes mhWalk { 0% { left: 0; } 44% { left: calc(100% - 74px); } 52% { left: calc(100% - 74px); }
                    96% { left: 0; } 100% { left: 0; } }
@keyframes mhFlip { 0% { transform: scaleX(1); } 48% { transform: scaleX(-1); } 100% { transform: scaleX(1); } }
@keyframes mhBob { from { transform: translateY(0); } to { transform: translateY(-2px); } }
@keyframes mhLeg { from { transform: rotate(14deg); } to { transform: rotate(-14deg); } }
@keyframes mhAutoHop { 0%, 82%, 100% { transform: translateY(0); } 88% { transform: translateY(-14px); } 93% { transform: translateY(0); }
                       96% { transform: translateY(-5px); } }
@keyframes mhHop { 0% { transform: translateY(0) scale(1); } 40% { transform: translateY(-16px) scale(1.06) rotate(-6deg); }
                   100% { transform: translateY(0) scale(1); } }
@keyframes mhShadow { 0%, 82%, 100% { transform: scaleX(1); opacity: 1; } 88% { transform: scaleX(.6); opacity: .5; } }
@keyframes mhWave { 0%, 44%, 52%, 96%, 100% { transform: rotate(0); } 46%, 50% { transform: rotate(-22deg); } 48% { transform: rotate(8deg); }
                    97%, 99% { transform: rotate(-22deg); } 98% { transform: rotate(8deg); } }
@keyframes mhBlink { 0%, 92%, 100% { transform: scaleY(1); } 95% { transform: scaleY(.1); } }
@keyframes mhBubble { 0%, 44%, 53%, 100% { opacity: 0; } 46%, 51% { opacity: 1; } }
@media (max-width: 900px) { .mh-lane { display: none; } }
"""

def render_mascot(enabled):
    """Gọi ngay TRƯỚC dải băng đầu trang. Không chiếm chỗ (cao 0), nhân vật nổi lên khoảng trống phía trên."""
    if not enabled: return
    st.markdown(f"<style>{_MASCOT_CSS}</style><div class='mh-lane'><div class='mh-lane-inner'><div class='mh-walker'>"
                f"<div class='mh-bubble'>Lồng tiếng thôi! 🎙️</div><div class='mh-shadow'></div>"
                f"<div class='mh-hop'><div class='mh-flip'><div class='mh-bob'>{_MASCOT_SVG}</div></div></div>"
                f"</div></div></div>", unsafe_allow_html=True)


def inject_effects(enabled, spinner_logo_b64, dark=False):
    nav_vars = (":root { --mh-nav-bg: rgba(14, 17, 23, 0.72); --mh-nav-accent: #8DB4F0; }" if dark
                else ":root { --mh-nav-bg: rgba(255, 255, 255, 0.74); --mh-nav-accent: #2A63BA; }")
    css = _BASE_CSS + nav_vars + _NAV_CSS + (_NAV_MOTION_CSS if enabled else "")
    if enabled:
        css += f":root {{ --mh-logo: url('data:image/png;base64,{spinner_logo_b64}'); }}\n" + _MOTION_CSS
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
