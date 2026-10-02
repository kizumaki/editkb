"""Linh vật: bé chibi lồng tiếng (tự vẽ, không dùng hình của ai).

- Có "bộ não" nhỏ chạy trong trình duyệt: tự chọn NGẪU NHIÊN hành động (đứng thở, ngó nghiêng, vẫy tay, hát vào micro,
  nhún nhảy, vươn vai, ngủ gật, đi vài bước, nhảy...) và nơi đứng: khoảng trống phía trên dải băng, mép dưới thanh menu
  (ngồi vắt vẻo đung đưa chân), phần trống bên phải dải băng xanh.
- Phần lớn thời gian đứng yên làm điệu bộ, ít đi lại (đỡ rối mắt). Rê chuột vào thì bé vui, nhảy lên và nói 1 câu.
- Chỉ hiện khi bật "Hiệu ứng chuyển động", màn hình rộng từ 900px. Tắt công tắc là bé biến mất.
- Kỹ thuật: 1 khung ẩn (components.html) gắn đoạn mã vào trang chính 1 lần duy nhất; mỗi lần trang chạy lại chỉ báo bật/tắt.
"""
import json
import streamlit as st
import streamlit.components.v1 as components

_SVG = """<svg viewBox="0 0 100 124" width="88" height="109" xmlns="http://www.w3.org/2000/svg" shape-rendering="geometricPrecision">
<defs>
 <radialGradient id="mhSkin" cx="42%" cy="36%" r="66%"><stop offset="0" stop-color="#FFEDE0"/><stop offset="1" stop-color="#F3BB97"/></radialGradient>
 <linearGradient id="mhShirt" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#4583E0"/><stop offset="1" stop-color="#1B4590"/></linearGradient>
 <radialGradient id="mhHair" cx="38%" cy="22%" r="85%"><stop offset="0" stop-color="#6A4636"/><stop offset="1" stop-color="#2A1812"/></radialGradient>
 <radialGradient id="mhCup" cx="35%" cy="30%" r="75%"><stop offset="0" stop-color="#7DAAF0"/><stop offset="1" stop-color="#1D4A98"/></radialGradient>
</defs>
<g class="mh-leg mh-leg-l"><rect x="38" y="95" width="10" height="20" rx="5" fill="#26365E" stroke="#1A2340" stroke-width="1.6"/><ellipse cx="42" cy="116.5" rx="7.5" ry="4.2" fill="#162039"/></g>
<g class="mh-leg mh-leg-r"><rect x="52" y="95" width="10" height="20" rx="5" fill="#26365E" stroke="#1A2340" stroke-width="1.6"/><ellipse cx="58" cy="116.5" rx="7.5" ry="4.2" fill="#162039"/></g>
<g class="mh-upper">
 <path d="M29 101 Q29 75 50 75 Q71 75 71 101 Z" fill="url(#mhShirt)" stroke="#1A2340" stroke-width="1.8"/>
 <path d="M50 95.5 C42.5 89.5 40.5 85 44 82.5 C46.5 80.8 49 82 50 84 C51 82 53.5 80.8 56 82.5 C59.5 85 57.5 89.5 50 95.5 Z" fill="#FF7E98" stroke="#fff" stroke-width="1.3"/>
 <g class="mh-head">
  <circle cx="50" cy="45" r="31" fill="url(#mhSkin)" stroke="#1A2340" stroke-width="1.8"/>
  <path d="M19 43 Q18 13 50 12 Q82 13 81 43 Q74 29 62 27 Q60 35 50 32 Q42 35 37 28 Q26 31 19 43 Z" fill="url(#mhHair)" stroke="#1A2340" stroke-width="1.4"/>
  <path d="M16.5 46 Q13.5 8.5 50 8 Q86.5 8.5 83.5 46" fill="none" stroke="#2A63BA" stroke-width="6" stroke-linecap="round"/>
  <ellipse cx="17.5" cy="49" rx="7.5" ry="10.5" fill="url(#mhCup)" stroke="#1A2340" stroke-width="1.6"/>
  <ellipse cx="82.5" cy="49" rx="7.5" ry="10.5" fill="url(#mhCup)" stroke="#1A2340" stroke-width="1.6"/>
  <g class="mh-eyes-open"><g class="mh-pupils">
   <ellipse cx="38" cy="50" rx="5.2" ry="6.8" fill="#23253A"/><ellipse cx="62" cy="50" rx="5.2" ry="6.8" fill="#23253A"/>
   <circle cx="36.3" cy="47.2" r="2.2" fill="#fff"/><circle cx="60.3" cy="47.2" r="2.2" fill="#fff"/>
   <circle cx="39.6" cy="53" r="1" fill="#fff" opacity=".8"/><circle cx="63.6" cy="53" r="1" fill="#fff" opacity=".8"/></g></g>
  <g class="mh-eyes-closed" fill="none" stroke="#23253A" stroke-width="2.4" stroke-linecap="round">
   <path d="M32.5 51 Q38 54.5 43.5 51"/><path d="M56.5 51 Q62 54.5 67.5 51"/></g>
  <g class="mh-eyes-happy" fill="none" stroke="#23253A" stroke-width="2.6" stroke-linecap="round">
   <path d="M32.5 52 Q38 45 43.5 52"/><path d="M56.5 52 Q62 45 67.5 52"/></g>
  <ellipse cx="29" cy="60" rx="5.5" ry="3.2" fill="#FF8FA3" opacity=".6"/><ellipse cx="71" cy="60" rx="5.5" ry="3.2" fill="#FF8FA3" opacity=".6"/>
  <path class="mh-mouth-smile" d="M44.5 62 Q50 67.5 55.5 62" fill="none" stroke="#7A3B2E" stroke-width="2.3" stroke-linecap="round"/>
  <ellipse class="mh-mouth-o" cx="50" cy="64" rx="3.4" ry="4.2" fill="#8E3A3A" stroke="#5A2020" stroke-width="1"/>
 </g>
 <g class="mh-arm mh-arm-l"><ellipse cx="30.5" cy="86" rx="5.5" ry="9" fill="url(#mhShirt)" stroke="#1A2340" stroke-width="1.6" transform="rotate(22 30.5 86)"/>
   <circle cx="27" cy="94" r="4.2" fill="url(#mhSkin)" stroke="#1A2340" stroke-width="1.4"/></g>
 <g class="mh-arm mh-arm-r"><ellipse cx="69.5" cy="85" rx="5.5" ry="9" fill="url(#mhShirt)" stroke="#1A2340" stroke-width="1.6" transform="rotate(-32 69.5 85)"/>
   <rect x="72" y="65" width="6.5" height="17" rx="3.2" fill="#3A3F4B" stroke="#1A2340" stroke-width="1.3"/>
   <circle cx="75.2" cy="62.5" r="6.6" fill="#5C6273" stroke="#1A2340" stroke-width="1.4"/><circle cx="73.4" cy="60.8" r="2.1" fill="#B9C0D0"/>
   <circle cx="75" cy="80" r="4.2" fill="url(#mhSkin)" stroke="#1A2340" stroke-width="1.4"/></g>
</g>
</svg>"""

_CSS = """
#mh-mascot { position: fixed; left: 0; top: 0; width: 88px; height: 109px; z-index: 999991; pointer-events: none;
             opacity: 0; transition: opacity .5s ease; will-change: transform; }
#mh-mascot.on { opacity: 1; }
#mh-mascot .mh-svgwrap { position: absolute; inset: 0; pointer-events: auto; cursor: pointer; }
#mh-mascot svg { display: block; overflow: visible; width: 88px; height: 109px; }
#mh-mascot.small, #mh-mascot.small .mh-svgwrap { width: 40px; height: 50px; }
#mh-mascot.small svg { width: 40px; height: 50px; }
#mh-mascot.small .mh-bubble { bottom: 48px; } #mh-mascot.small .mh-fx { transform: scale(.6); transform-origin: top left; }
#mh-mascot.small .mh-shadow { left: 10px; width: 20px; height: 4px; bottom: -1px; }
#mh-mascot.flip svg { transform: scaleX(-1); }
#mh-mascot .mh-shadow { position: absolute; left: 22px; bottom: -3px; width: 44px; height: 7px; border-radius: 50%;
                        background: rgba(15,23,42,.18); transition: transform .3s, opacity .3s; }
#mh-mascot.pose-jump .mh-shadow { opacity: 0; }
#mh-mascot .mh-bubble { position: absolute; bottom: 104px; left: 50%; transform: translateX(-50%) scale(.6); transform-origin: bottom center;
             white-space: nowrap; background: #fff; color: #183E80; font: 700 12px/1.3 'Source Sans Pro', Arial, sans-serif;
             padding: 4px 10px; border-radius: 12px; box-shadow: 0 6px 16px -6px rgba(24,62,128,.5); opacity: 0;
             transition: opacity .25s, transform .25s cubic-bezier(.3,1.5,.5,1); }
#mh-mascot .mh-bubble::after { content: ""; position: absolute; left: 50%; bottom: -5px; margin-left: -5px;
             border: 5px solid transparent; border-top-color: #fff; border-bottom: 0; }
#mh-mascot .mh-bubble.show { opacity: 1; transform: translateX(-50%) scale(1); }
#mh-mascot .mh-fx { position: absolute; left: 0; top: 0; width: 88px; height: 40px; pointer-events: none; }
#mh-mascot .mh-fx span { position: absolute; font-weight: 800; animation: mhFloat 1.8s ease-out forwards; }
@keyframes mhFloat { from { transform: translate(0, 0) rotate(0); opacity: 0; } 15% { opacity: 1; }
                     to { transform: translate(var(--dx), -38px) rotate(var(--rot)); opacity: 0; } }

#mh-mascot .mh-eyes-closed, #mh-mascot .mh-eyes-happy, #mh-mascot .mh-mouth-o { display: none; }
#mh-mascot .mh-leg, #mh-mascot .mh-upper, #mh-mascot .mh-arm, #mh-mascot .mh-head, #mh-mascot .mh-pupils, #mh-mascot .mh-eyes-open {
  transform-box: view-box; transition: transform .35s ease; }
#mh-mascot .mh-leg-l { transform-origin: 43px 96px; } #mh-mascot .mh-leg-r { transform-origin: 57px 96px; }
#mh-mascot .mh-arm-l { transform-origin: 34px 79px; } #mh-mascot .mh-arm-r { transform-origin: 66px 79px; }
#mh-mascot .mh-head { transform-origin: 50px 74px; } #mh-mascot .mh-upper { transform-origin: 50px 100px; }
#mh-mascot .mh-eyes-open { transform-origin: 50px 50px; }
#mh-mascot.blink .mh-eyes-open { transform: scaleY(.08); transition: transform .06s; }

/* đứng thở */
#mh-mascot.pose-idle .mh-upper { animation: mhBreath 2.6s ease-in-out infinite; }
@keyframes mhBreath { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(1.2px); } }
/* đi */
#mh-mascot.pose-walk .mh-leg-l { animation: mhLeg .5s ease-in-out infinite alternate; }
#mh-mascot.pose-walk .mh-leg-r { animation: mhLeg .5s ease-in-out infinite alternate-reverse; }
#mh-mascot.pose-walk .mh-upper { animation: mhBob .25s ease-in-out infinite alternate; }
#mh-mascot.pose-walk .mh-arm-l { animation: mhSwing .5s ease-in-out infinite alternate; }
@keyframes mhLeg { from { transform: rotate(16deg); } to { transform: rotate(-16deg); } }
@keyframes mhBob { from { transform: translateY(0); } to { transform: translateY(-2px); } }
@keyframes mhSwing { from { transform: rotate(-14deg); } to { transform: rotate(14deg); } }
/* ngó nghiêng */
#mh-mascot.pose-look .mh-head { animation: mhLook 2.8s ease-in-out; }
#mh-mascot.pose-look .mh-pupils { animation: mhPupil 2.8s ease-in-out; }
@keyframes mhLook { 0%, 100% { transform: rotate(0); } 25%, 40% { transform: rotate(-7deg); } 60%, 80% { transform: rotate(7deg); } }
@keyframes mhPupil { 0%, 100% { transform: translateX(0); } 25%, 40% { transform: translateX(-2.5px); } 60%, 80% { transform: translateX(2.5px); } }
/* vẫy tay */
#mh-mascot.pose-wave .mh-arm-l { animation: mhWave .38s ease-in-out infinite alternate; }
@keyframes mhWave { from { transform: rotate(118deg); } to { transform: rotate(158deg); } }
#mh-mascot.pose-wave .mh-head { transform: rotate(-5deg); }
/* hát */
#mh-mascot.pose-sing .mh-arm-r { transform: rotate(-28deg) translate(-4px, -2px); }
#mh-mascot.pose-sing .mh-mouth-smile { display: none; } #mh-mascot.pose-sing .mh-mouth-o { display: inline; }
#mh-mascot.pose-sing .mh-upper { animation: mhSway 1.1s ease-in-out infinite alternate; }
#mh-mascot.pose-sing .mh-eyes-open { display: none; } #mh-mascot.pose-sing .mh-eyes-happy { display: inline; }
@keyframes mhSway { from { transform: rotate(-4deg); } to { transform: rotate(4deg); } }
/* nhún nhảy */
#mh-mascot.pose-dance .mh-upper { animation: mhDance .5s ease-in-out infinite alternate; }
#mh-mascot.pose-dance .mh-arm-l { animation: mhArmUpL .5s ease-in-out infinite alternate; }
#mh-mascot.pose-dance .mh-leg-l { animation: mhLeg .5s ease-in-out infinite alternate; }
#mh-mascot.pose-dance .mh-eyes-open { display: none; } #mh-mascot.pose-dance .mh-eyes-happy { display: inline; }
@keyframes mhDance { from { transform: rotate(-8deg) translateY(0); } to { transform: rotate(8deg) translateY(-3px); } }
@keyframes mhArmUpL { from { transform: rotate(60deg); } to { transform: rotate(150deg); } }
/* vươn vai */
#mh-mascot.pose-stretch .mh-arm-l { transform: rotate(160deg); }
#mh-mascot.pose-stretch .mh-arm-r { transform: rotate(-150deg); }
#mh-mascot.pose-stretch .mh-upper { transform: translateY(-3px); }
#mh-mascot.pose-stretch .mh-eyes-open { display: none; } #mh-mascot.pose-stretch .mh-eyes-closed { display: inline; }
#mh-mascot.pose-stretch .mh-mouth-smile { display: none; } #mh-mascot.pose-stretch .mh-mouth-o { display: inline; }
/* ngủ gật */
#mh-mascot.pose-sleep .mh-eyes-open { display: none; } #mh-mascot.pose-sleep .mh-eyes-closed { display: inline; }
#mh-mascot.pose-sleep .mh-head { animation: mhNod 2.4s ease-in-out infinite; }
@keyframes mhNod { 0%, 100% { transform: rotate(10deg); } 50% { transform: rotate(16deg) translateY(1.5px); } }
/* ngồi bệt duỗi chân, lắc lư nhẹ */
#mh-mascot.pose-sit .mh-leg { animation: mhSitLeg 1.6s ease-in-out infinite alternate; }
#mh-mascot.pose-sit .mh-leg-r { animation-delay: -.8s; }
#mh-mascot.pose-sit .mh-upper { animation: mhSway 2.2s ease-in-out infinite alternate; }
@keyframes mhSitLeg { from { transform: rotate(-84deg); } to { transform: rotate(-72deg); } }
/* nhảy */
#mh-mascot.pose-jump .mh-leg-l { transform: rotate(28deg); } #mh-mascot.pose-jump .mh-leg-r { transform: rotate(-22deg); }
#mh-mascot.pose-jump .mh-arm-l { transform: rotate(120deg); } #mh-mascot.pose-jump .mh-arm-r { transform: rotate(-60deg); }
/* vui (khi rê chuột) */
#mh-mascot.pose-happy .mh-eyes-open { display: none; } #mh-mascot.pose-happy .mh-eyes-happy { display: inline; }
#mh-mascot.pose-happy .mh-arm-l { transform: rotate(140deg); }
@media (prefers-reduced-motion: reduce) { #mh-mascot { display: none !important; } }
"""

_ENGINE = r"""
(function () {
  if (window.__mhEngine) return; window.__mhEngine = true;
  const st = document.createElement('style'); st.id = 'mh-mascot-style'; st.textContent = __CSS__; document.head.appendChild(st);
  const el = document.createElement('div'); el.id = 'mh-mascot';
  el.innerHTML = '<div class="mh-bubble"></div><div class="mh-fx"></div><div class="mh-shadow"></div><div class="mh-svgwrap">' + __SVG__ + '</div>';
  document.body.appendChild(el);
  const bubble = el.querySelector('.mh-bubble'), fx = el.querySelector('.mh-fx'), hit = el.querySelector('.mh-svgwrap');
  const W = 88, FEET = 106, SIT = 89, SMALL = 40 / 88;
  const PHRASES = ['Lồng tiếng thôi! 🎙️', 'Cố lên team ơi! 💪', 'Nghỉ tay uống nước nha 🧋', 'Deadline nhớ chưa? ⏰',
                   'Hôm nay thu âm gì nè?', 'Mai Han số 1! 💙', 'Nhớ lưu file nha 💾', 'Hí hí 😆', 'Ai gọi tui đó? 👀'];
  const S = { plat: 'gap', u: Math.random() * 0.8 + 0.1, facing: 1, pose: 'idle', jump: null, walk: null, poke: false, shown: false };
  const rnd = (a, b) => a + Math.random() * (b - a);
  const pick = (opts) => { let t = opts.reduce((s, o) => s + o[1], 0) * Math.random(); for (const o of opts) { if ((t -= o[1]) <= 0) return o[0]; } return opts[0][0]; };
  const sleep = (ms) => new Promise((res) => { const t0 = performance.now(); (function w() { if (S.poke || performance.now() - t0 >= ms) res(); else setTimeout(w, 80); })(); });

  function plats() {
    const hd = document.querySelector('header[data-testid="stHeader"]'), hero = document.querySelector('.hero-container');
    if (!hd || !hero) return null;
    const h = hd.getBoundingClientRect(), r = hero.getBoundingClientRect(), vw = window.innerWidth;
    if (r.width < 300 || r.top < h.bottom - 4) return null;               // dải băng đã cuộn khuất dưới thanh menu
    const P = {};
    if (r.top - h.bottom > 75) P.gap = { y: r.top, x0: r.left + r.width * 0.40, x1: r.right - 50, sit: false };
    const navs = document.querySelectorAll('[data-testid="stTopNavSection"]');
    const navRight = navs.length ? navs[navs.length - 1].getBoundingClientRect().right : vw * 0.4;
    P.ledge = { y: h.bottom - 4, x0: navRight + 40, x1: vw - 300, small: true };
    P.banner = { y: r.bottom - 5, x0: r.left + r.width * 0.64, x1: r.right - 55, sit: false };
    if (P.ledge.x1 < P.ledge.x0 + 80) delete P.ledge;
    return P;
  }
  const visible = () => window.__mhEnabled !== false && window.innerWidth >= 900;
  function setPose(p) { S.pose = p; const sm = el.classList.contains('small'); el.className = 'pose-' + p + (S.facing < 0 ? ' flip' : '') + (S.shown ? ' on' : '') + (sm ? ' small' : ''); }
  function say(text, ms) { bubble.textContent = text; bubble.classList.add('show'); setTimeout(() => bubble.classList.remove('show'), ms || 2600); }
  function spawn(chars, color) {
    const s = document.createElement('span'); s.textContent = chars[Math.floor(Math.random() * chars.length)];
    s.style.left = rnd(48, 66) + 'px'; s.style.top = '18px'; s.style.color = color; s.style.fontSize = rnd(13, 18) + 'px';
    s.style.setProperty('--dx', rnd(-14, 18) + 'px'); s.style.setProperty('--rot', rnd(-30, 30) + 'deg');
    fx.appendChild(s); setTimeout(() => s.remove(), 1900);
  }

  function frame(now) {
    const P = plats();
    const show = visible() && P && P[S.plat];
    if (!show) { el.classList.remove('on'); S.shown = false; requestAnimationFrame(frame); return; }
    if (!S.shown) { S.shown = true; el.classList.add('on'); }
    let x, y, anchor = S.pose === 'sit' ? SIT : FEET, small = !!P[S.plat].small;
    const pl = P[S.plat];
    if (S.jump) {
      const j = S.jump, tp = P[j.to] || pl, k = Math.min(1, (now - j.t0) / j.T);
      const tx = tp.x0 + j.u * (tp.x1 - tp.x0), ty = tp.y;
      x = j.x + (tx - j.x) * k; y = j.y + (ty - j.y) * k - Math.sin(Math.PI * k) * j.h; anchor = FEET;
      small = k < 0.5 ? !!pl.small : !!tp.small;
      if (k >= 1) { S.plat = j.to; S.u = j.u; S.jump = null; j.done(); }
    } else {
      if (S.walk) { const w = S.walk, k = Math.min(1, (now - w.t0) / w.T); S.u = w.u0 + (w.u1 - w.u0) * k; if (k >= 1) { S.walk = null; w.done(); } }
      S.u = Math.max(0, Math.min(1, S.u));
      x = pl.x0 + S.u * (pl.x1 - pl.x0); y = pl.y;
    }
    S.lastX = x; S.lastY = y;
    const sc = small ? SMALL : 1;
    if (small !== el.classList.contains('small')) el.classList.toggle('small', small);
    el.style.transform = 'translate3d(' + Math.round(x - W * sc / 2) + 'px,' + Math.round(y - anchor * sc) + 'px,0)';
    requestAnimationFrame(frame);
  }

  function walkTo(u1) {
    const P = plats(); if (!P || !P[S.plat]) return sleep(500);
    const pl = P[S.plat], dist = Math.abs(u1 - S.u) * (pl.x1 - pl.x0);
    S.facing = u1 >= S.u ? 1 : -1; setPose('walk');
    return new Promise((res) => { S.walk = { u0: S.u, u1, t0: performance.now(), T: Math.max(400, dist / 38 * 1000), done: res }; });
  }
  function jumpTo(to, u, h, T) {
    const P = plats(); if (!P || !P[to]) return sleep(300);
    const tp = P[to], tx = tp.x0 + u * (tp.x1 - tp.x0);
    S.facing = tx >= (S.lastX || tx) ? 1 : -1; setPose('jump');
    return new Promise((res) => { S.jump = { to, u, x: S.lastX, y: S.lastY, h, T, t0: performance.now(), done: res }; });
  }

  async function act() {
    const P = plats();
    if (!P || !P[S.plat]) { if (P) { S.plat = P.gap ? 'gap' : 'banner'; } await sleep(600); return; }
    if (S.poke) {
      S.poke = false; setPose('happy'); say(PHRASES[Math.floor(Math.random() * PHRASES.length)], 2400);
      if (S.plat !== 'ledge') await jumpTo(S.plat, S.u, 16, 520);
      setPose('happy'); await sleep(1400); return;
    }
    if (S.plat === 'ledge') {
      const a = pick([['idle', 24], ['look', 14], ['wave', 12], ['sit', 16], ['walk', 18], ['leave', 16]]);
      if (a === 'leave') { const to = P.gap && Math.random() < 0.6 ? 'gap' : 'banner'; await jumpTo(to, rnd(0.1, 0.9), 26, 820); setPose('idle'); await sleep(rnd(600, 1200)); return; }
      if (a !== 'walk') { setPose(a); if (a === 'wave' && Math.random() < 0.4) say(PHRASES[Math.floor(Math.random() * PHRASES.length)]);
                          await sleep(a === 'sit' ? rnd(3500, 7000) : a === 'look' ? 2900 : a === 'wave' ? 2200 : rnd(2000, 4500)); return; }
    }
    const a = pick([['idle', 30], ['look', 15], ['wave', 8], ['sing', 8], ['dance', 5], ['stretch', 6], ['sleep', 4],
                    ['walk', 12], ['hop', 4], ['sit', 5], ['climb', P.ledge ? 6 : 0], ['switch', 5]]);
    if (S.plat === 'ledge') { const pl = P.ledge, maxU = Math.min(1, 140 / Math.max(pl.x1 - pl.x0, 1));
                              let u1 = S.u + (Math.random() < 0.5 ? -1 : 1) * rnd(0.3, 1) * maxU; await walkTo(Math.max(0, Math.min(1, u1))); setPose('idle'); await sleep(rnd(500, 1200)); }
    else if (a === 'idle') { setPose('idle'); await sleep(rnd(2500, 6500)); }
    else if (a === 'sit') { setPose('sit'); await sleep(rnd(4000, 8000)); setPose('idle'); await sleep(700); }
    else if (a === 'look') { setPose('look'); await sleep(2900); setPose('idle'); await sleep(rnd(500, 1500)); }
    else if (a === 'wave') { setPose('wave'); if (Math.random() < 0.45) say(PHRASES[Math.floor(Math.random() * PHRASES.length)]); await sleep(2200); }
    else if (a === 'sing') { setPose('sing'); const t = setInterval(() => spawn(['♪', '♫', '♬'], '#2A63BA'), 520); await sleep(rnd(3200, 4800)); clearInterval(t); }
    else if (a === 'dance') { setPose('dance'); const t = setInterval(() => spawn(['♪', '✦'], '#FF7E98'), 700); await sleep(rnd(2400, 3600)); clearInterval(t); }
    else if (a === 'stretch') { setPose('stretch'); await sleep(1900); setPose('idle'); await sleep(800); }
    else if (a === 'sleep') { setPose('sleep'); const t = setInterval(() => spawn(['z', 'Z', 'z'], '#7C8DB5'), 1100); await sleep(rnd(5500, 9000)); clearInterval(t);
                              setPose('happy'); say('Ơ... tui đâu có ngủ 😳', 2200); await sleep(1500); }
    else if (a === 'walk') { const pl = P[S.plat], span = pl.x1 - pl.x0, maxU = Math.min(1, 170 / Math.max(span, 1));
                             let u1 = S.u + (Math.random() < 0.5 ? -1 : 1) * rnd(0.25, 1) * maxU; u1 = Math.max(0, Math.min(1, u1));
                             await walkTo(u1); setPose('idle'); await sleep(rnd(400, 1200)); }
    else if (a === 'hop') { await jumpTo(S.plat, S.u, rnd(10, 16), 480); setPose('idle'); await sleep(600); }
    else if (a === 'climb') { await jumpTo('ledge', rnd(0.1, 0.9), 40, 900); setPose('happy'); await sleep(900); setPose('idle'); await sleep(rnd(1200, 2500)); }
    else if (a === 'switch') { const to = S.plat === 'gap' ? 'banner' : (P.gap ? 'gap' : 'banner'); if (to !== S.plat) { await jumpTo(to, rnd(0.1, 0.9), 34, 820); }
                               setPose('idle'); await sleep(rnd(800, 1600)); }
  }
  async function brain() { for (;;) { try { await act(); } catch (e) { await sleep(1000); } } }
  (function blinkLoop() { setTimeout(() => { if (!['sleep', 'stretch', 'sing', 'dance', 'happy'].includes(S.pose)) {
      el.classList.add('blink'); setTimeout(() => el.classList.remove('blink'), 140); } blinkLoop(); }, rnd(2200, 5200)); })();
  hit.addEventListener('mouseenter', () => { S.poke = true; });
  setPose('idle'); requestAnimationFrame(frame); brain();
})();
"""


def render_mascot(enabled):
    """Gọi 1 lần mỗi lần trang chạy (ở bất kỳ đâu). Lần đầu gắn bé vào trang; các lần sau chỉ báo bật/tắt."""
    engine = _ENGINE.replace("__CSS__", json.dumps(_CSS)).replace("__SVG__", json.dumps(_SVG))
    html = ("<script>(function(){try{var P=window.parent;P.__mhEnabled=" + ("true" if enabled else "false") + ";"
            "if(P.__mhEngine)return;var s=P.document.createElement('script');s.textContent=" + json.dumps(engine) + ";"
            "P.document.head.appendChild(s);}catch(e){}})();</script>")
    with st.container(key="mh_mascot_host"):
        components.html(html, height=0)
