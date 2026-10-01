"""Soát xưng hô trong kịch bản tiếng Việt (ưu tiên khẩu ngữ & tiếng lóng miền Nam).

Cách làm:
- Nhận diện cả cụm nhiều chữ ("tại hạ", "mấy ní", "tụi bây"...), ưu tiên cụm dài trước để "người ta" không bị hiểu là "ta".
- Chỉ dùng ĐẠI TỪ THUẦN (tui/tao/mày/bồ...) để bắt lỗi, vì nghĩa của chúng rõ ràng.
- Từ thân tộc (anh/chị/em/ông/bà/con...) vừa dùng để xưng vừa để gọi, lại hay là danh từ thường
  ("con chó", "ông bà") -> chỉ thống kê cho người soát xem, KHÔNG tự báo lỗi.
- Nếu đã lưu "Bảng quan hệ xưng hô" (VD: TYLER xưng "tui", gọi "ông") -> câu nào trái quy tắc được báo đỏ.
"""
import re
from collections import Counter

# Ngôi thứ nhất (xưng) — đại từ thuần
SELF_TERMS = [
    "tôi", "tui", "tao", "tớ", "mình", "ta", "tại hạ", "bổn cô nương", "mị", "qua đây",
    "tao đây", "ta đây", "tui đây", "bố mày", "bà mày", "ông mày", "ông đây", "bà đây", "chị đây", "anh đây",
    "chúng tôi", "chúng ta", "chúng mình", "chúng tao", "chúng tớ", "tụi tui", "tụi mình", "tụi tao", "tụi tớ",
    "bọn tao", "bọn mình", "bọn tớ", "bọn tui", "tụi này", "bọn này", "tụi em", "bọn em",
]
# Ngôi thứ hai (gọi) — đại từ thuần & tiếng lóng
TARGET_TERMS = [
    "bạn", "cậu", "mày", "mầy", "mi", "bây", "bồ", "đằng ấy", "ní", "mấy ní", "cưng", "mấy cưng", "nhóc", "bro",
    "mấy bro", "các bạn", "tụi mày", "bọn mày", "tụi bây", "bọn bây", "chúng mày", "mấy người", "mấy đứa",
    "quý vị", "bà con", "cô nương", "huynh", "đệ", "sư huynh", "sư đệ", "sư tỷ", "sư muội", "tỷ", "muội",
    "cha nội", "mấy ông", "mấy bà", "mấy cha", "mấy má",
]
# Ngôi thứ ba kiểu miền Nam & thường gặp — chỉ để nhận diện, không bắt lỗi
THIRD_TERMS = ["ổng", "bả", "ẻm", "người ta", "tụi nó", "bọn nó", "mấy ổng", "mấy bả", "hắn", "nó", "họ", "thằng đó", "con nhỏ đó"]
# Từ thân tộc: dùng linh hoạt để xưng hoặc gọi -> chỉ thống kê
KINSHIP_TERMS = ["anh", "chị", "em", "ông", "bà", "cô", "chú", "bác", "dì", "thím", "mợ", "dượng", "con", "cháu",
                 "ba", "bố", "mẹ", "má", "cha", "tía", "ngoại", "anh hai", "chị hai", "út", "bé", "anh em", "chị em",
                 "ông nội", "bà nội", "ông ngoại", "bà ngoại"]
# Cụm có chứa đại từ nhưng KHÔNG phải xưng hô -> bỏ qua để tránh báo nhầm
IGNORE_TERMS = ["bạn bè", "bạn trai", "bạn gái", "bạn thân", "một mình", "mình mẩy", "ta-xi",
                "ông bà", "cha mẹ", "ba mẹ", "bố mẹ", "con cái", "con người", "cô dâu", "chú rể", "bà con xa", "em bé"]

CATEGORY = {}
for _cat, _terms in (("self", SELF_TERMS), ("target", TARGET_TERMS), ("third", THIRD_TERMS),
                     ("kin", KINSHIP_TERMS), ("ignore", IGNORE_TERMS)):
    for _t in _terms:
        CATEGORY.setdefault(_t, _cat)  # từ đã có ở nhóm trước thì giữ nhóm trước
_ALL = sorted(CATEGORY, key=len, reverse=True)
_PATTERN = re.compile(r"(?<!\w)(" + "|".join(re.escape(t) for t in _ALL) + r")(?!\w)", re.IGNORECASE)

def find_terms(text):
    """Trả về danh sách (từ, nhóm) tìm thấy trong câu, cụm dài được ưu tiên."""
    return [(m.group(1).lower(), CATEGORY[m.group(1).lower()]) for m in _PATTERN.finditer(text or "")]

def _rules_for(speaker, rules):
    """Từ bảng quan hệ đã lưu: các từ được phép xưng / gọi của 1 nhân vật."""
    allowed_self, allowed_target = set(), set()
    for key, val in (rules or {}).items():
        a = key.split("|")[0].strip().upper()
        if a == speaker.strip().upper() and isinstance(val, dict):
            if val.get("self"): allowed_self.add(str(val["self"]).strip().lower())
            if val.get("target"): allowed_target.add(str(val["target"]).strip().lower())
    return allowed_self, allowed_target

def analyze(lines, rules=None, min_uses=3, dominant_share=0.5, rare_share=0.25):
    """lines: danh sách dict {speaker, text, timecode}. Trả về (danh sách câu đã soát, bảng thống kê nhân vật)."""
    stats = {}
    found_per_line = []
    for ln in lines:
        terms = find_terms(ln["text"])
        found_per_line.append(terms)
        s = stats.setdefault(ln["speaker"], {"self": Counter(), "target": Counter(), "kin": Counter(), "lines": 0})
        s["lines"] += 1
        for w, cat in terms:
            if cat in ("self", "target", "kin"): s[cat][w] += 1

    results = []
    for ln, terms in zip(lines, found_per_line):
        spk = ln["speaker"]; s = stats[spk]
        allowed_self, allowed_target = _rules_for(spk, rules)
        issues, level = [], 0  # 0 = ổn, 1 = nghi vấn, 2 = trái quy tắc đã lưu
        for cat, allowed, verb in (("self", allowed_self, "xưng"), ("target", allowed_target, "gọi")):
            used = [w for w, c in terms if c == cat]
            if not used: continue
            # 1) Trái quy tắc đã lưu: dùng đại từ thuần khác với cách xưng/gọi đã quy định
            if allowed:
                bad = [w for w in used if w not in allowed]
                if bad:
                    issues.append(f"{verb} '{', '.join(sorted(set(bad)))}' — quy tắc đã lưu: {verb} '{', '.join(sorted(allowed))}'")
                    level = 2
                    continue
            # 2) Lệch thói quen của chính nhân vật trong kịch bản này
            total = sum(s[cat].values())
            if total >= min_uses:
                top, top_n = s[cat].most_common(1)[0]
                if top_n / total >= dominant_share:
                    rare = [w for w in used if w != top and s[cat][w] / total <= rare_share]
                    if rare:
                        issues.append(f"{verb} '{', '.join(sorted(set(rare)))}' — thường {verb} '{top}' ({top_n}/{total} lần)")
                        level = max(level, 1)
        status = {0: "🟢 Ổn", 1: "🟡 Nghi vấn lệch thói quen", 2: "🔴 Trái quy tắc đã lưu"}[level]
        results.append({**ln, "found": ", ".join(w for w, c in terms if c != "ignore"), "status": status,
                        "detail": "; ".join(issues)})

    summary = []
    for spk, s in stats.items():
        def top3(c): return ", ".join(f"{w} ({n})" for w, n in c.most_common(3)) or "—"
        summary.append({"Nhân vật": spk, "Số câu": s["lines"], "Xưng (đại từ)": top3(s["self"]),
                        "Gọi (đại từ)": top3(s["target"]), "Từ thân tộc hay dùng": top3(s["kin"])})
    return results, summary
