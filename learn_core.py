"""Học từ kịch bản team đã biên tập (không dùng AI, chỉ đếm & quy tắc). Mọi kết quả là GỢI Ý, người duyệt mới lưu.

Nguồn:
- File Word đã biên tập: phần "VAI:" (nhân vật: DIỄN VIÊN), các dòng "Nhân vật: lời thoại", cặp phiên âm trong ngoặc
  (team viết cả 2 chiều: "Stiv (Steve)" và "Jimmy (chim-mi)").
- Phụ đề gốc tiếng Anh (.srt): các cụm "XXX:" đầu dòng KHÔNG phải nhân vật -> gợi ý "không phải tên nhân vật".
"""
import difflib
import re
import unicodedata
from collections import Counter, defaultdict

TIMECODE = re.compile(r"^\s*(?:\d{2}:)?\d{2}:\d{2}[,.]\d{3}\s*-->")
VN_MARK = re.compile(r"[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]", re.IGNORECASE)
ENG_OK = re.compile(r"^[A-Za-z][A-Za-z0-9'’&.\- ]*[A-Za-z0-9]$")
PAREN = re.compile(r"\(([^()]{2,40})\)")
SPEAKER_LINE = re.compile(r"^\s*([^:\t/()\"“]{1,40}?)\s*:\s*(.*)$", re.S)
NOTE_MARKS = {"c", "k", "nl", "s", "-", "x", "cười", "khóc", "hát", "thì thầm"}
ACTOR_NAME = re.compile(r"^[A-ZÀ-Ỹ][A-ZÀ-Ỹ.]{1,14}(?: [A-ZÀ-Ỹ][A-ZÀ-Ỹ.]{1,14}){0,2}$")
# Chữ chức năng tiếng Anh: cụm "XXX:" chứa các chữ này gần như chắc chắn KHÔNG phải tên nhân vật ("I was like:", "But remember:")
FUNCTION_WORDS = {"I", "I'M", "IM", "YOU", "HE", "SHE", "IT", "WE", "THEY", "HE'S", "SHE'S", "IT'S", "THEY'RE", "WE'RE",
                  "IS", "ARE", "WAS", "WERE", "BE", "THE", "A", "AN", "AND", "BUT", "OR", "SO", "LIKE", "SAID", "SAYS", "SAY",
                  "THAT", "THIS", "THERE", "HERE", "WHAT", "WHICH", "WHO", "WHEN", "IF", "OF", "TO", "IN", "ON", "FOR", "WITH",
                  "REMEMBER", "UPDATE", "NOTE", "WARNING", "STEP", "ROUND", "NUMBER", "PART", "LEVEL", "DAY", "QUESTION",
                  "ALL", "IS:", "MY", "YOUR", "OUR", "THEIR", "HIS", "HER", "JUST", "ONLY", "NOT", "NEXT", "FIRST", "LAST"}

def is_actor_name(s):
    s = s.strip().upper()
    return bool(ACTOR_NAME.match(s)) and not s.endswith(".")


def strip_marks(s):
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

def has_vn(s): return bool(VN_MARK.search(s))

def _pho_score(s, vn_syllables):
    sy = [x for x in re.split(r"[-\s]+", s.lower()) if x]
    sc = 2 if has_vn(s) else 0
    if "-" in s: sc += 1
    if sy and all(strip_marks(x) in vn_syllables for x in sy): sc += 1
    return sc

def _similar(pho, eng):
    a = re.sub(r"[^a-z]", "", strip_marks(pho).lower())
    b = re.sub(r"[^a-z]", "", eng.lower())
    return difflib.SequenceMatcher(None, a, b).ratio() if a and b else 0

def _clean_tok(t): return t.strip(".,!?;:\"“”'‘’…*/")

def _is_name(s):
    s = s.strip()
    return 0 < len(s.split()) <= 4 and s[0].isalpha() and not TIMECODE.match(s) and "-->" not in s


# ---------- quy chuẩn phiên âm: viết hoa chữ đầu mỗi từ, gạch nối liền ("chim-mi" -> "Chim-mi", "pít- tông" -> "Pít-tông") ----------
def normalize_pho(s, proper=None):
    """proper=True: tên riêng -> hoa chữ đầu mỗi từ ("Chim-mi"); False: danh từ thường -> viết thường ("lốc");
    None: giữ kiểu hoa/thường của chữ đầu mỗi từ như đã lưu, chỉ sửa phần còn lại ("ÍT-PI" -> "Ít-pi")."""
    s = unicodedata.normalize("NFC", str(s or "")).strip()
    s = re.sub(r"\s*-\s*", "-", s)          # bỏ dấu cách quanh gạch nối
    s = re.sub(r"-{2,}", "-", s).strip("-")
    s = re.sub(r"\s+", " ", s)
    if proper is True: return " ".join(w[:1].upper() + w[1:].lower() for w in s.split(" ") if w)
    if proper is False: return s.lower()
    return " ".join(w[:1] + w[1:].lower() for w in s.split(" ") if w)

def _cap_first(s): return s[:1].upper() + s[1:]

def norm_eng_key(s):
    """Khoá tiếng Anh thống nhất: IN HOA, dấu nháy thẳng, 1 dấu cách ("McDonald’s " -> "MCDONALD'S")."""
    s = unicodedata.normalize("NFC", str(s or "")).replace("’", "'").replace("‘", "'").replace("`", "'")
    return re.sub(r"\s+", " ", s).strip().upper()

def pho_missing(eng, pho):
    """Chưa có phiên âm thật (trống, hoặc chỉ chép lại NGUYÊN chữ tiếng Anh: PIZZA -> "Pizza") -> không chèn vào kịch bản.
    KHÔNG bỏ dấu tiếng Việt / gạch nối khi so: "A-ni-mê (Anime)", "Ca-mê-ra (Camera)" là phiên âm thật
    (bản cũ bỏ dấu + gạch nối nên coi chúng là "chưa có" -> ẩn khỏi bảng kho và không chèn vào kịch bản)."""
    def core(s): return norm_eng_key(s).replace("'", "").replace(" ", "")
    p = core(pho)
    return not p or p == core(eng)

def near_keys(eng, keys):
    """Từ gần giống đã có trong kho: khác nhau chỉ ở số nhiều / 's / dấu cách / gạch nối."""
    def core(k): return re.sub(r"('S|S)$", "", re.sub(r"[\s\-'.]", "", k))
    c = core(eng)
    return [k for k in keys if k != eng and core(k) == c]

def norm_pair_key(a, b):
    def n(x): return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(x or ""))).strip().rstrip(":").strip().upper()
    return f"{n(a)}|{n(b)}" if n(a) and n(b) else ""

def norm_term(t): return re.sub(r"\s+", " ", unicodedata.normalize("NFC", str(t or ""))).strip().lower()

def merge_pronoun_rels(kho, items, known_terms=()):
    """items: [(người nói, người nghe, xưng, gọi)]. Giống phiên âm: mới -> thêm; trùng -> bỏ qua; khác -> KHÔNG ghi đè.
    Trả về added / same / conflicts [(khoá, cũ, mới)] / odd [(khoá, từ lạ)]."""
    out = {"added": [], "same": [], "conflicts": [], "odd": []}
    known = {norm_term(t) for t in known_terms}
    for a, b, s, t in items:
        k = norm_pair_key(a, b)
        new = {"self": norm_term(s), "target": norm_term(t)}
        if not k or not (new["self"] or new["target"]): continue
        if known:
            out["odd"] += [(k, w) for w in (new["self"], new["target"]) if w and w not in known]
        cur = kho.get(k)
        if not isinstance(cur, dict): kho[k] = new; out["added"].append((k, new))
        elif {"self": norm_term(cur.get("self")), "target": norm_term(cur.get("target"))} == new: out["same"].append((k, new))
        else: out["conflicts"].append((k, cur, new))
    return out

def merge_phonetics(kho, items):
    """items: [(english, phiên âm đã chuẩn hoá)]. Thêm từ mới vào kho NGAY; từ trùng hệt bỏ qua;
    từ có rồi mà khác -> KHÔNG ghi đè, trả về để người duyệt chọn. Trả về dict added/same/conflicts/near."""
    out = {"added": [], "same": [], "conflicts": [], "near": []}
    for eng, pho in items:
        k = norm_eng_key(eng)
        if not k or not pho: continue
        cur = kho.get(k)
        if cur is None or pho_missing(k, cur):
            for n in near_keys(k, kho): out["near"].append((k, n, kho[n]))
            kho[k] = pho; out["added"].append((k, pho))
        elif cur == pho: out["same"].append((k, pho))
        else: out["conflicts"].append((k, cur, pho))
    return out

def _sentence_start(before):
    b = before.rstrip()
    return not b or b[-1] in ".!?/…:\t" or bool(re.search(r"^\s*[^\s:]{1,30}:\s*$", b))

def standardize_pairs(text, phonetic_db):
    """Cặp "phiên âm (English)" đã có trong câu: nếu kho có từ đó -> sửa phiên âm theo kho; nếu không -> chỉ chuẩn hoá cách viết."""
    def repl(m):
        eng = m.group(2).strip()
        before = m.group(1)
        known = phonetic_db.get(eng.upper())
        toks = before.rstrip().split(" ")
        if known:
            std = normalize_pho(known)
            n = len(std.split())
            old = toks[-n:]
            # Chỉ thay khi phần cũ trông như phiên âm (có gạch nối hoặc gần giống), tránh nuốt chữ thường ("chơi (Minecraft)")
            if not (any("-" in t for t in old) or _similar(" ".join(old), std) >= 0.5): return m.group(0)
            keep = toks[:-n]
        elif toks and "-" in toks[-1] and not ENG_OK.match(toks[-1]):  # chưa có trong kho: chỉ sửa cách viết
            std, keep = normalize_pho(toks[-1]), toks[:-1]
        else:
            return m.group(0)
        if _sentence_start(" ".join(keep)): std = _cap_first(std)  # đầu câu thì viết hoa như chính tả
        return (" ".join(keep) + " " if keep else "") + f"{std} ({eng})"
    return re.sub(r"([^()]*?\S)\s*\(([A-Za-z][A-Za-z0-9'’&.\- ]{0,38}[A-Za-z0-9])\)", repl, text)


# ---------- thống nhất cách ghi: "phiên âm (English)" ----------
_REVERSED = re.compile(r"((?:[A-Za-z][A-Za-z0-9'’&.]*\s+){0,2}[A-Za-z][A-Za-z0-9'’&.]*)\s*\(([^()]{2,40})\)")

def flip_reversed_pairs(text, phonetic_db=None):
    """Đổi "Jimmy (chim-mi)" thành "chim-mi (Jimmy)" (chuẩn của team). Chỉ đổi khi chắc chắn bên trong ngoặc là phiên âm."""
    phonetic_db = phonetic_db or {}
    def repl(m):
        words, pho = m.group(1).split(), m.group(2).strip()
        if pho.lower() in NOTE_MARKS or ":" in pho or not (has_vn(pho) or "-" in pho): return m.group(0)
        if ENG_OK.match(pho) and "-" not in pho: return m.group(0)
        best, best_k = 0, 0
        n_syl = len([x for x in re.split(r"[-\s]+", pho) if x])
        for k in range(1, len(words) + 1):  # chọn số chữ tiếng Anh phía trước khớp phiên âm nhất
            eng = " ".join(words[-k:])
            known = phonetic_db.get(eng.upper())
            if known and strip_marks(known).lower() == strip_marks(pho).lower(): s = 1.0
            elif eng.isupper() and eng.isalpha() and 2 <= len(eng) <= 5 and n_syl == len(eng): s = 0.9  # chữ viết tắt: XP -> ít-pi
            else: s = _similar(pho, eng)
            # Cụm viết hoa đầu chữ ("The End") được ưu tiên lấy trọn nếu điểm chỉ kém chút ít
            if s > best or (k > best_k and words[-k][:1].isupper() and s >= best - 0.15 and s >= 0.4): best, best_k = max(s, best), k
        if best < 0.4: return m.group(0)
        keep, eng = words[:-best_k], " ".join(words[-best_k:])
        return (" ".join(keep) + " " if keep else "") + f"{pho} ({eng})"
    return _REVERSED.sub(repl, text)


# ---------- 1 file Word đã biên tập ----------
def parse_edited_docx(paragraph_texts):
    """Trả về dict: cast {NHÂN VẬT: DIỄN VIÊN}, speakers [tên], turns [(nhân vật, lời)], lines [câu thoại]."""
    cast, speakers, turns, lines = {}, [], [], []
    header_actors = set()
    in_header = False
    for raw in (x for para in paragraph_texts for x in str(para).split("\n")):  # 1 đoạn Word có thể chứa nhiều dòng
        t = raw.strip()
        if not t: continue
        if t.upper().startswith("VAI:"):
            in_header = True; continue
        if TIMECODE.match(t):
            in_header = False; continue
        m = SPEAKER_LINE.match(t)
        if in_header:
            if m and _is_name(m.group(1)):
                actor = m.group(2).split("\t")[0].strip().upper()
                if is_actor_name(actor):
                    cast[m.group(1).strip().upper()] = actor; header_actors.add(actor)
                speakers.append(m.group(1).strip())
            continue
        if m and _is_name(m.group(1)) and ("\t" in raw or m.group(2).strip()):
            name = m.group(1).strip()
            rest = m.group(2)
            if "\t" in rest:
                before, after = rest.split("\t", 1)
                if is_actor_name(before) and before.strip().upper() == before.strip():  # "Preston: KHÁNH<tab>lời" -> bỏ tên diễn viên
                    cast.setdefault(name.upper(), before.strip().upper())
                    rest = after
            text = rest.strip("\t ").strip()
            speakers.append(name)
            turns.append((name, text))
            lines.append(text)
        else:
            lines.append(t)
            if turns: turns[-1] = (turns[-1][0], turns[-1][1] + " " + t)
    return {"cast": cast, "header_actors": header_actors, "speakers": speakers, "turns": turns, "lines": lines}


def find_phonetic_pairs(text, vn_syllables, keep_case=False):
    """[(TIẾNG ANH, phiên âm)] tìm thấy trong 1 câu, nhận cả 2 chiều viết."""
    out = []
    for m in PAREN.finditer(text):
        inner = m.group(1).strip()
        if inner.lower() in NOTE_MARKS or ":" in inner or inner.startswith("*") or inner.replace(" ", "").isdigit(): continue
        before = text[:m.start()].rstrip()
        toks = [x for x in before.split() if _clean_tok(x)]
        if not toks: continue
        if ENG_OK.match(inner) and not has_vn(inner) and _pho_score(inner, vn_syllables) < 2:
            # Chiều 1: "phiên âm (English)" -> lấy số chữ phía trước bằng số chữ tiếng Anh
            k = len(inner.split())
            prev = [_clean_tok(x) for x in toks[-k:]]
            pho = " ".join(prev)
            if ENG_OK.match(pho) and not has_vn(pho) and "-" not in pho:
                # phía trước cũng là tiếng Anh -> có thể là chiều 2 ("Jimmy (chim-mi)")
                eng, pho = pho, inner
                if _pho_score(pho, vn_syllables) <= _pho_score(eng, vn_syllables): continue
            else:
                eng = inner
        else:
            # Chiều 2: "English (phiên âm)" -> lấy các chữ tiếng Anh liền trước
            pho = inner
            prev = []
            for x in reversed(toks[-4:]):
                c = _clean_tok(x)
                if not c or has_vn(c) or not ENG_OK.match(c + "x") or (c.islower() and strip_marks(c) in vn_syllables): break
                prev.insert(0, c)
            if not prev: continue
            eng = " ".join(prev)
        if not ENG_OK.match(eng) or len(re.sub(r"[^A-Za-z]", "", eng)) < 2: continue
        if not (has_vn(pho) or "-" in pho or _pho_score(pho, vn_syllables) >= 1): continue
        if pho.strip().lower() == eng.strip().lower(): continue
        if _similar(pho, eng) < 0.3: continue
        out.append((eng if keep_case else eng.upper(), pho))
    return out


# ---------- gộp nhiều file ----------
def learn(docx_files, srt_files, vn_syllables, pronoun_find_terms):
    """docx_files: [(tên file, [đoạn văn])]; srt_files: [(tên file, nội dung)]. Trả về dict các gợi ý thô."""
    pho = defaultdict(Counter); pho_src = {}; pho_eng_case = defaultdict(lambda: [0, 0])
    spk = Counter(); spk_src = {}
    cast = defaultdict(Counter); known_actors = Counter()
    pair_self = defaultdict(Counter); pair_target = defaultdict(Counter)
    for fname, paras in docx_files:
        d = parse_edited_docx(paras)
        known_actors.update(d["header_actors"])  # Counter: đếm số kịch bản có tên diễn viên này ở phần "VAI:"
        for ch, ac in d["cast"].items(): cast[ch][ac] += 1
        for s in set(d["speakers"]):
            spk[s.strip()] += 1; spk_src.setdefault(s.strip(), fname)
        for line in d["lines"]:
            for eng_raw, p in find_phonetic_pairs(line, vn_syllables, keep_case=True):
                eng = norm_eng_key(eng_raw)
                pho[eng][normalize_pho(p)] += 1; pho_src.setdefault(eng, fname)
                if not eng_raw.isupper(): pho_eng_case[eng][0 if eng_raw[:1].isupper() else 1] += 1
        turns = d["turns"]
        for i, (a, text) in enumerate(turns):
            b = next((turns[j][0] for j in (i - 1, i + 1) if 0 <= j < len(turns) and turns[j][0].upper() != a.upper()), None)
            if not b: continue
            key = f"{a.strip().upper()}|{b.strip().upper()}"
            for term, cat in pronoun_find_terms(text):
                if cat == "self": pair_self[key][term] += 1
                elif cat == "target": pair_target[key][term] += 1
    # Chỉ giữ tên diễn viên đã từng xuất hiện ở phần "VAI:" (loại câu cảm thán viết hoa bị nhận nhầm)
    real_actors = {a for a, n in known_actors.items() if n >= 2}
    cast = {ch: Counter({a: n for a, n in cnt.items() if a in real_actors}) for ch, cnt in cast.items()}
    cast = {ch: cnt for ch, cnt in cast.items() if cnt}
    speaker_upper = {s.upper() for s in spk}
    ns = Counter(); ns_src = {}
    srt_names = Counter(); srt_names_src = {}
    caps = defaultdict(lambda: [0, 0])  # từ tiếng Anh -> [số lần viết hoa giữa câu, số lần viết thường]
    for fname, content in srt_files:
        for line in content.splitlines():
            for mm in re.finditer(r"(?<=[A-Za-z0-9,;'\"] )([A-Za-z][A-Za-z'’]+)", line):  # chỉ đếm chữ GIỮA câu
                w = mm.group(1)
                if w.isupper(): continue  # viết tắt / la hét: không dùng để đoán
                caps[w.upper()][0 if w[0].isupper() else 1] += 1
            m = re.match(r"^\s*([A-Za-z0-9][^:]{0,45}?)\s*:(\s|$)", line)
            if not m or TIMECODE.match(line) or re.match(r"^\d+:\d+", line.strip()): continue
            p = m.group(1).strip()
            if "http" in p.lower() or p.upper() in speaker_upper: continue
            words = p.replace("’", "'").upper().split()
            if re.search(r"[?!]|\w\.\s", p) and not re.match(r"^(MS|MR|MRS|DR)\.", p.upper()): continue  # câu dính tên người nói: bỏ
            # "Bri's Dad", "Guy with phone": vẫn là nhân vật (mô tả người nói)
            descriptor = bool(re.search(r"\w'S\b", p.replace("’", "'").upper())) or bool(re.match(
                r"^(GUY|MAN|WOMAN|BOY|GIRL|KID|LADY|DUDE|PERSON|WORKER|EMPLOYEE)\b.*\b(WITH|IN)\b", p.upper()))
            looks_phrase = not descriptor and (any(w in FUNCTION_WORDS for w in words)
                                               or any(w[:1].islower() for w in p.split()) or len(words) > 4)
            if looks_phrase:   # "I was like:", "But remember:" -> không phải tên nhân vật
                ns[p.upper()] += 1; ns_src.setdefault(p.upper(), (fname, line.strip()[:90]))
            else:              # "Cameraman:", "Bri's Dad:" -> tên nhân vật tiếng Anh (bản Việt đã đổi tên)
                srt_names[p] += 1; srt_names_src.setdefault(p, fname)
    return {"pho": pho, "pho_src": pho_src, "pho_eng_case": pho_eng_case, "spk": spk, "spk_src": spk_src, "cast": cast,
            "pair_self": pair_self, "pair_target": pair_target, "ns": ns, "ns_src": ns_src,
            "srt_names": srt_names, "srt_names_src": srt_names_src, "caps": dict(caps)}


def is_proper(eng, caps, eng_case, pho_variants):
    """Đoán tên riêng: 1) phụ đề gốc viết hoa giữa câu; 2) team viết hoa chữ tiếng Anh trong ngoặc; 3) team viết hoa phiên âm."""
    words = eng.split()
    if len(words) == 1 or not eng.isupper():
        c = caps.get(words[0].upper() if words else "", [0, 0])
        if sum(c) >= 3:
            r = c[0] / sum(c)
            if r >= 0.7: return True
            if r <= 0.3: return False
    up, low = eng_case.get(eng, [0, 0])
    if up + low >= 2 and not all(w.isupper() for w in words):
        if up / (up + low) >= 0.7: return True
        if low / (up + low) >= 0.7: return False
    total = sum(pho_variants.values())
    capn = sum(n for v, n in pho_variants.items() if v[:1].isupper())
    return total > 0 and capn / total >= 0.5


def dominant(counter, min_uses=3, share=0.6):
    """Từ chiếm ưu thế rõ ràng, hoặc None."""
    total = sum(counter.values())
    if total < min_uses: return None
    term, n = counter.most_common(1)[0]
    return term if n / total >= share else None
