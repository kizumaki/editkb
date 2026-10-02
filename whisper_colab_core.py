"""Lõi công cụ "Tạo phụ đề tiếng Anh bằng Whisper" (chạy trên Google Colab; cũng chạy được trên máy để thử).
Notebook `Tao_phu_de_Whisper.ipynb` tải file này từ GitHub rồi gọi `run()`.

- Nhận file TIẾNG hoặc VIDEO (bản sao người dùng tải lên phiên Colab). Video: tự tách tiếng rồi XOÁ BẢN SAO ngay — file gốc trên máy người dùng không bị đụng tới.
- faster-whisper: thử card đồ hoạ (GPU) trước; lỗi/không có thì tự chuyển sang CPU (chậm hơn) và báo rõ.
- Chia phụ đề theo chuẩn: tối đa 42 ký tự/dòng, 2 dòng/khung, 6 giây/khung; ngắt ở dấu câu hoặc chỗ ngừng nói.
- Xong: tự tải file .srt về máy, xoá hết file tạm trong phiên.
"""
import os
import re
import shutil
import subprocess
import time

AUDIO_EXT = (".wav", ".mp3", ".m4a", ".aac", ".flac", ".ogg", ".opus", ".wma")
MAX_LINE, MAX_LINES, MAX_DUR, MIN_DUR, PAUSE = 42, 2, 6.0, 0.8, 0.6


def fmt_ts(t):
    t = max(0.0, t); ms = int(round(t * 1000))
    h, ms = divmod(ms, 3600000); m, ms = divmod(ms, 60000); s, ms = divmod(ms, 1000)
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


STYLE_HINT = "Hello, everyone! Welcome back to the channel. Today, we're trying something new, and I love it."  # giữ dấu câu + chữ hoa


# Tên quen theo kênh -> gợi ý cho Whisper (chọn ở ô KENH trong notebook). Chỉ tên người/nhân vật công khai — KHÔNG có nội dung kịch bản.
# Nhóm Preston/Brianna/Keeley: rút từ 84 kịch bản team đã biên tập (tên xuất hiện ở >= 3 kịch bản). Kênh khác: tra cứu công khai
# 2026-10-02 (Wikipedia/fandom: thành viên Dude Perfect, nhóm MrBeast của Karl, Ethan Schulteis = The Amagi kênh giải thích anime,
# bạn bè IShowSpeed; Ethan theo trang kênh người dùng gửi) + bản chính thức McDonald's (Manny, Remy của Nick). Team bổ sung thì sửa ở đây.
_NHOM_PRESTON = ("Preston, TBNRfrags, Brianna, Bri, Keeley, Chase, Stephen, Scott, Larry, Yomi, Caleb, Riley, Courtney, Josh, Alan, "
                 "Vince, Johnny, Joe, Kat, Ben, ZHC")
KENH = {
    "Không rõ / kênh khác": "",
    "Preston (prestonyt / PrestonPlayz)": _NHOM_PRESTON + ", Minecraft, Creeper, Enderman, Villager",
    "Brianna (BriannaPlayz / BriannaYT)": _NHOM_PRESTON,
    "Keeley (ItsKeeleyElise)": _NHOM_PRESTON,
    "Nick DiGiovanni": "Nick DiGiovanni, Nick, Manny, Remy",
    "Dude Perfect": "Dude Perfect, Tyler, Cory, Coby, Garrett, Cody, Panda, Sparky",
    "IShowSpeed": "IShowSpeed, Speed, Kai Cenat, KSI, Jamal, Ronaldo, Messi",
    "Karl (Karl Jacobs)": "Karl, MrBeast, Jimmy, Chandler, Chris, Nolan, Tareq, Sapnap, GeorgeNotFound, Minecraft",
    # @EthanSchulteis = kênh "Ethan" thử thách/sinh tồn (KHÔNG phải The Amagi trùng tên — tra web nhầm, người dùng sửa 2026-10-02).
    # Chỉ có tên chắc chắn từ trang kênh (video "Ryan Trahan Mystery Country Challenge"); chờ team bổ sung.
    "Ethan (EthanSchulteis)": "Ethan, Ryan Trahan, Ryan",
}

_NAME_JUNK = {"mp4", "mov", "mkv", "wav", "mp3", "m4a", "final", "edit", "raw", "copy", "en", "eng", "vi", "vn", "sub", "subs",
              "hd", "fhd", "uhd", "4k", "tieng", "audio", "video"}


def title_from_filename(path):
    """Tên file -> gợi ý cho Whisper (tự động, không ai phải gõ): "I_Tried_McDonald's_From_Every_Country_1080p.mp4"
    -> "I Tried McDonald's From Every Country". Bỏ mã số, độ phân giải, chữ rác thường gặp trong tên file."""
    name = os.path.splitext(os.path.basename(path))[0]
    name = re.sub(r"\[[^\]]*\]|\([^)]*\)", " ", name)  # [mã], (bản 2)
    words = [w for w in re.split(r"[\s_.]+", name) if w]
    keep = [w for w in words if not re.fullmatch(r"\d+|\d+p|v\d+|\d+fps", w, re.I) and w.lower() not in _NAME_JUNK]
    return " ".join(keep)[:150]


def split_two(text, width=MAX_LINE):
    """Chia thành 2 dòng cân đối, CHỈ ngắt ở dấu cách. Không chia được -> None."""
    words = text.split(); best = None
    for i in range(1, len(words)):
        a, b = " ".join(words[:i]), " ".join(words[i:])
        if len(a) <= width and len(b) <= width:
            score = abs(len(a) - len(b)) - (8 if a.endswith((",", ".", "?", "!")) else 0)
            if best is None or score < best[0]: best = (score, a, b)
    return [best[1], best[2]] if best else None


def fits(text):
    return len(text) <= MAX_LINE or split_two(text) is not None


def wrap_lines(text, width=MAX_LINE):
    """Chia 1 khung thành tối đa 2 dòng cân đối (không bao giờ cắt giữa chữ)."""
    if len(text) <= width: return [text]
    two = split_two(text, width)
    if two: return two
    lines, cur = [], ""  # khung quá dài (hiếm): xuống dòng theo chữ
    for w in text.split():
        if cur and len(cur) + 1 + len(w) > width: lines.append(cur); cur = w
        else: cur = (cur + " " + w).strip()
    return lines + [cur]


def join_words(ws):
    """Nối chữ Whisper trả về: mỗi chữ đã kèm dấu cách đầu (" Nori", "-dusted", " 9", "2") -> nối thẳng,
    KHÔNG thêm dấu cách (bản cũ thêm -> "Nori -dusted", "9 2", "McDonald' s")."""
    return re.sub(r"\s+", " ", "".join(w[2] for w in ws)).strip()


def build_cues(words):
    """words: [(start, end, text)] theo thứ tự. Trả về [(start, end, text)] khung phụ đề."""
    cues, cur = [], []
    def flush():
        if cur:
            txt = join_words(cur)
            if txt: cues.append([cur[0][0], cur[-1][1], txt])
        cur.clear()
    for w in words:
        if not w[2].strip(): continue
        if cur:
            gap = w[0] - cur[-1][1]
            new_txt = join_words(cur + [w])
            too_long = not fits(new_txt) or (w[1] - cur[0][0]) > MAX_DUR
            sentence_end = cur[-1][2].strip().endswith((".", "?", "!")) and len(new_txt) > 20
            if too_long or gap >= PAUSE or sentence_end:
                flush()
        cur.append(w)
        if w[2].strip().endswith((".", "?", "!")) and len(join_words(cur)) >= MAX_LINE * MAX_LINES * 0.75:
            flush()
    flush()
    for i, c in enumerate(cues):  # khung quá ngắn -> kéo dài (không chồng khung sau)
        if c[1] - c[0] < MIN_DUR:
            nxt = cues[i + 1][0] if i + 1 < len(cues) else c[0] + MIN_DUR
            c[1] = min(c[0] + MIN_DUR, max(c[1], nxt - 0.05))
    return [tuple(c) for c in cues]


def to_srt(cues):
    out = []
    for i, (s, e, t) in enumerate(cues, 1):
        out.append(f"{i}\n{fmt_ts(s)} --> {fmt_ts(e)}\n" + "\n".join(wrap_lines(t)) + "\n")
    return "\n".join(out)


def extract_audio(path):
    """Video/tiếng bất kỳ -> wav 16kHz 1 kênh. Video gốc bị XOÁ ngay sau khi tách."""
    out = os.path.splitext(path)[0] + "_tieng.wav"
    r = subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", path, "-vn", "-ac", "1", "-ar", "16000", out],
                       capture_output=True, text=True)
    if r.returncode != 0: raise RuntimeError("Không tách được tiếng: " + (r.stderr or "")[-300:])
    if not path.lower().endswith(AUDIO_EXT): os.remove(path)  # xoá video ngay
    return out


CUDA_DIR = "/content/mh_cuda12"  # thư viện CUDA 12 tải riêng (không đụng thư viện có sẵn của Colab)
CUDA_LIBS = ("libcublasLt.so.12", "libcublas.so.12", "libcudnn*.so.9")


def preload_cuda_libs(log=print):
    """faster-whisper (ctranslate2) cần thư viện CUDA 12 + cuDNN 9. Colab bản mới có thể không còn sẵn
    (lỗi "libcublas.so.12 is not found") -> tải gói pip nvidia-*-cu12 vào thư mục riêng rồi nạp trước."""
    import ctypes, glob, sys
    try:
        ctypes.CDLL("libcublas.so.12"); ctypes.CDLL("libcudnn_ops.so.9")
        return True  # máy đã có sẵn
    except OSError:
        pass
    if not glob.glob(os.path.join(CUDA_DIR, "nvidia", "*", "lib")):
        log("⏳ Máy Colab thiếu thư viện cho card đồ hoạ — đang tải thêm (khoảng 5 phút, chỉ lần đầu mỗi phiên)...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", "--target", CUDA_DIR,
                        "nvidia-cublas-cu12", "nvidia-cudnn-cu12==9.*"], check=False)
    files = [f for d in glob.glob(os.path.join(CUDA_DIR, "nvidia", "*", "lib"))
             for pat in CUDA_LIBS for f in glob.glob(os.path.join(d, pat))]
    pending = sorted(set(files))
    for _ in range(4):  # nạp nhiều lượt: thư viện phụ thuộc nhau, thứ tự không chắc
        left = []
        for f in pending:
            try: ctypes.CDLL(f, mode=ctypes.RTLD_GLOBAL)
            except OSError: left.append(f)
        if not left or len(left) == len(pending): break
        pending = left
    try:
        ctypes.CDLL("libcublas.so.12"); return True
    except OSError:
        return False


def gpu_works(model):
    """Chạy thử 1 giây trên card đồ hoạ. Lỗi thư viện CUDA chỉ lộ ra lúc nhận dạng thật, không lộ lúc nạp mô hình."""
    import numpy as np
    segs, _ = model.transcribe(np.zeros(16000, dtype=np.float32), language="en", beam_size=1)
    list(segs)
    return True


def load_model(name="large-v3", log=print):
    from faster_whisper import WhisperModel
    try:
        import ctranslate2
        if ctranslate2.get_cuda_device_count() > 0:
            preload_cuda_libs(log)
            m = WhisperModel(name, device="cuda", compute_type="float16")
            gpu_works(m)
            log("⚡ Đang dùng CARD ĐỒ HOẠ (nhanh).")
            return m, "gpu"
    except Exception as e:
        log(f"⚠️ Không dùng được card đồ hoạ ({str(e)[:120]}) — chuyển sang bộ xử lý thường.")
    log("🐢 Đang chạy bằng bộ xử lý thường, sẽ CHẬM (video 10 phút có thể mất 30–60 phút). "
        "Nếu chưa bật card: Thời gian chạy → Thay đổi loại thời gian chạy → T4 GPU rồi chạy lại.")
    return WhisperModel(name, device="cpu", compute_type="int8"), "cpu"


def is_cuda_error(e):
    s = str(e).lower()
    return any(k in s for k in ("cuda", "cublas", "cudnn", "libcu"))


def read_wav(wav):
    """Đọc wav 16kHz 1 kênh -> mảng số (không cần thư viện đọc âm thanh riêng, tránh lỗi phiên bản)."""
    import wave
    import numpy as np
    with wave.open(wav, "rb") as f:
        data = f.readframes(f.getnframes())
    return np.frombuffer(data, dtype=np.int16).astype(np.float32) / 32768.0


def transcribe_file(model, wav, language="en", log=print, keywords=""):
    """keywords: tên riêng / từ khó của video ("Nick, Manny, McFlurry") -> gợi ý cho Whisper nghe đúng.
    Gợi ý (hotwords) được gắn vào MỌI đoạn 30 giây, kèm 1 câu mẫu có dấu câu để Whisper không bỏ dấu chấm / chữ hoa
    (bản cũ: có đoạn dài toàn chữ thường, không dấu)."""
    t0 = time.time()
    hint = (STYLE_HINT + " " + keywords.strip()).strip() if language in ("en", "auto") else keywords.strip()
    segs, info = model.transcribe(read_wav(wav), language=None if language == "auto" else language, beam_size=5,
                                  vad_filter=True, vad_parameters={"min_silence_duration_ms": 400},
                                  word_timestamps=True, condition_on_previous_text=False,
                                  hotwords=hint or None)
    words, last = [], 0
    for seg in segs:
        for w in (seg.words or []):
            words.append((w.start, w.end, w.word))
        if info.duration and seg.end - last > 30:
            last = seg.end; log(f"   … {int(seg.end // 60)}:{int(seg.end % 60):02d} / {int(info.duration // 60)}:{int(info.duration % 60):02d}")
    log(f"   ✔ xong trong {int(time.time() - t0)} giây")
    return build_cues(words)


def run(paths, model_name="large-v3", language="en", log=print, keywords=""):
    """Chạy cho danh sách file đã tải vào phiên. Trả về danh sách file .srt đã tạo."""
    model, dev = load_model(model_name, log)
    made = []
    for p in paths:
        name = os.path.basename(p); log(f"\n🎬 {name}")
        wav = None
        title = title_from_filename(p)  # tên video thường chứa đúng tên riêng khó nghe (VD "McDonald's")
        kw = ", ".join(x for x in (title, keywords.strip()) if x)
        if kw: log(f"   💡 Gợi ý cho Whisper (tên file + kênh + từ khoá): {kw}")
        try:
            wav = extract_audio(p); log("   ✔ đã tách tiếng" + ("" if p.lower().endswith(AUDIO_EXT) else " và xoá bản sao video trên Colab (video gốc trên máy bạn vẫn còn)"))
            try:
                cues = transcribe_file(model, wav, language, log, kw)
            except Exception as e:
                if dev != "gpu" or not is_cuda_error(e): raise
                log(f"   ⚠️ Card đồ hoạ lỗi ({str(e)[:120]}) — chuyển sang bộ xử lý thường và làm lại (CHẬM hơn).")
                from faster_whisper import WhisperModel
                model, dev = WhisperModel(model_name, device="cpu", compute_type="int8"), "cpu"
                cues = transcribe_file(model, wav, language, log, kw)
            srt = os.path.splitext(p)[0] + "_EN_whisper.srt"
            with open(srt, "w", encoding="utf-8") as f: f.write(to_srt(cues))
            made.append(srt); log(f"   ✔ {len(cues)} khung phụ đề → {os.path.basename(srt)}")
        except Exception as e:
            log(f"   ❌ Lỗi: {e}")
        finally:  # dù thành công hay lỗi: xoá video/tiếng gốc và file tiếng tạm (bảo mật)
            for f in (p, wav):
                if f and os.path.exists(f) and not f.lower().endswith(".srt"):
                    try: os.remove(f)
                    except OSError: pass
    return made


def cleanup(folder):
    """Xoá sạch thư mục làm việc trong phiên Colab (ra khỏi thư mục trước, vì không xoá được thư mục đang đứng)."""
    try: os.chdir(os.path.dirname(os.path.abspath(folder)))
    except OSError: pass
    shutil.rmtree(folder, ignore_errors=True)
