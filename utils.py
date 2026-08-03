import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_COLOR_INDEX
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io
import os
import re
import random
from collections import Counter
import time
import pandas as pd
import json
from gtts import gTTS
import zipfile

# ==========================================
# CẤU HÌNH DATABASE & FILE PATH
# ==========================================
NON_SPEAKER_DB_FILE = "custom_non_speakers.json"
SPEAKER_DB_FILE = "custom_speakers.json"
PHONETIC_DB_FILE = "custom_phonetics.json"
CAST_DB_FILE = "custom_cast_mapping.json"
TRACKER_DB_FILE = "dubbing_tracker.json"
RATES_DB_FILE = "payroll_rates.json"
PRONOUN_REL_DB_FILE = "custom_pronoun_relationships.json"
SPEAKER_COLOR_DB_FILE = "custom_speaker_colors.json"

DEFAULT_CAST_MAPPING = {
    "BRI": "TRÚC", "CHASE": "THIỆN", "PRESTON": "KHÁNH", "SCOTT": "THÔNG",
    "KEELEY": "TÚ", "DAKA": "QUANG", "STEPHEN": "QUANG", "VINCE": "QUANG",
    "JOSH": "THÔNG", "YOMI": "TÚ", "LARRY": "A.TRUNG", "CHLOE": "TÚ",
    "STEVEN": "QUANG (HOẶC THÔNG)", "LOGAN": "THÔNG", "NATHAN": "A.TRUNG",
    "RILEY": "PHỤNG", "ZHONG": "THÔNG", "MICHELLE (YOUTUBER)": "TÚ",
    "BEN AZELART": "QUANG", "ZONG": "THÔNG", "BA BRI": "THÔNG", "TYLER": "PHÁT",
    "COBY": "THIỆN", "CORY": "KHÁNH", "SPARKY": "A.TRUNG", "GARRETT": "C.DŨNG",
    "CODY": "CƯỜNG", "BETHANY": "TRÚC", "KRISTIN": "PHỤNG", "AMY": "TÚ",
    "ALLISON": "TÚ", "AUBREY": "PHỤNG", "TY JACKSON": "THÔNG", "HLV": "CƯỜNG",
    "JACKSON OLSON": "THÔNG", "DANNY": "QUANG", "KYLE": "QUANG", "BILL": "HITA", "TIM": "HITA"
}

DEFAULT_FIXED_SPEAKER_COLORS = {
    "ALL": {"text_color": (255, 0, 0), "highlight_color": (255, 255, 0)},
    "BEN AZELART": {"text_color": (21, 96, 130), "highlight_color": None},
    "BETHANY": {"text_color": (255, 192, 0), "highlight_color": None},
    "BRI": {"text_color": (116, 166, 123), "highlight_color": None},
    "CALEB": {"text_color": (116, 56, 25), "highlight_color": None},
    "CHASE": {"text_color": (255, 0, 255), "highlight_color": None},
    "CHUNKZ": {"text_color": (21, 96, 130), "highlight_color": None},
    "COBY": {"text_color": (255, 0, 255), "highlight_color": None},
    "CODY": {"text_color": (233, 113, 50), "highlight_color": None},
    "CORY": {"text_color": (255, 0, 0), "highlight_color": None},
    "DAKA": {"text_color": (110, 196, 229), "highlight_color": None},
    "GARRETT": {"text_color": (243, 243, 243), "highlight_color": None},
    "JOSH": {"text_color": (160, 43, 147), "highlight_color": None},
    "KEELEY": {"text_color": (255, 192, 0), "highlight_color": None},
    "LARRY": {"text_color": (243, 243, 243), "highlight_color": None},
    "LOGAN": {"text_color": (216, 170, 211), "highlight_color": None},
    "NATHAN": {"text_color": (21, 96, 130), "highlight_color": None},
    "NICK": {"text_color": (255, 0, 0), "highlight_color": None},
    "PRESTON": {"text_color": (255, 0, 0), "highlight_color": None},
    "RILEY": {"text_color": (116, 166, 123), "highlight_color": None},
    "SCOTT": {"text_color": (233, 113, 50), "highlight_color": None},
    "SPARKY": {"text_color": (116, 56, 25), "highlight_color": None},
    "STEPHEN": {"text_color": (0, 51, 204), "highlight_color": None},
    "STEVEN": {"text_color": (0, 255, 255), "highlight_color": None},
    "TYLER": {"text_color": (160, 43, 147), "highlight_color": None},
    "YOMI": {"text_color": (0, 153, 153), "highlight_color": None}
}

DEFAULT_SOUTH_VIETNAM_PHONETICS = {
    "MCDONALD": "Mắc-đô-nồ", "MCDONALD'S": "Mắc-đô-nồ", "LONDON": "Luân Đôn", "TNT": "Ti-èn-ti",
    "NOOB": "Núp", "GOLEM": "Gô-lùm", "GOOGLE": "Gú-gồ", "FACEBOOK": "Phây-sbúc",
    "KFC": "Ke-ép-xi", "STARBUCKS": "Xì-ta-bắc", "APPLE": "Ép-pồ", "BURGER": "Bơ-gơ",
    "PARIS": "Ba Lê", "WASHINGTON": "Hoa Thịnh Đốn", "CALIFORNIA": "Cả-li", "PRO": "Prồ",
    "LAG": "Lác", "BUFF": "Bớp", "VIP": "Vi-ai-pi", "FBI": "Ép-bi-ai"
}

DEFAULT_NON_SPEAKER_PHRASES = {
    "AND REMEMBER", "OFFICIAL DISTANCE", "GOOD NEWS FOR THEIR TEAMMATES", 
    "LL BE HONEST", "FIRST AND FOREMOST", "I SAID", "THE ONLY THING LEFT TO SETTLE", 
    "QUESTION IS", "FINALISTS", "WHISPERS", "SRT CONVERSION", 
    "WILL RED THRIVE OR WILL RED BE DEAD", "BUT REMEMBER", "THE RESULTS ARE IN", 
    "WE CHALLENGED", "I THINK", "IN THEIR DEFENSE", "THE PEAK OF HIS LIFE WAS DOING THE SPACETHING",
    "THE ROCKETS ARE BIGGER", "THE DISTANCE SHOULD BE FURTHER", "GET CRAFTY", "THAT WAS SO SICK",
    "OUT OF 100 CONTESTANTS", "THE FIRST ROUND IS BRUTAL", "YOU KNOW WHICH END GOES",
    "THE GAME IS ON", "THAT'S A GOOD THROW", "HE'S GOING FOR IT", "WE GOT THIS",
    "LAUNCH", "OH NO", "OH", "AH", "YEP", "WAIT", "YEAH", "WOO", "OKAY", "YES", "I ANH", "O BRI", "NG", "THE ONLY PROBLEM", "NOTE", "WARNING", "THINGS"
}

TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
ENGLISH_WORD_REGEX = re.compile(r"\b[A-Za-z][A-Za-z0-9'-]*\b")
RED_COLOR = RGBColor(255, 0, 0)

# ==========================================
# HÀM ĐỌC/GHI JSON DATABASE
# ==========================================
def load_json_db(filepath, default_data=None):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                return set(data) if isinstance(data, list) and isinstance(default_data, set) else data
        except Exception: pass
    return default_data if default_data is not None else (set() if not isinstance(default_data, dict) else {})

def save_json_db(filepath, data_container):
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            if isinstance(data_container, set): json.dump(list(data_container), f, ensure_ascii=False, indent=2)
            else: json.dump(data_container, f, ensure_ascii=False, indent=2)
    except Exception as e: st.error(f"Không thể lưu vào Database: {e}")

def hex_to_rgb(hex_str):
    if not hex_str: return None
    hex_str = str(hex_str).strip().lstrip('#')
    if len(hex_str) == 6:
        try: return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))
        except ValueError: return None
    return None

def generate_vibrant_rgb_colors_excluding(excluded_rgbs, count=200):
    colors = []
    used_set = set(tuple(x) for x in excluded_rgbs if x)
    attempts = 0
    while len(colors) < count and attempts < 10000:
        attempts += 1
        h = random.random(); s = 0.9; v = 0.8
        i = int(h * 6.0); f = h * 6.0 - i; p = v * (1.0 - s); q = v * (1.0 - s * f); t = v * (1.0 - s * (1.0 - f))
        if i % 6 == 0: r, g, b = v, t, p
        elif i % 6 == 1: r, g, b = q, v, p
        elif i % 6 == 2: r, g, b = p, v, t
        elif i % 6 == 3: r, g, b = p, q, v
        elif i % 6 == 4: r, g, b = t, p, v
        else: r, g, b = v, p, q
        r_int, g_int, b_int = int(r * 255), int(g * 255), int(b * 255)
        rgb_tuple = (r_int, g_int, b_int)
        if rgb_tuple not in used_set:
            used_set.add(rgb_tuple)
            colors.append(rgb_tuple)
    return colors

def get_speaker_color_config(speaker_name, speaker_color_map, available_rgb_tuples):
    spk_upper = speaker_name.strip().upper()
    fixed_colors = st.session_state.get('fixed_speaker_colors', DEFAULT_FIXED_SPEAKER_COLORS)
    
    if spk_upper in fixed_colors:
        cfg = fixed_colors[spk_upper]
        if isinstance(cfg, dict): return cfg
        elif isinstance(cfg, (tuple, list)): return {"text_color": tuple(cfg), "highlight_color": None}
            
    if spk_upper not in speaker_color_map:
        if available_rgb_tuples: r, g, b = available_rgb_tuples.pop()
        else: r, g, b = (random.randint(50, 220), random.randint(50, 220), random.randint(50, 220))
        speaker_color_map[spk_upper] = {"text_color": (r, g, b), "highlight_color": None}
        
    return speaker_color_map[spk_upper]

def apply_speaker_styling_to_run(run, text_color_tuple, highlight_color_tuple):
    if text_color_tuple:
        r, g, b = text_color_tuple
        run.font.color.rgb = RGBColor(r, g, b)
    if highlight_color_tuple:
        hr, hg, hb = highlight_color_tuple
        hex_fill = f"{hr:02X}{hg:02X}{hb:02X}"
        shd_xml = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_fill}"/>')
        run._r.get_or_add_rPr().append(shd_xml)

def clean_and_normalize_text(text, strip_all_tags=False, fix_punctuation=True, normalize_spaces=True, capitalize_first=True, remove_leading_dash=True):
    if not text or not isinstance(text, str): return ""
    res = text
    res = re.sub(r'\{\\[^}]*\}', '', res)
    res = re.sub(r'\\N', '\n', res, flags=re.IGNORECASE)
    
    if strip_all_tags: res = re.sub(r'<[^>]*>', '', res)
    else: res = re.sub(r'<(?!/?(i|b|u)\b)[^>]*>', '', res, flags=re.IGNORECASE)
        
    if remove_leading_dash:
        res = re.sub(r'^\s*[-–—]\s*', '', res)
        res = re.sub(r'(\n)\s*[-–—]\s*', r'\1', res)
        
    if fix_punctuation:
        res = re.sub(r'\s+([,!?:;\.\)])', r'\1', res)
        res = re.sub(r'([,!?:;])([A-Za-zÀ-ỹ0-9])', r'\1 \2', res)
        res = re.sub(r'(\.\.\.)([A-Za-zÀ-ỹ0-9])', r'\1 \2', res)
        res = re.sub(r'"\s*(.*?)\s*"', r'"\1"', res)
        
    if normalize_spaces:
        res = re.sub(r'[ \t]+', ' ', res)
        res = re.sub(r'\s*\n\s*', '\n', res)
        res = res.strip()
        
    if capitalize_first and res:
        lines = res.split('\n')
        cap_lines = []
        for line in lines:
            m = re.match(r'^([^:]+:\s*)([a-zà-ỹ])(.*)$', line)
            if m: line = m.group(1) + m.group(2).upper() + m.group(3)
            elif line and line[0].islower(): line = line[0].upper() + line[1:]
            cap_lines.append(line)
        res = '\n'.join(cap_lines)
        
    return res

def calculate_duration_sec(timecode_line):
    m = re.match(r"(\d{2}):(\d{2}):(\d{2})[,.](\d{3})\s+-->\s+(\d{2}):(\d{2}):(\d{2})[,.](\d{3})", timecode_line.strip())
    if not m: return 1.0, 0.0, 0.0
    h1, m1, s1, ms1, h2, m2, s2, ms2 = map(int, m.groups())
    t1 = h1 * 3600 + m1 * 60 + s1 + ms1 / 1000.0
    t2 = h2 * 3600 + m2 * 60 + s2 + ms2 / 1000.0
    dur = t2 - t1
    return dur if dur > 0 else 0.5, t1, t2

def timecode_to_sec(tc):
    m = re.match(r"(\d{2}):(\d{2}):(\d{2})[,.](\d{3})", str(tc).strip())
    if not m: return 0.0
    h, mins, s, ms = map(int, m.groups())
    return h * 3600 + mins * 60 + s + ms / 1000.0

def calculate_time_overlap(s1, e1, s2, e2):
    latest_start = max(s1, s2)
    earliest_end = min(e1, e2)
    return max(0.0, earliest_end - latest_start)

def is_valid_speaker_boundary(prefix, start_idx):
    if start_idx == 0: return True
    prev_char = prefix[start_idx - 1]
    curr_char = prefix[start_idx]
    if not prev_char.isalnum(): return True
    if prev_char.islower() and curr_char.isupper(): return True
    return False

def find_all_speaker_tags(text, custom_speakers=None, non_speakers=None):
    if custom_speakers is None: custom_speakers = st.session_state.get('custom_speakers', set())
    if non_speakers is None: non_speakers = st.session_state.get('custom_non_speakers', set())
    non_speakers_upper = {s.upper() for s in non_speakers}.union({s.upper() for s in DEFAULT_NON_SPEAKER_PHRASES})
    
    custom_names_upper = {s.upper(): s for s in custom_speakers if s.strip()}
    pattern = r"([A-Za-z0-9À-ỹ \t&\-\(\)\.]{1,35}):\s*"
    
    matches = []
    for m in re.finditer(pattern, text):
        raw_prefix = m.group(1)
        extracted_spk = None
        
        for spk_upper, orig_spk in sorted(custom_names_upper.items(), key=lambda x: len(x[0]), reverse=True):
            if raw_prefix.upper().endswith(spk_upper):
                start_idx = len(raw_prefix) - len(spk_upper)
                if is_valid_speaker_boundary(raw_prefix, start_idx):
                    extracted_spk = orig_spk
                    spk_start = m.start(1) + start_idx
                    spk_end = m.end()
                    matches.append((spk_start, spk_end, extracted_spk, m.group(0)))
                    break
                    
        if extracted_spk: continue
            
        match_tail = re.search(r"(?:^|[^\w]|(?<=[a-zà-ỹ]))([A-ZÀ-Ỹ0-9][A-Za-z0-9À-ỹ \t&\-\(\)\.]{0,24})$", raw_prefix)
        if match_tail:
            cand = match_tail.group(1).strip(".,!?:;- ")
            if cand and len(cand) <= 25 and not cand.isdigit():
                if not (cand.startswith('(') or cand.endswith(')')):
                    if cand.upper() not in non_speakers_upper:
                        if len(cand.split()) <= 4:
                            start_idx = len(raw_prefix) - len(match_tail.group(1))
                            if is_valid_speaker_boundary(raw_prefix, start_idx):
                                spk_start = m.start(1) + start_idx
                                spk_end = m.end()
                                matches.append((spk_start, spk_end, cand, m.group(0)))

    matches.sort(key=lambda x: x[0])
    filtered_matches = []
    last_end = -1
    for start, end, spk, raw_m in matches:
        if start >= last_end:
            filtered_matches.append((start, end, spk, raw_m))
            last_end = end
            
    return filtered_matches

def is_valid_speaker_name(name):
    clean = name.strip()
    if not clean or len(clean) > 25 or clean.isdigit() or re.match(r'^\d+[\d\s:]*$', clean): return False
    if clean.startswith('(') or clean.endswith(')') or clean.upper() in DEFAULT_NON_SPEAKER_PHRASES: return False
    if any(char in clean for char in ['/', '?', '!', ',', '.', '-->', '(', ')']): return False
    if len(clean.split()) > 4: return False
    return True

def get_paragraph_text_with_html(paragraph):
    text = ""
    for run in paragraph.runs:
        r_text = run.text
        if not r_text: continue
        if run.italic and not ("<i>" in r_text or "</i>" in r_text): text += f"<i>{r_text}</i>"
        else: text += r_text
    text = text.replace("</i><i>", "")
    return text

def round_seconds_to_int_minutes(total_sec):
    if total_sec <= 0: return 1
    mins = int(total_sec // 60); secs = int(round(total_sec % 60))
    if secs >= 30: mins += 1
    return max(1, mins)

def srt_timecode_to_ass(timecode_line):
    if not timecode_line or not isinstance(timecode_line, str): return None, None
    m = re.match(r"(\d{2}):(\d{2}):(\d{2})[,.](\d{3})\s+-->\s+(\d{2}):(\d{2}):(\d{2})[,.](\d{3})", timecode_line.strip())
    if not m: return None, None
    h1, m1, s1, ms1, h2, m2, s2, ms2 = m.groups()
    ass_start = f"{int(h1)}:{m1}:{s1}.{int(ms1)//10:02d}"
    ass_end = f"{int(h2)}:{m2}:{s2}.{int(ms2)//10:02d}"
    return ass_start, ass_end

def rgb_to_ass_hex(rgb_obj):
    if not rgb_obj: return "&H00FFFFFF&"
    try:
        r, g, b = rgb_obj[0], rgb_obj[1], rgb_obj[2]
        return f"&H00{b:02X}{g:02X}{r:02X}&"
    except Exception: return "&H00FFFFFF&"

def clean_file_name_for_output(original_filename, tag="_edit", ext=".docx"):
    name_without_ext = os.path.splitext(original_filename)[0]
    cleaned = re.sub(r'(CONVERTED_|FORMATTED_|\s*\(.*\)$|_edit$|_resync$|_final$)', '', name_without_ext, flags=re.IGNORECASE).strip()
    return f"{cleaned}{tag}{ext}"

def generate_actor_docx(video_title, actor_name, dialogue_list, font_size_pt=12):
    doc = Document()
    p_title = doc.add_paragraph(f"KỊCH BẢN THU ÂM - DIỄN VIÊN: {actor_name}")
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.runs[0].font.name = 'Times New Roman'; p_title.runs[0].font.size = Pt(16); p_title.runs[0].bold = True
    
    p_sub = doc.add_paragraph(f"Video: {video_title}")
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.runs[0].font.name = 'Times New Roman'; p_sub.runs[0].font.size = Pt(12); p_sub.runs[0].font.italic = True
    
    doc.add_paragraph()
    TAB_STOP = Inches(1.0)
    for item in dialogue_list:
        if item.get("timecode"):
            p_tc = doc.add_paragraph(item["timecode"])
            p_tc.runs[0].font.name = 'Times New Roman'; p_tc.runs[0].font.size = Pt(font_size_pt); p_tc.runs[0].bold = True
            p_tc.paragraph_format.space_before = Pt(0); p_tc.paragraph_format.space_after = Pt(0)
            
        p_line = doc.add_paragraph()
        p_line.paragraph_format.left_indent = TAB_STOP
        p_line.paragraph_format.first_line_indent = Inches(-1.0)
        p_line.paragraph_format.tab_stops.add_tab_stop(TAB_STOP, WD_TAB_ALIGNMENT.LEFT)
        
        r_spk = p_line.add_run(f"{item['speaker']}:")
        r_spk.font.name = 'Times New Roman'; r_spk.font.size = Pt(font_size_pt); r_spk.font.bold = True
        r_spk.font.color.rgb = RGBColor(79, 70, 229)
        p_line.add_run("\t")
        r_text = p_line.add_run(item['text'])
        r_text.font.name = 'Times New Roman'; r_text.font.size = Pt(font_size_pt)
        p_line.paragraph_format.space_before = Pt(0); p_line.paragraph_format.space_after = Pt(4)
        
    for p in doc.paragraphs: p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf
