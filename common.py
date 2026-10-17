# -*- coding: utf-8 -*-
"""common.py — ฟังก์ชันและค่าคงที่ที่ใช้ร่วมกันใน exam-code.py และ gen-report-*.py"""

import os
import re
import pickle
from datetime import datetime

import numpy as np
import pandas as pd

from config import *   # noqa: F401,F403  (DATA_DIR, OUT_DIR, EXTYPE, ...)

pd.set_option('display.max_columns', None)

extype = EXTYPE        # ชื่อเดิมที่โค้ดใช้


# ── State handoff: exam-code.py -> gen-report-*.py ────────────────────────────

STATE_PATH = os.path.join(OUT_DIR, STATE_FILE)


def save_state(**objs) -> str:
    """Pickle the given objects (e.g. dfX, dfST, dfslot, dtlist) to STATE_PATH."""
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(STATE_PATH, 'wb') as f:
        pickle.dump(objs, f, protocol=pickle.HIGHEST_PROTOCOL)
    print(f"Saved state: {STATE_PATH}  keys={list(objs)}")
    return STATE_PATH


def load_state() -> dict:
    """Load the dict written by save_state()."""
    if not os.path.exists(STATE_PATH):
        raise FileNotFoundError(f"{STATE_PATH} not found — run exam-code.py first")
    with open(STATE_PATH, 'rb') as f:
        state = pickle.load(f)
    print(f"Loaded state: {STATE_PATH}  keys={list(state)}")
    return state


def out_path(filename: str) -> str:
    """Path of a report file inside OUT_DIR (flat, no sub folders)."""
    os.makedirs(OUT_DIR, exist_ok=True)
    return os.path.join(OUT_DIR, filename)


# ── Thai date / session constants ─────────────────────────────────────────────

THAI_DAYS_SHORT = {
    0: 'จ.', 1: 'อ.', 2: 'พ.',
    3: 'พฤ.', 4: 'ศ.', 5: 'ส.', 6: 'อา.',
}
THAI_MONTHS_SHORT = {
    1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.',
    5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.',
    9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.',
}
SESSION_DURATION = {
    '09': '09:30-12:30 น.',
    '13': '13:30-16:30 น.',
}
SESSION_LABEL = {
    '09': 'เช้า',
    '13': 'บ่าย',
}



def parse_ex_dt(ex_dt_str: str):
    """Return (datetime, hour_str) from 'YYYY-MM-DD HH'."""
    date_str, hour_str = ex_dt_str.strip().split()
    return datetime.strptime(date_str, "%Y-%m-%d"), hour_str


def slot_sheet_name(ex_dt_str: str) -> str:
    """'2026-03-16 09' → '16 มี.ค. (เช้า)'"""
    dt, hour_str = parse_ex_dt(ex_dt_str)
    session = SESSION_LABEL.get(hour_str, hour_str)
    return f"{dt.day} {THAI_MONTHS_SHORT[dt.month]} ({session})"



def exam_semester(ex_date) -> str:
    d = pd.Timestamp(ex_date)
    if 7 <= d.month <= 11:
        return f"1/{d.year + 543}"
    else:
        return f"2/{(d.year - 1) + 543}"


def exam_type_th(extype: str) -> str:
    return 'ปลายภาค' if extype == 'final' else 'กลางภาค'


# ── Reading exam schedule files ───────────────────────────────────────────────

def process_exam_report(file_path):
    """
    Read the new single-sheet exam Excel format.
    Row 0  : title text  → ignored
    Row 1  : column headers
    Row 2+ : data rows   (rows where Remark == 'เพิ่มแถว' are skipped)
    """
    df_raw = pd.read_excel(file_path, header=1)   # row index 1 = header row

    # ── Drop extra/blank rows ─────────────────────────────────────────────────
    #df_raw = df_raw[df_raw['Remark'].isna() | (df_raw['Remark'].str.strip() == '')]
    df_raw = df_raw.dropna(subset=['วันที่', 'เวลา', 'วิชา'])

    # ── Parse exam date from Thai short format e.g. 'จ. 17 ส.ค. 69 ' ──────────
    THAI_MONTHS_REV = {v.replace('.', ''): k for k, v in THAI_MONTHS_SHORT.items()}

    def parse_thai_date(s: str):
        s = str(s).strip()
        # pattern: '<day_abbr> DD MMM YY'  e.g. 'จ. 17 ส.ค. 69'
        match = re.search(r'(\d{1,2})\s+([ก-์\.]+)\s+(\d{2})', s)
        if not match:
            return None
        day        = int(match.group(1))
        month_s    = match.group(2).replace('.', '').strip()
        year_short = int(match.group(3))
        month      = THAI_MONTHS_REV.get(month_s)
        if month is None:
            return None
        year = year_short + 2500 - 543     # พ.ศ. 2-digit → ค.ศ.
        return datetime(year, month, day).date()

    df_raw['ex_date'] = df_raw['วันที่'].apply(parse_thai_date)

    # ── Parse exam duration / start hour from 'HH:MM:SS - HH:MM:SS' ──────────
    def parse_ex_dur(s: str) -> tuple[str, str]:
        """Return (duration_str '09:30-12:30', hour_str '09')."""
        s = str(s).strip()
        match = re.search(r'(\d{2}):(\d{2}).*?(\d{2}):(\d{2})', s)
        if match:
            dur  = f"{match.group(1)}:{match.group(2)}-{match.group(3)}:{match.group(4)}"
            hour = match.group(1)
            return dur, hour
        return s, s[:2]

    df_raw[['ex_dur', 'ex_time']] = df_raw['เวลา'].apply(
        lambda x: pd.Series(parse_ex_dur(x))
    )

    # ── Rename columns to match downstream code ───────────────────────────────
    df_out = df_raw.rename(columns={
        'ลำดับ':    'order',
        'วิชา':     'subj',
        'กลุ่ม':    'sec',
        'ชั้นปี':   'st_year',
        'นศ.':      'st_max',
        'ผู้สอน':   'teacher_str',
        'อาคาร':    'building',
        'ห้อง':     'room',
        'หมายเหตุ': 'note',
    })

    # ── Normalise text fields ──────────────────────────────────────────────────
    for col in ['subj', 'building', 'room', 'teacher_str']:
        if col in df_out.columns:
            df_out[col] = df_out[col].astype(str).str.strip()

    df_out['code4']  = df_out['subj'].str.split().str[0].str.strip()
    df_out['ex_dt']  = df_out['ex_date'].astype(str) + ' ' + df_out['ex_time']
    df_out['st_num'] = (df_out['note'].astype(str)
                        .str.split('/').str[0]
                        .str.strip()
                        .replace('nan', '0')
                        .apply(pd.to_numeric, errors='coerce')
                        .fillna(0).astype(int))

    df_out['teacher_str'] = df_out['teacher_str'].astype(str)
    df_out['teachers']    = df_out['teacher_str'].apply(
        lambda x: clean_name_list(x.split(','))
    )

    return df_out.reset_index(drop=True)

def clean_name_list(name_list):
    cleaned_names = []
    # Define a list of common titles to remove
    # Order matters for specificity: longer/more complex patterns first.
    titles_to_remove = [
        r'ผศ\.\s*ดร\.\s*',  # Assistant Prof. Dr.
        r'รศ\.\s*ดร\.\s*',  # Associate Prof. Dr.
        r'ศ\.\s*ดร\.\s*',   # Prof. Dr.
        r'อาจารย์พิเศษ\s*',
        r'อาจารย์\s*',
        r'ผศ\.\s*',          # Assistant Prof.
        r'รศ\.\s*',          # Associate Prof.
        r'ศ\.\s*',           # Professor
        r'ดร\.\s*',          # Doctor
        r'อ\.\s*',           # Ajarn (Teacher)
        r'นาย\s*',            # Mr.
        r'นาง\s*',            # Mrs.
        r'นส\.\s*',          # Ms.
        r'Dr\.\s*',          # Dr. (English)
        r'TA\s*'             # Teaching Assistant
    ]

    # Create a single regex pattern that matches any of the titles at the beginning of the string.
    # The (?:...)+ ensures that multiple titles (e.g., 'ผศ. ดร.') are matched and removed.
    title_pattern = r'^(?:' + '|'.join(titles_to_remove) + r')+'

    for name in name_list:
        # Use re.sub to replace the matched titles with an empty string
        cleaned_name = re.sub(title_pattern, '', name).strip()
        # Replace multiple spaces with a single space, then strip again to remove any new leading/trailing spaces
        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
        cleaned_name = cleaned_name.lower()
        cleaned_names.append(cleaned_name)
    return cleaned_names
