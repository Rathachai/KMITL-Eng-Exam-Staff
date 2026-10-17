# -*- coding: utf-8 -*-
"""หลักฐานการจ่ายเงินค่ากรรมการคุมสอบ -> reports/for-payment.xlsx"""

import openpyxl

from common import *   # pd, np, re, os, datetime, config values, helpers

_state = load_state()
extype = _state["extype"]
dfX, dfST = _state["dfX"], _state["dfST"]

# ─────────────────────────────── EXPORT SIGN SHEET (รับ-ส่งข้อสอบ) ──────────

print('########## EXPORT PAYMENT SHEET (หลักฐานการจ่ายเงินค่ากรรมการคุมสอบ)  ##########')
# ─────────────────────────────── EXPORT PAYMENT SHEET (หลักฐานการจ่ายเงิน) ───


# ── Settings: ดู config.py (PAY_ORDER_TEXT, PAY_HOURS, PAY_AMOUNT, SHOW_TOTAL, FONT_SIZE)

PAY_COLS_TH = ['ลำดับ', 'ชื่อ - สกุล', 'ห้องสอบ', 'เวลา', 'จำนวน ชม.',
               'จำนวนเงิน', 'ลายมือชื่อผู้ปฏิบัติงาน', 'ลายมือชื่อผู้รับเงิน', 'วัน/เดือน/ปี']

PAY_COL_WIDTHS = {
    'ลำดับ': 8, 'ชื่อ - สกุล': 36, 'ห้องสอบ': 26, 'เวลา': 16, 'จำนวน ชม.': 12,
    'จำนวนเงิน': 14, 'ลายมือชื่อผู้ปฏิบัติงาน': 26, 'ลายมือชื่อผู้รับเงิน': 26,
    'วัน/เดือน/ปี': 18,
}

THAI_DAYS_FULL = {0: 'จันทร์', 1: 'อังคาร', 2: 'พุธ', 3: 'พฤหัสบดี',
                  4: 'ศุกร์', 5: 'เสาร์', 6: 'อาทิตย์'}
THAI_MONTHS_FULL = {1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
                    5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
                    9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'}


# ── Master lookup for titles (match by name without title, case-insensitive) ──
# Longest prefixes first, compared in lower case; stacked titles such as
# 'ผศ. ดร.' are removed one prefix at a time by strip_title().
TITLE_PREFIXES = sorted([
    'ศ.', 'รศ.', 'ผศ.', 'ดร.', 'อ.', 'นางสาว', 'นาง', 'นาย', 'น.ส.',
    'asst. prof.', 'assoc. prof.', 'prof.', 'dr.', 'mr.', 'mrs.', 'ms.',
], key=len, reverse=True)

def strip_title(s) -> str:
    """'อ. อมตะ อนันต์พินิจวัฒนา' → 'อมตะ อนันต์พินิจวัฒนา'"""
    s = ' '.join(str(s).split())
    changed = True
    while changed:                          # repeat for stacked titles
        changed = False
        low = s.casefold()
        for p in TITLE_PREFIXES:
            if low.startswith(p):
                s = s[len(p):].strip()
                changed = True
                break
    return s

def norm_key(s) -> str:
    """Matching key: no title, no extra spaces, lower-case."""
    return strip_title(s).casefold()

dfST_k       = dfST.copy()
dfST_k.index = dfST_k.index.map(norm_key)            # master index → same key format
dfST_k       = dfST_k[~dfST_k.index.duplicated(keep='first')]

def staff_with_title(name) -> str:
    """Display name = master 'title' + name as written in dfX (original case)."""
    key = norm_key(name)
    if key in dfST_k.index:
        t = dfST_k.at[key, 'title']
        t = '' if pd.isna(t) else str(t).strip()
        return f"{t} {strip_title(name)}".strip()
    return ' '.join(str(name).split())

_names     = dfX['staff'].dropna().unique()
_unmatched = [n for n in _names if norm_key(n) not in dfST_k.index]
print(f"master match: {len(_names) - len(_unmatched)} / {len(_names)}")
print("unmatched:", _unmatched)


def build_header_lines_pay(ex_date, session: str, building: str) -> list[tuple[str, str]]:
    dt    = pd.Timestamp(ex_date)
    sem   = exam_semester(ex_date)                      # e.g. '1/2569'
    etype = exam_type_th(extype)
    date_full = (f"วัน{THAI_DAYS_FULL[dt.weekday()]}ที่ {dt.day} "
                 f"{THAI_MONTHS_FULL[dt.month]} {dt.year + 543}")
    return [
        (f"หลักฐานการจ่ายเงินค่ากรรมการคุมสอบ{etype} ประจำภาคเรียนที่ {sem}", "center"),
        (PAY_ORDER_TEXT,                                                        "center"),
        ("ข้าพเจ้าผู้มีรายนามข้างท้ายนี้ได้รับเงินจากคณะวิศวกรรมศาสตร์ สจล."
         "ไปเป็นการถูกต้องแล้ว จึงลงลายมือชื่อไว้เป็นสำคัญ",                    "center"),
        (date_full,                                                             "center"),
        (f"ตึก {building} {session}",                                           "left"),
    ]


def make_pay_sheet(ws, df_slot: pd.DataFrame, ex_date, session: str, building: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.page import PageMargins

    LONG_ROOM_SIZE = FONT_SIZE - 3           # smaller font for long room names

    ncols       = len(df_slot.columns)
    header_rows = build_header_lines_pay(ex_date, session, building)
    n_header    = len(header_rows)

    thin  = Side(style='thin',   color='000000')
    thick = Side(style='medium', color='000000')     # use 'thick' for even bolder
    full  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Header text block ─────────────────────────────────────────────────────
    for i, (text, align) in enumerate(header_rows, start=1):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=ncols)
        cell           = ws.cell(row=i, column=1, value=text)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=FONT_SIZE)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        ws.row_dimensions[i].height = 28

    # ── Column headers (thick top + bottom border) ────────────────────────────
    col_row  = n_header + 1
    hdr_bord = Border(left=thin, right=thin, top=thick, bottom=thick)
    for c, name in enumerate(df_slot.columns, start=1):
        cell           = ws.cell(row=col_row, column=c, value=name)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=FONT_SIZE)
        cell.fill      = PatternFill('solid', start_color='FFFFFF')
        cell.border    = hdr_bord
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[col_row].height = 32

    # ── Last row of each room group (thick bottom border) ─────────────────────
    ROW_HEIGHT     = 46          # tall enough for a 2-line wrapped name at size 16
    data_start_row = col_row + 1
    CENTER_COLS    = {'ลำดับ', 'ห้องสอบ', 'เวลา', 'จำนวน ชม.', 'จำนวนเงิน'}

    rooms          = df_slot['ห้องสอบ'].tolist()
    room_last_rows = {
        data_start_row + i
        for i, rm in enumerate(rooms)
        if i == len(rooms) - 1 or rooms[i + 1] != rm
    }

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r, row in enumerate(df_slot.itertuples(index=False), start=data_start_row):
        ws.row_dimensions[r].height = ROW_HEIGHT
        border = Border(
            left=thin, right=thin,
            top=thick if r == data_start_row else thin,        # meets header's thick bottom
            bottom=thick if r in room_last_rows else thin,     # separates rooms
        )
        for c, value in enumerate(row, start=1):
            col_name    = df_slot.columns[c - 1]
            cell        = ws.cell(row=r, column=c, value=value)
            cell.border = border

            # long room names (e.g. 'ห้องคอมพิวเตอร์ E12-201') → smaller font
            size = FONT_SIZE
            if col_name == 'ห้องสอบ' and len(str(value)) > 12:
                size = LONG_ROOM_SIZE
            cell.font = Font(name='TH Sarabun New', size=size)

            if col_name == 'จำนวนเงิน':
                cell.number_format = '#,##0.00'

            cell.alignment = Alignment(
                horizontal = 'center' if col_name in CENTER_COLS else 'left',
                vertical   = 'center',
                wrap_text  = (col_name == 'ชื่อ - สกุล'),   # word wrap for names
            )

    last_data_row = data_start_row + len(df_slot) - 1

    # ── Optional total row ────────────────────────────────────────────────────
    if SHOW_TOTAL and len(df_slot) > 0:
        tr    = last_data_row + 1
        amt_c = list(df_slot.columns).index('จำนวนเงิน') + 1
        ws.row_dimensions[tr].height = 32
        ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=amt_c - 1)
        for c in range(1, ncols + 1):
            ws.cell(row=tr, column=c).border = full
        lbl           = ws.cell(row=tr, column=1, value='รวม')
        lbl.font      = Font(name='TH Sarabun New', bold=True, size=FONT_SIZE)
        lbl.alignment = Alignment(horizontal='right', vertical='center')
        col_l         = get_column_letter(amt_c)
        tot           = ws.cell(row=tr, column=amt_c,
                                value=f"=SUM({col_l}{data_start_row}:{col_l}{last_data_row})")
        tot.font          = Font(name='TH Sarabun New', bold=True, size=FONT_SIZE)
        tot.number_format = '#,##0.00'
        tot.alignment     = Alignment(horizontal='center', vertical='center')

    # ── Column widths ─────────────────────────────────────────────────────────
    for c, name in enumerate(df_slot.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = PAY_COL_WIDTHS.get(name, 14)

    # ── Print setup ───────────────────────────────────────────────────────────
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_margins           = PageMargins(left=0.5, right=0.5, top=0.75,
                                            bottom=0.75, header=0.3, footer=0.3)
    ws.print_title_rows       = f'1:{col_row}'
    ws.oddFooter.right.text   = 'หน้า &P/&N'
    ws.oddFooter.right.size   = 12
    ws.oddFooter.right.font   = 'TH Sarabun New'


# ── Build workbook ────────────────────────────────────────────────────────────
wb_pay = openpyxl.Workbook()
wb_pay.remove(wb_pay.active)

all_dates     = sorted(dfX['ex_date'].dropna().unique())
all_buildings = sorted(dfX['building'].dropna().unique())

for ex_date in all_dates:
    df_date = dfX[dfX['ex_date'] == ex_date]

    for hour_key, session, session_en in [('09', 'เช้า', 'am'), ('13', 'บ่าย', 'pm')]:
        df_s = df_date[df_date['ex_time'] == hour_key]
        if df_s.empty:
            continue

        pay_time = SESSION_DURATION.get(hour_key, '').replace(' น.', '')   # '09:30-12:30'

        for building in all_buildings:
            df_b = df_s[df_s['building'] == building].copy()

            # only rows that actually have a proctor; one row per person per session
            df_b = df_b[df_b['staff'].notna() & (df_b['staff'].astype(str).str.strip() != '')]
            df_b = (df_b.sort_values(['room', 'staff'])
                        .drop_duplicates(subset='staff', keep='first')
                        .reset_index(drop=True))
            if df_b.empty:
                continue

            dt         = pd.Timestamp(ex_date)
            month_th   = THAI_MONTHS_SHORT[dt.month].replace('.', '')
            sheet_name = f"{dt.day:02d}-{month_th}-{session_en} {building}"
            for ch in '\\/*?:[]':                      # characters Excel forbids in sheet names
                sheet_name = sheet_name.replace(ch, '-')
            sheet_name = sheet_name[:31]

            n = len(df_b)
            df_out = pd.DataFrame({
                'ลำดับ':                    range(1, n + 1),
                'ชื่อ - สกุล':              df_b['staff'].map(staff_with_title),
                'ห้องสอบ':                  df_b['room'].values,
                'เวลา':                     pay_time,
                'จำนวน ชม.':                PAY_HOURS,
                'จำนวนเงิน':                PAY_AMOUNT,
                'ลายมือชื่อผู้ปฏิบัติงาน':  '',
                'ลายมือชื่อผู้รับเงิน':      '',
                'วัน/เดือน/ปี':             '',
            })

            ws = wb_pay.create_sheet(title=sheet_name)
            make_pay_sheet(ws, df_out, ex_date, session, building)

# Guard: Excel needs at least one visible sheet
if not wb_pay.sheetnames:
    ws = wb_pay.create_sheet("ไม่มีข้อมูล")
    ws['A1'] = "ไม่พบข้อมูลผู้คุมสอบ"

out_pay = out_path('for-payment.xlsx')
wb_pay.save(out_pay)
print(f"Saved: {out_pay}  sheets={len(wb_pay.sheetnames)}")


print(' ******* END JOB ********')





 

