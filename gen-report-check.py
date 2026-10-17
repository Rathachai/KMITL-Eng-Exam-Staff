# -*- coding: utf-8 -*-
"""รายงานตรวจสอบ -> reports/for-cheking.xlsx"""

import openpyxl

from common import *   # pd, np, re, os, datetime, config values, helpers

_state = load_state()
dfX, dfST = _state["dfX"], _state["dfST"]

######## FOR CHECKING ############

dfx_cols = ['ID', 'subj', 'sec', 'program_type', 
       'teacher_str', 'building', 'room', 'note', 'ex_date', 'ex_time', 'ex_dur', 
       'ex_dt', 'st_num', 'staff_index', 'reason', 'Remark']
dfZ = dfX[dfx_cols].copy()
dfZ['staff_role'] = dfZ['staff_index'].apply(lambda x: dfST.loc[x, 'role'] if x in dfST.index else '')
dfZ['staff_dept'] = dfZ['staff_index'].apply(lambda x: dfST.loc[x, 'dept'] if x in dfST.index else '')

#dfZ.to_excel(out_path('for-cheking.xlsx'), index=False)


# วันซ้ำ
g1 = dfZ.groupby('staff_index')[['ex_dt']].agg(['count', 'nunique']).sort_values(('ex_dt', 'count'), ascending=False)

# คนที่คุมเยอะ
dfZ.groupby('staff_index')[['ex_dt']].count().sort_values('ex_dt', ascending=False).head(20)

# คนที่คุมน้อย
dfZ.groupby('staff_index')[['ex_dt']].count().sort_values('ex_dt', ascending=True).head(20)


# จำนวนที่ผูบริหารคุม
g4 = dfZ[dfZ['staff_index'].isin(dfST.loc[dfST['role']=='ผู้บริหาร'].index)].groupby('staff_index')[['ex_dt']].count().sort_values('ex_dt', ascending=False)


# จำนวนคนคุมต่อวิชา
g3 = dfZ.groupby(['subj', 'ex_dt', 'room'])[['staff_index']].count().sort_values('staff_index', ascending=False)

# ห้องต่อคนคุม
#g2 = dfZ.groupby(['room', 'ex_dt'])[['staff_index']].count().sort_values('staff_index')
g2 = pd.pivot_table(dfZ, index=['room','ex_dt'], columns='staff_role', values='staff_index', aggfunc='count', fill_value=0, margins=True, margins_name='รวม').reset_index().pipe(lambda d: d[['room','ex_dt','รวม'] + [c for c in d.columns if c not in ('room','ex_dt','รวม')]]).sort_values('รวม', ascending=False)
g2['จำนวนอาจารย์'] = g2['ผู้บริหาร'] + g2['หัวหน้าภาค'] + g2['อาจารย์']
g2.sort_values('จำนวนอาจารย์', inplace=True)

## รวมคน
g5 = dfST[[ 'dept','role','lang', 'todo', 'ex_dt']].copy()
g5['num_ex'] = g5['ex_dt'].apply(len)
g5['ex_dt'] = g5['ex_dt'].str.join(', ')
g5.sort_values('num_ex', inplace=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHECKS — ตรวจความถูกต้องของการจัดตาราง  (ผลรวมที่ชีต "สรุปตรวจสอบ")
# ══════════════════════════════════════════════════════════════════════════════

MAJORITY = 0.5          # "ส่วนใหญ่" = อย่างน้อย 50%

dfoff = _state.get("dfoff")
if dfoff is None:       # state รุ่นเก่าไม่มี staff_cannot -> อ่านจาก master
    _mp = os.environ.get('EXAM_MASTER_PATH') or os.path.join(DATA_DIR, MASTER_FILE)
    dfoff = pd.read_excel(_mp, sheet_name="staff_cannot")

dfST_u = dfST[~dfST.index.duplicated(keep='first')]
R = dfX.copy()
R['si'] = R['staff_index'].fillna('').astype(str).str.strip()
R['assigned'] = R['si'] != ''
R['room_key'] = R['building'].astype(str).str.strip() + ' | ' + R['room'].astype(str).str.strip()
R['ex_date_s'] = R['ex_date'].astype(str)
R['s_role'] = R['si'].map(dfST_u['role'].astype(str).str.strip()).fillna('')
R['s_main'] = R['si'].map(dfST_u['main_role']).fillna('')
R['s_lang'] = R['si'].map(dfST_u['lang']).fillna('')
R['s_dept'] = R['si'].map(dfST_u['dept']).fillna('')
R['s_group'] = R['si'].map(dfST_u['group_dept']).fillna('')
R['is_teacher'] = R['s_main'] == 'teacher'
R['own_subj'] = [s in t if s else False for s, t in zip(R['si'], R['teachers'])]
R['own_dept'] = R['is_teacher'] & (R['s_dept'] == R['dept'])
RA = R[R['assigned']]
ROW_COLS = ['ID', 'ex_dt', 'subj', 'sec', 'program_type', 'dept', 'building', 'room']
summary = []          # (รายการ, เกณฑ์, ผลที่พบ, ผ่าน?, ชีต)


def add(item, rule, found, ok, sheet=''):
    summary.append([item, rule, found, 'ผ่าน' if ok else 'ไม่ผ่าน', sheet])


def names_of(g):
    return ', '.join(sorted(g['si'][g['si'] != '']))


# 1) ทุก row มีผู้คุมสอบ -------------------------------------------------------
c1 = R.loc[~R['assigned'], ROW_COLS + ['reason']]
add('ทุก row ถูก assign ผู้คุมสอบ', 'ไม่มี row ที่ว่าง', f'{len(c1)} / {len(R)} row ว่าง', c1.empty, 'ck1_ไม่มีผู้คุม')

# 2) ทุกห้อง-เวลา มีอาจารย์/ผู้บริหารอย่างน้อย 1 คน --------------------------------
rooms = R.groupby(['ex_dt', 'room_key'])
c2 = rooms.apply(lambda g: pd.Series({
        'subj': ' / '.join(sorted(set(g['subj'].astype(str)))),
        'จำนวนผู้คุม': int(g['assigned'].sum()),
        'อาจารย์/ผู้บริหาร': int(g['is_teacher'].sum()),
        'ผู้คุม': names_of(g)})).reset_index()
c2 = c2[c2['อาจารย์/ผู้บริหาร'] == 0]
add('ทุกห้อง-เวลา มีอาจารย์หรือผู้บริหาร ≥ 1 คน', 'main_role = teacher (อาจารย์/ผู้บริหาร/หัวหน้าภาค)',
    f'{len(c2)} / {rooms.ngroups} ห้อง-เวลา ไม่ผ่าน', c2.empty, 'ck2_ห้องไม่มีอาจารย์')

# 3) วิชา Inter: ต้องมี lang=en หรืออาจารย์ ≥ 1 คนในห้อง -------------------------------
inter = R[R['program_type'] == 'Inter'].groupby(['ex_dt', 'room_key'])
c3 = inter.apply(lambda g: pd.Series({
        'subj': ' / '.join(sorted(set(g['subj'].astype(str)))),
        'มี lang=en': bool((g['s_lang'] == 'en').any()),
        'มีอาจารย์': bool(g['is_teacher'].any()),
        'ผู้คุม': names_of(g)})).reset_index()
c3 = c3[~(c3['มี lang=en'] | c3['มีอาจารย์'])]
add('วิชา Inter มี lang=en หรืออาจารย์ ≥ 1 คนในห้อง', 'ทุกห้อง-เวลาของวิชา Inter',
    f'{len(c3)} / {inter.ngroups} ห้อง-เวลา ไม่ผ่าน', c3.empty, 'ck3_Inter')

# 4-5) อาจารย์คุมวิชาตัวเอง / ภาคตัวเอง ---------------------------------------------
T = RA[RA['is_teacher']]
c45 = T.groupby('si').agg(role=('s_role', 'first'), dept=('s_dept', 'first'), คุมทั้งหมด=('si', 'size'),
                          วิชาตัวเอง=('own_subj', 'sum'), ภาคตัวเอง=('own_dept', 'sum')).reset_index()
c45['วิชาตัวเอง%'] = (100 * c45['วิชาตัวเอง'] / c45['คุมทั้งหมด']).round(1)
c45['ภาคตัวเอง%'] = (100 * c45['ภาคตัวเอง'] / c45['คุมทั้งหมด']).round(1)
c45_dept = c45.groupby('dept').agg(คน=('si', 'size'), คุมทั้งหมด=('คุมทั้งหมด', 'sum'),
                                   วิชาตัวเอง=('วิชาตัวเอง', 'sum'), ภาคตัวเอง=('ภาคตัวเอง', 'sum')).reset_index()
c45_dept['วิชาตัวเอง%'] = (100 * c45_dept['วิชาตัวเอง'] / c45_dept['คุมทั้งหมด']).round(1)
c45_dept['ภาคตัวเอง%'] = (100 * c45_dept['ภาคตัวเอง'] / c45_dept['คุมทั้งหมด']).round(1)
p_dept = T['own_dept'].mean() if len(T) else 0

# 4) รายวิชา: แต่ละ section-ห้อง-เวลา มีผู้สอนวิชานั้นอยู่ในกลุ่มผู้คุมหรือไม่ -------------
def own_teacher_names(g):
    return ', '.join(sorted(set(g.loc[g['own_subj'], 'si'])))

# เฉพาะผู้สอนที่อยู่ใน master และ todo=yes (ผู้สอนภายนอก/ไม่คุม จะไม่ถูกนับ)
eligible = set(dfST_u.index[dfST_u['todo'] == True])
sec_g = R.groupby(['subj', 'sec', 'ex_dt', 'room_key'])
c4_sec = sec_g.apply(lambda g: pd.Series({
        'ผู้สอน': ', '.join(sorted({t for ts in g['teachers'] for t in ts if t in eligible})),
        'ผู้คุม': names_of(g),
        'มีผู้สอนเป็นผู้คุม': bool(g['own_subj'].any()),
        'ผู้สอนที่คุม': own_teacher_names(g)})).reset_index()
n_sec_all = len(c4_sec)
c4_sec = c4_sec[c4_sec['ผู้สอน'] != '']                    # ตัดห้องที่ไม่มีผู้สอนที่คุมได้
c4_subj = c4_sec.groupby('subj').agg(ห้อง_เวลา=('sec', 'size'), มีผู้สอนคุม=('มีผู้สอนเป็นผู้คุม', 'sum')).reset_index()
c4_subj['ไม่มีผู้สอนคุม'] = c4_subj['ห้อง_เวลา'] - c4_subj['มีผู้สอนคุม']
c4_subj['สัดส่วน%'] = (100 * c4_subj['มีผู้สอนคุม'] / c4_subj['ห้อง_เวลา']).round(1)
c4_subj = c4_subj.sort_values(['สัดส่วน%', 'subj'])
p_subj = c4_sec['มีผู้สอนเป็นผู้คุม'].mean() if len(c4_sec) else 0
n_subj_ok = int((c4_subj['สัดส่วน%'] >= 100 * MAJORITY).sum())
add('วิชาส่วนใหญ่ได้ผู้สอนของวิชานั้นเป็นผู้คุม', f'≥ {MAJORITY:.0%} ของ section-ห้อง-เวลา (เฉพาะห้องที่ผู้สอนอยู่ใน master และ todo=yes)',
    f'{p_subj:.1%} ({int(c4_sec["มีผู้สอนเป็นผู้คุม"].sum())}/{len(c4_sec)} ห้อง-เวลา; ตัด {n_sec_all - len(c4_sec)} ห้องที่ไม่มีผู้สอนคุมได้) | {n_subj_ok}/{len(c4_subj)} วิชา ≥ {MAJORITY:.0%}',
    p_subj >= MAJORITY, 'ck4_รายวิชา')
add('อาจารย์ส่วนใหญ่คุมวิชาของภาควิชาตนเอง', f'≥ {MAJORITY:.0%} ของ duty อาจารย์ทั้งหมด',
    f'{p_dept:.1%} ({int(T["own_dept"].sum())}/{len(T)} duty)', p_dept >= MAJORITY, 'ck4-5_วิชา-ภาคตัวเอง')

# 6) คุมมากกว่า 1 ที่ในเวลาเดียวกัน ---------------------------------------------------
g6 = RA.groupby(['si', 'ex_dt'])['room_key'].agg(['nunique', lambda s: ' ; '.join(sorted(set(s)))])
g6.columns = ['จำนวนห้อง', 'ห้อง']
c6 = g6[g6['จำนวนห้อง'] > 1].reset_index()
add('ไม่มีใครคุมมากกว่า 1 ที่ในเวลาเดียวกัน', 'ห้องไม่ซ้ำ ต่อ คน ต่อ ex_dt', f'{len(c6)} กรณี', c6.empty, 'ck6_คุมซ้อน')

# 7) การกระจายของแต่ละภาค ตามวันสอบ -------------------------------------------------
pv_date = pd.pivot_table(RA, index='s_group', columns='ex_date_s', values='si', aggfunc='count',
                         fill_value=0, margins=True, margins_name='รวม')
pv_date = pv_date.drop(columns='รวม')      # column ละวันเท่านั้น (เก็บแถว รวม ไว้)
pv_date.index.name = 'group_dept'
pv_date.columns.name = None

add('การกระจาย duty ของแต่ละภาคตามวันสอบ', 'ดูในชีต (ไม่มีเกณฑ์ผ่าน/ไม่ผ่าน)', f'{len(pv_date) - 1} กลุ่ม × {pv_date.shape[1]} วัน', True,
    'ck7_กระจายตามวัน')

# 8) คุมเกิน/น้อยกว่าเกณฑ์ (load = โควตาของคน) --------------------------------------
duty = RA.groupby('si').size()
c8 = dfST_u[['title', 'first', 'last', 'dept', 'role', 'todo', 'load']].copy()
c8['คุมจริง'] = duty.reindex(c8.index).fillna(0).astype(int)
c8['ส่วนต่าง'] = c8['คุมจริง'] - c8['load']
c8['สถานะ'] = np.where(c8['ส่วนต่าง'] > 0, 'เกิน', np.where(c8['ส่วนต่าง'] < 0, 'น้อยกว่า', 'ครบ'))
c8 = c8[(c8['สถานะ'] != 'ครบ') & ((c8['todo'] == True) | (c8['คุมจริง'] > 0))]
c8 = c8.reset_index().rename(columns={'sindex': 'ชื่อ', 'index': 'ชื่อ'}).sort_values(['สถานะ', 'dept'])
n_over, n_under = (c8['สถานะ'] == 'เกิน').sum(), (c8['สถานะ'] == 'น้อยกว่า').sum()
add('จำนวนครั้งที่คุม ตรงตามเกณฑ์ (load)', 'คุมจริง = load ของแต่ละคน (todo=yes)', f'เกิน {n_over} คน, น้อยกว่า {n_under} คน',
    c8.empty, 'ck8_เกิน-น้อยกว่าเกณฑ์')

# 9) คุมตรงวันลา (staff_cannot) ------------------------------------------------------
off = dfoff.dropna(subset=['fullname', 'cannot']).copy()
off['si'] = off['fullname'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip().str.lower()
off['ex_dt'] = off['cannot'].astype(str).str.strip()
c9 = RA.merge(off[['si', 'ex_dt']].drop_duplicates(), on=['si', 'ex_dt'])[['si', 'ex_dt', 'subj', 'sec', 'room_key']]
add('ไม่คุมสอบตรงวันลา (staff_cannot)', f'ตรวจ {len(off)} รายการวันลา', f'{len(c9)} กรณีชนวันลา', c9.empty, 'ck9_ชนวันลา')

# 10) ผู้บริหาร ----------------------------------------------------------------------
ex_ids = dfST_u[dfST_u['role'].astype(str).str.strip() == 'ผู้บริหาร']
EX = RA[RA['s_role'] == 'ผู้บริหาร']
c10 = ex_ids[['title', 'first', 'last', 'dept', 'todo', 'load']].copy()
c10['คุมจริง'] = EX.groupby('si').size().reindex(c10.index).fillna(0).astype(int)
c10['คุมวิชาตัวเอง'] = EX.groupby('si')['own_subj'].sum().reindex(c10.index).fillna(0).astype(int)
c10['สอนวิชาในรอบนี้ (row)'] = [sum(i in t for t in R['teachers']) for i in c10.index]
c10['คุมตรงวันลา'] = c9.groupby('si').size().reindex(c10.index).fillna(0).astype(int)
c10['คุมซ้อน'] = c6.groupby('si').size().reindex(c10.index).fillna(0).astype(int) if len(c6) else 0
c10['สถานะ'] = np.where(c10['คุมจริง'] == c10['load'], 'ครบตามเกณฑ์',
                        np.where(c10['คุมจริง'] > c10['load'], 'เกิน', 'น้อยกว่า'))
c10 = c10.reset_index().rename(columns={'sindex': 'ชื่อ', 'index': 'ชื่อ'})
c10_det = EX[['si', 'ex_dt', 'subj', 'sec', 'room_key', 'own_subj', 'reason']].sort_values(['si', 'ex_dt'])
bad10 = c10[(c10['todo'] == True) & ((c10['สถานะ'] != 'ครบตามเกณฑ์') | (c10['คุมตรงวันลา'] > 0) | (c10['คุมซ้อน'] > 0))]
add('ผู้บริหารคุมตามเกณฑ์', 'จำนวนที่คุม = load (= min(วิชาที่สอน, โควตา)), ไม่ชนวันลา/ไม่คุมซ้อน',
    f'{len(bad10)} / {int((c10["todo"] == True).sum())} คน ไม่ตรงเกณฑ์', bad10.empty, 'ck10_ผู้บริหาร')

# ── Save ─────────────────────────────────────────────────────────────────────

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

with pd.ExcelWriter(out_path('for-cheking.xlsx'), engine='openpyxl') as w:
    sm = pd.DataFrame(summary, columns=['รายการตรวจ', 'เกณฑ์', 'ผลที่พบ', 'สถานะ', 'ดูรายละเอียดชีต'])
    sm.insert(0, 'ข้อ', range(1, len(sm) + 1))
    sm.to_excel(w, sheet_name='สรุปตรวจสอบ', index=False)
    c1.to_excel(w, sheet_name='ck1_ไม่มีผู้คุม', index=False)
    c2.to_excel(w, sheet_name='ck2_ห้องไม่มีอาจารย์', index=False)
    c3.to_excel(w, sheet_name='ck3_Inter', index=False)
    c4_subj.to_excel(w, sheet_name='ck4_รายวิชา', index=False)
    c4_sec.to_excel(w, sheet_name='ck4_รายวิชา', index=False, startrow=len(c4_subj) + 3)
    c45_dept.to_excel(w, sheet_name='ck4-5_วิชา-ภาคตัวเอง', index=False)
    c45.to_excel(w, sheet_name='ck4-5_วิชา-ภาคตัวเอง', index=False, startrow=len(c45_dept) + 3)
    c6.to_excel(w, sheet_name='ck6_คุมซ้อน', index=False)
    pv_date.to_excel(w, sheet_name='ck7_กระจายตามวัน')
    c8.to_excel(w, sheet_name='ck8_เกิน-น้อยกว่าเกณฑ์', index=False)
    c9.to_excel(w, sheet_name='ck9_ชนวันลา', index=False)
    c10.to_excel(w, sheet_name='ck10_ผู้บริหาร', index=False)
    c10_det.to_excel(w, sheet_name='ck10_ผู้บริหาร', index=False, startrow=len(c10) + 3)
    dfZ.to_excel(w, sheet_name="DATA", index=False)
    g1.to_excel(w, sheet_name="by_staff")
    g2.to_excel(w, sheet_name="by_room_exdur_role")
    g3.to_excel(w, sheet_name="by_subj")
    g4.to_excel(w, sheet_name="vip_only")
    g5.to_excel(w, sheet_name="all_staffs")

    # ปรับหน้าตา: หัวตารางหนา, ความกว้างคอลัมน์, สีผ่าน/ไม่ผ่าน
    for ws in w.sheets.values():
        for idx, col in enumerate(ws.iter_cols(), start=1):
            width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 8), 60)
    ws = w.sheets['สรุปตรวจสอบ']
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = 'A2'
    for row in ws.iter_rows(min_row=2):
        cell = row[4]
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='2E7D32' if cell.value == 'ผ่าน' else 'C62828')
        cell.alignment = Alignment(horizontal='center')

print(sm.to_string(index=False))
