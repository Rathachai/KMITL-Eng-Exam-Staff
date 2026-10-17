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

# save

with pd.ExcelWriter(out_path('for-cheking.xlsx')) as w:
    dfZ.to_excel(w, sheet_name="DATA", index=False)
    g1.to_excel(w, sheet_name="by_staff")
    g2.to_excel(w, sheet_name="by_room_exdur_role")
    g3.to_excel(w, sheet_name="by_subj")
    g4.to_excel(w, sheet_name="vip_only")
    g5.to_excel(w, sheet_name="all_staffs")
 
    


