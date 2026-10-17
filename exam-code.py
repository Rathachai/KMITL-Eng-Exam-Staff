# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re
from datetime import datetime
import glob
import random
import os
import pickle


pd.set_option('display.max_columns', None)



######## CONFIG #############################

#PATH_DIR = "/Users/rathachai/Library/CloudStorage/GoogleDrive-rathachai.ch@kmitl.ac.th/My Drive/RC-KMITL/Research/2026 Exam Staff"
_BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(_BASE, "exam-data")   # input: master table + exam schedule xlsx
OUT_DIR  = os.path.join(_BASE, "reports")     # output: results/

THAI_MONTHS_SHORT = {
    1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.',
    5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.',
    9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.',
}

def exam_semester(ex_date) -> str:
    d = pd.Timestamp(ex_date)
    if 7 <= d.month <= 11:
        return f"1/{d.year + 543}"
    else:
        return f"2/{(d.year - 1) + 543}"

def exam_type_th(extype: str) -> str:
    return 'ปลายภาค' if extype == 'final' else 'กลางภาค'

extype = 'final'
#extype = 'mid'


######### FUNCTIONS #############################

def process_exam_report_bkp(file_path):
    # 1. อ่านไฟล์ Excel (แบบไม่ระบุ Header เพื่อให้เห็นโครงสร้างทั้งหมด)
    # ใช้ engine='xlrd' สำหรับ .xls หรือ 'openpyxl' สำหรับ .xlsx
    df_raw = pd.read_excel(file_path, header=None)

    # --- ส่วนที่ 1: ดึงและแปลงวันที่สอบ ---
    def extract_exam_date(df):
        # ค้นหาเซลล์ที่มีคำว่า "วันสอบ"
        mask = df.applymap(lambda x: "วันสอบ" in str(x))
        found_cells = df[mask].dropna(how='all').dropna(axis=1, how='all')

        if found_cells.empty:
            return None

        raw_text = str(found_cells.iloc[0, 0])

        # Regex ดึงวันที่ (1-2 หลัก), เดือน (ภาษาไทย), และปี (2 หลัก)
        #match = re.search(r'(\d{1,2})\s+([ก-ฮ\.]+)\s+(\d{2})', raw_text)
        match = re.search(r'(\d{1,2})\s+([ก-ฮะ-ูเ-์\.]+)\s+(\d{2})', raw_text)

        if match:
            day = int(match.group(1))
            month_th = match.group(2).replace('.', '').strip()
            year_short = int(match.group(3))

            # Mapping เดือนไทย
            months = {
                'มค': 1, 'กพ': 2, 'มีค': 3, 'เมย': 4, 'พค': 5, 'มิย': 6,
                'กค': 7, 'สค': 8, 'กย': 9, 'ตค': 10, 'พย': 11, 'ธค': 12
            }

            month = months.get(month_th)
            # แปลงปี พ.ศ. 2 หลัก เป็น ค.ศ. (พ.ศ. 2569 -> ค.ศ. 2026)
            year = year_short + 2500 - 543
            return datetime(year, month, day).date()
        return None

    exam_date = extract_exam_date(df_raw)

    # --- ส่วนที่ 2: รวมตารางทั้งหมดเข้าด้วยกัน ---
    # ค้นหาตำแหน่งแถวที่เป็น Header (ใช้คำว่า 'ลำดับ' เป็นหลัก)
    header_name = 'ลำดับ'
    header_indices = df_raw[df_raw[0] == header_name].index.tolist()

    if not header_indices:
        print("ไม่พบหัวตารางในไฟล์")
        return None

    # ใช้หัวตารางจากตำแหน่งแรกที่พบ
    cols = df_raw.iloc[header_indices[0]].values

    # ดึงข้อมูลตั้งแต่หัวตารางแรกเป็นต้นไป
    df_main = df_raw.iloc[header_indices[0] + 1:].copy()
    df_main.columns = cols

    # ลบแถวที่ซ้ำกับ Header (กรณีมีหลายหน้า) และลบแถวที่เป็นค่าว่าง
    df_main = df_main[df_main[header_name] != header_name]

    # จัดการกรณี Merged Cells (Forward Fill ข้อมูล ลำดับ และ วิชา)
    # แทนที่ช่องว่าง (string ว่าง) ด้วย NaN เพื่อให้ ffill() ทำงานได้
    df_main = df_main.replace(r'^\s*$', np.nan, regex=True)
    df_main[['ลำดับ', 'วิชา']] = df_main[['ลำดับ', 'วิชา']].ffill()

    # ลบแถวขยะ (เช่น แถวสรุปท้ายหน้า) โดยเช็คจากคอลัมน์ 'กลุ่ม' หรือ 'เวลา' ที่ต้องมีข้อมูล
    df_main = df_main.dropna(subset=['กลุ่ม', 'เวลา'])

    # เพิ่มคอลัมน์วันที่สอบที่ดึงมาได้
    df_main['วันที่สอบ'] = exam_date

    return df_main.reset_index(drop=True)




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

######### READ MASTER #############################

dfst = pd.read_excel(open(DATA_DIR + "/Exam-Master-Table.xlsx", "rb"), sheet_name="exam_staff")
dfst['todo'] = dfst['todo'].map({'yes': True, 'no': False})
dfst['first'].fillna('',inplace=True)
dfst['last'].fillna('',inplace=True)
dfst['fullname'] = dfst['first'].str.strip() + ' ' + dfst['last'].str.strip()
dfst['fullname'] = dfst['fullname'].str.replace(r'\s+', ' ', regex=True)
dfst['fullname'] = dfst['fullname'].str.strip()
dfst['fullname'] = dfst['fullname'].str.lower()
dfst['sindex'] = dfst['fullname'] 
dfst.set_index('sindex', inplace=True)
#dfst['main_dept'] =  dfst['dept'].str.split('-').str[0]
aj_roles = ['อาจารย์', 'ผู้บริหาร', 'หัวหน้าภาค']
sp_roles = ['เจ้าหน้าที่', 'หัวหน้างาน']
dfst.loc[dfst['role'].isin(aj_roles), 'main_role'] = 'teacher'
dfst.loc[dfst['role'].isin(sp_roles), 'main_role'] = 'support'

dfdp = pd.read_excel(open(DATA_DIR + "/Exam-Master-Table.xlsx", "rb"), sheet_name="department", dtype={'code4': str, 0: str})

dfoff = pd.read_excel(open(DATA_DIR + "/Exam-Master-Table.xlsx", "rb"), sheet_name="staff_cannot")
stoff = (dfoff['fullname'] + ' ' + dfoff['cannot']).tolist()

dfslot = pd.read_excel(open(DATA_DIR + "/Exam-Master-Table.xlsx", "rb"), sheet_name="staff_slot")
dfslot['self_subj'] = dfslot['self_subj'].map({'yes': True, 'no': False})
dfslot.set_index('role', inplace=True)


######### READ EXAM TIME #############################

'''
file_name = DATA_DIR + '/17 บ่าย.xls'
df_final = process_exam_report(file_name)
df_final.columns = ['order', 'subj', 'sec', 'st_year', 'st_num', 'ex_time', 'teachers', 'building', 'room', 'note', 'ex_date']

df_final['code4'] = df_final['subj'].str[:5]
'''

fnames = glob.glob(DATA_DIR + '/ร*.xlsx')

df_ = []

for fname in fnames:
    _df = process_exam_report(fname)
    df_.append(_df)

dfx = pd.concat(df_)
#dfx.columns = ['order', 'subj', 'sec', 'st_year', 'st_max', 'ex_dur', 'teacher_str', 'building', 'room', 'note', 'ex_date']
'''
dfx.columns = ['eid', 'ex_date_th', 'ex_dur_str', 'order', 'subj', 'sec', 'st_year', 'st_max',
       'teacher_str', 'building', 'room', 'note', 'Remark', 'คำนำหน้า', 'ชื่อ',
       'นามสกุล', 'หมายเหตุ.1', 'ให้ภาค', 'Unnamed: 18', 'Unnamed: 19',
       'Unnamed: 20', 'ex_date', 'ex_dur', 'ex_time', 'code4', 'ex_dt',
       'st_num', 'teachers']
'''

dfx['subj'] =  dfx['subj'].str.strip()
dfx['code4'] = dfx['subj'].str[:4]
dfx['ex_time'] = dfx['ex_dur'].str[:2]
dfx['st_num'] = dfx['note'].str.split('/').str[0].fillna(0).astype(int)
dfx['ex_dt'] = dfx['ex_date'].astype(str) + ' ' + dfx['ex_time']

dfx['teacher_str'] = dfx['teacher_str'].astype(str)
dfx['teachers'] = dfx['teacher_str'].apply(lambda x: x.split(','))
dfx['teachers'] = dfx['teachers'].apply(lambda x: clean_name_list(x))

dfx.dropna(subset=['room'], inplace=True)    


dfx.reset_index(drop=True, inplace=True)



'''
# test
inames = list(set(dfx['teachers'].sum()))
tnames = dfst['fullname'].tolist()

for iname in inames:
    if iname in tnames:
        pass
    else:
        print(iname)

'; '.join(nonnames)
'''


####### DEPT LOAD #########################

# department loads
depts = dfdp.groupby('code4')['department'].first()
deptstypes = dfdp.groupby('code4')['program_type'].first()

dfx['dept'] = dfx['code4'].apply(lambda x: depts[x] if x  in depts else 'none')
dfx['program_type'] = dfx['code4'].apply(lambda x: deptstypes[x] if x  in deptstypes else 'none')

exload = dfx.groupby('dept')['order'].count()

# init load

dfst['load'] = 0

# ajarn loads
role = 'อาจารย์'
dfst.loc[(dfst['role']==role) & (dfst['todo']==True), 'load'] = dfslot.loc[role, extype+'_num']


# head dept loads

role = 'หัวหน้าภาค'
dfst.loc[(dfst['role']==role) & (dfst['todo']==True), 'load'] = dfslot.loc[role, extype+'_num']



# vip loads
st_subj = dfx['teachers'].sum()
role = 'ผู้บริหาร'
vipcon = (dfst['role']==role) & (dfst['todo']==True)
dfst.loc[vipcon,'load'] = dfst.loc[vipcon,'fullname'].apply(lambda x: min(st_subj.count(x), dfslot.loc[role, extype+'_num']))


# support staff loads
role = 'เจ้าหน้าที่'
dfst.loc[(dfst['role']==role), 'load'] = dfslot.loc[role, extype+'_num']

# support staff loads
role = 'หัวหน้างาน'
dfst.loc[(dfst['role']==role), 'load'] = dfslot.loc[role, extype+'_num']


# department load
dpload = dfst[dfst['load']>0].groupby('dept')['load'].sum()


# concat
dfload = pd.DataFrame(data={'dpload':dpload, 'exload':exload}).fillna(0).astype(int)

dfload['dpfree'] = dfload['dpload'] - dfload['exload']



####### CLEAN TEACHER #########################

for index, row in dfx.iterrows():
    teachers = row['teachers']
    
    for teacher in teachers:
        if teacher not in dfst.index:
            dfx.loc[index, 'teachers'].remove(teacher)
            print(teacher)
    
    


####### ASSIGN STAFF #########################


### ASSIGN Teachers

# clean teachers


dfx['staff'] = ''
dfst['free'] = dfst['load']
dfst['ex_dt'] = [list() for x in range(len(dfst))]
dfload['used'] = 0

dfload.sort_values('dpfree', ascending=False, inplace=True)

dplist = dfload.index
dtlist = dfx.sort_values('ex_dt')['ex_dt'].unique()

# for Raw and Extra room

dfx.rename({'ผู้คุมสอบ':'fix_staff'}, inplace=True, axis=1)
dfx['reason'] = ''
dfx['in_same_room'] = 0

raw_assign = dfx['Remark'].isna() | (dfx['Remark'].str.strip() == '')

dfX = dfx[raw_assign].copy()
dfXtra = dfx[~raw_assign].copy()

#Process

dfST = dfst.copy()
#dfX = dfx.copy()


dfLOAD = dfload.copy()

dfLOAD['dpfree2'] = dfLOAD['dpload']


# force assign

dfF = dfX.dropna(subset='fix_staff')
dfF = dfF[dfF['fix_staff'].str.len()>1]

for index, row in dfF.iterrows():
    teacher = row['fix_staff']
    dt = row['ex_dt']
    if (teacher in dfST[dfST['main_role']=='teacher'].sort_values('role').index) \
        and (dfST.loc[teacher, 'todo']==True) \
        and (teacher + ' ' + dt not in stoff) \
        and (dt not in dfST.loc[teacher, 'ex_dt']) :
            
        dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
        dfX.loc[index, "staff_index"] = teacher
        dfX.loc[index, "reason"] = 'บังคับลง'
        #dfX.loc[index, 'staff'] = teacher
        dfST.loc[teacher, 'free'] -= 1
        dfST.loc[teacher, 'ex_dt'].append(dt)
        dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
        dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
        
        #count in same room
        dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
        
        print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                    

print('', '-- END --', len(dfX[dfX['staff']=='']))

# assign อ.ประจำวิชา ในภาคเดียวกัน



print('-- BEGIN --', 'assign teachers,', 'same subject,', 'same department,', 'free exams:', len(dfX[dfX['staff']=='']))

#for dt in dtlist:
    
#dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
#dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('in_same_room')
dfdt = dfX[dfX['staff']==''].sample(frac=1).sort_values('in_same_room')

#dfdt = dfdt.sample(frac=1)

for index, row in dfdt.iterrows():
    
    dt = row['ex_dt']
    
    teachers = row['teachers']
    
    if len(teachers)>1:
        random.shuffle(teachers)
    
    for teacher in teachers:
        
        
        if (teacher in dfST[dfST['main_role']=='teacher'].sort_values('role').index) \
            and (dfST.loc[teacher, 'free']>0) \
            and (dfST.loc[teacher, 'todo']==True) \
            and (dfST.loc[teacher, 'dept'] == dfX.loc[index,'dept']) \
            and (teacher + ' ' + dt not in stoff) \
            and (dt not in dfST.loc[teacher, 'ex_dt']) :
                
            dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
            dfX.loc[index, "staff_index"] = teacher
            dfX.loc[index, "reason"] = 'อ.วิชา ภาคเดียวกัน'
            #dfX.loc[index, 'staff'] = teacher
            dfST.loc[teacher, 'free'] -= 1
            dfST.loc[teacher, 'ex_dt'].append(dt)
            dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
            dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
            
            #count in same room
            dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
            
            print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt']) 
            
            break # when teacher found
            

print('', '-- END --', len(dfX[dfX['staff']=='']))



# assign คนคุมวิชา inter  (ไม่สนในภาค)
for aj_round in range(3):
    print('-- BEGIN --', 'assign teachers,', 'Inter,', 'free exams:', len(dfX[dfX['staff']=='']))
    
    random.shuffle(dtlist)
    
    for dt in dtlist:
        
        #dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['program_type']=='Inter') & (dfX['staff']=='')].sort_values('st_num', ascending=False)
        dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['program_type']=='Inter') & (dfX['staff']=='')].sample(frac=1).sort_values('in_same_room')
    
        #dfdt = dfdt.sample(frac=1)
        
        #loop subjects
        for index, row in dfdt.iterrows():
            
            #loop teacher
            dfST['rnd_order'] = np.random.rand(len(dfST))
            cond = (dfST['main_role']=='teacher') & (dfST['lang']=='en')
            for teacher, strow in dfST[cond].sort_values(['lang', 'free', 'rnd_order'], ascending=False).iterrows():
                
                if  (dfST.loc[teacher, 'free']>0) \
                    and (dfST.loc[teacher, 'todo']==True) \
                    and (teacher + ' ' + dt not in stoff) \
                    and (dt not in dfST.loc[teacher, 'ex_dt']) :
                        
                    # Skip English speaking in Th room
                    if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                        continue
                        
                    dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                    dfX.loc[index, "staff_index"] = teacher
                    dfX.loc[index, "reason"] = 'อ. หรือ lang=en'
                    #dfX.loc[index, 'staff'] = teacher
                    dfST.loc[teacher, 'free'] -= 1
                    dfST.loc[teacher, 'ex_dt'].append(dt)
                    dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
                    dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
                    print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                    
                    #count in same room
                    dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
                    
                    break # when teacher found
    
    print('', '-- END --', len(dfX[dfX['staff']=='']))



# assign อ.ประจำวิชา ไม่สนใจภาค

print('-- BEGIN --', 'assign teachers,', 'same subject,',  'free exams:', len(dfX[dfX['staff']=='']))

#for dt in dtlist:
    
#dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
dfdt = dfX[dfX['staff']==''].sample(frac=1).sort_values('in_same_room')
#dfdt = dfdt.sample(frac=1)

for index, row in dfdt.iterrows():
    dt = row['ex_dt']
    
    teachers = row['teachers']
    
    if len(teachers)>1:
        random.shuffle(teachers)
    
    for teacher in teachers:
        
        
        if (teacher in dfST[dfST['main_role']=='teacher'].sort_values('role').index) \
            and (dfST.loc[teacher, 'free']>0) \
            and (dfST.loc[teacher, 'todo']==True) \
            and (teacher + ' ' + dt not in stoff) \
            and (dt not in dfST.loc[teacher, 'ex_dt']) :
                
            dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
            dfX.loc[index, "staff_index"] = teacher
            dfX.loc[index, "reason"] = 'อ.วิชา คนละภาค'
            #dfX.loc[index, 'staff'] = teacher
            dfST.loc[teacher, 'free'] -= 1
            dfST.loc[teacher, 'ex_dt'].append(dt)
            dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
            dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
            print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
            
            #count in same room
            dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
            
            break # when teacher found
            

print('', '-- END --', len(dfX[dfX['staff']=='']))


# reset ผู้ริหาร
dfST.loc[dfST['role']=='ผู้บริหาร', 'free'] = 0

# assign คนสอนในภาคเดียวกัน
for aj_round in range(3):
    print('-- BEGIN --', 'assign teachers,', 'same department,', 'free exams:', len(dfX[dfX['staff']=='']))
    
    #for dt in dtlist:
        
    #dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
    #dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('in_same_room')
    dfdt = dfX[dfX['staff']==''].sample(frac=1).sort_values('in_same_room')
    
    #dfdt = dfdt.sample(frac=1)
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        dt = row['ex_dt']
        
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[dfST['main_role']=='teacher'].sort_values('rnd_order').iterrows():
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (dfST.loc[teacher, 'todo']==True) \
                and (dfST.loc[teacher, 'dept'] == dfX.loc[index,'dept']) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                #if (len(dfX[dfX['subj']==row['subj']])>1) and (row['st_num']<10):
                #    continue
                
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                dfX.loc[index, "reason"] = 'อ. ภาคเดียวกัน'
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
                dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                #count in same room
                dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
                
                break # when teacher found
                    
    print('', '-- END --', len(dfX[dfX['staff']=='']))
 
        




# assign อาจารย์ที่เหลือ (ไม่สนใจภาค)

for aj_round in range(3):
    print('-- BEGIN --', 'assign teachers,', 'free exams:', len(dfX[dfX['staff']=='']))
    
    #for dt in dtlist[::-1]:
        
    #dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
    #dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('in_same_room')
    dfdt = dfX[dfX['staff']==''].sample(frac=1).sort_values('in_same_room')
    
    #dfdt = dfdt.sample(frac=1)  
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        dt = row['ex_dt']
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[dfST['main_role']=='teacher'].sort_values(['free','rnd_order'], ascending=False).iterrows():
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (dfST.loc[teacher, 'todo']==True) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                dfX.loc[index, "reason"] = 'อ. ที่เหลือ'
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                dfLOAD.loc[dfST.loc[teacher, 'dept'], 'used'] +=1
                dfLOAD.loc[dfST.loc[teacher, 'dept'], 'dpfree2'] -=1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                #count in same room
                dfX.loc[(dfX['ex_dt']==dt)&(dfX['room']==dfX.loc[index, "room"]), 'in_same_room'] +=1
                
                break # when teacher found        
    
    print('', '-- END --', len(dfX[dfX['staff']=='']))

### ASSIGN Support Staffs

dfSLOAD = dfST[(dfST['main_role']=='support')].groupby('dept')[['free']].count()

dfSLOAD.columns = ['num_staff']

#dfSLOAD = dfSLOAD[dfSLOAD['num_staff']>1].copy()

TTB = pd.DataFrame(index=dfSLOAD.index, columns=dtlist)
for dept in dfSLOAD.index:
    TTB.loc[dept] = dfSLOAD.loc[dept, 'num_staff']
    

# assign ฝ่ายสนับสนุนในภาคเดียวกัน
print('-- BEGIN --', 'assign support staffs,', 'same department,', 'free exams:', len(dfX[dfX['staff']=='']))

random.shuffle(dtlist)

for dt in dtlist[::-1]:
    
    dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
    dfdt = dfdt.sample(frac=1)    
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[dfST['main_role']=='support'].sort_values('rnd_order').iterrows():
            
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (TTB.loc[dfST.loc[teacher, 'dept'], dt] > 1) \
                and (dfST.loc[teacher, 'main_dept'] == dfX.loc[index,'dept']) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                dfX.loc[index, "reason"] = 'จนท. วิชาตามภาค'
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                TTB.loc[dfST.loc[teacher, 'dept'], dt] -= 1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                break # when teacher found
                
print('', '-- END --', len(dfX[dfX['staff']=='']))


# assign ฝ่ายสนับสนุน (ไม่สนในภาค)
print('-- BEGIN --', 'assign support staffs,', 'free exams:', len(dfX[dfX['staff']=='']))

random.shuffle(dtlist)


for dt in dtlist[::-1]:
    
    dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
    dfdt = dfdt.sample(frac=1)    
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[dfST['main_role']=='support'].sort_values(['free', 'rnd_order'], ascending=False).iterrows():
            
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (TTB.loc[dfST.loc[teacher, 'dept'], dt] > 1) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                dfX.loc[index, "reason"] = 'จนท. ไม่สน dept'
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                TTB.loc[dfST.loc[teacher, 'dept'], dt] -= 1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                break # when teacher found

print('', '-- END --', len(dfX[dfX['staff']=='']))



# add dummy

print('-- BEGIN --', 'DUMMY', 'assign support staffs,', 'free exams:', len(dfX[dfX['staff']=='']))


# dfXXX = dfX.copy()
# dfX = dfXXX.copy()

"""
dfX['loc_index'] = dfX.index
                                              
#dfdummy = dfX.groupby(['ex_dt', 'room']).agg({'loc_index':'first', 'staff':'count', 'st_num':'sum'}).reset_index()
dfdummy = dfX.groupby(['ex_dt', 'room']).agg({'loc_index':'first', 'staff':'count', 'st_num':'sum', 'staff':'nunique'}).reset_index()

#dfdummyx = dfdummy[((dfdummy['staff']==1) & (dfdummy['st_num']>30)) | (dfdummy['st_num']>75) ].sort_values('st_num', ascending=False)
dfdummyx = dfdummy[(dfdummy['staff']==1) | (dfdummy['st_num']>75) & (dfdummy['staff']<3) ].sort_values(['staff', 'st_num'], ascending=[True, False])


for index, row in dfdummyx.iterrows():
    break
    idx = row['loc_index']
    dummy_row = dfX.loc[idx].copy()
    dummy_row['order'] = 0
    dummy_row['subj'] = "เสริม"
    dummy_row['sec'] = 0
    dummy_row['st_year'] = 0
    dummy_row['st_max'] = 0
    dummy_row['teacher_str'] = ""
    dummy_row['note'] = ""
    dummy_row['code4'] = ""
    dummy_row['st_num'] = 0
    dummy_row['teachers'] = []
    dummy_row['dept'] = "none"
    dummy_row['program_type'] = "none"
    dummy_row['staff'] = ""
    dummy_row['staff_index'] = ""
    
    dfX.loc[len(dfX)] = dummy_row
"""

####### EXTRA
dfX = pd.concat([dfX, dfXtra], ignore_index=True)

# inter only

random.shuffle(dtlist)


for dt in dtlist[::-1]:
    
    dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='') & (dfX['program_type']=='Inter')].sort_values('st_num', ascending=False)
    dfdt = dfdt.sample(frac=1)
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[(dfST['main_role']=='support')&(dfST['lang']=='en')].sort_values(['free', 'rnd_order'], ascending=False).iterrows():
            
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (TTB.loc[dfST.loc[teacher, 'dept'], dt] > 1) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                TTB.loc[dfST.loc[teacher, 'dept'], dt] -= 1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                break # when teacher found


# normal

for dt in dtlist[::-1]:
    
    dfdt = dfX[(dfX['ex_dt']==dt) & (dfX['staff']=='')].sort_values('st_num', ascending=False)
    dfdt = dfdt.sample(frac=1)
    
    #loop subjects
    for index, row in dfdt.iterrows():
        
        #loop teacher
        dfST['rnd_order'] = np.random.rand(len(dfST))
        for teacher, strow in dfST[dfST['main_role']=='support'].sort_values(['free', 'rnd_order'], ascending=False).iterrows():
            
            
            if  (dfST.loc[teacher, 'free']>0) \
                and (TTB.loc[dfST.loc[teacher, 'dept'], dt] > 1) \
                and (teacher + ' ' + dt not in stoff) \
                and (dt not in dfST.loc[teacher, 'ex_dt']) :
                    
                # Skip English speaking in Th room
                if (dfX.loc[index, "program_type"] != 'Inter') and (dfST.loc[teacher, "lang"]=='en'):
                    continue
                    
                dfX.loc[index, "staff"] = ' '.join(dfST.loc[teacher, ['formal_title', 'first', 'last']].astype(str))
                dfX.loc[index, "staff_index"] = teacher
                #dfX.loc[index, 'staff'] = teacher
                dfST.loc[teacher, 'free'] -= 1
                dfST.loc[teacher, 'ex_dt'].append(dt)
                TTB.loc[dfST.loc[teacher, 'dept'], dt] -= 1
                print('     - ', dt, teacher, dfST.loc[teacher, 'ex_dt'])
                
                break # when teacher found

print('', '-- END --', len(dfX[dfX['staff']=='']))


dfX['st_group_dept'] =  dfX['staff_index'].apply(lambda x: dfST.loc[x,'group_dept'] if x in dfST.index else 'none')      



# Save to csv

dfEXST = dfX[['staff_index', 'subj', 'sec', 'ex_date', 'ex_dur', 'room']].sort_values(['staff_index', 'ex_date', 'ex_dur']) 
dfEXST['st_dept'] =  dfEXST['staff_index'].apply(lambda x: dfST.loc[x,'main_dept'] if x in dfST.index else 'none')  


#dfEXST.to_csv(OUT_DIR+'/results/for-checking.csv', index=False, encoding='utf-8-sig')
        
'''
dfST[dfST['load']>dfST['free']]

dfST[(dfST['free']>0)]

dfST[(dfST['free']>0) & (dfST['todo'])]
dfST[(dfST['load']==1) & (dfST['todo'])]

dfST[(dfST['todo']==False)]


dfLOAD

dfST.loc['รัฐชัย ชาวอุทัย']
dfST.loc['รัฐชัย ชาวอุทัย','ex_dt']

dfEXST[dfEXST['staff']=='รัฐชัย ชาวอุทัย']

dfX[dfX['staff']=='']

dfX[dfX['staff']==''].groupby('dept')['order'].count()

'''
########################################################################
# SAVE REPORT
########################################################################

def save_df(df, name: str = 'dfX', directory: str = OUT_DIR) -> str:
    """Pickle df to '{directory}/{name}-yyyy-mm-dd hh-MM-ss.save'; return the path."""
    ts   = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    path = os.path.join(directory, f'{name}-{ts}.save')
    with open(path, 'wb') as f:
        pickle.dump(df, f, protocol=pickle.HIGHEST_PROTOCOL)
    print(f"Saved: {path}")
    return path


def load_latest_df(name: str = 'dfX', directory: str = OUT_DIR):
    """Load the most recent '{name}-*.save' from directory."""
    files = sorted(glob.glob(os.path.join(directory, f'{name}-*.save')))
    if not files:
        raise FileNotFoundError(f"No {name}-*.save found in {directory}")
    with open(files[-1], 'rb') as f:
        df = pickle.load(f)
    print(f"Loaded: {files[-1]}  shape={df.shape}")
    return df


# ── Usage ─────────────────────────────────────────────────────────────────────
save_df(dfX, OUT_DIR + '/results/bkp')

#dfX = load_latest_df(OUT_DIR + '/results/bkp')
        
        
########################################################################
# EXPORT PER-SLOT EXCEL 
########################################################################



EXPORT_COLS    = ['order', 'subj', 'sec', 'st_year', 'st_max', 'ex_dur',
                  'teacher_str', 'building', 'room', 'note', 'staff']
EXPORT_COLS_TH = ['ลำดับ', 'วิชา', 'กลุ่ม', 'ชั้นปี', 'นศ.', 'เวลา',
                  'ผู้สอน', 'อาคาร', 'ห้อง', 'หมายเหตุ', 'ผู้คุมสอบ']

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


def build_header_lines(ex_dt_str: str) -> list[tuple[str, str]]:
    """
    Return list of (text, alignment) for the header block.
    alignment: 'center' | 'right'
    All rows are bold.
    """
    dt, hour_str = parse_ex_dt(ex_dt_str)
    day_s    = THAI_DAYS_SHORT[dt.weekday()]
    month_s  = THAI_MONTHS_SHORT[dt.month]
    year_be  = (dt.year + 543) % 100
    session  = SESSION_DURATION.get(hour_str, f'{hour_str}:00 น.')
    date_line = f"วันสอบ  {day_s} {dt.day} {month_s} {year_be}   เวลา {session}"

    return [
        ("สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง", "center"),
        ("รายงานการจัดสอบปลายภาค",                          "center"),
        ("ประจำภาคเรียนที่ 2 ปีการศึกษา 2568",              "center"),
        ("คณะวิศวกรรมศาสตร์",                               "center"),
        ("",                                                  "center"),
        (date_line,                                           "right"),
    ]


def make_slot_sheet(ws, df_slot: pd.DataFrame, ex_dt_str: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    ncols       = len(df_slot.columns)
    header_rows = build_header_lines(ex_dt_str)
    n_header    = len(header_rows)

    thin    = Side(style='thin', color='000000')
    no_side = Side(style=None)

    # ── Header text block ─────────────────────────────────────────────────────
    for i, (text, align) in enumerate(header_rows, start=1):
        if ncols > 1:
            ws.merge_cells(start_row=i, start_column=1,
                           end_row=i,   end_column=ncols)
        cell           = ws.cell(row=i, column=1, value=text)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=14)
        cell.alignment = Alignment(horizontal=align, vertical='center')

    # ── Column headers ────────────────────────────────────────────────────────
    col_row  = n_header + 1
    hdr_font = Font(name='TH Sarabun New', bold=True, size=14, color='000000')
    hdr_fill = PatternFill('solid', start_color='FFFFFF')
    for c, name in enumerate(df_slot.columns, start=1):
        cell           = ws.cell(row=col_row, column=c, value=name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=False)
    ws.row_dimensions[col_row].height = 24

    # ── Deduplicate: blank subj/order for repeated subjects ───────────────────
    subj_col  = 'วิชา'
    order_col = 'ลำดับ'
    df_out = df_slot.copy()
    if subj_col in df_out.columns:
        seen = set()
        for idx in df_out.index:
            subj = df_out.at[idx, subj_col]
            if subj in seen:
                df_out.at[idx, subj_col]  = ''
                if order_col in df_out.columns:
                    df_out.at[idx, order_col] = ''
            else:
                seen.add(subj)

    # ── Pre-compute last row of each subject group ────────────────────────────
    subj_last_rows = set()
    if subj_col in df_slot.columns:
        orig_subj = df_slot[subj_col].tolist()
        for i, subj in enumerate(orig_subj):
            is_last = (i == len(orig_subj) - 1) or (orig_subj[i + 1] != subj)
            if is_last:
                subj_last_rows.add(col_row + 1 + i)

    # ── Column index map ──────────────────────────────────────────────────────
    col_idx = {name: i + 1 for i, name in enumerate(df_out.columns)}
    subj_c  = col_idx.get(subj_col)
    order_c = col_idx.get(order_col)

    def make_border(excel_row: int, col: int) -> Border:
        if col in (subj_c, order_c):
            if excel_row in subj_last_rows:
                return Border(left=thin, right=thin, top=no_side, bottom=thin)
            else:
                return Border(left=thin, right=thin, top=no_side, bottom=no_side)
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r, row in enumerate(df_out.itertuples(index=False), start=col_row + 1):
        for c, value in enumerate(row, start=1):
            if isinstance(value, list):
                value = ', '.join(str(v) for v in value)
            is_subj        = (c == subj_c)
            cell           = ws.cell(row=r, column=c, value=value)
            cell.font      = Font(name='TH Sarabun New', size=14 if is_subj else 13)
            cell.border    = make_border(r, c)
            cell.alignment = Alignment(wrap_text=False, vertical='top')

    # ── Column widths ─────────────────────────────────────────────────────────
    col_widths = {
        'ลำดับ': 7, 'วิชา': 38, 'กลุ่ม': 8, 'ชั้นปี': 8,
        'นศ.': 7, 'เวลา': 16, 'ผู้สอน': 28,
        'อาคาร': 10, 'ห้อง': 10, 'หมายเหตุ': 14, 'ผู้คุมสอบ': 22,
    }
    for c, name in enumerate(df_out.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = col_widths.get(name, 12)


    # ── Print setup: fit to page width ───────────────────────────────────────
    from openpyxl.worksheet.page import PageMargins, PrintPageSetup

    ws.page_setup.orientation     = 'landscape'
    ws.page_setup.fitToPage       = True
    ws.page_setup.fitToWidth      = 1
    ws.page_setup.fitToHeight     = 0   # 0 = ไม่จำกัดความสูง (ขยายลงได้กี่หน้าก็ได้)
    ws.page_setup.paperSize       = ws.PAPERSIZE_A4

    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75,
        header=0.3, footer=0.3
    )

    # ── Repeat header rows on every page ─────────────────────────────────────
    # แถว 1 ถึง col_row (header block + หัว column)
    ws.print_title_rows = f'1:{col_row}'

    # ── Footer: page number ───────────────────────────────────────────────────
    ws.oddFooter.center.text = 'หน้า &P/&N'
    ws.oddFooter.center.size = 12
    ws.oddFooter.center.font = 'TH Sarabun New'


# ── Build workbook ────────────────────────────────────────────────────────────
import openpyxl

dfX2 = dfX[EXPORT_COLS].copy()
dfX2.columns = EXPORT_COLS_TH

wb_subj = openpyxl.Workbook()
wb_subj.remove(wb_subj.active)

for dt in sorted(dtlist):                          # เรียงตามวันเวลา
    df_slot = dfX2[dfX['ex_dt'] == dt].sort_values('วิชา').reset_index(drop=True)
    if df_slot.empty:
        continue

    ws = wb_subj.create_sheet(title=slot_sheet_name(dt))
    make_slot_sheet(ws, df_slot, dt)

out_path = OUT_DIR + '/results/by-subjects.xlsx'
wb_subj.save(out_path)
print(f"Saved: {out_path}")


########################################################################
# EXPORT BY ROOM
########################################################################

# ─────────────────────────────── EXPORT BY ROOM ──────────────────────────────

ROOM_COLS    = ['room', 'subj', 'st_num', 'staff']
ROOM_COLS_TH = ['ห้อง', 'วิชา', 'จำนวน', 'ผู้คุมสอบ']


def build_header_lines_room(ex_dt_str: str, building: str) -> list[tuple[str, str]]:
    dt, hour_str = parse_ex_dt(ex_dt_str)
    day_s    = THAI_DAYS_SHORT[dt.weekday()]
    month_s  = THAI_MONTHS_SHORT[dt.month]
    year_be  = (dt.year + 543) % 100
    session  = SESSION_DURATION.get(hour_str, f'{hour_str}:00 น.')
    date_line = f"อาคาร {building}   วันสอบ  {day_s} {dt.day} {month_s} {year_be}   เวลา {session}"

    return [
        ("สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง", "center"),
        ("รายงานการจัดสอบปลายภาค",                          "center"),
        ("ประจำภาคเรียนที่ 2 ปีการศึกษา 2568",              "center"),
        ("คณะวิศวกรรมศาสตร์",                               "center"),
        ("",                                                  "center"),
        (date_line,                                           "right"),
    ]


def make_room_sheet(ws, df_slot: pd.DataFrame, ex_dt_str: str, building: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    ncols       = len(df_slot.columns)
    header_rows = build_header_lines_room(ex_dt_str, building)
    n_header    = len(header_rows)

    thin    = Side(style='thin', color='000000')
    no_side = Side(style=None)

    # ── Header text block ─────────────────────────────────────────────────────
    for i, (text, align) in enumerate(header_rows, start=1):
        if ncols > 1:
            ws.merge_cells(start_row=i, start_column=1,
                           end_row=i,   end_column=ncols)
        cell           = ws.cell(row=i, column=1, value=text)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=14)
        cell.alignment = Alignment(horizontal=align, vertical='center')

    # ── Column headers ────────────────────────────────────────────────────────
    col_row  = n_header + 1
    hdr_font = Font(name='TH Sarabun New', bold=True, size=14, color='000000')
    hdr_fill = PatternFill('solid', start_color='FFFFFF')
    for c, name in enumerate(df_slot.columns, start=1):
        cell           = ws.cell(row=col_row, column=c, value=name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=False)
    ws.row_dimensions[col_row].height = 24

    # ── Pre-compute last row of each room group ───────────────────────────────
    room_col      = 'ห้อง'
    room_last_rows = set()
    if room_col in df_slot.columns:
        orig_room = df_slot[room_col].tolist()
        for i, room in enumerate(orig_room):
            is_last = (i == len(orig_room) - 1) or (orig_room[i + 1] != room)
            if is_last:
                room_last_rows.add(col_row + 1 + i)

    # ── Deduplicate room: show only on first row of each group ────────────────
    df_out = df_slot.copy()
    if room_col in df_out.columns:
        seen = set()
        for idx in df_out.index:
            room = df_out.at[idx, room_col]
            if room in seen:
                df_out.at[idx, room_col] = ''
            else:
                seen.add(room)

    # ── Column index map ──────────────────────────────────────────────────────
    col_idx  = {name: i + 1 for i, name in enumerate(df_out.columns)}
    room_c   = col_idx.get('ห้อง')
    subj_c   = col_idx.get('วิชา')
    staff_c  = col_idx.get('ผู้คุมสอบ')

    def make_border(excel_row: int, col: int) -> Border:
        is_last = excel_row in room_last_rows
        if col in (room_c, subj_c):
            # เส้นขอบล่างเฉพาะแถวสุดท้ายของห้อง
            return Border(left=thin, right=thin, top=no_side,
                          bottom=thin if is_last else no_side)
        # ผู้คุมสอบ และ column อื่น: full box ทุกแถว
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r, row in enumerate(df_out.itertuples(index=False), start=col_row + 1):
        for c, value in enumerate(row, start=1):
            if isinstance(value, list):
                value = ', '.join(str(v) for v in value)
            is_subj        = (df_out.columns[c - 1] == 'วิชา')
            cell           = ws.cell(row=r, column=c, value=value)
            cell.font      = Font(name='TH Sarabun New', size=14 if is_subj else 13)
            cell.border    = make_border(r, c)
            cell.alignment = Alignment(wrap_text=False, vertical='top')

    # ── Column widths ─────────────────────────────────────────────────────────
    col_widths = {'ห้อง': 12, 'วิชา': 60, 'จำนวน':10, 'ผู้คุมสอบ': 40}
    for c, name in enumerate(df_out.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = col_widths.get(name, 14)

    # ── Print setup ───────────────────────────────────────────────────────────
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation  = 'landscape'
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.page_setup.paperSize    = ws.PAPERSIZE_A4
    ws.page_margins            = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75, header=0.3, footer=0.3
    )
    ws.print_title_rows        = f'1:{col_row}'
    ws.oddFooter.center.text   = 'หน้า &P/&N'
    ws.oddFooter.center.size   = 12
    ws.oddFooter.center.font   = 'TH Sarabun New'


# ── Build workbook ────────────────────────────────────────────────────────────
wb_room = openpyxl.Workbook()
wb_room.remove(wb_room.active)

all_buildings = sorted(dfX['building'].dropna().unique())

for building in all_buildings:
    for dt in sorted(dtlist):
        df_dt = dfX[(dfX['ex_dt'] == dt) & (dfX['building'] == building)].copy()
        if df_dt.empty:
            continue

        df_b = (df_dt.sort_values(['room', 'subj'])[ROOM_COLS]
                .reset_index(drop=True))
        df_b.columns = ROOM_COLS_TH

        sheet_name = f"{building} {slot_sheet_name(dt)}"[:31]
        ws = wb_room.create_sheet(title=sheet_name)
        make_room_sheet(ws, df_b, dt, building)

out_room = OUT_DIR + '/results/by-room.xlsx'
wb_room.save(out_room)
print(f"Saved: {out_room}")        



########################################################################
# EXPORT BY STAFF
########################################################################


# ─────────────────────────────── EXPORT BY STAFF ─────────────────────────────
import re

STAFF_COLS    = ['staff', 'ex_date', 'ex_dur', 'room', 'building']
STAFF_COLS_TH = ['ผู้คุมสอบ', 'วันที่', 'เวลา', 'ห้อง', 'อาคาร']

# ── Column alignment map ──────────────────────────────────────────────────
CENTER_COLS = {'วันที่', 'เวลา'}

# ── Title stripping (key for matching names with the master file) ─────────
TITLE_PAT = re.compile(
    r'^(ศ\.\s*ดร\.|รศ\.\s*ดร\.|ผศ\.\s*ดร\.|ศ\.|รศ\.|ผศ\.|ดร\.|อ\.|'
    r'นางสาว|นาง|นาย|น\.ส\.)\s*'
)

def strip_title(s) -> str:
    """'อ. อมตะ อนันต์พินิจวัฒนา' → 'อมตะ อนันต์พินิจวัฒนา'"""
    s = ' '.join(str(s).split())
    prev = None
    while prev != s:                       # stacked titles e.g. 'ผศ. ดร.'
        prev = s
        s = TITLE_PAT.sub('', s).strip()
    return s


def format_thai_date_short(d) -> str:
    """date → 'พฤ. 19 มี.ค. 69'"""
    dt = pd.Timestamp(d)
    return f"{THAI_DAYS_SHORT[dt.weekday()]} {dt.day} {THAI_MONTHS_SHORT[dt.month]} {(dt.year + 543) % 100}"


def build_header_lines_staff(dept: str) -> list[tuple[str, str]]:
    sem_num, sem_year = exam_semester(dfX['ex_date'].dropna().min()).split('/')
    etype             = exam_type_th(extype)
    return [
        ("สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง", "center"),
        (f"รายงานการจัดสอบ{etype}",                          "center"),
        (f"ประจำภาคเรียนที่ {sem_num} ปีการศึกษา {sem_year}", "center"),
        ("คณะวิศวกรรมศาสตร์",                               "center"),
        (f"หน่วยงาน: {dept}",                                "center"),
        ("",                                                  "center"),
    ]


def make_staff_sheet(ws, df_slot: pd.DataFrame, dept: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    ncols       = len(df_slot.columns)
    header_rows = build_header_lines_staff(dept)
    n_header    = len(header_rows)

    thin    = Side(style='thin', color='000000')
    no_side = Side(style=None)

    # ── Header text block ─────────────────────────────────────────────────────
    for i, (text, align) in enumerate(header_rows, start=1):
        if ncols > 1:
            ws.merge_cells(start_row=i, start_column=1,
                           end_row=i,   end_column=ncols)
        cell           = ws.cell(row=i, column=1, value=text)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=14)
        cell.alignment = Alignment(horizontal=align, vertical='center')

    # ── Column headers ────────────────────────────────────────────────────────
    col_row  = n_header + 1
    hdr_font = Font(name='TH Sarabun New', bold=True, size=14, color='000000')
    hdr_fill = PatternFill('solid', start_color='FFFFFF')
    for c, name in enumerate(df_slot.columns, start=1):
        cell           = ws.cell(row=col_row, column=c, value=name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=False)
    ws.row_dimensions[col_row].height = 24

    # ── Pre-compute last row of each staff group ──────────────────────────────
    staff_col       = 'ผู้คุมสอบ'
    staff_last_rows = set()
    if staff_col in df_slot.columns:
        orig_staff = df_slot[staff_col].tolist()
        for i, name in enumerate(orig_staff):
            is_last = (i == len(orig_staff) - 1) or (orig_staff[i + 1] != name)
            if is_last:
                staff_last_rows.add(col_row + 1 + i)

    # ── Deduplicate: show staff name only on first row of each group ──────────
    df_out = df_slot.copy()
    if staff_col in df_out.columns:
        seen = set()
        for idx in df_out.index:
            s = df_out.at[idx, staff_col]
            if s in seen:
                df_out.at[idx, staff_col] = ''
            else:
                seen.add(s)

    # ── Column index map ──────────────────────────────────────────────────────
    col_idx = {name: i + 1 for i, name in enumerate(df_out.columns)}
    staff_c = col_idx.get(staff_col)

    def make_border(excel_row: int, col: int) -> Border:
        if col == staff_c:
            is_last = excel_row in staff_last_rows
            return Border(left=thin, right=thin, top=no_side,
                          bottom=thin if is_last else no_side)
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for r, row in enumerate(df_out.itertuples(index=False), start=col_row + 1):
        for c, value in enumerate(row, start=1):
            col_name = df_out.columns[c - 1]
            if isinstance(value, list):
                value = ', '.join(str(v) for v in value)
            cell           = ws.cell(row=r, column=c, value=value)
            cell.font      = Font(name='TH Sarabun New', size=13)
            cell.border    = make_border(r, c)
            cell.alignment = Alignment(
                horizontal = 'center' if col_name in CENTER_COLS else 'left',
                wrap_text  = False,
                vertical   = 'top'
            )

    # ── Column widths ─────────────────────────────────────────────────────────
    col_widths = {'ผู้คุมสอบ': 30, 'วันที่': 16, 'เวลา': 15, 'ห้อง': 12, 'อาคาร': 20}
    for c, name in enumerate(df_out.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = col_widths.get(name, 14)

    # ── Print setup ───────────────────────────────────────────────────────────
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation  = 'portrait'
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.page_setup.paperSize    = ws.PAPERSIZE_A4
    ws.page_margins            = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75, header=0.3, footer=0.3
    )
    ws.print_title_rows        = f'1:{col_row}'
    ws.oddFooter.center.text   = 'หน้า &P/&N'
    ws.oddFooter.center.size   = 12
    ws.oddFooter.center.font   = 'TH Sarabun New'


# ── Master lookup (title / role), keyed by name without title ─────────────────
dfST_k       = dfST.copy()
dfST_k.index = dfST_k.index.map(strip_title)
dfST_k       = dfST_k[~dfST_k.index.duplicated(keep='first')]

role_order = dfslot['sort_order'].to_dict()      # role → sort order

def staff_info(name):
    """Return (display_name, sort_order) for one staff name from dfX."""
    key = strip_title(name)
    if key in dfST_k.index:
        title = dfST_k.at[key, 'formal_title']
        title = '' if pd.isna(title) else str(title).strip()
        role  = dfST_k.at[key, 'role']
        return f"{title} {key}".strip(), role_order.get(role, 99)
    return ' '.join(str(name).split()), 99        # not in master: keep as is

# Diagnostics
names = dfX['staff'].dropna().unique()
n_ok  = sum(strip_title(n) in dfST_k.index for n in names)
print(f"master match: {n_ok} / {len(names)}")

# ── Build workbook ────────────────────────────────────────────────────────────
wb_staff = openpyxl.Workbook()
wb_staff.remove(wb_staff.active)

all_depts = sorted(dfX['st_group_dept'].dropna().unique())
if 'none' in all_depts:
    all_depts.remove('none')

for dept in all_depts:
    df_dept = dfX[dfX['st_group_dept'] == dept].copy()
    if df_dept.empty:
        continue

    info = df_dept['staff'].map(staff_info)
    df_dept['_display'] = info.map(lambda t: t[0])
    df_dept['_sort']    = info.map(lambda t: t[1])

    # order: role (sort_order) → name → exam date/time (as before)
    df_dept = (df_dept
               .sort_values(['_sort', '_display', 'ex_dt'])
               .reset_index(drop=True))

    df_dept['staff']   = df_dept['_display']
    df_dept['ex_date'] = df_dept['ex_date'].apply(format_thai_date_short)

    df_dept = df_dept[STAFF_COLS]
    df_dept.columns = STAFF_COLS_TH

    ws = wb_staff.create_sheet(title=dept[:31])
    make_staff_sheet(ws, df_dept, dept)

# Guard: Excel needs at least one visible sheet
if not wb_staff.sheetnames:
    ws = wb_staff.create_sheet("ไม่มีข้อมูล")
    ws['A1'] = "ไม่พบข้อมูลผู้คุมสอบ (ตรวจสอบ st_group_dept)"

out_staff = OUT_DIR + '/results/by-staffs.xlsx'
wb_staff.save(out_staff)
print(f"Saved: {out_staff}")


############################################################
#  EXPORT BY BUILDING & TIME ───────────────────
############################################################



# ─────────────────────────────── EXPORT BY BUILDING-TIME ─────────────────────

BLDG_COL_WIDTHS = {
    'เวลา': 20, 'รหัสวิชา': 14, 'ชื่อวิชา': 36, 'กลุ่ม': 7,
    'ห้องสอบ': 14, 'ผู้รับข้อสอบ': 16, 'ผู้ส่งข้อสอบ': 16,
    'ซอง': 8, 'รับข้อสอบไปตรวจ': 18, 'หมายเหตุ': 14,
}

BIG_COLS    = {'รหัสวิชา', 'ห้องสอบ'}     # larger + bold
CENTER_COLS = {'เวลา', 'รหัสวิชา', 'ห้องสอบ'}   # horizontally centered


def exam_semester(ex_date) -> str:
    d = pd.Timestamp(ex_date)
    if 7 <= d.month <= 11:
        return f"1/{d.year + 543}"
    else:
        return f"2/{(d.year - 1) + 543}"


def exam_type_th(extype: str) -> str:
    return 'ปลายภาค' if extype == 'final' else 'กลางภาค'


def build_header_lines_bldg(building: str, session: str,
                             ex_date, extype: str) -> list[tuple[str, str]]:
    dt      = pd.Timestamp(ex_date)
    day_s   = THAI_DAYS_SHORT[dt.weekday()]
    month_s = THAI_MONTHS_SHORT[dt.month]
    year_be = dt.year + 543
    sem     = exam_semester(ex_date)
    etype   = exam_type_th(extype)
    time_s  = SESSION_DURATION.get('09' if session == 'เช้า' else '13', '')

    return [
        ("สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง", "center"),
        (f"ตารางรับ-ส่งข้อสอบ  คณะวิศวกรรมศาสตร์ การสอบ{etype} ภาคการศึกษาที่ {sem}", "center"),
        (f"อาคาร{building}   วัน{day_s}ที่ {dt.day} {month_s} {year_be} ช่วง{session}   เวลา {time_s}", "center"),
    ]


def make_bldg_sheet(ws, df_slot: pd.DataFrame, building: str,
                    session: str, ex_date, extype: str):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    ncols       = len(df_slot.columns)
    header_rows = build_header_lines_bldg(building, session, ex_date, extype)
    n_header    = len(header_rows)

    thin  = Side(style='thin',   color='000000')
    thick = Side(style='medium', color='000000')     # use 'thick' for even bolder

    # ── Header text block ─────────────────────────────────────────────────────
    for i, (text, align) in enumerate(header_rows, start=1):
        if ncols > 1:
            ws.merge_cells(start_row=i, start_column=1,
                           end_row=i,   end_column=ncols)
        cell           = ws.cell(row=i, column=1, value=text)
        cell.font      = Font(name='TH Sarabun New', bold=True, size=16)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        ws.row_dimensions[i].height = 28

    # ── Blank row between header and column headers ───────────────────────────
    blank_row = n_header + 1
    ws.row_dimensions[blank_row].height = 10

    # ── Column headers (thick top + bottom border) ────────────────────────────
    col_row  = n_header + 2
    hdr_font = Font(name='TH Sarabun New', bold=True, size=14, color='000000')
    hdr_fill = PatternFill('solid', start_color='FFFFFF')
    hdr_bord = Border(left=thin, right=thin, top=thick, bottom=thick)
    for c, name in enumerate(df_slot.columns, start=1):
        cell           = ws.cell(row=col_row, column=c, value=name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.border    = hdr_bord
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=False)
    ws.row_dimensions[col_row].height = 28

    # ── Last row of each room group (thick bottom border) ─────────────────────
    data_start_row = col_row + 1
    rooms          = df_slot['ห้องสอบ'].tolist()
    room_last_rows = {
        data_start_row + i
        for i, rm in enumerate(rooms)
        if i == len(rooms) - 1 or rooms[i + 1] != rm
    }

    # ── Data rows ─────────────────────────────────────────────────────────────
    ROW_HEIGHT = 52
    for r, row in enumerate(df_slot.itertuples(index=False), start=data_start_row):
        ws.row_dimensions[r].height = ROW_HEIGHT
        border = Border(
            left=thin, right=thin,
            top=thick if r == data_start_row else thin,   # match header's thick bottom
            bottom=thick if r in room_last_rows else thin,
        )
        for c, value in enumerate(row, start=1):
            col_name = df_slot.columns[c - 1]
            if isinstance(value, list):
                value = ', '.join(str(v) for v in value)
            cell        = ws.cell(row=r, column=c, value=value)
            cell.border = border
            cell.font   = (Font(name='TH Sarabun New', size=20, bold=True)
                           if col_name in BIG_COLS
                           else Font(name='TH Sarabun New', size=14))
            cell.alignment = Alignment(
                horizontal = 'center' if col_name in CENTER_COLS else None,
                vertical   = 'center',
                wrap_text  = (col_name == 'ชื่อวิชา'),
            )

    # ── Column widths ─────────────────────────────────────────────────────────
    for c, name in enumerate(df_slot.columns, start=1):
        ws.column_dimensions[get_column_letter(c)].width = BLDG_COL_WIDTHS.get(name, 12)

    # ── Print setup ───────────────────────────────────────────────────────────
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation  = 'landscape'
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.page_setup.paperSize    = ws.PAPERSIZE_A4
    ws.page_margins            = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75, header=0.3, footer=0.3
    )
    ws.print_title_rows        = f'1:{col_row}'
    ws.oddFooter.right.text    = 'หน้า &P/&N'
    ws.oddFooter.right.size    = 12
    ws.oddFooter.right.font    = 'TH Sarabun New'


# ── Prepare subject columns ───────────────────────────────────────────────────
dfX['subj_code'] = dfX['subj'].str.split().str[0].str.strip()
dfX['subj_name'] = dfX['subj'].str.split().str[1:].str.join(' ').str.strip()
dfX['is_extra']  = ~dfX['subj_code'].str.match(r'^\d+$')

# ── Build workbook ────────────────────────────────────────────────────────────
wb_bldg = openpyxl.Workbook()
wb_bldg.remove(wb_bldg.active)

all_dates     = sorted(dfX['ex_date'].dropna().unique())
all_buildings = sorted(dfX['building'].dropna().unique())

for ex_date in all_dates:
    df_date = dfX[dfX['ex_date'] == ex_date]

    for hour_key, session, session_en in [('09', 'เช้า', 'am'), ('13', 'บ่าย', 'pm')]:
        df_s = df_date[df_date['ex_time'] == hour_key]
        if df_s.empty:
            continue

        for building in all_buildings:
            df_day = df_s[df_s['building'] == building].copy()
            df_day = df_day[~df_day['is_extra']]
            if df_day.empty:
                continue

            df_day = df_day.sort_values(['room', 'ex_dur'])

            dt         = pd.Timestamp(ex_date)
            month_en   = THAI_MONTHS_SHORT[dt.month].replace('.', '')
            day_str    = f"{dt.day:02d}-{month_en}-{session}"
            sheet_name = f"{day_str} {building}"[:31]

            df_out = pd.DataFrame({
                'เวลา':             df_day['ex_dur'].values,
                'รหัสวิชา':         df_day['subj_code'].values,
                'ชื่อวิชา':         df_day['subj_name'].values,
                'กลุ่ม':            df_day['sec'].values,
                'ห้องสอบ':          df_day['room'].values,
                'ผู้รับข้อสอบ':     '',
                'ผู้ส่งข้อสอบ':     '',
                'ซอง':              '',
                'รับข้อสอบไปตรวจ': '',
                'หมายเหตุ':         '',
            }).reset_index(drop=True)

            ws = wb_bldg.create_sheet(title=sheet_name)
            make_bldg_sheet(ws, df_out, building, session, ex_date, extype)

# Guard: Excel needs at least one visible sheet
if not wb_bldg.sheetnames:
    ws = wb_bldg.create_sheet("ไม่มีข้อมูล")
    ws['A1'] = "ไม่พบข้อมูลการสอบ"

out_bldg = OUT_DIR + '/results/by-building-time.xlsx'
wb_bldg.save(out_bldg)
print(f"Saved: {out_bldg}")






######## FOR CHECKING ############

dfx_cols = ['ID', 'subj', 'sec', 'program_type', 
       'teacher_str', 'building', 'room', 'note', 'ex_date', 'ex_time', 'ex_dur', 
       'ex_dt', 'st_num', 'staff_index', 'reason', 'Remark']
dfZ = dfX[dfx_cols].copy()
dfZ['staff_role'] = dfZ['staff_index'].apply(lambda x: dfST.loc[x, 'role'] if x in dfST.index else '')
dfZ['staff_dept'] = dfZ['staff_index'].apply(lambda x: dfST.loc[x, 'dept'] if x in dfST.index else '')

#dfZ.to_excel(OUT_DIR + '/results/for-cheking.xlsx', index=False)


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

with pd.ExcelWriter(OUT_DIR + '/results/for-cheking.xlsx') as w:
    dfZ.to_excel(w, sheet_name="DATA", index=False)
    g1.to_excel(w, sheet_name="by_staff")
    g2.to_excel(w, sheet_name="by_room_exdur_role")
    g3.to_excel(w, sheet_name="by_subj")
    g4.to_excel(w, sheet_name="vip_only")
    g5.to_excel(w, sheet_name="all_staffs")
 
    


# ─────────────────────────────── EXPORT SIGN SHEET (รับ-ส่งข้อสอบ) ──────────

print('########## EXPORT PAYMENT SHEET (หลักฐานการจ่ายเงินค่ากรรมการคุมสอบ)  ##########')
# ─────────────────────────────── EXPORT PAYMENT SHEET (หลักฐานการจ่ายเงิน) ───


# ── Settings (edit these) ─────────────────────────────────────────────────────
PAY_ORDER_TEXT = "ตามคำสั่งคณะวิศวกรรมศาสตร์ ที่ 414/2569  ลว. 24 กรกฎาคม 2569"
PAY_HOURS      = 3          # จำนวน ชม. per session
PAY_AMOUNT     = 150.0      # จำนวนเงิน per person per session
SHOW_TOTAL     = False      # True → add a "รวม" row (SUM formula) under each table
FONT_SIZE      = 16         # base font size for the whole sheet

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

out_pay = OUT_DIR + '/results/for-payment.xlsx'
wb_pay.save(out_pay)
print(f"Saved: {out_pay}  sheets={len(wb_pay.sheetnames)}")


print(' ******* END JOB ********')





 
