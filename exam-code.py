# -*- coding: utf-8 -*-
"""
exam-code.py — จัดตารางผู้คุมสอบ (ส่วนหลัก)

อ่าน exam-data/ -> จัดคนคุมสอบ -> pickle state ไปที่ reports/exam-state.pkl
จากนั้นรัน gen-report-*.py เพื่อสร้างรายงานแต่ละตัว
ค่าที่ผู้ใช้แก้ได้อยู่ใน config.py
"""

import glob
import random

from common import *   # pd, np, re, os, datetime, config values, helpers


######### READ MASTER / EXAM TIME / ASSIGN #############################

######### READ MASTER #############################

dfst = pd.read_excel(open(os.path.join(DATA_DIR, MASTER_FILE), "rb"), sheet_name="exam_staff")
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

dfdp = pd.read_excel(open(os.path.join(DATA_DIR, MASTER_FILE), "rb"), sheet_name="department", dtype={'code4': str, 0: str})

dfoff = pd.read_excel(open(os.path.join(DATA_DIR, MASTER_FILE), "rb"), sheet_name="staff_cannot")
stoff = (dfoff['fullname'] + ' ' + dfoff['cannot']).tolist()

dfslot = pd.read_excel(open(os.path.join(DATA_DIR, MASTER_FILE), "rb"), sheet_name="staff_slot")
dfslot['self_subj'] = dfslot['self_subj'].map({'yes': True, 'no': False})
dfslot.set_index('role', inplace=True)


######### READ EXAM TIME #############################

'''
file_name = DATA_DIR + '/17 บ่าย.xls'
df_final = process_exam_report(file_name)
df_final.columns = ['order', 'subj', 'sec', 'st_year', 'st_num', 'ex_time', 'teachers', 'building', 'room', 'note', 'ex_date']

df_final['code4'] = df_final['subj'].str[:5]
'''

fnames = glob.glob(os.path.join(DATA_DIR, EXAM_FILE_GLOB))

df_ = []

for fname in fnames:
    _df = process_exam_report(fname)
    df_.append(_df)

dfx = pd.concat(df_)

extype = detect_extype(dfx['ex_date'])   # 'mid' / 'final' จากเดือนสอบ (config.EXAM_MONTHS)
print('exam type:', extype)
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


#dfEXST.to_csv(out_path('for-checking.csv'), index=False, encoding='utf-8-sig')
        
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
# SAVE STATE  (ส่งต่อให้ gen-report-*.py)
########################################################################

save_state(dfX=dfX, dfST=dfST, dfslot=dfslot, dtlist=dtlist, extype=extype)
