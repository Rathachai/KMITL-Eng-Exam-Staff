# KMITL-Eng-Exam-Staff
สำหรับจัดตารางคุมสอบ

## โครงสร้าง

| ไฟล์ | หน้าที่ |
|---|---|
| `config.py` | ค่าที่ผู้ใช้แก้ก่อนรัน (โฟลเดอร์, ปลายภาค/กลางภาค, ค่าหลักฐานจ่ายเงิน) |
| `common.py` | ฟังก์ชัน/ค่าคงที่ที่ใช้ร่วมกัน, การเซฟ/โหลด state |
| `exam-code.py` | จัดตารางผู้คุมสอบ แล้ว pickle ผลเป็น `exam-state.pkl` |
| `gen-report-*.py` | สร้างรายงานอย่างละตัว (subjects, room, staff, building, check, payment) |
| `gen-exam-room.py` | ตัวควบคุมสำหรับรันจาก command line: เรียกสคริปต์ข้างบนให้ |
| `app.py` | Web UI (Flask) port 8089 |

## วิธีใช้

1. วางไฟล์ใน `exam-data/`: `Exam-Master-Table.xlsx` และไฟล์ตารางสอบ (`ร*.xlsx`)
2. แก้ `config.py` (เดือนสอบ `EXAM_MONTHS` ที่ใช้ตรวจ mid/final อัตโนมัติ หรือบังคับด้วย `EXTYPE`, ข้อความคำสั่งและอัตราค่าตอบแทน)
3. รัน

```bash
python3 gen-exam-room.py                                # จัดตารางใหม่ + สร้างรายงานทั้งหมด
python3 gen-exam-room.py report-only                    # สร้างรายงานทั้งหมดใหม่จากรอบล่าสุด
python3 gen-exam-room.py report-only 2026-10-17_14-26   # สร้างใหม่จากรอบที่ระบุ
```

บน Windows ใช้ `py` (หรือ `python`) แทน `python3`

```bat
py gen-exam-room.py
py gen-exam-room.py report-only
py gen-exam-room.py report-only 2026-10-17_14-26
```

ไลบรารีที่ต้องใช้อยู่ใน `requirements.txt` — `gen-exam-room.py` จะรัน `pip install -r requirements.txt` ให้เองถ้ายังไม่ได้ติดตั้ง หรือรันเองก่อน: `pip install -r requirements.txt` (Windows: `py -m pip install -r requirements.txt`)

> การจัดตารางมีการสุ่ม รันใหม่ทุกครั้งได้ผลไม่เหมือนเดิม
> ถ้าแก้แค่หน้าตา/ค่าใน config ของรายงาน ให้ใช้ `report-only` เพื่อคงตารางเดิมไว้

## Web UI

```bash
python3 app.py        # Windows: py app.py
```

เปิด <http://localhost:8089> (ไม่มี login)

- อัปโหลด **Master file** และ **ตารางสอบ** แยกกัน (ตารางสอบเลือกได้หลายไฟล์) — ถ้าไม่เลือกอันไหน ใช้ไฟล์เริ่มต้นใน `exam-data/`
- กด **เริ่ม gen** → ได้ 1 record ต่อ 1 รอบ (= โฟลเดอร์ `reports/yyyy-mm-dd_HH-MM/`) แสดงสถานะและ progress ทีละขั้น (จัดตาราง + 6 รายงาน)
- **ดาวน์โหลด zip** ของรายงาน (ไม่รวม `exam-state.pkl` และไฟล์ที่อัปโหลด)
- **gen report ใหม่** ของ record เดิม ใช้ตารางคนเดิม สร้างรายงานทับในโฟลเดอร์เดิม
- **log** ดูผลการรันของ record นั้น
- กด gen ซ้อนกันได้ ระบบรันทีละรอบตามคิว
- ค่าเริ่มต้นเปิดเฉพาะเครื่องนี้ (`127.0.0.1`) เพราะไม่มี login ถ้าจะให้เครื่องอื่นใน LAN เข้าได้: `EXAM_HOST=0.0.0.0 python3 app.py` (Windows PowerShell: `$env:EXAM_HOST="0.0.0.0"; py app.py`)

## ผลลัพธ์

แต่ละรอบอยู่ใน `reports/yyyy-mm-dd_HH-MM/`

- `by-subjects.xlsx` ตามวิชา
- `by-room.xlsx` ตามห้องสอบ
- `by-staffs.xlsx` ตามผู้คุมสอบ
- `by-building-time.xlsx` ตามอาคาร/เวลา (รับ-ส่งข้อสอบ)
- `for-cheking.xlsx` ตรวจสอบการจัดคน
- `for-payment.xlsx` หลักฐานการจ่ายเงิน
- `exam-state.pkl` ข้อมูลที่ส่งต่อให้รายงาน
