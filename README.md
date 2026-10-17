# KMITL-Eng-Exam-Staff
สำหรับจัดตารางคุมสอบ

## โครงสร้าง

| ไฟล์ | หน้าที่ |
|---|---|
| `config.py` | ค่าที่ผู้ใช้แก้ก่อนรัน (โฟลเดอร์, ปลายภาค/กลางภาค, ค่าหลักฐานจ่ายเงิน) |
| `common.py` | ฟังก์ชัน/ค่าคงที่ที่ใช้ร่วมกัน, การเซฟ/โหลด state |
| `exam-code.py` | จัดตารางผู้คุมสอบ แล้ว pickle ผลเป็น `exam-state.pkl` |
| `gen-report-*.py` | สร้างรายงานอย่างละตัว (subjects, room, staff, building, check, payment) |
| `gen-exam-room.py` | ตัวควบคุม: เรียกสคริปต์ข้างบนให้ |

## วิธีใช้

1. วางไฟล์ใน `exam-data/`: `Exam-Master-Table.xlsx` และไฟล์ตารางสอบ (`ร*.xlsx`)
2. แก้ `config.py` (เช่น `EXTYPE = 'final'` หรือ `'mid'`, ข้อความคำสั่งและอัตราค่าตอบแทน)
3. รัน

```bash
python3 gen-exam-room.py                                # จัดตารางใหม่ + สร้างรายงานทั้งหมด
python3 gen-exam-room.py report-only                    # สร้างรายงานทั้งหมดใหม่จากรอบล่าสุด
python3 gen-exam-room.py report-only 2026-10-17_14-26   # สร้างใหม่จากรอบที่ระบุ
```

> การจัดตารางมีการสุ่ม รันใหม่ทุกครั้งได้ผลไม่เหมือนเดิม
> ถ้าแก้แค่หน้าตา/ค่าใน config ของรายงาน ให้ใช้ `report-only` เพื่อคงตารางเดิมไว้

## ผลลัพธ์

แต่ละรอบอยู่ใน `reports/yyyy-mm-dd_HH-MM/`

- `by-subjects.xlsx` ตามวิชา
- `by-room.xlsx` ตามห้องสอบ
- `by-staffs.xlsx` ตามผู้คุมสอบ
- `by-building-time.xlsx` ตามอาคาร/เวลา (รับ-ส่งข้อสอบ)
- `for-cheking.xlsx` ตรวจสอบการจัดคน
- `for-payment.xlsx` หลักฐานการจ่ายเงิน
- `exam-state.pkl` ข้อมูลที่ส่งต่อให้รายงาน
