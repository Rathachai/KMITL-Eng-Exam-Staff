# -*- coding: utf-8 -*-
"""
config.py — ค่าตั้งต้นที่ผู้ใช้แก้ได้ก่อนรัน (แก้ไฟล์นี้ก่อน gen)

ลำดับการรัน:
    python3 exam-code.py            # จัดตารางคุมสอบ -> reports/exam-state.pkl
    python3 gen-report-subjects.py  # รายงานแต่ละตัว (รันตัวไหนก่อนก็ได้)
    python3 gen-report-room.py
    python3 gen-report-staff.py
    python3 gen-report-building.py
    python3 gen-report-check.py
    python3 gen-report-payment.py
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── โฟลเดอร์ ────────────────────────────────────────────────────────────────
DATA_DIR = os.path.join(BASE_DIR, "exam-data")   # input: master table + ตารางสอบ
OUT_DIR  = os.path.join(BASE_DIR, "reports")     # output: reports/yyyy-mm-dd_HH-MM/ (หนึ่งโฟลเดอร์ต่อการรัน exam-code.py)

# ── ไฟล์นำเข้า ───────────────────────────────────────────────────────────────
MASTER_FILE      = "Exam-Master-Table.xlsx"   # ใน DATA_DIR
EXAM_FILE_GLOB   = "ร*.xlsx"                  # ไฟล์ตารางสอบ ใน DATA_DIR

# ── ประเภทการสอบ ─────────────────────────────────────────────────────────────
EXTYPE = 'final'      # 'final' = ปลายภาค, 'mid' = กลางภาค

# ── ไฟล์ state ที่ exam-code.py ส่งต่อให้ gen-report-*.py ──────────────────────
STATE_FILE = "exam-state.pkl"                 # ในโฟลเดอร์ของรอบนั้น

# ── หลักฐานการจ่ายเงิน (gen-report-payment.py) ────────────────────────────────
PAY_ORDER_TEXT = "ตามคำสั่งคณะวิศวกรรมศาสตร์ ที่ 414/2569  ลว. 24 กรกฎาคม 2569"
PAY_HOURS      = 3          # จำนวน ชม. per session
PAY_AMOUNT     = 150.0      # จำนวนเงิน per person per session
SHOW_TOTAL     = False      # True → เพิ่มแถว "รวม" (SUM) ใต้แต่ละตาราง
FONT_SIZE      = 16         # ขนาดฟอนต์พื้นฐานของทั้งชีต
