# -*- coding: utf-8 -*-
"""
gen-exam-room.py — ตัวควบคุมการรัน

    python3 gen-exam-room.py
        จัดตารางคนใหม่ (exam-code.py) -> สร้างโฟลเดอร์ใหม่ reports/yyyy-mm-dd_HH-MM/
        แล้วสร้างรายงานครบทั้งหมด

    python3 gen-exam-room.py report-only
        ใช้ exam-state ของโฟลเดอร์ล่าสุด สร้างรายงานทั้งหมดใหม่ในโฟลเดอร์เดิม

    python3 gen-exam-room.py report-only 2026-10-17_14-26
        เหมือนข้างบน แต่ใช้โฟลเดอร์ที่ระบุ
"""

import importlib.util
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def ensure_requirements():
    """ถ้ายังไม่ได้ติดตั้งไลบรารี ให้ pip install -r requirements.txt ให้อัตโนมัติ"""
    needed = ['pandas', 'numpy', 'openpyxl']
    if all(importlib.util.find_spec(m) for m in needed):
        return
    print("ไม่พบไลบรารีบางตัว -> pip install -r requirements.txt")
    subprocess.run([sys.executable, '-m', 'pip', 'install', '-r',
                    os.path.join(HERE, 'requirements.txt')], check=True)


ensure_requirements()

from common import find_run_dir  # noqa: E402  (import หลังติดตั้งไลบรารีแล้ว)

REPORTS = ['subjects', 'room', 'staff', 'building', 'pickup', 'check', 'payment']


def run(script, *args):
    print(f"\n>>> python3 {script} {' '.join(args)}")
    subprocess.run([sys.executable, os.path.join(HERE, script), *args], check=True)


def main(argv):
    if not argv:
        run('exam-code.py')
        run_name = os.path.basename(find_run_dir())      # โฟลเดอร์ที่เพิ่งสร้าง
    elif argv[0] == 'report-only' and len(argv) <= 2:
        run_name = os.path.basename(find_run_dir(argv[1] if len(argv) == 2 else None))
    else:
        sys.exit("usage:\n  gen-exam-room.py\n  gen-exam-room.py report-only [yyyy-mm-dd_HH-MM]")

    print(f"\n=== reports for run: {run_name} ===")
    for r in REPORTS:
        run(f'gen-report-{r}.py', run_name)
    print(f"\nDONE -> reports/{run_name}/")


if __name__ == '__main__':
    main(sys.argv[1:])
