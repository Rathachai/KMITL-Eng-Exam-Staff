# -*- coding: utf-8 -*-
"""
app.py — web UI (Flask) สำหรับจัดตารางคุมสอบ   http://localhost:8089

- อัปโหลด master file / ตารางสอบ (แยกกัน) ถ้าไม่เลือกใช้ค่าเริ่มต้นใน exam-data/
- 1 การ gen = 1 record = 1 โฟลเดอร์ reports/yyyy-mm-dd_HH-MM/
- ดาวน์โหลด zip ของรายงาน (ไม่รวม exam-state.pkl และไฟล์นำเข้า)
- gen report ใหม่ของ record เดิมได้ (ใช้ state เดิม)

ไม่มี login: ค่าเริ่มต้นเปิดเฉพาะเครื่องนี้ (127.0.0.1)
เปิดให้เครื่องอื่นใน LAN:  EXAM_HOST=0.0.0.0 python3 app.py
"""

import importlib.util
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

if not all(importlib.util.find_spec(m) for m in ('flask', 'pandas', 'numpy', 'openpyxl')):
    print("ไม่พบไลบรารีบางตัว -> pip install -r requirements.txt")
    subprocess.run([sys.executable, '-m', 'pip', 'install', '-r',
                    os.path.join(HERE, 'requirements.txt')], check=True)

import io
import json
import queue
import re
import threading
import zipfile
from datetime import datetime

from flask import Flask, abort, jsonify, request, send_file

from config import OUT_DIR, STATE_FILE

PORT = 8089
REPORTS = ['subjects', 'room', 'staff', 'building', 'check', 'payment']
REPORT_LABEL = {'subjects': 'ตามวิชา', 'room': 'ตามห้องสอบ', 'staff': 'ตามผู้คุมสอบ',
                'building': 'ตามอาคาร/เวลา', 'check': 'ตรวจสอบ', 'payment': 'หลักฐานจ่ายเงิน'}
RUN_RE = re.compile(r'^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}(-\d+)?$')
STATUS_FILE = 'status.json'
LOG_FILE = 'run.log'

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024
jobs = queue.Queue()
lock = threading.Lock()


# ── run folder helpers ────────────────────────────────────────────────────────

def run_path(name, *parts):
    if not RUN_RE.match(name):
        abort(404)
    return os.path.join(OUT_DIR, name, *parts)


def read_status(name):
    p = run_path(name, STATUS_FILE)
    if os.path.exists(p):
        with lock, open(p, encoding='utf-8') as f:
            return json.load(f)
    return None


def write_status(name, st):
    with lock:
        tmp = run_path(name, STATUS_FILE + '.tmp')
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(st, f, ensure_ascii=False)
        os.replace(tmp, run_path(name, STATUS_FILE))


def report_files(name):
    d = run_path(name)
    return sorted(f for f in os.listdir(d)
                  if f.endswith('.xlsx') and not f.startswith('input-'))


def new_run_name():
    base = datetime.now().strftime('%Y-%m-%d_%H-%M')
    name, n = base, 1
    while os.path.exists(os.path.join(OUT_DIR, name)):
        n += 1
        name = f'{base}-{n}'
    os.makedirs(os.path.join(OUT_DIR, name))
    return name


def list_runs():
    if not os.path.isdir(OUT_DIR):
        return []
    out = []
    for name in sorted(os.listdir(OUT_DIR), reverse=True):
        if not RUN_RE.match(name) or not os.path.isdir(os.path.join(OUT_DIR, name)):
            continue
        st = read_status(name)
        if st is None:       # รอบที่สร้างจาก gen-exam-room.py (ไม่มี status.json)
            has_state = os.path.exists(run_path(name, STATE_FILE))
            st = {'status': 'done' if has_state else 'error', 'step': '',
                  'done': len(REPORTS) + 1 if has_state else 0, 'total': len(REPORTS) + 1,
                  'master': 'default', 'schedule': ['default'], 'error': '' if has_state else 'ไม่มี state'}
        st['name'] = name
        st['files'] = report_files(name)
        out.append(st)
    return out


# ── worker ────────────────────────────────────────────────────────────────────

def run_step(name, script, args, env_extra, log):
    env = dict(os.environ, PYTHONIOENCODING='utf-8', **env_extra)
    log.write(f"\n>>> python3 {script} {' '.join(args)}\n")
    log.flush()
    r = subprocess.run([sys.executable, os.path.join(HERE, script), *args], cwd=HERE,
                       env=env, stdout=log, stderr=subprocess.STDOUT)
    return r.returncode == 0


def worker():
    while True:
        job = jobs.get()
        name, mode = job['name'], job['mode']
        st = read_status(name)
        st.update(status='running', error='', failed=[])
        steps = (['exam-code'] if mode == 'full' else []) + REPORTS
        st['total'] = len(steps)
        st['done'] = 0
        write_status(name, st)
        env_extra = {}
        if mode == 'full':
            env_extra['EXAM_RUN_NAME'] = name
            if st['master'] != 'default':
                env_extra['EXAM_MASTER_PATH'] = run_path(name, 'input-master.xlsx')
            if st['schedule'] != ['default']:
                env_extra['EXAM_SCHEDULE_FILES'] = os.pathsep.join(
                    run_path(name, f'input-schedule-{i + 1}.xlsx') for i in range(len(st['schedule'])))
        try:
            with open(run_path(name, LOG_FILE), 'a', encoding='utf-8') as log:
                log.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} {mode} =====\n")
                for step in steps:
                    st['step'] = step
                    write_status(name, st)
                    if step == 'exam-code':
                        ok = run_step(name, 'exam-code.py', [], env_extra, log)
                        if not ok:
                            st.update(status='error', error='จัดตารางไม่สำเร็จ (ดู log)')
                            break
                    else:
                        if not run_step(name, f'gen-report-{step}.py', [name], env_extra, log):
                            st.setdefault('failed', []).append(step)
                    st['done'] += 1
                    write_status(name, st)
                else:
                    if st.get('failed'):
                        st.update(status='error', error='รายงานที่ล้มเหลว: ' + ', '.join(st['failed']))
                    else:
                        st['status'] = 'done'
                    st['step'] = ''
        except Exception as e:        # noqa: BLE001
            st.update(status='error', error=str(e))
        write_status(name, st)
        jobs.task_done()


def recover_interrupted():
    if not os.path.isdir(OUT_DIR):
        return
    for name in os.listdir(OUT_DIR):
        if RUN_RE.match(name):
            st = read_status(name)
            if st and st.get('status') in ('running', 'queued'):
                st.update(status='error', error='ถูกขัดจังหวะ (server ถูกปิดระหว่างทำงาน)')
                write_status(name, st)


# ── routes ────────────────────────────────────────────────────────────────────

def is_xlsx(f):
    return f and f.filename and f.filename.lower().endswith('.xlsx')


@app.post('/api/runs')
def create_run():
    master = request.files.get('master')
    schedules = [f for f in request.files.getlist('schedule') if f and f.filename]
    if master and master.filename and not is_xlsx(master):
        return jsonify(error='Master file ต้องเป็น .xlsx'), 400
    if any(not is_xlsx(f) for f in schedules):
        return jsonify(error='ตารางสอบต้องเป็น .xlsx'), 400
    has_master = bool(master and master.filename)
    name = new_run_name()
    if has_master:
        master.save(run_path(name, 'input-master.xlsx'))
    for i, f in enumerate(schedules):
        f.save(run_path(name, f'input-schedule-{i + 1}.xlsx'))
    st = {'status': 'queued', 'step': '', 'done': 0, 'total': len(REPORTS) + 1, 'error': '',
          'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
          'master': master.filename if has_master else 'default',
          'schedule': [f.filename for f in schedules] or ['default']}
    write_status(name, st)
    jobs.put({'name': name, 'mode': 'full'})
    return jsonify(name=name), 201


@app.post('/api/runs/<name>/regen')
def regen(name):
    st = read_status(name)
    if not os.path.exists(run_path(name, STATE_FILE)):
        return jsonify(error='ไม่มี state ของรอบนี้ จึงสร้างรายงานใหม่ไม่ได้'), 400
    if st is None:
        st = {'master': 'default', 'schedule': ['default']}
    if st.get('status') in ('running', 'queued'):
        return jsonify(error='รอบนี้กำลังทำงานอยู่'), 409
    st.update(status='queued', done=0, error='', step='')
    write_status(name, st)
    jobs.put({'name': name, 'mode': 'report-only'})
    return jsonify(name=name)


@app.get('/api/runs')
def api_runs():
    return jsonify(runs=list_runs())


@app.get('/api/runs/<name>/zip')
def download_zip(name):
    files = report_files(name)
    if not files:
        abort(404)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(run_path(name, f), arcname=f)
    buf.seek(0)
    return send_file(buf, mimetype='application/zip', as_attachment=True,
                     download_name=f'reports_{name}.zip')


@app.get('/api/runs/<name>/log')
def get_log(name):
    p = run_path(name, LOG_FILE)
    if not os.path.exists(p):
        abort(404)
    with open(p, encoding='utf-8', errors='replace') as f:
        return f.read()[-20000:], 200, {'Content-Type': 'text/plain; charset=utf-8'}


@app.get('/')
def index():
    return PAGE


PAGE = r'''<!doctype html>
<html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>จัดตารางคุมสอบ</title>
<style>
 :root{--bg:#f6f6f4;--card:#fff;--bd:#dcdcd6;--tx:#1f1f1e;--mu:#6b6b66;--ac:#2563eb;--ok:#15803d;--er:#b91c1c;--wa:#a16207}
 @media(prefers-color-scheme:dark){:root{--bg:#1b1b1a;--card:#262625;--bd:#3d3d3a;--tx:#ececea;--mu:#9a9a94;--ac:#60a5fa;--ok:#4ade80;--er:#f87171;--wa:#facc15}}
 *{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.6 system-ui,"Sarabun",sans-serif}
 main{max-width:860px;margin:0 auto;padding:20px 16px 48px}h1{font-size:20px;font-weight:600;margin:0 0 16px}
 .card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:16px;margin-bottom:16px}
 h2{font-size:15px;font-weight:600;margin:0 0 12px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}
 label{display:block;font-size:13px;color:var(--mu);margin-bottom:4px}input[type=file]{width:100%;font-size:13px}
 .hint{font-size:12px;color:var(--mu);margin-top:4px}
 button{font:inherit;font-size:14px;padding:6px 14px;border-radius:8px;border:1px solid var(--bd);background:var(--card);color:var(--tx);cursor:pointer}
 button:hover{border-color:var(--ac)}button.pri{background:var(--ac);color:#fff;border-color:var(--ac)}button:disabled{opacity:.5;cursor:default}
 .run{border:1px solid var(--bd);border-radius:10px;padding:12px;margin-bottom:10px}.top{display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap;align-items:center}
 .name{font-weight:600}.badge{font-size:12px;padding:1px 8px;border-radius:8px;margin-left:8px;border:1px solid currentColor}
 .done{color:var(--ok)}.error{color:var(--er)}.running,.queued{color:var(--wa)}
 .bar{height:6px;background:var(--bg);border-radius:3px;margin:8px 0}.bar>div{height:100%;border-radius:3px;background:var(--ac);transition:width .3s}
 .meta{font-size:12px;color:var(--mu)}.err{font-size:13px;color:var(--er);margin:4px 0}.acts{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
 #msg{font-size:13px;margin-left:10px}
</style></head><body><main>
<h1>จัดตารางคุมสอบ</h1>
<form class="card" id="f">
 <h2>สร้างรอบใหม่</h2>
 <div class="grid">
  <div><label for="master">Master file (Exam-Master-Table.xlsx)</label><input type="file" id="master" name="master" accept=".xlsx">
   <div class="hint">ไม่เลือก = ใช้ไฟล์เริ่มต้นใน exam-data/</div></div>
  <div><label for="schedule">รายการตารางสอบ (.xlsx เลือกได้หลายไฟล์)</label><input type="file" id="schedule" name="schedule" accept=".xlsx" multiple>
   <div class="hint">ไม่เลือก = ใช้ไฟล์เริ่มต้นใน exam-data/</div></div>
 </div>
 <div style="margin-top:14px"><button class="pri" id="go" type="submit">เริ่ม gen</button><span id="msg"></span></div>
</form>
<div class="card"><h2>รายการที่ gen</h2><div id="runs"><div class="meta">กำลังโหลด…</div></div></div>
</main>
<script>
const LB={queued:'รอคิว',running:'กำลังทำ',done:'เสร็จแล้ว',error:'ผิดพลาด'};
const $=(t,c,x)=>{const e=document.createElement(t);if(c)e.className=c;if(x!==undefined)e.textContent=x;return e};
function render(runs){
 const box=document.getElementById('runs');box.replaceChildren();
 if(!runs.length){box.append($('div','meta','ยังไม่มีรายการ'));return}
 for(const r of runs){
  const d=$('div','run'),top=$('div','top'),l=$('div');
  l.append($('span','name',r.name),$('span','badge '+r.status,LB[r.status]||r.status));
  top.append(l,$('span','meta',r.done+'/'+r.total+' ขั้นตอน'+(r.step?' · '+r.step:'')));
  const bar=$('div','bar'),fill=$('div');fill.style.width=(r.total?100*r.done/r.total:0)+'%';bar.append(fill);
  d.append(top,bar);
  d.append($('div','meta','master: '+r.master+' · ตารางสอบ: '+(r.schedule||[]).join(', ')));
  if(r.error)d.append($('div','err',r.error));
  const a=$('div','acts'),busy=r.status==='running'||r.status==='queued';
  if(r.files.length){const z=$('a');z.href='/api/runs/'+r.name+'/zip';const b=$('button',null,'ดาวน์โหลด zip');z.append(b);a.append(z)}
  const g=$('button',null,'gen report ใหม่');g.disabled=busy;g.onclick=()=>regen(r.name);a.append(g);
  const lg=$('a');lg.href='/api/runs/'+r.name+'/log';lg.target='_blank';lg.append($('button',null,'log'));a.append(lg);
  d.append(a);box.append(d);
 }
}
async function load(){try{const j=await (await fetch('/api/runs')).json();render(j.runs)}catch(e){}}
async function regen(n){const r=await fetch('/api/runs/'+n+'/regen',{method:'POST'});if(!r.ok)alert((await r.json()).error);load()}
document.getElementById('f').onsubmit=async e=>{
 e.preventDefault();const m=document.getElementById('msg'),b=document.getElementById('go');
 b.disabled=true;m.textContent='กำลังอัปโหลด…';
 const r=await fetch('/api/runs',{method:'POST',body:new FormData(e.target)});
 const j=await r.json();b.disabled=false;
 if(!r.ok){m.textContent=j.error;return}
 m.textContent='เริ่มรอบ '+j.name;e.target.reset();load()};
load();setInterval(load,1500);
</script></body></html>'''


if __name__ == '__main__':
    os.makedirs(OUT_DIR, exist_ok=True)
    recover_interrupted()
    threading.Thread(target=worker, daemon=True).start()
    host = os.environ.get('EXAM_HOST', '127.0.0.1')
    print(f"เปิด http://localhost:{PORT}  (host={host})")
    app.run(host=host, port=PORT, threaded=True)
