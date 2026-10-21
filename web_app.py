# -*- coding: utf-8 -*-
"""
صفحه وب محلی برای ساخت گزارش پایش POS و IPG.
فایل‌ها فقط روی همین کامپیوتر پردازش می‌شوند و به هیچ سرور دیگری ارسال نمی‌شوند.
"""
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file, render_template_string

import pos_analysis as pa

app = Flask(__name__)

PAGE = """<!doctype html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>گزارش پایش پایانه‌های فروش</title>
<style>
  body { font-family: Tahoma, Vazirmatn, sans-serif; background:#f3f6fa; margin:0; color:#1f2d3d; }
  .wrap { max-width: 760px; margin: 40px auto; background:#fff; border-radius:16px;
          box-shadow:0 6px 24px rgba(31,78,120,.12); padding:32px; }
  h1 { color:#1f4e78; font-size:24px; margin:0 0 8px; }
  p.hint { color:#555; font-size:15px; line-height:1.9; margin:0 0 22px; }
  #drop { border:2px dashed #2e75b6; border-radius:14px; padding:34px; text-align:center;
          background:#f7fbff; cursor:pointer; transition:.2s; }
  #drop.over { background:#e2efda; border-color:#3a8f3a; }
  #drop b { font-size:18px; color:#1f4e78; }
  #files { display:none; }
  ul { list-style:none; padding:0; margin:18px 0; }
  li { background:#eef4fb; border-radius:8px; padding:8px 14px; margin:6px 0; font-size:15px; }
  button { background:#1f4e78; color:#fff; border:0; border-radius:10px; padding:14px 28px;
           font-size:18px; cursor:pointer; }
  button:disabled { background:#9bb0c4; cursor:not-allowed; }
  #status { margin-top:18px; font-size:16px; min-height:24px; }
  .ok { color:#1e7b34; } .err { color:#c00000; }
  a.dl { display:inline-block; margin-top:14px; background:#1e7b34; color:#fff; text-decoration:none;
         padding:14px 28px; border-radius:10px; font-size:18px; }
  .foot { margin-top:26px; font-size:13px; color:#777; line-height:1.8; }
</style>
</head>
<body>
<div class="wrap">
  <h1>گزارش پایش و پرستاری پایانه‌ها</h1>
  <p class="hint">هر تعداد فایل اکسل ماهانه (POS و IPG) را اینجا بکشید یا انتخاب کنید.
    گزارش با همه فایل‌ها ساخته می‌شود. فایل‌ها روی همین کامپیوتر پردازش می‌شوند.</p>

  <div id="drop"><b>فایل‌های اکسل را اینجا بکشید</b><br><span>یا برای انتخاب کلیک کنید</span></div>
  <input type="file" id="files" accept=".xlsx" multiple>
  <ul id="list"></ul>
  <button id="go" disabled>ساخت گزارش</button>
  <div id="status"></div>

  <div class="foot">فایل‌های اکسل با نام دلخواه قابل قبول هستند؛ دوره و نوع فایل (POS یا IPG) از عنوان داخل فایل خوانده می‌شود.</div>
</div>

<script>
  const drop = document.getElementById('drop');
  const input = document.getElementById('files');
  const list = document.getElementById('list');
  const go = document.getElementById('go');
  const status = document.getElementById('status');
  let chosen = [];

  function render() {
    list.innerHTML = '';
    chosen.forEach(f => {
      const li = document.createElement('li');
      li.textContent = f.name;
      list.appendChild(li);
    });
    go.disabled = chosen.length === 0;
  }
  function addFiles(fileList) {
    for (const f of fileList) {
      if (f.name.toLowerCase().endsWith('.xlsx')) chosen.push(f);
    }
    render();
  }
  drop.onclick = () => input.click();
  input.onchange = e => addFiles(e.target.files);
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); addFiles(e.dataTransfer.files); };

  go.onclick = async () => {
    const form = new FormData();
    chosen.forEach(f => form.append('files', f, f.name));
    go.disabled = true;
    document.querySelectorAll('a.dl').forEach(a => a.remove());
    status.className = '';
    status.textContent = 'در حال پردازش ' + chosen.length + ' فایل... این کار چند ثانیه طول می‌کشد.';
    try {
      const res = await fetch('/analyze', { method: 'POST', body: form });
      if (!res.ok) {
        const data = await res.json().catch(() => ({ error: 'خطای نامشخص' }));
        throw new Error(data.error || 'خطا در ساخت گزارش');
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'pos_monitoring_report.xlsx';
      a.className = 'dl';
      a.textContent = 'دانلود گزارش اکسل';
      status.className = 'ok';
      status.textContent = 'گزارش آماده شد.';
      status.after(a);
    } catch (err) {
      status.className = 'err';
      status.textContent = err.message;
    } finally {
      go.disabled = chosen.length === 0;
    }
  };
</script>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.post("/analyze")
def analyze():
    uploads = [f for f in request.files.getlist("files") if f.filename.lower().endswith(".xlsx")]
    if not uploads:
        return jsonify(error="هیچ فایل .xlsx انتخاب نشده است."), 400

    work = Path(tempfile.mkdtemp(prefix="pos_report_"))
    paths = []
    for i, f in enumerate(uploads):
        p = work / f"{i:03d}_{Path(f.filename).name}"
        f.save(p)
        paths.append(str(p))

    out = work / "pos_monitoring_report.xlsx"
    try:
        pa.run(paths, str(out))
    except SystemExit as e:
        return jsonify(error=str(e)), 400
    except Exception as e:  # نمایش خطا به کاربر
        return jsonify(error=f"خطا در پردازش فایل‌ها: {e}"), 500
    return send_file(out, as_attachment=True, download_name="pos_monitoring_report.xlsx")
