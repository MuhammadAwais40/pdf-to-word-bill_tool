(() => {
  const $ = (s) => document.querySelector(s);
  const dz = $('#dropzone'), input = $('#fileInput');
  const MAX_MB = 50;
  const AUTO_RESET_DELAY = 2000;
  let current = { id: null, name: 'document' };
  let toastTimer, resetTimer;
  function isDesktop() {
    return !!(window.pywebview && window.pywebview.api && typeof window.pywebview.api.save_file === 'function');
  }

  function setStep(n) {
    document.querySelectorAll('.step').forEach((el, i) => {
      el.classList.toggle('active', i + 1 === n);
      el.classList.toggle('done', i + 1 < n);
    });
  }
  function show(id) {
    ['uploadSection', 'progressSection', 'resultSection'].forEach(s => { $('#' + s).hidden = s !== id; });
  }
  function toast(msg, type = 'error') {
    const t = $('#toast');
    t.textContent = msg;
    t.className = 'toast' + (type === 'success' ? ' success' : '');
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 4500);
  }
  function discard(id) {
    if (!id) return;
    try { fetch('/api/discard/' + id, { method: 'POST', keepalive: true }).catch(() => {}); } catch (_) {}
  }
  function reset() {
    clearTimeout(resetTimer);
    discard(current.id);
    current = { id: null, name: 'document' };
    input.value = '';
    $('#barFill').classList.remove('indeterminate');
    $('#barFill').style.width = '0%';
    $('#preview').innerHTML = '';
    show('uploadSection');
    setStep(1);
  }
  function fail(msg) { toast(msg); reset(); }

  // Download ke baad auto refresh
  function scheduleReset() {
    if (AUTO_RESET_DELAY <= 0) return;
    const id = current.id;
    clearTimeout(resetTimer);
    resetTimer = setTimeout(() => { if (current.id === id) reset(); }, AUTO_RESET_DELAY);
  }

  // ---------- Upload handling ----------
  function handleFile(file) {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.pdf') && file.type !== 'application/pdf') {
      return toast('Only PDF files are allowed.');
    }
    if (file.size > MAX_MB * 1024 * 1024) return toast('File too large (max ' + MAX_MB + ' MB).');
    upload(file);
  }

  function upload(file) {
    current.name = file.name.replace(/\.pdf$/i, '') || 'document';
    $('#progFileName').textContent = file.name;
    $('#progStatus').textContent = 'Uploading… 0%';
    show('progressSection');
    setStep(2);

    const fd = new FormData();
    fd.append('file', file);
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/upload');

    xhr.upload.onprogress = (e) => {
      if (!e.lengthComputable) return;
      const pct = Math.round((e.loaded / e.total) * 100);
      $('#barFill').style.width = pct + '%';
      if (pct >= 100) {
        $('#barFill').classList.add('indeterminate');
        $('#progStatus').textContent = 'Reading PDF & extracting text and tables…';
      } else {
        $('#progStatus').textContent = 'Uploading… ' + pct + '%';
      }
    };
    xhr.onload = () => {
      let res = {};
      try { res = JSON.parse(xhr.responseText); } catch (_) {}
      if (xhr.status >= 200 && xhr.status < 300 && res.id) renderResult(res, file.name);
      else fail(res.error || 'Conversion failed. Please try another file.');
    };
    xhr.onerror = () => fail('Network error. Please check your connection.');
    xhr.send(fd);
  }

  // ---------- Result ----------
  function renderResult(res, name) {
    current.id = res.id;
    $('#resFileName').textContent = name;
    $('#stPages').textContent = res.stats.pages;
    $('#stParas').textContent = res.stats.paragraphs;
    $('#stTables').textContent = res.stats.tables;
    $('#stWords').textContent = res.stats.words;

    const box = $('#preview');
    box.innerHTML = '';
    res.preview.forEach((b) => {
      if (b.type === 'para') {
        const d = document.createElement('div');
        d.className = 'pv-para' + (b.heading ? ' h' : '') + (b.bold ? ' b' : '');
        d.textContent = b.text;
        box.appendChild(d);
      } else {
        const wrap = document.createElement('div');
        wrap.className = 'pv-wrap';
        const tbl = document.createElement('table');
        tbl.className = 'pv-table';
        b.rows.forEach((row) => {
          const tr = document.createElement('tr');
          row.forEach((cell) => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
          });
          tbl.appendChild(tr);
        });
        wrap.appendChild(tbl);
        box.appendChild(wrap);
      }
    });

    $('#barFill').classList.remove('indeterminate');
    show('resultSection');
    setStep(3);
  }

  // ---------- Download ----------
  async function download(fmt, btn) {
    if (!current.id) return;
    clearTimeout(resetTimer);           // pehle se chalta hua auto reset rok do
    btn.disabled = true;
    btn.classList.add('loading');
    const ext = fmt === 'word' ? 'docx' : 'xlsx';
    let success = false;
    try {
      if (isDesktop()) {
        // Desktop app: Python "Save As" dialog kholta hai aur file save karta hai
        const res = await window.pywebview.api.save_file(current.id, fmt);
        if (res && res.ok) {
          toast(`${ext.toUpperCase()} file saved!`, 'success');
          success = true;
        } else if (res && res.cancelled) {
          toast('Save cancelled.');
        } else {
          throw new Error((res && res.error) || 'Save failed.');
        }
      } else {
        // Normal browser: purana blob download
        const resp = await fetch(`/api/download/${current.id}/${fmt}`);
        if (!resp.ok) {
          let msg = 'Download failed.';
          try { msg = (await resp.json()).error || msg; } catch (_) {}
          throw new Error(msg);
        }
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${current.name}.${ext}`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 2000);
        toast(`${ext.toUpperCase()} file downloaded!`, 'success');
        success = true;
      }
    } catch (err) {
      toast(err.message);
    } finally {
      btn.disabled = false;
      btn.classList.remove('loading');
    }
    if (success) scheduleReset();       // download hone ke baad tool refresh
  }

  // ---------- Events ----------
  dz.addEventListener('click', () => input.click());
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => handleFile(input.files[0]));
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => handleFile(e.dataTransfer.files[0]));
  $('#btnWord').addEventListener('click', (e) => download('word', e.currentTarget));
  $('#btnExcel').addEventListener('click', (e) => download('excel', e.currentTarget));
  $('#btnReset').addEventListener('click', reset);
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
})();
