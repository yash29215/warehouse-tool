/* ─────────────────────────────────────────────
   Warehouse Tools — Frontend SPA Logic
   ───────────────────────────────────────────── */

// ── Sidebar ──────────────────────────────────

let sidebarCollapsed = false;

function toggleSidebar() {
  const sb = document.getElementById('sidebar');
  const icon = document.getElementById('toggleIcon');
  sidebarCollapsed = !sidebarCollapsed;
  sb.classList.toggle('collapsed', sidebarCollapsed);
  icon.textContent = sidebarCollapsed ? '›' : '‹';
}

// ── Tab navigation ────────────────────────────

function showTab(tabId) {
  document.querySelectorAll('.tab-page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById(`tab-${tabId}`).classList.add('active');
  document.querySelector(`[data-tab="${tabId}"]`).classList.add('active');
  setStatus(`Tab: ${tabId.replace(/-/g, ' ')}`);
  // Resize canvas if showing case-pick
  if (tabId === 'case-pick' && cpMap.points.length) {
    setTimeout(cpResizeCanvas, 50);
  }
  if (tabId === 'point-to-point' && p2pMap.points.length) {
    setTimeout(p2pResizeCanvas, 50);
  }
}

// ── Status bar ────────────────────────────────

function setStatus(msg) {
  document.getElementById('status-text').textContent = msg;
}

// ── File input labels ─────────────────────────

function wireFileInput(inputId, nameId, onChange) {
  const input = document.getElementById(inputId);
  const nameEl = document.getElementById(nameId);
  if (!input || !nameEl) return;
  input.addEventListener('change', () => {
    nameEl.textContent = input.files.length ? input.files[0].name : 'No file selected';
    if (onChange && input.files.length) onChange(input.files[0]);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  wireFileInput('ep-file', 'ep-filename');
  wireFileInput('cp-file', 'cp-filename', cpOnFileSelected);
  wireFileInput('cp-template', 'cp-template-name');
  wireFileInput('p2p-json', 'p2p-json-name', p2pOnFileSelected);
  wireFileInput('p2p-excel', 'p2p-excel-name');
  wireFileInput('mv-file', 'mv-filename');
  wireFileInput('of-file', 'of-filename', ofOnFileSelected);
  wireFileInput('ms-file', 'ms-filename', msUpdatePreview);

  document.getElementById('cp-cross-aisle').addEventListener('change', function () {
    document.getElementById('cp-sensitivity-wrap').style.display =
      this.checked ? '' : 'none';
  });

  cpInitMap();
  p2pInitMap();
  window.addEventListener('resize', () => {
    if (cpMap.points.length)  cpResizeCanvas();
    if (p2pMap.points.length) p2pResizeCanvas();
  });
});

// ── Log helpers ───────────────────────────────

function clearLog(id) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = '';
}

function appendLog(id, msg, level = 'info') {
  const el = document.getElementById(id);
  if (!el) return;
  const line = document.createElement('span');
  line.className = `log-line ${level}`;
  const ts = new Date().toLocaleTimeString('en-GB', { hour12: false });
  line.textContent = `[${ts}] ${msg}`;
  el.appendChild(line);
  el.appendChild(document.createElement('br'));
  el.scrollTop = el.scrollHeight;
}

function showCard(id, visible = true) {
  const el = document.getElementById(id);
  if (el) el.style.display = visible ? '' : 'none';
}

function setDownload(linkId, rowId, href, filename) {
  const link = document.getElementById(linkId);
  if (link) { link.href = href; link.download = filename; }
  showCard(rowId);
}

function setBtnLoading(id, loading, label = '▶ Run') {
  const btn = document.getElementById(id);
  if (!btn) return;
  btn.disabled = loading;
  btn.innerHTML = loading ? '<span class="spinner"></span> Working…' : label;
}

// ── Each Pick ────────────────────────────────────────────────

async function runEachPick() {
  const fileEl = document.getElementById('ep-file');
  if (!fileEl.files.length) { alert('Please select a JSON or SMAP file.'); return; }

  setBtnLoading('ep-run-btn', true, '▶ Extract');
  clearLog('ep-log');
  showCard('ep-result-card');
  showCard('ep-download-row', false);
  setStatus('Each Pick: extracting…');

  const fd = new FormData();
  fd.append('file', fileEl.files[0]);
  fd.append('include_work', document.getElementById('ep-work').checked);
  fd.append('include_resource', document.getElementById('ep-resource').checked);

  try {
    const res = await fetch('/api/each-pick/process', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      appendLog('ep-log', `Error: ${data.error}`, 'err');
      setStatus('Each Pick: error');
      return;
    }
    (data.log || []).forEach(m => appendLog('ep-log', m, detectLevel(m)));
    setDownload('ep-download-link', 'ep-download-row', data.download, data.filename);
    setStatus('Each Pick: complete');
  } catch (e) {
    appendLog('ep-log', `Request failed: ${e}`, 'err');
    setStatus('Each Pick: failed');
  } finally {
    setBtnLoading('ep-run-btn', false, '▶ Extract');
  }
}

/* ─────────────────────────────────────────────
   CASE PICK — MAP CANVAS (point tagging)
   ───────────────────────────────────────────── */

const cpMap = {
  canvas: null,
  ctx: null,
  points: [],           // [{name, type:"AP"|"LM", x, y}]
  loading: new Set(),
  unloading: new Set(),
  crossAisle: new Set(),
  excluded: new Set(),
  tagMode: 'loading',
  // view transform
  scale: 1, panX: 0, panY: 0, worldCx: 0, worldCy: 0,
  // interaction
  dragMode: null,       // 'pan' | 'select' | null
  dragStart: null,
  dragLast: null,
  selActive: false,
};

function cpInitMap() {
  const canvas = document.getElementById('cp-map-canvas');
  if (!canvas) return;
  cpMap.canvas = canvas;
  cpMap.ctx = canvas.getContext('2d');

  canvas.addEventListener('contextmenu', e => e.preventDefault());
  canvas.addEventListener('mousedown', cpOnMouseDown);
  canvas.addEventListener('mousemove', cpOnMouseMove);
  canvas.addEventListener('mouseup',   cpOnMouseUp);
  canvas.addEventListener('mouseleave', cpOnMouseUp);
  canvas.addEventListener('wheel', cpOnWheel, { passive: false });

  cpSetTagMode('loading');
}

function cpResizeCanvas() {
  const canvas = cpMap.canvas;
  if (!canvas) return;
  const rect = canvas.getBoundingClientRect();
  // Use device-independent pixels
  canvas.width  = Math.max(1, Math.floor(rect.width));
  canvas.height = Math.max(1, Math.floor(rect.height));
  cpRedraw();
}

function cpOnFileSelected(file) {
  // Auto-fill the Map Path field. Browsers strip the full disk path for
  // security — we put the filename in so the user knows what was picked, and
  // they can prepend the directory to enable the "save-next-to-map" flow.
  const mapPathEl = document.getElementById('cp-map-path');
  if (mapPathEl) {
    const cur = (mapPathEl.value || '').trim();
    const curBase = cur.split(/[\\/]/).pop();
    if (curBase !== file.name) mapPathEl.value = file.name;
  }
  // Parse the JSON client-side and load points into the map
  const reader = new FileReader();
  reader.onload = (e) => {
    let data;
    try { data = JSON.parse(e.target.result); }
    catch (err) {
      alert('File is not valid JSON: ' + err.message);
      return;
    }
    cpLoadPointsFromJson(data);
  };
  reader.onerror = () => alert('Failed to read file');
  reader.readAsText(file);
}

function cpLoadPointsFromJson(data) {
  const pts = [];
  (data.advancedPointList || []).forEach(p => {
    const name = p.instanceName;
    const pos = p.pos || {};
    if (name && 'x' in pos && 'y' in pos) {
      pts.push({
        name,
        type: p.className === 'ActionPoint' ? 'AP' : 'LM',
        x: parseFloat(pos.x),
        y: parseFloat(pos.y),
      });
    }
  });
  cpMap.points = pts;
  cpMap.loading.clear();
  cpMap.unloading.clear();
  cpMap.crossAisle.clear();
  cpMap.excluded.clear();
  cpUpdateBadges();

  showCard('cp-map-card');
  cpResizeCanvas();
  cpFitAll();
  cpRedraw();
  setStatus(`Loaded ${pts.length} points from map`);
}

function cpFitAll() {
  if (!cpMap.points.length) return;
  const xs = cpMap.points.map(p => p.x);
  const ys = cpMap.points.map(p => p.y);
  cpMap.worldCx = (Math.min(...xs) + Math.max(...xs)) / 2;
  cpMap.worldCy = (Math.min(...ys) + Math.max(...ys)) / 2;
  const w = Math.max(cpMap.canvas.width,  100);
  const h = Math.max(cpMap.canvas.height, 100);
  const spanX = Math.max(Math.max(...xs) - Math.min(...xs), 0.001);
  const spanY = Math.max(Math.max(...ys) - Math.min(...ys), 0.001);
  cpMap.scale = Math.min(w / spanX, h / spanY) * 0.85;
  cpMap.panX = 0;
  cpMap.panY = 0;
}

function cpResetView() {
  cpFitAll();
  cpRedraw();
}

function cpW2C(wx, wy) {
  const w = Math.max(cpMap.canvas.width,  1);
  const h = Math.max(cpMap.canvas.height, 1);
  const cx = (wx - cpMap.worldCx) * cpMap.scale + w / 2 + cpMap.panX;
  const cy = -(wy - cpMap.worldCy) * cpMap.scale + h / 2 + cpMap.panY;
  return [cx, cy];
}

function cpRedraw() {
  const ctx = cpMap.ctx;
  if (!ctx) return;
  const w = cpMap.canvas.width;
  const h = cpMap.canvas.height;
  ctx.clearRect(0, 0, w, h);

  if (!cpMap.points.length) {
    ctx.fillStyle = '#908fa0';
    ctx.font = '12px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Select a map file to view points', w / 2, h / 2);
    return;
  }

  const showAP  = document.getElementById('cp-show-ap').checked;
  const showLM  = document.getElementById('cp-show-lm').checked;
  const showLbl = document.getElementById('cp-show-labels').checked;
  const rBase = Math.max(4, Math.min(10, Math.round(cpMap.scale * 0.6)));

  for (const pt of cpMap.points) {
    if (pt.type === 'AP' && !showAP) continue;
    if (pt.type === 'LM' && !showLM) continue;

    const [cx, cy] = cpW2C(pt.x, pt.y);
    const isLoading  = cpMap.loading.has(pt.name);
    const isUnload   = cpMap.unloading.has(pt.name);
    const isCA       = cpMap.crossAisle.has(pt.name);
    const isExcl     = cpMap.excluded.has(pt.name);

    let fill, stroke, marker = null;
    if (pt.type === 'AP') {
      if (isCA)         { fill = '#06b6d4'; stroke = '#ffffff'; marker = '+'; }
      else if (isLoading) { fill = '#4edea3'; stroke = '#ffffff'; marker = 'L'; }
      else if (isUnload)  { fill = '#f59e0b'; stroke = '#ffffff'; marker = 'U'; }
      else if (isExcl)    { fill = '#ffb4ab'; stroke = '#ffffff'; marker = 'X'; }
      else                { fill = '#494bd6'; stroke = '#8083ff'; }
      ctx.beginPath();
      ctx.arc(cx, cy, rBase, 0, Math.PI * 2);
      ctx.fillStyle = fill;
      ctx.fill();
      ctx.strokeStyle = stroke;
      ctx.lineWidth = 1;
      ctx.stroke();
      if (marker) {
        ctx.fillStyle = '#ffffff';
        ctx.font = `bold ${Math.max(7, rBase - 1)}px Inter, sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(marker, cx, cy);
      }
    } else {
      const r = Math.max(3, rBase - 1);
      if (isCA)        { fill = '#06b6d4'; stroke = '#ffffff'; marker = '+'; }
      else if (isExcl) { fill = '#ffb4ab'; stroke = '#ffffff'; }
      else             { fill = '#4edea3'; stroke = '#34d399'; }
      ctx.beginPath();
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx + r, cy);
      ctx.lineTo(cx, cy + r);
      ctx.lineTo(cx - r, cy);
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.fill();
      ctx.strokeStyle = stroke;
      ctx.lineWidth = 1;
      ctx.stroke();
      if (marker) {
        ctx.fillStyle = '#ffffff';
        ctx.font = `bold ${Math.max(7, r - 1)}px Inter, sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(marker, cx, cy);
      }
    }

    if (showLbl) {
      ctx.fillStyle = '#908fa0';
      ctx.font = '8px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      ctx.fillText(pt.name, cx, cy + rBase + 4);
    }
  }

  // Rubber-band selection overlay
  if (cpMap.dragMode === 'select' && cpMap.selActive && cpMap.dragLast && cpMap.dragStart) {
    const [x0, y0] = cpMap.dragStart;
    const [x1, y1] = cpMap.dragLast;
    ctx.strokeStyle = '#494bd6';
    ctx.lineWidth = 1.5;
    ctx.setLineDash([5, 4]);
    ctx.strokeRect(Math.min(x0, x1), Math.min(y0, y1),
                   Math.abs(x1 - x0), Math.abs(y1 - y0));
    ctx.setLineDash([]);
  }
}

function cpMousePos(e) {
  const rect = cpMap.canvas.getBoundingClientRect();
  return [e.clientX - rect.left, e.clientY - rect.top];
}

function cpOnMouseDown(e) {
  const pos = cpMousePos(e);
  if (e.button === 2) {
    cpMap.dragMode = 'pan';
    cpMap.dragLast = pos;
    cpMap.canvas.style.cursor = 'grabbing';
  } else if (e.button === 0) {
    cpMap.dragMode = 'select';
    cpMap.dragStart = pos;
    cpMap.dragLast = pos;
    cpMap.selActive = false;
  }
}

function cpOnMouseMove(e) {
  if (!cpMap.dragMode) return;
  const pos = cpMousePos(e);
  if (cpMap.dragMode === 'pan') {
    const dx = pos[0] - cpMap.dragLast[0];
    const dy = pos[1] - cpMap.dragLast[1];
    cpMap.panX += dx;
    cpMap.panY += dy;
    cpMap.dragLast = pos;
    cpRedraw();
  } else if (cpMap.dragMode === 'select') {
    cpMap.dragLast = pos;
    const [x0, y0] = cpMap.dragStart;
    if (Math.hypot(pos[0] - x0, pos[1] - y0) > 5) cpMap.selActive = true;
    if (cpMap.selActive) cpRedraw();
  }
}

function cpOnMouseUp() {
  if (!cpMap.dragMode) return;
  if (cpMap.dragMode === 'pan') {
    cpMap.dragMode = null;
    cpMap.canvas.style.cursor = 'crosshair';
    return;
  }
  if (cpMap.dragMode === 'select') {
    if (cpMap.selActive) {
      const [x0, y0] = cpMap.dragStart;
      const [x1, y1] = cpMap.dragLast;
      const rx0 = Math.min(x0, x1), rx1 = Math.max(x0, x1);
      const ry0 = Math.min(y0, y1), ry1 = Math.max(y0, y1);
      const showAP = document.getElementById('cp-show-ap').checked;
      const showLM = document.getElementById('cp-show-lm').checked;
      let changed = false;
      for (const pt of cpMap.points) {
        if (pt.type === 'AP' && !showAP) continue;
        if (pt.type === 'LM' && !showLM) continue;
        const [cx, cy] = cpW2C(pt.x, pt.y);
        if (cx >= rx0 && cx <= rx1 && cy >= ry0 && cy <= ry1) {
          cpApplyTag(pt.name, pt.type);
          changed = true;
        }
      }
      if (changed) cpUpdateBadges();
    } else {
      // Click — find nearest visible point within threshold
      const threshold = Math.max(10, Math.round(cpMap.scale * 0.8));
      let best = null, bestD = threshold;
      const showAP = document.getElementById('cp-show-ap').checked;
      const showLM = document.getElementById('cp-show-lm').checked;
      for (const pt of cpMap.points) {
        if (pt.type === 'AP' && !showAP) continue;
        if (pt.type === 'LM' && !showLM) continue;
        const [cx, cy] = cpW2C(pt.x, pt.y);
        const d = Math.hypot(cpMap.dragLast[0] - cx, cpMap.dragLast[1] - cy);
        if (d < bestD) { bestD = d; best = pt; }
      }
      if (best) {
        cpApplyTag(best.name, best.type);
        cpUpdateBadges();
      }
    }
    cpMap.dragMode = null;
    cpMap.selActive = false;
    cpMap.dragStart = null;
    cpRedraw();
  }
}

function cpOnWheel(e) {
  e.preventDefault();
  const pos = cpMousePos(e);
  const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12;
  const w = cpMap.canvas.width, h = cpMap.canvas.height;
  const mxWorld = (pos[0] - w / 2 - cpMap.panX) / cpMap.scale + cpMap.worldCx;
  const myWorld = -(pos[1] - h / 2 - cpMap.panY) / cpMap.scale + cpMap.worldCy;
  cpMap.scale *= factor;
  cpMap.panX = pos[0] - w / 2 - (mxWorld - cpMap.worldCx) * cpMap.scale;
  cpMap.panY = pos[1] - h / 2 + (myWorld - cpMap.worldCy) * cpMap.scale;
  cpRedraw();
}

function cpApplyTag(name, ptType) {
  const mode = cpMap.tagMode;
  if ((mode === 'loading' || mode === 'unloading') && ptType !== 'AP') return;
  // Clear any existing tag for this point
  cpMap.loading.delete(name);
  cpMap.unloading.delete(name);
  cpMap.crossAisle.delete(name);
  cpMap.excluded.delete(name);
  if      (mode === 'loading')     cpMap.loading.add(name);
  else if (mode === 'unloading')   cpMap.unloading.add(name);
  else if (mode === 'cross_aisle') cpMap.crossAisle.add(name);
  else if (mode === 'remove')      cpMap.excluded.add(name);
}

function cpSetTagMode(mode) {
  cpMap.tagMode = mode;
  ['loading', 'unloading', 'cross_aisle', 'remove'].forEach(m => {
    const btn = document.getElementById('tag-btn-' + m);
    if (!btn) return;
    btn.classList.remove('active-loading', 'active-unloading', 'active-cross_aisle', 'active-remove');
    if (m === mode) btn.classList.add('active-' + m);
  });
}

function cpUpdateBadges() {
  document.getElementById('cp-load-count').textContent   = cpMap.loading.size;
  document.getElementById('cp-unload-count').textContent = cpMap.unloading.size;
  document.getElementById('cp-ca-count').textContent     = cpMap.crossAisle.size;
  document.getElementById('cp-excl-count').textContent   = cpMap.excluded.size;
}

function cpClearAll() {
  cpMap.loading.clear();
  cpMap.unloading.clear();
  cpMap.crossAisle.clear();
  cpMap.excluded.clear();
  cpUpdateBadges();
  cpRedraw();
}

/* ─────────────────────────────────────────────
   POINT TO POINT — MAP PREVIEW (read-only, pan/zoom)
   Mirrors the Case Pick canvas, minus tag-mode tooling.
   ───────────────────────────────────────────── */

const p2pMap = {
  canvas: null,
  ctx: null,
  points: [],
  subzoneTargetIdx: null,  // non-null = drawing subzones for this zone index
  scale: 1, panX: 0, panY: 0, worldCx: 0, worldCy: 0,
  dragMode: null, dragLast: null,
  // Zone-drawing state
  drawModeOn: false,
  drawStart: null,   // canvas px [x,y]
  drawCurrent: null, // canvas px [x,y]
  pendingBox: null,  // {minX,maxX,minY,maxY} pending zone-form confirmation
  zones: [],         // [{name, entry, exit, minX, maxX, minY, maxY}]
};

function p2pInitMap() {
  const canvas = document.getElementById('p2p-map-canvas');
  if (!canvas) return;
  p2pMap.canvas = canvas;
  p2pMap.ctx = canvas.getContext('2d');

  canvas.addEventListener('contextmenu', e => e.preventDefault());
  canvas.addEventListener('mousedown', p2pOnMouseDown);
  canvas.addEventListener('mousemove', p2pOnMouseMove);
  canvas.addEventListener('mouseup',   p2pOnMouseUp);
  canvas.addEventListener('mouseleave', p2pOnMouseUp);
  canvas.addEventListener('wheel', p2pOnWheel, { passive: false });
}

function p2pResizeCanvas() {
  const canvas = p2pMap.canvas;
  if (!canvas) return;
  const rect = canvas.getBoundingClientRect();
  canvas.width  = Math.max(1, Math.floor(rect.width));
  canvas.height = Math.max(1, Math.floor(rect.height));
  p2pRedraw();
}

function p2pOnFileSelected(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    let data;
    try { data = JSON.parse(e.target.result); }
    catch (err) { alert('File is not valid JSON: ' + err.message); return; }
    p2pLoadPointsFromJson(data);
  };
  reader.onerror = () => alert('Failed to read file');
  reader.readAsText(file);
}

function p2pLoadPointsFromJson(data) {
  const pts = [];
  (data.advancedPointList || []).forEach(p => {
    const name = p.instanceName;
    const pos = p.pos || {};
    if (name && 'x' in pos && 'y' in pos) {
      pts.push({
        name,
        type: p.className === 'ActionPoint' ? 'AP' : 'LM',
        x: parseFloat(pos.x),
        y: parseFloat(pos.y),
      });
    }
  });
  p2pMap.points = pts;
  showCard('p2p-map-card');
  p2pResizeCanvas();
  p2pFitAll();
  p2pRedraw();
  setStatus(`Point to Point: ${pts.length} points loaded on map`);
}

function p2pFitAll() {
  if (!p2pMap.points.length) return;
  const xs = p2pMap.points.map(p => p.x);
  const ys = p2pMap.points.map(p => p.y);
  p2pMap.worldCx = (Math.min(...xs) + Math.max(...xs)) / 2;
  p2pMap.worldCy = (Math.min(...ys) + Math.max(...ys)) / 2;
  const w = Math.max(p2pMap.canvas.width,  100);
  const h = Math.max(p2pMap.canvas.height, 100);
  const spanX = Math.max(Math.max(...xs) - Math.min(...xs), 0.001);
  const spanY = Math.max(Math.max(...ys) - Math.min(...ys), 0.001);
  p2pMap.scale = Math.min(w / spanX, h / spanY) * 0.85;
  p2pMap.panX = 0;
  p2pMap.panY = 0;
}

function p2pResetView() {
  p2pFitAll();
  p2pRedraw();
}

function p2pW2C(wx, wy) {
  const w = Math.max(p2pMap.canvas.width,  1);
  const h = Math.max(p2pMap.canvas.height, 1);
  const cx = (wx - p2pMap.worldCx) * p2pMap.scale + w / 2 + p2pMap.panX;
  const cy = -(wy - p2pMap.worldCy) * p2pMap.scale + h / 2 + p2pMap.panY;
  return [cx, cy];
}

function p2pC2W(cx, cy) {
  const w = Math.max(p2pMap.canvas.width,  1);
  const h = Math.max(p2pMap.canvas.height, 1);
  const wx = (cx - w / 2 - p2pMap.panX) / p2pMap.scale + p2pMap.worldCx;
  const wy = -(cy - h / 2 - p2pMap.panY) / p2pMap.scale + p2pMap.worldCy;
  return [wx, wy];
}

function p2pRedraw() {
  const ctx = p2pMap.ctx;
  if (!ctx) return;
  const w = p2pMap.canvas.width;
  const h = p2pMap.canvas.height;
  ctx.clearRect(0, 0, w, h);

  if (!p2pMap.points.length) {
    ctx.fillStyle = '#908fa0';
    ctx.font = '12px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Select a map file to view points', w / 2, h / 2);
    return;
  }

  const showAP  = document.getElementById('p2p-show-ap').checked;
  const showLM  = document.getElementById('p2p-show-lm').checked;
  const showLbl = document.getElementById('p2p-show-labels').checked;
  const rBase = Math.max(4, Math.min(10, Math.round(p2pMap.scale * 0.6)));

  for (const pt of p2pMap.points) {
    if (pt.type === 'AP' && !showAP) continue;
    if (pt.type === 'LM' && !showLM) continue;
    const [cx, cy] = p2pW2C(pt.x, pt.y);

    if (pt.type === 'AP') {
      ctx.beginPath();
      ctx.arc(cx, cy, rBase, 0, Math.PI * 2);
      ctx.fillStyle = '#494bd6';
      ctx.fill();
      ctx.strokeStyle = '#8083ff';
      ctx.lineWidth = 1;
      ctx.stroke();
    } else {
      const r = Math.max(3, rBase - 1);
      ctx.beginPath();
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx + r, cy);
      ctx.lineTo(cx, cy + r);
      ctx.lineTo(cx - r, cy);
      ctx.closePath();
      ctx.fillStyle = '#4edea3';
      ctx.fill();
      ctx.strokeStyle = '#34d399';
      ctx.lineWidth = 1;
      ctx.stroke();
    }

    if (showLbl) {
      ctx.fillStyle = '#908fa0';
      ctx.font = '8px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'top';
      ctx.fillText(pt.name, cx, cy + rBase + 4);
    }
  }

  // ── Saved zones ──
  p2pMap.zones.forEach((zone, zIdx) => {
    const [cx0, cy0] = p2pW2C(zone.minX, zone.maxY);
    const [cx1, cy1] = p2pW2C(zone.maxX, zone.minY);
    const x = Math.min(cx0, cx1), y = Math.min(cy0, cy1);
    const w2 = Math.abs(cx1 - cx0), h2 = Math.abs(cy1 - cy0);
    const isTarget = p2pMap.subzoneTargetIdx === zIdx;
    ctx.fillStyle   = isTarget ? 'rgba(73,75,214,0.07)' : 'rgba(73,75,214,0.12)';
    ctx.strokeStyle = isTarget ? '#F7941D' : '#8083ff';
    ctx.lineWidth   = isTarget ? 2.5 : 1.5;
    if (isTarget) ctx.setLineDash([6, 3]);
    ctx.fillRect(x, y, w2, h2);
    ctx.strokeRect(x, y, w2, h2);
    ctx.setLineDash([]);
    ctx.fillStyle = isTarget ? '#F7941D' : '#c0c1ff';
    ctx.font = 'bold 11px Inter, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
    ctx.fillText(zone.name, x + 5, y + 4);

    // ── Subzones ──
    (zone.subzones || []).forEach(sz => {
      const [sx0, sy0] = p2pW2C(sz.minX, sz.maxY);
      const [sx1, sy1] = p2pW2C(sz.maxX, sz.minY);
      const sx = Math.min(sx0, sx1), sy = Math.min(sy0, sy1);
      const sw = Math.abs(sx1 - sx0), sh = Math.abs(sy1 - sy0);
      ctx.fillStyle   = 'rgba(247,148,29,0.10)';
      ctx.strokeStyle = '#F7941D';
      ctx.lineWidth   = 1;
      ctx.fillRect(sx, sy, sw, sh);
      ctx.strokeRect(sx, sy, sw, sh);
      ctx.fillStyle = '#F7941D';
      ctx.font = '10px Inter, sans-serif';
      ctx.fillText(sz.name, sx + 3, sy + 3);
    });
  });

  // ── In-progress draw rectangle ──
  if (p2pMap.drawModeOn && p2pMap.drawStart && p2pMap.drawCurrent) {
    const [x0, y0] = p2pMap.drawStart;
    const [x1, y1] = p2pMap.drawCurrent;
    const x = Math.min(x0, x1), y = Math.min(y0, y1);
    const w2 = Math.abs(x1 - x0), h2 = Math.abs(y1 - y0);
    ctx.fillStyle = 'rgba(78, 222, 163, 0.15)';
    ctx.strokeStyle = '#4edea3';
    ctx.lineWidth = 2;
    ctx.setLineDash([5, 4]);
    ctx.fillRect(x, y, w2, h2);
    ctx.strokeRect(x, y, w2, h2);
    ctx.setLineDash([]);
  }
}

function p2pMousePos(e) {
  const rect = p2pMap.canvas.getBoundingClientRect();
  return [e.clientX - rect.left, e.clientY - rect.top];
}

function p2pOnMouseDown(e) {
  const pos = p2pMousePos(e);
  if (p2pMap.drawModeOn && e.button === 0) {
    // Left-click in draw mode → start drawing a zone rectangle
    p2pMap.dragMode = 'draw';
    p2pMap.drawStart = pos;
    p2pMap.drawCurrent = pos;
    p2pMap.canvas.style.cursor = 'crosshair';
    return;
  }
  // Otherwise: pan with any button
  p2pMap.dragMode = 'pan';
  p2pMap.dragLast = pos;
  p2pMap.canvas.style.cursor = 'grabbing';
}

function p2pOnMouseMove(e) {
  if (!p2pMap.dragMode) return;
  const pos = p2pMousePos(e);
  if (p2pMap.dragMode === 'pan') {
    p2pMap.panX += pos[0] - p2pMap.dragLast[0];
    p2pMap.panY += pos[1] - p2pMap.dragLast[1];
    p2pMap.dragLast = pos;
    p2pRedraw();
  } else if (p2pMap.dragMode === 'draw') {
    p2pMap.drawCurrent = pos;
    p2pRedraw();
  }
}

function p2pOnMouseUp() {
  if (p2pMap.dragMode === 'draw' && p2pMap.drawStart && p2pMap.drawCurrent) {
    const [x0, y0] = p2pMap.drawStart;
    const [x1, y1] = p2pMap.drawCurrent;
    // Reject tiny accidental clicks
    if (Math.abs(x1 - x0) > 6 && Math.abs(y1 - y0) > 6) {
      const [wx0, wy0] = p2pC2W(x0, y0);
      const [wx1, wy1] = p2pC2W(x1, y1);
      zfOpenForm({
        minX: Math.min(wx0, wx1), maxX: Math.max(wx0, wx1),
        minY: Math.min(wy0, wy1), maxY: Math.max(wy0, wy1),
      });
    }
    p2pMap.drawStart = null;
    p2pMap.drawCurrent = null;
  }
  p2pMap.dragMode = null;
  if (p2pMap.canvas) {
    p2pMap.canvas.style.cursor = p2pMap.drawModeOn ? 'crosshair' : 'grab';
  }
  p2pRedraw();
}

function p2pOnWheel(e) {
  e.preventDefault();
  const pos = p2pMousePos(e);
  const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12;
  const w = p2pMap.canvas.width, h = p2pMap.canvas.height;
  const mxWorld = (pos[0] - w / 2 - p2pMap.panX) / p2pMap.scale + p2pMap.worldCx;
  const myWorld = -(pos[1] - h / 2 - p2pMap.panY) / p2pMap.scale + p2pMap.worldCy;
  p2pMap.scale *= factor;
  p2pMap.panX = pos[0] - w / 2 - (mxWorld - p2pMap.worldCx) * p2pMap.scale;
  p2pMap.panY = pos[1] - h / 2 + (myWorld - p2pMap.worldCy) * p2pMap.scale;
  p2pRedraw();
}

// ── Draw-mode toggle ──
function p2pToggleDrawMode() {
  p2pMap.drawModeOn = !p2pMap.drawModeOn;
  if (!p2pMap.drawModeOn) p2pMap.subzoneTargetIdx = null;
  const btn = document.getElementById('p2p-draw-btn');
  if (btn) btn.classList.toggle('active-cross_aisle', p2pMap.drawModeOn);
  if (p2pMap.canvas) {
    p2pMap.canvas.style.cursor = p2pMap.drawModeOn ? 'crosshair' : 'grab';
  }
  setStatus(p2pMap.drawModeOn
    ? 'Draw Zone ON — left-drag a rectangle to add a zone'
    : 'Draw Zone OFF');
  p2pRedraw();
}

function p2pStartSubzoneDraw(idx) {
  p2pMap.subzoneTargetIdx = idx;
  p2pMap.drawModeOn = true;
  const btn = document.getElementById('p2p-draw-btn');
  if (btn) btn.classList.add('active-cross_aisle');
  if (p2pMap.canvas) p2pMap.canvas.style.cursor = 'crosshair';
  setStatus(`Subzone draw mode — drag inside "${p2pMap.zones[idx].name}" to add subzones · click ⬚ Draw Zone to exit`);
  p2pRedraw();
}

// ── Zone / Subzone form modal ──
function zfOpenForm(bbox) {
  p2pMap.pendingBox = bbox;
  const isSubzone = p2pMap.subzoneTargetIdx !== null;
  const parentZone = isSubzone ? p2pMap.zones[p2pMap.subzoneTargetIdx] : null;

  document.getElementById('zone-form-overlay').style.display = 'flex';
  document.getElementById('zf-title').textContent = isSubzone
    ? `Add Subzone — ${parentZone.name}`
    : 'Add Zone';
  document.getElementById('zf-name').value = isSubzone
    ? `${parentZone.name}_${(parentZone.subzones || []).length + 1}`
    : `Zone_${p2pMap.zones.length + 1}`;

  const cx = (bbox.minX + bbox.maxX) / 2;
  const cy = (bbox.minY + bbox.maxY) / 2;
  const lms = p2pMap.points
    .filter(p => p.type === 'LM')
    .map(p => ({ ...p, d: Math.hypot(p.x - cx, p.y - cy) }))
    .sort((a, b) => a.d - b.d);

  const fill = (selId) => {
    const sel = document.getElementById(selId);
    sel.innerHTML = '';
    lms.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.name;
      opt.textContent = `${p.name}  (d=${p.d.toFixed(1)})`;
      sel.appendChild(opt);
    });
  };
  fill('zf-entry');
  fill('zf-exit');

  document.getElementById('zf-bbox-display').textContent =
    `bbox: x [${bbox.minX.toFixed(2)} → ${bbox.maxX.toFixed(2)}]  ·  ` +
    `y [${bbox.minY.toFixed(2)} → ${bbox.maxY.toFixed(2)}]`;

  document.getElementById('zf-name').focus();
  document.getElementById('zf-name').select();
}

function zfClose(ev) {
  if (ev && ev.target && !ev.target.classList.contains('picker-overlay')) return;
  document.getElementById('zone-form-overlay').style.display = 'none';
  p2pMap.pendingBox = null;
  p2pRedraw();
}

function zfConfirm() {
  const bbox = p2pMap.pendingBox;
  if (!bbox) { zfClose(); return; }
  const name = (document.getElementById('zf-name').value || '').trim();
  if (!name) { alert('Name is required.'); return; }
  const entry = document.getElementById('zf-entry').value;
  const exit  = document.getElementById('zf-exit').value;
  if (!entry || !exit) {
    alert('Pick both Entry LM and Exit LM (load a map with LMs first).');
    return;
  }
  if (p2pMap.subzoneTargetIdx !== null) {
    // Subzone mode — push into the parent zone's subzones array
    const zone = p2pMap.zones[p2pMap.subzoneTargetIdx];
    if (!zone.subzones) zone.subzones = [];
    zone.subzones.push({ name, entry, exit,
      minX: bbox.minX, maxX: bbox.maxX, minY: bbox.minY, maxY: bbox.maxY });
  } else {
    // Zone mode
    p2pMap.zones.push({ name, entry, exit,
      minX: bbox.minX, maxX: bbox.maxX, minY: bbox.minY, maxY: bbox.maxY,
      subzones: [] });
  }
  p2pMap.pendingBox = null;
  document.getElementById('zone-form-overlay').style.display = 'none';
  p2pRenderZonesTable();
  p2pRedraw();
}

// ── Zone list rendering ──
function p2pRenderZonesTable() {
  const lmNames = p2pMap.points.filter(p => p.type === 'LM').map(p => p.name);

  // Keep a single shared datalist in the document, refreshed each render
  let dl = document.getElementById('p2p-lm-datalist');
  if (!dl) {
    dl = document.createElement('datalist');
    dl.id = 'p2p-lm-datalist';
    document.body.appendChild(dl);
  }
  dl.innerHTML = lmNames.map(n => `<option value="${escHtml(n)}">`).join('');

  // Inline LM input cell — scroll dropdown + free typing
  const lmInp = (cur, attrs) => {
    const attrsStr = Object.entries(attrs).map(([k, v]) => `data-${k}="${v}"`).join(' ');
    return `<td><input type="text" list="p2p-lm-datalist" class="text-input lm-inp"
      value="${escHtml(cur)}"
      style="font-size:11px;padding:2px 6px;min-width:80px;width:100px"
      ${attrsStr} /></td>`;
  };

  const tbody = document.getElementById('p2p-zones-tbody');
  tbody.innerHTML = '';

  p2pMap.zones.forEach((z, zIdx) => {
    const szCount = (z.subzones || []).length;
    const isTarget = p2pMap.subzoneTargetIdx === zIdx;
    const tr = document.createElement('tr');
    if (isTarget) tr.style.outline = '1px solid #F7941D';
    tr.innerHTML = `
      <td>${zIdx + 1}</td>
      <td>${escHtml(z.name)}</td>
      ${lmInp(z.entry, { type:'zone', zi: zIdx, field:'entry' })}
      ${lmInp(z.exit,  { type:'zone', zi: zIdx, field:'exit'  })}
      <td>${z.minX.toFixed(2)}</td>
      <td>${z.maxX.toFixed(2)}</td>
      <td>${z.minY.toFixed(2)}</td>
      <td>${z.maxY.toFixed(2)}</td>
      <td style="white-space:nowrap">
        <button class="btn-ghost" data-add-sz="${zIdx}"
                style="padding:2px 7px;font-size:10px;color:${isTarget ? '#F7941D' : ''}">
          ⊕ ${szCount ? szCount + ' sz' : 'Subzone'}
        </button>
        <button class="btn-ghost" data-del-zone="${zIdx}"
                style="padding:2px 7px;font-size:10px;margin-left:2px">✕</button>
      </td>
    `;
    tbody.appendChild(tr);

    // Subzone rows
    (z.subzones || []).forEach((sz, sIdx) => {
      const szTr = document.createElement('tr');
      szTr.style.background = 'rgba(247,148,29,0.05)';
      szTr.innerHTML = `
        <td style="padding-left:18px;color:var(--text-lo);font-size:10px">↳ ${zIdx+1}.${sIdx+1}</td>
        <td style="font-size:11px;color:#F7941D">${escHtml(sz.name)}</td>
        ${lmInp(sz.entry, { type:'sz', zi: zIdx, si: sIdx, field:'entry' })}
        ${lmInp(sz.exit,  { type:'sz', zi: zIdx, si: sIdx, field:'exit'  })}
        <td style="font-size:11px">${sz.minX.toFixed(2)}</td>
        <td style="font-size:11px">${sz.maxX.toFixed(2)}</td>
        <td style="font-size:11px">${sz.minY.toFixed(2)}</td>
        <td style="font-size:11px">${sz.maxY.toFixed(2)}</td>
        <td><button class="btn-ghost" data-del-sz-zone="${zIdx}" data-del-sz-idx="${sIdx}"
                    style="padding:2px 7px;font-size:10px">✕</button></td>
      `;
      tbody.appendChild(szTr);
    });
  });

  // Wire LM inputs — update zones array on change (fires on pick or blur after typing)
  tbody.querySelectorAll('input.lm-inp').forEach(inp => {
    const commit = () => {
      const { type, zi, si, field } = inp.dataset;
      const val = inp.value.trim();
      if (!val) return;
      if (type === 'zone') {
        p2pMap.zones[+zi][field] = val;
      } else {
        p2pMap.zones[+zi].subzones[+si][field] = val;
      }
    };
    inp.addEventListener('change', commit);
    inp.addEventListener('blur',   commit);
  });

  tbody.querySelectorAll('button[data-add-sz]').forEach(btn => {
    btn.addEventListener('click', () => p2pStartSubzoneDraw(+btn.dataset.addSz));
  });
  tbody.querySelectorAll('button[data-del-zone]').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = +btn.dataset.delZone;
      if (p2pMap.subzoneTargetIdx === i) p2pMap.subzoneTargetIdx = null;
      p2pMap.zones.splice(i, 1);
      p2pRenderZonesTable();
      p2pRedraw();
    });
  });
  tbody.querySelectorAll('button[data-del-sz-zone]').forEach(btn => {
    btn.addEventListener('click', () => {
      p2pMap.zones[+btn.dataset.delSzZone].subzones.splice(+btn.dataset.delSzIdx, 1);
      p2pRenderZonesTable();
      p2pRedraw();
    });
  });

  document.getElementById('p2p-zone-count').textContent = p2pMap.zones.length;
  showCard('p2p-zones-card', p2pMap.zones.length > 0);
}

function p2pClearZones() {
  if (!p2pMap.zones.length) return;
  if (!confirm(`Remove all ${p2pMap.zones.length} drawn zone(s)?`)) return;
  p2pMap.zones = [];
  p2pRenderZonesTable();
  p2pRedraw();
}

/* ─────────────────────────────────────────────
   CASE PICK — Analyze & Generate
   ───────────────────────────────────────────── */

let cpAisles = [];
let cpSortAxis = 'x';
let cpApNames = [];
let cpAnalyzed = false;

async function analyzeCasePick() {
  const fileEl = document.getElementById('cp-file');
  if (!fileEl.files.length) { alert('Please select a JSON or SMAP file.'); return; }

  setBtnLoading('cp-analyze-btn', true, '⟠ Analyze');
  clearLog('cp-log');
  showCard('cp-result-card', false);
  showCard('cp-aisles-card', false);
  showCard('cp-output-card', false);
  showCard('cp-action-card', false);
  setStatus('Case Pick: analyzing…');

  const fd = new FormData();
  fd.append('file', fileEl.files[0]);
  fd.append('orientation', document.getElementById('cp-orientation').value);
  fd.append('cross_aisle', document.getElementById('cp-cross-aisle').checked);
  fd.append('sensitivity', document.getElementById('cp-sensitivity').value);
  fd.append('cross_aisle_markers', JSON.stringify([...cpMap.crossAisle]));

  try {
    const res = await fetch('/api/case-pick/analyze', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      showCard('cp-result-card');
      appendLog('cp-log', `Error: ${data.error}`, 'err');
      setStatus('Case Pick: error');
      return;
    }

    cpAisles = (data.aisles || []).map(a => ({ ...a, direction: 'f2l' }));
    cpSortAxis = data.sort_axis || 'x';
    cpApNames = data.ap_names || [];
    cpAnalyzed = true;

    document.getElementById('cp-aisle-count').textContent = data.aisle_count;
    document.getElementById('cp-ap-count').textContent = data.ap_count;
    renderAisleTable();
    showCard('cp-aisles-card');

    showCard('cp-output-card');
    showCard('cp-action-card');
    showCard('cp-result-card');
    clearLog('cp-log');
    (data.log || []).forEach(m => appendLog('cp-log', m, detectLevel(m)));
    setStatus(`Case Pick: ${data.aisle_count} aisles, ${data.ap_count} APs`);
  } catch (e) {
    showCard('cp-result-card');
    appendLog('cp-log', `Request failed: ${e}`, 'err');
    setStatus('Case Pick: failed');
  } finally {
    setBtnLoading('cp-analyze-btn', false, '⟠ Analyze');
  }
}

function renderAisleTable() {
  const tbody = document.getElementById('cp-aisle-tbody');
  tbody.innerHTML = '';
  cpAisles.forEach((a, idx) => {
    const tr = document.createElement('tr');
    const dirLabel = a.direction === 'f2l' ? 'First→Last' : 'Last→First';
    const nameVal = a.name || '';
    tr.innerHTML = `
      <td>${a.index}</td>
      <td><input type="text" class="aisle-name-input" data-idx="${idx}"
                 value="${escHtml(nameVal)}" placeholder="e.g. PLE" /></td>
      <td>${escHtml(a.first)}</td>
      <td>${escHtml(a.last)}</td>
      <td>${a.count}</td>
      <td class="dir-cell" data-idx="${idx}">${dirLabel}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('.dir-cell').forEach(cell => {
    cell.addEventListener('click', () => {
      const idx = parseInt(cell.dataset.idx);
      cpAisles[idx].direction = cpAisles[idx].direction === 'f2l' ? 'l2f' : 'f2l';
      cell.textContent = cpAisles[idx].direction === 'f2l' ? 'First→Last' : 'Last→First';
      _switchToManualMode();
    });
  });
  tbody.querySelectorAll('.aisle-name-input').forEach(inp => {
    inp.addEventListener('input', () => {
      const idx = parseInt(inp.dataset.idx);
      cpAisles[idx].name = inp.value;
    });
  });
}

function _switchToManualMode() {
  // No-op now that the Mode dropdown has been removed (sequence always runs
  // in manual mode honouring the table's per-aisle direction toggles).
}

function setAllDirections(dir) {
  cpAisles.forEach(a => a.direction = dir);
  renderAisleTable();
  _switchToManualMode();
}

function onModeChange() {
  // No-op — Mode dropdown removed; sequence is always manual.
}

function onDupToggle() {
  const on = document.getElementById('cp-dup-enabled').checked;
  document.getElementById('cp-dup-count-wrap').style.display = on ? '' : 'none';
}

async function generateCasePick() {
  if (!cpAnalyzed) { alert('Run Analyze first.'); return; }
  const fileEl = document.getElementById('cp-file');
  const mapPath = (document.getElementById('cp-map-path').value || '').trim();
  if (!fileEl.files.length && !mapPath) {
    alert('Please pick a map file (Browse) or enter a Map Path.');
    return;
  }

  setBtnLoading('cp-gen-btn', true, '▶ Generate');
  clearLog('cp-log');
  showCard('cp-result-card');
  showCard('cp-download-row', false);
  setStatus('Case Pick: generating sequence…');

  const dupEnabled = document.getElementById('cp-dup-enabled').checked;
  const dup = dupEnabled ? parseInt(document.getElementById('cp-dup-count').value) : 0;
  const templateEl = document.getElementById('cp-template');

  const fd = new FormData();
  const outputPath = (document.getElementById('cp-output-path').value || '').trim();
  // The map file is optional if a valid Map Path is given (server reads from disk)
  if (fileEl.files.length) fd.append('file', fileEl.files[0]);
  if (mapPath)             fd.append('map_path', mapPath);
  if (outputPath)          fd.append('output_path', outputPath);
  if (templateEl.files.length) fd.append('template', templateEl.files[0]);
  // Sequence always runs in manual mode; per-aisle direction comes from the
  // AISLES DETECTED table (default "f2l" if the user did not toggle).
  fd.append('mode', 'manual');
  fd.append('orientation', document.getElementById('cp-orientation').value);
  fd.append('cross_aisle', document.getElementById('cp-cross-aisle').checked);
  fd.append('sensitivity', document.getElementById('cp-sensitivity').value);
  fd.append('duplication', dup);
  fd.append('aisle_directions',
    JSON.stringify(cpAisles.map(a => ({ index: a.index, direction: a.direction }))));
  // Send the EXACT aisle structure (AP names + direction + name) the user saw
  // in the table — keeps per-aisle settings aligned with the aisles configured.
  fd.append('manual_aisles',
    JSON.stringify(cpAisles.map(a => ({
      aps: a.aps,
      direction: a.direction,
      name: (a.name || '').trim(),
    }))));
  fd.append('excluded_aps',         JSON.stringify([...cpMap.excluded]));
  fd.append('loading_aps',          JSON.stringify([...cpMap.loading]));
  fd.append('unloading_aps',        JSON.stringify([...cpMap.unloading]));
  fd.append('cross_aisle_markers',  JSON.stringify([...cpMap.crossAisle]));

  try {
    const res = await fetch('/api/case-pick/generate', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      appendLog('cp-log', `Error: ${data.error}`, 'err');
      setStatus('Case Pick: error');
      return;
    }
    (data.log || []).forEach(m => appendLog('cp-log', m, detectLevel(m)));
    if (data.saved_to) {
      // Server wrote the Excel to the user's disk directly — no download needed
      appendLog('cp-log', `✅ Saved to: ${data.saved_to}`, 'ok');
      showCard('cp-download-row', false);
      setStatus(`Case Pick: saved to ${data.saved_to}`);
    } else if (data.download) {
      setDownload('cp-download-link', 'cp-download-row', data.download,
                  data.filename || 'case_pick_sequence.xlsx');
      setStatus(`Case Pick: ${data.ap_count} APs written`);
    }
  } catch (e) {
    appendLog('cp-log', `Request failed: ${e}`, 'err');
    setStatus('Case Pick: failed');
  } finally {
    setBtnLoading('cp-gen-btn', false, '▶ Generate');
  }
}

/* ─────────────────────────────────────────────
   POINT TO POINT
   ───────────────────────────────────────────── */

function onNdeepToggle() {
  const on = document.getElementById('p2p-ndeep').checked;
  document.getElementById('p2p-ndeep-opts').style.display = on ? '' : 'none';
  document.getElementById('p2p-standard-opts').style.display = on ? 'none' : '';
}

async function generateP2P() {
  const jsonEl = document.getElementById('p2p-json');
  const excelEl = document.getElementById('p2p-excel');
  const useDrawn = document.getElementById('p2p-use-drawn').checked;

  if (!jsonEl.files.length) {
    alert('Please select a JSON / SMAP file.');
    return;
  }
  if (useDrawn) {
    if (!p2pMap.zones.length) {
      alert('No drawn zones. Toggle Draw Zone on the map and drag rectangles to define zones, or uncheck "Use drawn zones".');
      return;
    }
  } else if (!excelEl.files.length) {
    alert('Please select a Zone Config Excel file (or enable "Use drawn zones").');
    return;
  }

  setBtnLoading('p2p-gen-btn', true, '▶ Generate Zone Config');
  clearLog('p2p-log');
  showCard('p2p-result-card');
  showCard('p2p-download-row', false);
  setStatus('Point to Point: generating…');

  const ndeep = document.getElementById('p2p-ndeep').checked;
  const fd = new FormData();
  fd.append('json_file', jsonEl.files[0]);
  if (useDrawn) {
    fd.append('drawn_zones', JSON.stringify(p2pMap.zones));
  } else {
    fd.append('excel_file', excelEl.files[0]);
  }
  fd.append('entry_exit_type', document.getElementById('p2p-entry-type').value);
  fd.append('ndeep', ndeep);
  fd.append('drop_seq', document.getElementById('p2p-drop-seq').value);
  fd.append('zone_scope', document.getElementById('p2p-zone-scope').value);
  fd.append('sub_scope', document.getElementById('p2p-sub-scope').value);
  fd.append('loc_scope', document.getElementById('p2p-loc-scope').value);
  fd.append('create_sequence', document.getElementById('p2p-create-seq').checked);

  try {
    const res = await fetch('/api/point-to-point/generate', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      appendLog('p2p-log', `Error: ${data.error}`, 'err');
      setStatus('Point to Point: error');
      return;
    }
    (data.log || []).forEach(m => appendLog('p2p-log', m, detectLevel(m)));
    setDownload('p2p-download-link', 'p2p-download-row', data.download, data.filename);
    setStatus(`Point to Point: ${data.zones} zone(s) generated`);
  } catch (e) {
    appendLog('p2p-log', `Request failed: ${e}`, 'err');
    setStatus('Point to Point: failed');
  } finally {
    setBtnLoading('p2p-gen-btn', false, '▶ Generate Zone Config');
  }
}

/* ─────────────────────────────────────────────
   MAP VALIDATOR (SSE)
   ───────────────────────────────────────────── */

let mvEventSource = null;

async function runMapValidator() {
  const fileEl = document.getElementById('mv-file');
  if (!fileEl.files.length) { alert('Please select a JSON or SMAP file.'); return; }

  setBtnLoading('mv-run-btn', true, '▶ Validate');
  clearLog('mv-log');
  document.getElementById('mv-issues-tbody').innerHTML = '';
  showCard('mv-summary', false);
  document.getElementById('mv-live-dot').style.display = '';
  setStatus('Map Validator: uploading…');

  if (mvEventSource) { mvEventSource.close(); mvEventSource = null; }

  const fd = new FormData();
  fd.append('file', fileEl.files[0]);

  let streamId;
  try {
    const res = await fetch('/api/map-validator/validate', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      appendLog('mv-log', `Error: ${data.error}`, 'err');
      setBtnLoading('mv-run-btn', false, '▶ Validate');
      return;
    }
    streamId = data.stream_id;
  } catch (e) {
    appendLog('mv-log', `Upload failed: ${e}`, 'err');
    setBtnLoading('mv-run-btn', false, '▶ Validate');
    return;
  }

  setStatus('Map Validator: validating…');
  mvEventSource = new EventSource(`/api/map-validator/stream/${streamId}`);

  mvEventSource.onmessage = (ev) => {
    let msg;
    try { msg = JSON.parse(ev.data); } catch { return; }

    if (msg.type === 'log') {
      appendLog('mv-log', msg.msg, msg.level || 'info');
    }
    if (msg.type === 'result') {
      renderValidatorResult(msg.data);
    }
    if (msg.type === 'done') {
      mvEventSource.close();
      mvEventSource = null;
      document.getElementById('mv-live-dot').style.display = 'none';
      setBtnLoading('mv-run-btn', false, '▶ Validate');
      setStatus('Map Validator: complete');
    }
  };

  mvEventSource.onerror = () => {
    appendLog('mv-log', 'Stream connection lost', 'err');
    mvEventSource.close();
    mvEventSource = null;
    document.getElementById('mv-live-dot').style.display = 'none';
    setBtnLoading('mv-run-btn', false, '▶ Validate');
    setStatus('Map Validator: error');
  };
}

function renderValidatorResult(data) {
  const tbody = document.getElementById('mv-issues-tbody');
  tbody.innerHTML = '';
  const MAX_ROWS = 200;
  let rowCount = 0;

  (data.issues || []).forEach(([name, ptType, issue]) => {
    if (rowCount >= MAX_ROWS) return;
    const tr = document.createElement('tr');
    const level = issue.includes('exit') && !issue.includes('entry') ? 'warn' : 'err';
    tr.innerHTML = `
      <td class="${level}">${escHtml(name)}</td>
      <td>${escHtml(ptType)}</td>
      <td class="${level}">${escHtml(issue)}</td>
    `;
    tbody.appendChild(tr);
    rowCount++;
  });

  if ((data.conn_rows || []).length) {
    const sep = document.createElement('tr');
    sep.innerHTML = `<td class="conn" colspan="3">── Connectivity: ${data.n_conn} unreachable direction(s) · ${data.n_sccs} group(s)</td>`;
    tbody.appendChild(sep);
    (data.conn_rows || []).forEach(row => {
      if (rowCount >= MAX_ROWS) return;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="conn">${escHtml(row.src_label)}</td>
        <td class="conn">${escHtml(row.src_type)}</td>
        <td class="conn">${escHtml(row.issue)}</td>
      `;
      tbody.appendChild(tr);
      rowCount++;
    });
    if (rowCount >= MAX_ROWS) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td colspan="3" style="color:var(--text-lo);text-align:center">… results capped at ${MAX_ROWS} rows</td>`;
      tbody.appendChild(tr);
    }
  }

  const total = data.total || 0;
  const nIssues = data.n_issues || 0;
  const nConn = data.n_conn || 0;
  const nOk = total - nIssues;
  document.getElementById('mv-ok-badge').textContent = `✓ ${nOk} OK`;
  document.getElementById('mv-warn-badge').textContent = `⚠ ${nIssues} path issue(s)`;
  document.getElementById('mv-conn-badge').textContent = `⬡ ${nConn} conn gap(s)`;
  document.getElementById('mv-total-badge').textContent = `◎ ${total} total`;
  showCard('mv-summary');
}

/* ─────────────────────────────────────────────
   Utilities
   ───────────────────────────────────────────── */

function detectLevel(msg) {
  if (/error|fail|❌/i.test(msg)) return 'err';
  if (/warn|⚠/i.test(msg)) return 'warn';
  if (/✅|success|complete|done/i.test(msg)) return 'ok';
  return 'info';
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ─────────────────────────────────────────────
   SERVER FILE PICKER MODAL
   Browses the local filesystem on the machine running the Flask app and
   returns a real absolute path. Used to populate the Map Path field with
   the user's actual on-disk path (which the browser refuses to expose).
   ───────────────────────────────────────────── */

const _picker = {
  targetInputId: null,
  ext: '',
  currentPath: '',
  parentPath: null,
};

function openServerPicker(targetInputId, ext = '') {
  _picker.targetInputId = targetInputId;
  _picker.ext = ext;
  document.getElementById('picker-overlay').style.display = 'flex';
  // Seed from existing value's directory if it looks absolute, else start at root
  const cur = (document.getElementById(targetInputId).value || '').trim();
  let startPath = '';
  if (cur && /^[A-Za-z]:[\\/]/.test(cur)) {
    // Windows absolute path — use its directory if it's a file
    const dirOnly = cur.replace(/[\\/][^\\/]*$/, '');
    startPath = dirOnly;
  } else if (cur.startsWith('/')) {
    startPath = cur.replace(/\/[^/]*$/, '');
  }
  _pickerLoad(startPath);
}

function closeServerPicker(ev) {
  if (ev && ev.target && !ev.target.classList.contains('picker-overlay')) return;
  document.getElementById('picker-overlay').style.display = 'none';
}

async function _pickerLoad(path) {
  const list = document.getElementById('picker-list');
  list.innerHTML = '<li class="is-empty"><span class="pi-icon">⌛</span>Loading…</li>';
  const url = `/api/fs/list?path=${encodeURIComponent(path)}&ext=${encodeURIComponent(_picker.ext)}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (!res.ok) {
      list.innerHTML = `<li class="is-empty"><span class="pi-icon">⚠</span>${escHtml(data.error || 'Error')}</li>`;
      return;
    }
    _picker.currentPath = data.path || '';
    _picker.parentPath = data.parent;
    document.getElementById('picker-path-input').value = data.path || '';
    _pickerRender(data);
  } catch (e) {
    list.innerHTML = `<li class="is-empty"><span class="pi-icon">⚠</span>${escHtml(String(e))}</li>`;
  }
}

function _pickerRender(data) {
  const list = document.getElementById('picker-list');
  const items = [];
  (data.folders || []).forEach(name => {
    const li = document.createElement('li');
    li.className = 'is-folder';
    li.innerHTML = `<span class="pi-icon">📁</span>${escHtml(name)}`;
    li.addEventListener('click', () => {
      // On root (Windows drives) the entry already contains a trailing backslash
      const next = data.is_root
        ? name
        : (data.path.endsWith('/') || data.path.endsWith('\\')
            ? data.path + name
            : data.path + (data.path.includes('\\') ? '\\' : '/') + name);
      _pickerLoad(next);
    });
    items.push(li);
  });
  (data.files || []).forEach(name => {
    const li = document.createElement('li');
    li.className = 'is-file';
    li.innerHTML = `<span class="pi-icon">📄</span>${escHtml(name)}`;
    li.addEventListener('click', () => {
      const sep = data.path.includes('\\') ? '\\' : '/';
      const fullPath = data.path.endsWith(sep) ? data.path + name : data.path + sep + name;
      _pickerPickFile(fullPath);
    });
    items.push(li);
  });
  if (!items.length) {
    list.innerHTML = '<li class="is-empty"><span class="pi-icon">∅</span>No items at this level</li>';
    return;
  }
  list.innerHTML = '';
  items.forEach(li => list.appendChild(li));
}

function pickerGoUp() {
  if (_picker.parentPath || _picker.parentPath === '') {
    _pickerLoad(_picker.parentPath || '');
  }
}

function pickerNavigate(path) {
  _pickerLoad((path || '').trim());
}

function _pickerPickFile(fullPath) {
  const target = document.getElementById(_picker.targetInputId);
  if (target) target.value = fullPath;
  closeServerPicker();
}

/* ─────────────────────────────────────────────
   ORDERFILE GENERATOR
   ───────────────────────────────────────────── */

let ofRawStats = null;
let ofEventSource = null;

function ofToggleLimits() {
  const on = document.getElementById('of-use-limits').checked;
  document.getElementById('of-limits-fields').style.display = on ? '' : 'none';
}

async function ofOnFileSelected(file) {
  setStatus('Orderfile: analyzing data…');
  showCard('of-summary-card', false);
  showCard('of-targets-card', false);
  showCard('of-limits-card', false);
  showCard('of-progress-card', false);
  showCard('of-result-card', false);
  ofRawStats = null;

  const fd = new FormData();
  fd.append('file', file);

  try {
    const res = await fetch('/api/orderfile/analyze', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      setStatus(`Orderfile: ${data.error}`);
      alert(`Error analyzing file: ${data.error}`);
      return;
    }
    ofRawStats = data;
    document.getElementById('of-raw-orders').textContent = data.raw_orders.toLocaleString();
    document.getElementById('of-raw-lines').textContent  = data.raw_lines.toLocaleString();
    document.getElementById('of-raw-units').textContent  = Math.round(data.raw_units).toLocaleString();

    // Seed target inputs with sensible defaults
    const tOrd  = document.getElementById('of-target-orders');
    const tLin  = document.getElementById('of-target-lines');
    const tUnit = document.getElementById('of-target-units');
    tOrd.max  = data.raw_orders;
    tLin.max  = data.raw_lines;
    tUnit.max = Math.round(data.raw_units);
    tOrd.value  = Math.min(data.raw_orders, 100);
    tLin.value  = Math.min(data.raw_lines, 500);
    tUnit.value = Math.min(Math.round(data.raw_units), 5000);

    showCard('of-summary-card');
    showCard('of-targets-card');
    showCard('of-limits-card');
    setStatus(`Orderfile: ${data.raw_orders.toLocaleString()} orders loaded`);
  } catch (e) {
    setStatus(`Orderfile: failed — ${e}`);
    alert(`Failed to analyze file: ${e}`);
  }
}

async function generateOrderfile() {
  const fileEl = document.getElementById('of-file');
  if (!fileEl.files.length) { alert('Please select a raw order data file.'); return; }
  if (!ofRawStats)           { alert('File has not been analyzed yet.'); return; }

  const targetOrders = parseInt(document.getElementById('of-target-orders').value) || 0;
  const targetLines  = parseInt(document.getElementById('of-target-lines').value)  || 0;
  const targetUnits  = parseFloat(document.getElementById('of-target-units').value) || 0;
  if (!targetOrders || !targetLines || !targetUnits) {
    alert('All target values must be greater than zero.');
    return;
  }

  const useLimits   = document.getElementById('of-use-limits').checked;
  const maxQtyOrder = useLimits ? parseFloat(document.getElementById('of-max-order').value) : null;
  const minQtyOrder = useLimits ? parseFloat(document.getElementById('of-min-order').value) : null;
  const maxQtyLine  = useLimits ? parseFloat(document.getElementById('of-max-line').value)  : null;

  // Client-side feasibility check
  if (useLimits) {
    if (maxQtyOrder !== null && maxQtyLine !== null) {
      const maxAchievable = Math.min(maxQtyOrder, targetLines * maxQtyLine) * targetOrders;
      if (maxAchievable < targetUnits) {
        alert(`Target is unreachable: max achievable units (${Math.round(maxAchievable).toLocaleString()}) < target units (${Math.round(targetUnits).toLocaleString()}).\n\nFix: increase Max qty per order, reduce Target Units, or remove limits.`);
        return;
      }
    }
    if (minQtyOrder !== null && minQtyOrder * targetOrders > targetUnits) {
      alert(`Target is unreachable: min required units (${Math.round(minQtyOrder * targetOrders).toLocaleString()}) > target units (${Math.round(targetUnits).toLocaleString()}).\n\nFix: reduce Min qty per order or increase Target Units.`);
      return;
    }
  }

  setBtnLoading('of-gen-btn', true, '▶ Generate Orderfile');
  clearLog('of-log');
  showCard('of-progress-card');
  showCard('of-result-card', false);
  document.getElementById('of-live-dot').style.display = '';
  document.getElementById('of-progress-bar').style.width = '0%';
  document.getElementById('of-progress-text').textContent = 'Starting…';
  setStatus('Orderfile: optimizing…');

  if (ofEventSource) { ofEventSource.close(); ofEventSource = null; }

  const fd = new FormData();
  fd.append('file', fileEl.files[0]);
  fd.append('target_orders', targetOrders);
  fd.append('target_lines',  targetLines);
  fd.append('target_units',  targetUnits);
  fd.append('use_limits', useLimits);
  if (useLimits) {
    if (maxQtyOrder !== null) fd.append('max_qty_per_order', maxQtyOrder);
    if (minQtyOrder !== null) fd.append('min_qty_per_order', minQtyOrder);
    if (maxQtyLine  !== null) fd.append('max_qty_per_line',  maxQtyLine);
  }

  let streamId;
  try {
    const res = await fetch('/api/orderfile/generate', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) {
      appendLog('of-log', `Error: ${data.error}`, 'err');
      setBtnLoading('of-gen-btn', false, '▶ Generate Orderfile');
      document.getElementById('of-live-dot').style.display = 'none';
      setStatus('Orderfile: error');
      return;
    }
    streamId = data.stream_id;
  } catch (e) {
    appendLog('of-log', `Upload failed: ${e}`, 'err');
    setBtnLoading('of-gen-btn', false, '▶ Generate Orderfile');
    document.getElementById('of-live-dot').style.display = 'none';
    setStatus('Orderfile: failed');
    return;
  }

  ofEventSource = new EventSource(`/api/orderfile/stream/${streamId}`);

  ofEventSource.onmessage = (ev) => {
    let msg;
    try { msg = JSON.parse(ev.data); } catch { return; }

    if (msg.type === 'progress') {
      document.getElementById('of-progress-bar').style.width = `${msg.pct}%`;
      document.getElementById('of-progress-text').textContent = msg.msg || '';
      appendLog('of-log', msg.msg || '', 'info');
    }
    if (msg.type === 'log') {
      appendLog('of-log', msg.msg, msg.level || 'info');
    }
    if (msg.type === 'result') {
      renderOrderfileResult(msg);
    }
    if (msg.type === 'done') {
      ofEventSource.close();
      ofEventSource = null;
      document.getElementById('of-live-dot').style.display = 'none';
      document.getElementById('of-progress-bar').style.width = '100%';
      setBtnLoading('of-gen-btn', false, '▶ Generate Orderfile');
      setStatus('Orderfile: complete');
    }
  };

  ofEventSource.onerror = () => {
    appendLog('of-log', 'Stream connection lost', 'err');
    if (ofEventSource) ofEventSource.close();
    ofEventSource = null;
    document.getElementById('of-live-dot').style.display = 'none';
    setBtnLoading('of-gen-btn', false, '▶ Generate Orderfile');
    setStatus('Orderfile: error');
  };
}

function renderOrderfileResult(msg) {
  const s = msg.stats || {};
  const aOrd = s.actual_orders || 0;
  const aLin = s.actual_lines  || 0;
  const aQty = s.actual_qty    || 0;
  const tOrd = s.target_orders || 0;
  const tLin = s.target_lines  || 0;
  const tQty = s.target_units  || 0;

  const fmt  = n => Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });
  const fmtF = n => Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const delta = (a, t) => { const d = a - t; return (d >= 0 ? '+' : '') + fmt(d); };

  document.getElementById('of-r-orders').textContent = `Orders: ${fmt(aOrd)} (${delta(aOrd, tOrd)})`;
  document.getElementById('of-r-lines').textContent  = `Lines: ${fmt(aLin)} (${delta(aLin, tLin)})`;
  document.getElementById('of-r-units').textContent  = `Units: ${fmt(aQty)} (${delta(aQty, tQty)})`;

  const upl = aLin > 0 ? aQty / aLin : 0;
  const upo = aOrd > 0 ? aQty / aOrd : 0;
  const lpo = aOrd > 0 ? aLin / aOrd : 0;
  document.getElementById('of-r-upl').textContent = `Units/Line: ${fmtF(upl)}`;
  document.getElementById('of-r-upo').textContent = `Units/Order: ${fmtF(upo)}`;
  document.getElementById('of-r-lpo').textContent = `Lines/Order: ${fmtF(lpo)}`;

  if (msg.download) {
    setDownload('of-download-link', 'of-download-row', msg.download,
                msg.filename || 'orderfile_output.xlsx');
  }
  showCard('of-result-card');
}


// ── Robot Map Sync ────────────────────────────

let msEventSource = null;

function msUpdatePreview() {
  const baseDir   = (document.getElementById('ms-base-dir').value   || '').trim();
  const prefix    = (document.getElementById('ms-prefix').value      || '').trim();
  const subFolder = (document.getElementById('ms-subfolder').value   || '').trim();
  const start     = parseInt(document.getElementById('ms-start').value)  || 1;
  const count     = parseInt(document.getElementById('ms-count').value)  || 0;
  const hasFile   = document.getElementById('ms-file').files.length > 0;
  const card      = document.getElementById('ms-preview-card');

  if (!baseDir || !prefix || count <= 0 || !hasFile) { card.style.display = 'none'; return; }

  const sep = baseDir.includes('/') ? '/' : '\\';
  const buildPath = n => {
    const r = `${prefix}${n}`;
    return subFolder ? `${baseDir}${sep}${r}${sep}${subFolder}` : `${baseDir}${sep}${r}`;
  };

  document.getElementById('ms-p-count').textContent = count.toLocaleString();
  document.getElementById('ms-p-first').textContent = buildPath(start);
  document.getElementById('ms-p-last').textContent  = buildPath(start + count - 1);
  card.style.display = '';
}

async function msExecute() {
  const file      = document.getElementById('ms-file').files[0];
  const baseDir   = (document.getElementById('ms-base-dir').value   || '').trim();
  const prefix    = (document.getElementById('ms-prefix').value      || '').trim();
  const subFolder = (document.getElementById('ms-subfolder').value   || '').trim();
  const start     = parseInt(document.getElementById('ms-start').value)  || 1;
  const count     = parseInt(document.getElementById('ms-count').value)  || 0;

  if (!file || !baseDir || !prefix || count <= 0) {
    alert('Please fill all required fields and select a source file.'); return;
  }

  if (msEventSource) { msEventSource.close(); msEventSource = null; }

  document.getElementById('ms-log').innerHTML = '';
  document.getElementById('ms-progress-bar').style.width = '0%';
  document.getElementById('ms-progress-text').textContent = '0%';
  document.getElementById('ms-result-card').style.display = 'none';
  showCard('ms-progress-card');
  document.getElementById('ms-live-dot').style.display = '';
  setBtnLoading('ms-run-btn', true, '⏳ Running…');

  const fd = new FormData();
  fd.append('file',       file);
  fd.append('base_dir',   baseDir);
  fd.append('prefix',     prefix);
  fd.append('sub_folder', subFolder);
  fd.append('start',      start);
  fd.append('count',      count);

  let res;
  try { res = await fetch('/api/map-sync/execute', { method: 'POST', body: fd }); }
  catch (e) { alert(`Network error: ${e}`); setBtnLoading('ms-run-btn', false, '▶ Copy to Robots'); return; }

  const body = await res.json();
  if (!res.ok) {
    alert(body.error || 'Request failed');
    setBtnLoading('ms-run-btn', false, '▶ Copy to Robots');
    return;
  }

  msEventSource = new EventSource(`/api/map-sync/stream/${body.stream_id}`);
  msEventSource.onmessage = e => {
    const msg = JSON.parse(e.data);
    if (msg.type === 'progress') {
      document.getElementById('ms-progress-bar').style.width  = msg.pct + '%';
      document.getElementById('ms-progress-text').textContent = msg.pct + '%';
      appendLog('ms-log', msg.msg, msg.ok === false ? 'err' : 'ok');
    } else if (msg.type === 'result') {
      msRenderResult(msg);
      setBtnLoading('ms-run-btn', false, '▶ Copy to Robots');
      document.getElementById('ms-live-dot').style.display = 'none';
      setStatus(`Map Sync: ${msg.success}/${msg.total} copied`);
    } else if (msg.type === 'done') {
      if (msEventSource) { msEventSource.close(); msEventSource = null; }
      document.getElementById('ms-live-dot').style.display = 'none';
      setBtnLoading('ms-run-btn', false, '▶ Copy to Robots');
    }
  };
  msEventSource.onerror = () => {
    appendLog('ms-log', 'Stream connection lost', 'err');
    if (msEventSource) { msEventSource.close(); msEventSource = null; }
    document.getElementById('ms-live-dot').style.display = 'none';
    setBtnLoading('ms-run-btn', false, '▶ Copy to Robots');
    setStatus('Map Sync: error');
  };
}

function msRenderResult(msg) {
  document.getElementById('ms-result-grid').innerHTML = `
    <div class="stat-box">
      <div class="stat-val">${msg.total}</div>
      <div class="stat-key">Total</div>
    </div>
    <div class="stat-box">
      <div class="stat-val" style="color:var(--ok-color)">${msg.success}</div>
      <div class="stat-key">Copied</div>
    </div>
    <div class="stat-box">
      <div class="stat-val" style="color:${msg.failed ? 'var(--err-color)' : 'var(--ok-color)'}">${msg.failed}</div>
      <div class="stat-key">Failed</div>
    </div>
  `;
  showCard('ms-result-card');
}
