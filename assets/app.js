'use strict';

// ---------- DOM + utility ----------
const $ = sel => document.querySelector(sel);
const log = (...args) => { const el = $('#log'); el.textContent += args.join(' ') + "\n"; el.scrollTop = el.scrollHeight; };
const human = n => n.toLocaleString();

function normalizePath(p){ return p.replaceAll('\\','/').replace(/^\.\//,'').replace(/\/{2,}/g,'/'); }
function pathJoin(){ return normalizePath(Array.from(arguments).filter(Boolean).join('/')).replace(/(^\/+)|((?<=:)\/+)/g,''); }
function splitPath(p){ p = normalizePath(p); const parts = p.split('/'); const file = parts.pop(); return {parts,file}; }

// ---------- PN/DRW detection (deterministic mapper) ----------
let detectedRoot = '';
let lastAutoRoot = '';
let lastFileName = '';
let inZip = null;
let inEntries = [];   // {path, entry}
let lastManifest = null; // built from last preview/transform

function detectTopRoot(entries){
  const firstSegs = new Set(entries.map(e => normalizePath(e.path).split('/')[0]));
  if(firstSegs.size === 1){
    const seg = Array.from(firstSegs)[0];
    if(!/\.[a-z0-9]+$/i.test(seg)) return seg; // looks like a folder
  }
  return '';
}

function extractIdCandidate(text){
  if(!text) return '';
  const m = /(?:RFQ|PO)\s*([0-9]+)/i.exec(text);
  return m ? m[1] : '';
}

function autoRootForMode(mode){
  const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName);
  if(mode === 'po') return id ? `PO${id} Drawing package` : 'PO Drawing package';
  return id ? `RFQ${id} Drawing package` : 'RFQ Drawing package';
}

function findPnDrwIndices(dirs){
  const pnIdx = dirs.findIndex(s => /^PN[^/]+$/i.test(s));
  if(pnIdx === -1 || pnIdx+1 >= dirs.length) return {pnIdx:-1, drwIdx:-1};
  const drwIdx = /^DRW[^/]+$/i.test(dirs[pnIdx+1]) ? pnIdx+1 : -1;
  return {pnIdx, drwIdx};
}

function mapDeterministic(relPath, opts){
  const p = normalizePath(relPath);
  const sp = splitPath(p);
  // strip segments
  let dirs = sp.parts;
  if(opts.stripList?.length){
    const lower = new Set(opts.stripList.map(s=>String(s).toLowerCase()));
    dirs = dirs.filter(seg => !lower.has(String(seg).toLowerCase()));
  }
  // ignore OS cruft
  if(sp.file.endsWith('.DS_Store') || sp.file.endsWith('Thumbs.db')) return null;

  const {pnIdx, drwIdx} = findPnDrwIndices(dirs);
  if(pnIdx === -1 || drwIdx === -1) return opts.includeUnknown ? pathJoin(opts.root, 'attachments', sp.file) : null;

  const ext = (sp.file.split('.').pop() || '').toLowerCase();
  const keepPO = new Set(['pdf','dxf','stp','step']);
  const keepRFQ = new Set(['pdf']);
  const ok = opts.mode === 'po' ? keepPO.has(ext) : keepRFQ.has(ext);
  if(!ok) return opts.includeUnknown ? pathJoin(opts.root, 'attachments', sp.file) : null;

  // preserve PN/DRW path
  return pathJoin(opts.root, dirs[pnIdx], dirs[drwIdx], sp.file);
}

function collisionSafePath(destPath, seen){
  let p = destPath;
  if(!seen.has(p)) { seen.add(p); return p; }
  const m = /(.*?)(?: \((\d+)\))?(\.[^.]+)?$/.exec(destPath);
  let base = m[1], n = parseInt(m[2]||'1',10), ext = m[3]||'';
  do { n++; p = `${base} (${n})${ext}`; } while(seen.has(p));
  seen.add(p);
  return p;
}

// ---------- Load ZIP ----------
async function loadZip(file){
  if(!window.JSZip){ alert('JSZip failed to load. Check network or self-host JSZip.'); return; }
  $('#bar').style.width = '0%';
  $('#stats').textContent = 'Reading zip…';
  inZip = await JSZip.loadAsync(file);
  inEntries = [];
  inZip.forEach((relPath, entry) => { if(!entry.dir) inEntries.push({ path: normalizePath(relPath), entry }); });
  $('#stats').textContent = `Loaded ${human(inEntries.length)} files`;
  $('#btnGo').disabled = false; $('#btnPreview').disabled = false;
  $('#btnPdfRFQ').disabled = true; $('#btnPdfPO').disabled = true;
  $('#fileInfo').innerHTML = `<span class="filetag">${file.name}</span> <span class="badge">${(file.size/1e6).toFixed(1)} MB</span>`;
  lastFileName = file.name;
  log('Zip loaded:', file.name, `(${human(inEntries.length)} files)`);

  detectedRoot = detectTopRoot(inEntries);
  log('Detected top folder:', detectedRoot || '(none)');

  // auto-suggest root based on mode + detected id
  const mode = $('#mode').value;
  const suggested = autoRootForMode(mode);
  const cur = $('#rootName').value.trim();
  if(cur === '' || cur === lastAutoRoot || /^PO\b.*Drawing package$/i.test(cur) || /^RFQ\b.*Drawing package$/i.test(cur)){
    $('#rootName').value = suggested;
    lastAutoRoot = suggested;
  }
}

// ---------- Manifest builder ----------
function buildManifest(opts){
  const mapFn = (rp)=>mapDeterministic(rp, opts);
  const items = [];
  for(const e of inEntries){
    const dest = mapFn(e.path);
    if(!dest) continue;
    const parts = normalizePath(dest).split('/');
    // expect <root>/<PN>/<DRW>/<file>
    const pn = parts.length>=3 ? parts[parts.length-3] : null;
    const drw = parts.length>=2 ? parts[parts.length-2] : null;
    const name = parts[parts.length-1];
    const ext = (name.split('.').pop()||'').toLowerCase();
    items.push({ src:e.path, dest, pn, drw, name, ext });
  }
  // group by PN/DRW
  const groups = new Map(); // key: `${pn}/${drw}`
  for(const it of items){
    const key = `${it.pn}/${it.drw}`;
    if(!groups.has(key)) groups.set(key, { pn: it.pn, drw: it.drw, files: [] });
    groups.get(key).files.push(it);
  }
  const grouped = Array.from(groups.values()).sort((a,b)=> (a.pn+a.drw).localeCompare(b.pn+b.drw));
  return { root: opts.root, mode: opts.mode, items, grouped };
}

// ---------- Preview ----------
function buildTreeFromItems(items){
  const tree = {};
  for(const it of items){
    const parts = normalizePath(it.dest).split('/');
    let node = tree;
    for(let i=0;i<parts.length;i++){
      const part = parts[i];
      const isFile = i === parts.length-1;
      if(!node[part]) node[part] = isFile ? null : {};
      node = node[part]||{};
    }
  }
  function render(node, prefix=''){
    const keys = Object.keys(node).sort((a,b)=>{
      const isDirA = node[a] !== null, isDirB = node[b] !== null;
      if(isDirA !== isDirB) return isDirA? -1 : 1;
      return a.localeCompare(b);
    });
    let out = '';
    for(const k of keys){
      const v = node[k];
      out += `${prefix}${v===null?'├──':'└──'} ${k}\n`;
      if(v) out += render(v, prefix + (v===null? '': '    '));
    }
    return out;
  }
  return render(tree);
}

async function preview(){
  const opts = currentOpts();
  const manifest = buildManifest(opts);
  lastManifest = manifest;
  $('#preview').textContent = buildTreeFromItems(manifest.items);
  log('Preview built for', human(manifest.items.length), 'files across', human(manifest.grouped.length), 'drawings');
  // Enable PDF buttons now that we have a manifest
  $('#btnPdfRFQ').disabled = false;
  $('#btnPdfPO').disabled = false;
}

// ---------- Transform ----------
async function transform(){
  const t0 = performance.now();
  const opts = currentOpts();
  const manifest = buildManifest(opts);
  lastManifest = manifest;

  const out = new JSZip();
  const seen = new Set();
  let n = 0; const total = manifest.items.length;
  $('#bar').style.width = '0%';
  for(const it of manifest.items){
    const data = await inZip.file(it.src).async('arraybuffer');
    const dest = collisionSafePath(it.dest, seen);
    out.file(dest, data, {date: new Date()});
    n++;
    if(n%10===0 || n===total){ $('#bar').style.width = ((n/total)*100).toFixed(1)+'%'; await new Promise(r=>requestAnimationFrame(r)); }
  }

  $('#stats').textContent = `Packaging ${human(n)} files…`;
  const blob = await out.generateAsync({type:'blob', compression:'DEFLATE', compressionOptions:{level:6}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `${opts.root||'package'}.zip`;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(url), 15000);

  const dt = ((performance.now()-t0)/1000).toFixed(2);
  $('#stats').textContent = `Done: ${human(n)} files → ${(blob.size/1e6).toFixed(1)} MB in ${dt}s`;
  $('#bar').style.width = '100%';
  log('ok: zip ready for download');

  // Enable PDF buttons (manifest populated)
  $('#btnPdfRFQ').disabled = false;
  $('#btnPdfPO').disabled = false;
}

// ---------- Options ----------
function currentOpts(){
  const mode = $('#mode').value;
  const root = $('#rootName').value.trim() || (mode==='po'?'PO Drawing package':'RFQ Drawing package');
  const stripList = ['__macosx','export','out'];
  if($('#stripTop').checked && detectedRoot){ stripList.push(detectedRoot); }
  return { mode, root, stripList, includeUnknown: $('#includeUnknown').checked };
}

// ---------- Wire-up ----------
(function init(){
  $('#mode').addEventListener('change', ()=>{
    const newMode = $('#mode').value;
    const suggested = autoRootForMode(newMode);
    const cur = $('#rootName').value.trim();
    if(cur === '' || cur === lastAutoRoot || /^PO\b.*Drawing package$/i.test(cur) || /^RFQ\b.*Drawing package$/i.test(cur)){
      $('#rootName').value = suggested; lastAutoRoot = suggested;
    }
  });

  const drop = $('#drop');
  const input = $('#file');
  drop.addEventListener('dragover', e=>{ e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', ()=> drop.classList.remove('dragover'));
  drop.addEventListener('drop', e=>{
    e.preventDefault(); drop.classList.remove('dragover');
    const f = e.dataTransfer.files[0]; if(!f) return;
    if(!/\.zip$/i.test(f.name)){ log('error: not a .zip'); return; }
    loadZip(f).catch(err=>{ log('error:', err.message); });
  });
  drop.addEventListener('click', ()=> input.click());
  input.addEventListener('change', ()=>{
    const f = input.files[0]; if(!f) return;
    if(!/\.zip$/i.test(f.name)){ log('error: not a .zip'); return; }
    loadZip(f).catch(err=>{ log('error:', err.message); });
  });

  $('#btnPreview').addEventListener('click', ()=>{ if(inZip) preview().catch(err=>log('error:',err.message)); });
  $('#btnGo').addEventListener('click', ()=>{ if(inZip) transform().catch(err=>log('error:',err.message)); });

  // PDF buttons
  $('#btnPdfRFQ').addEventListener('click', ()=>{
    if(!lastManifest){ const opts=currentOpts(); lastManifest=buildManifest(opts); }
    window.pdfBuilder && window.pdfBuilder.buildRFQ(lastManifest, { titleFrom: $('#rootName').value });
  });
  $('#btnPdfPO').addEventListener('click', ()=>{
    if(!lastManifest){ const opts=currentOpts(); lastManifest=buildManifest(opts); }
    window.pdfBuilder && window.pdfBuilder.buildPO(lastManifest, { titleFrom: $('#rootName').value });
  });
})();
