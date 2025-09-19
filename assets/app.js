
'use strict';
const $ = sel => document.querySelector(sel);
const human = n => n.toLocaleString();
function normalizePath(p){ return p.replace(/\\/g,'/').replace(/^\.\//,'').replace(/\/{2,}/g,'/'); }
function splitPath(p){ p = normalizePath(p); const parts = p.split('/'); const file = parts.length?parts[parts.length-1]:''; if(parts.length) parts.pop(); return {parts,file}; }
function pathJoin(){ return Array.from(arguments).filter(Boolean).join('/').replace(/\/{2,}/g,'/'); }

let inZip = null, inEntries = [], detectedRoot = '', lastFileName = '', extractedTitle = '';

function detectTopRoot(entries){
  const firstSegs = new Set(entries.map(e => normalizePath(e.path).split('/')[0]));
  if(firstSegs.size === 1){ const seg = Array.from(firstSegs)[0]; if(!/\.[a-z0-9]+$/i.test(seg)) return seg; }
  return '';
}
function extractIdCandidate(text){ if(!text) return ''; const m = /(?:RFQ|PO)\s*([0-9]+(?:-\d+)*)/i.exec(text); return m ? m[1] : ''; }
function autoRootForMode(mode){
  const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName);
  if(mode === 'po') return id ? `PO${id} Drawing package` : 'PO Drawing package';
  return id ? `RFQ${id} Drawing package` : 'RFQ Drawing package';
}
function findPnDrwIndices(dirs){ const pnIdx = dirs.findIndex(s => /^PN[^/]+$/i.test(s)); if(pnIdx === -1 || pnIdx+1 >= dirs.length) return {pnIdx:-1, drwIdx:-1}; const drwIdx = /^DRW[^/]+$/i.test(dirs[pnIdx+1]) ? pnIdx+1 : -1; return {pnIdx, drwIdx}; }

function buildManifest(mode, root){
  const items = [];
  for(const e of inEntries){
    const p = normalizePath(e.path);
    const sp = splitPath(p); const dirs = sp.parts;
    const {pnIdx, drwIdx} = findPnDrwIndices(dirs);
    if(pnIdx === -1 || drwIdx === -1) continue;
    const afterDrw = dirs.slice(drwIdx+1);
    const ext = (sp.file.split('.').pop()||'').toLowerCase();
    if(mode === 'rfq' && ext !== 'pdf') continue;
    const dest = pathJoin(root, dirs[pnIdx], dirs[drwIdx], afterDrw.join('/'), sp.file);
    items.push({ src:e.path, dest, pn:dirs[pnIdx], drw:dirs[drwIdx], name:sp.file, ext });
  }
  const groups = []; const map = new Map();
  for(const it of items){ const key = `${it.pn}/${it.drw}`; if(!map.has(key)) map.set(key,{pn:it.pn,drw:it.drw,names:[]}); map.get(key).names.push(it.name); }
  for(const g of map.values()){ g.names.sort((a,b)=> a.localeCompare(b, undefined, {numeric:true, sensitivity:'base'})); groups.push(g); }
  groups.sort((a,b)=> (a.pn+a.drw).localeCompare(b.pn+b.drw));
  return { root, mode, items, groups };
}

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
    let out='';
    for(const k of keys){
      const v = node[k];
      out += `${prefix}${v===null?'├──':'└──'} ${k}\n`;
      if(v) out += render(v, prefix + (v===null? '': '    '));
    }
    return out;
  }
  return render(tree);
}

async function autoPreview(){
  if(!inZip) return;
  try{
    const mode = $('#mode').value;
    const root = ($('#rootName').value || '').trim() || autoRootForMode(mode);
    const manifest = buildManifest(mode, root);
    let tree = buildTreeFromItems(manifest.items) || '';
    tree = tree.replace(/\n+$/,''); 
    const totals = `Files: ${human(manifest.items.length)} · Drawings: ${human(manifest.groups.length)}`;
    $('#preview').textContent = (tree || '(no matches)') + '\n' + totals;
    $('#status').textContent = 'Preview updated.';
  }catch(e){
    $('#preview').textContent = '(preview error)'; $('#status').textContent = e.message || 'Preview error';
  }
}

async function buildDrawingZipFromManifest(manifest){
  const out = new JSZip(); let n = 0;
  for(const it of manifest.items){
    const data = await inZip.file(it.src).async('arraybuffer');
    out.file(it.dest, data, {date: new Date()}); n++;
  }
  const blob = await out.generateAsync({type:'blob', compression:'DEFLATE', compressionOptions:{level:6}});
  return { blob, count:n };
}

function listXlsxPaths(){
  const xs = [];
  for(const e of inEntries){
    const p = normalizePath(e.path);
    if(/\.xlsx$/i.test(p)) xs.push(p);
  }
  return xs;
}

function findWorkbookEntry(){
  const xs = listXlsxPaths();
  const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName) || '';
  if(id){
    const want1 = `${id}/${id}.xlsx`.toLowerCase();
    for(const e of inEntries){
      if(normalizePath(e.path).toLowerCase() === want1) return e;
    }
  }
  const pat = /^(?:[^/]*\/)?(RFQ[^/]*|PO[^/]*)\/\1\.xlsx$/i;
  for(const e of inEntries){
    const p = normalizePath(e.path);
    if(/\.xlsx$/i.test(p) && pat.test(p)) return e;
  }
  const cands = inEntries.filter(e => /\.xlsx$/i.test(e.path) && normalizePath(e.path).split('/').length >= 2);
  cands.sort((a,b)=>{
    const A = a.path.toLowerCase(), B = b.path.toLowerCase();
    const s = (x)=> (x.includes('/rfq')?-10:0) + (x.includes('/po')?-8:0) + (x.match(/\/rfq[^/]*\.xlsx$/)?-2:0) + (x.match(/\/po[^/]*\.xlsx$/)?-2:0);
    return (s(A)-s(B)) || A.length - B.length;
  });
  return cands[0] || null;
}

function a1ToColRow(a1){
  const m = /^([A-Z]+)(\d+)$/.exec(a1.toUpperCase());
  if(!m) return null;
  const letters = m[1], row = parseInt(m[2],10);
  let col = 0;
  for(let i=0;i<letters.length;i++){ col = col*26 + (letters.charCodeAt(i)-64); }
  return {col,row};
}

function withinC4toF4(a1){
  const cr = a1ToColRow(a1);
  if(!cr) return false;
  return cr.row === 4 && cr.col >= 3 && cr.col <= 6;
}

async function readC4F4_fromFallback(buf){
  const inner = await JSZip.loadAsync(buf);
  const getText = (p)=> inner.file(p) ? inner.file(p).async('string') : null;
  const sharedStringsXml = await getText("xl/sharedStrings.xml");
  let sst = [];
  if(sharedStringsXml){
    const doc = new DOMParser().parseFromString(sharedStringsXml, "application/xml");
    sst = Array.from(doc.getElementsByTagName("si")).map(si => {
      const tNodes = si.getElementsByTagName("t");
      let t = ""; for(let i=0;i<tNodes.length;i++){ t += tNodes[i].textContent; }
      return t.trim();
    });
  }
  const sheetPath = inner.file("xl/worksheets/sheet1.xml") ? "xl/worksheets/sheet1.xml" : null;
  if(!sheetPath) return '';
  const xml = await inner.file(sheetPath).async('string');
  const xdoc = new DOMParser().parseFromString(xml, "application/xml");
  const cells = Array.from(xdoc.getElementsByTagName("c"));
  const vals = [];
  for(const c of cells){
    const ref = c.getAttribute("r"); if(!ref) continue;
    if(!withinC4toF4(ref)) continue;
    const tAttr = c.getAttribute("t");
    const vNode = c.getElementsByTagName("v")[0];
    let val = "";
    if(tAttr === "s" && vNode){
      const idx = parseInt(vNode.textContent,10);
      if(!isNaN(idx) && sst[idx] != null) val = sst[idx];
    }else{
      const isNode = c.getElementsByTagName("is")[0];
      if(isNode){
        const tNodes = isNode.getElementsByTagName("t");
        for(let i=0;i<tNodes.length;i++){ val += tNodes[i].textContent; }
      }else if(vNode){
        val = vNode.textContent;
      }
    }
    if(val && val.trim()) vals.push(val.trim());
  }
  return vals.join(' ');
}

function readC4F4_fromSheetJS(buf){
  const wb = XLSX.read(buf, { type:'array' });
  const first = wb.SheetNames && wb.SheetNames.length ? wb.SheetNames[0] : null;
  if(!first) return '';
  const ws = wb.Sheets[first];
  const parts = [];
  for(const col of ['C','D','E','F']){
    const cell = ws[col+'4'];
    let text = '';
    if(cell && typeof cell.v !== 'undefined') text = String(cell.v);
    if(text && text.trim()) parts.push(text.trim());
  }
  return parts.join(' ');
}

async function extractTitleSmart(){
  const picked = findWorkbookEntry();
  const all = listXlsxPaths();
  if(!picked) return { title:'', picked:null, all, engine:'none' };
  const buf = await picked.entry.async('arraybuffer');
  if(typeof XLSX !== 'undefined'){
    try{
      const t = readC4F4_fromSheetJS(buf);
      if(t) return { title:t, picked, all, engine:'sheetjs-C4F4' };
    }catch(e){}
  }
  try{
    const t2 = await readC4F4_fromFallback(buf);
    return { title:t2, picked, all, engine:'fallback-C4F4' };
  }catch(e){
    return { title:'', picked, all, engine:'fallback-error' };
  }
}

function findTopLevelRFQPdf(){
  const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName) || '';
  if(id){
    const want = `${id}/${id}.pdf`.toLowerCase();
    for(const e of inEntries){
      const p = normalizePath(e.path).toLowerCase();
      if(p === want) return { path:e.path, entry:e.entry };
    }
  }
  const cands = [];
  const rx = /^([^/]+)\/\1\.pdf$/i;
  for(const e of inEntries){
    const p = normalizePath(e.path);
    if(!/\.pdf$/i.test(p)) continue;
    if(rx.test(p)) cands.push({ path:e.path, entry:e.entry });
  }
  if(cands.length){
    cands.sort((a,b)=> a.path.length - b.path.length);
    return cands[0];
  }
  return null;
}

function makeOverviewPdf(manifest){
  const { jsPDF } = window.jspdf || {}; if(!jsPDF){ alert('jsPDF not loaded.'); return null; }
  const doc = new jsPDF({unit:'mm', format:'a4'});
  const header = manifest.mode === 'rfq' ? 'RFQ Document list' : 'PO Document list';
  doc.setFont('helvetica','bold'); doc.setFontSize(14);
  doc.text(header, 20, 18);
  if(extractedTitle){ doc.setFont('helvetica','normal'); doc.setFontSize(11); doc.text(extractedTitle, 20, 25); }
  doc.setDrawColor(200); doc.line(20, 28, 190, 28);
  doc.setFont('helvetica','bold'); doc.setFontSize(10);
  const xPN=20,xDRW=70,xREF=130,yHead=36;
  doc.text('Part number', xPN, yHead);
  doc.text('Drawing folder', xDRW, yHead);
  doc.text('Drawing ref.', xREF, yHead);
  doc.setDrawColor(220); doc.line(20, yHead+2, 190, yHead+2);
  doc.setFont('helvetica','normal'); doc.setFontSize(10);
  let y=yHead+8, maxY=285;
  for(const g of manifest.groups){
    const refs = g.names.join(', ');
    const refLines = doc.splitTextToSize(refs, 60);
    const rowH = 5 * Math.max(1, refLines.length);
    if(y+rowH>maxY){
      doc.addPage();
      doc.setFont('helvetica','bold'); doc.setFontSize(14); doc.text(header, 20, 18);
      if(extractedTitle){ doc.setFont('helvetica','normal'); doc.setFontSize(11); doc.text(extractedTitle, 20, 25); }
      doc.setDrawColor(200); doc.line(20, 28, 190, 28);
      doc.setFont('helvetica','bold'); doc.setFontSize(10);
      doc.text('Part number', xPN, yHead); doc.text('Drawing folder', xDRW, yHead); doc.text('Drawing ref.', xREF, yHead);
      doc.setDrawColor(220); doc.line(20, yHead+2, 190, yHead+2);
      doc.setFont('helvetica','normal'); doc.setFontSize(10);
      y = yHead+8;
    }
    doc.text(g.pn||'', xPN, y);
    doc.text(g.drw||'', xDRW, y);
    doc.text(refLines, xREF, y);
    y += rowH;
  }
  return doc.output('blob');
}

function download(name, blob){
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob); a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 15000);
}

async function onDownload(){
  try{
    if(!window.JSZip){ alert('JSZip failed to load (CDN blocked?).'); return; }
    const mode = $('#mode').value;
    const root = ($('#rootName').value || '').trim() || autoRootForMode(mode);
    $('#status').textContent = 'Building…';
    const manifest = buildManifest(mode, root);
    if(manifest.items.length === 0){ $('#status').textContent = 'No files found in DRW folders.'; return; }

    const { blob: drawingBlob } = await buildDrawingZipFromManifest(manifest);
    const overviewBlob = makeOverviewPdf(manifest);
    if(!overviewBlob){ $('#status').textContent = 'jsPDF missing.'; return; }

    const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName) || '';
    const prefix = (mode==='rfq' ? (id?`RFQ${id}`:'RFQ') : (id?`PO${id}`:'PO'));
    const safeTitle = (extractedTitle||'').replace(/[\\/:*?"<>|]+/g, ' ').trim();
    const topName = safeTitle ? `${prefix} ${safeTitle}.zip` : `${prefix}.zip`;

    const top = new JSZip();
    top.file(`${root}.zip`, drawingBlob);
    top.file('File overview.pdf', overviewBlob);

    if(mode === 'rfq'){
      const rfq = findTopLevelRFQPdf();
      if(rfq){
        const rfqBlob = new Blob([await rfq.entry.async('arraybuffer')], {type:'application/pdf'});
        top.file(rfq.path.split('/').pop(), rfqBlob);
      }
    }

    const topBlob = await top.generateAsync({type:'blob', compression:'DEFLATE', compressionOptions:{level:6}});
    download(topName, topBlob);
    $('#status').textContent = `Downloaded: ${topName}`;
  }catch(e){
    $('#status').textContent = e.message || 'Build failed';
  }
}

async function loadZip(file){
  $('#btnGo').disabled = true; refreshEmailBtn(); refreshEmlBtn();
  try{
    if(!window.JSZip){ alert('JSZip failed to load (CDN blocked?).'); return; }
    $('#status').textContent = 'Reading…';
    inZip = await JSZip.loadAsync(file); inEntries = [];
    inZip.forEach((relPath, entry) => { if(!entry.dir) inEntries.push({ path: normalizePath(relPath), entry }); });
    detectedRoot = detectTopRoot(inEntries);
    lastFileName = file.name;
    $('#fileInfo').textContent = file.name;
    $('#rootName').value = autoRootForMode($('#mode').value);
    $('#status').textContent = `Ready (${human(inEntries.length)} files)`;

    extractedTitle = '';
    try{ 
      const res = await extractTitleSmart();
      extractedTitle = res.title || '';
    }catch(_){ extractedTitle=''; }
    $('#titleOut').textContent = extractedTitle || '(not detected)';

    $('#btnGo').disabled = false; refreshEmailBtn(); refreshEmlBtn();
    autoPreview();
  }catch(e){
    $('#status').textContent = e.message || 'Failed to read zip'; $('#btnGo').disabled = true; refreshEmailBtn(); refreshEmlBtn(); $('#preview').textContent = '(no preview)';
  }
}


// ---- RFQ Email Template ----
let rfqEmailTemplateCache = null;
async function loadRfqTemplate(){
  if(rfqEmailTemplateCache !== null) return rfqEmailTemplateCache;
  try{
    const res = await fetch('assets/rfq_email_template.txt', {cache:'no-store'});
    rfqEmailTemplateCache = await res.text();
  }catch(e){
    rfqEmailTemplateCache = 'Hello,\n\nPlease find the RFQ package attached.\n\nRegards,\n';
  }
  return rfqEmailTemplateCache;
}
function computeId(){
  const extract = (text)=>{ if(!text) return ''; const m = /(?:RFQ|PO)\s*([0-9]+(?:-\d+)*)/i.exec(text); return m ? m[1] : ''; };
  return extract(detectedRoot) || extract(lastFileName) || '';
}
async function onEmail(){
  const mode = document.querySelector('#mode').value;
  if(mode !== 'rfq') return; // safety
  const id = computeId();
  const title = (typeof extractedTitle === 'string' && extractedTitle.trim()) ? extractedTitle.trim() : '';
  const subject = `RFQ${id}${title ? ' - ' + title : ''}`.trim();
  const body = await loadRfqTemplate();
  const href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.location.href = href;
}
function refreshEmailBtn(){
  const btn = document.querySelector('#btnEmail');
  const ok = !!(inZip) && document.querySelector('#mode').value === 'rfq';
  btn.disabled = !ok;
}


// ===== EML utilities =====
function b64FromBlob(blob){
  return new Promise((resolve,reject)=>{
    const r = new FileReader();
    r.onload = ()=>{
      const s = String(r.result || '');
      const comma = s.indexOf(',');
      resolve(comma >= 0 ? s.slice(comma+1) : s);
    };
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}
function wrap76(s){
  const out = [];
  for(let i=0;i<s.length;i+=76) out.push(s.slice(i,i+76));
  return out.join('\r\n');
}
function rfc2047(subject){
  // minimal RFC2047 for UTF-8 text; fall back to raw if ASCII
  if (/^[\x00-\x7F]*$/.test(subject)) return subject;
  const enc = btoa(unescape(encodeURIComponent(subject)));
  return '=?UTF-8?B?' + enc + '?=';
}

async function buildRfqEml(manifest){
  const mode = manifest.mode;
  if(mode !== 'rfq') throw new Error('EML is RFQ-only');

  // Build artifacts
  const { blob: drawingBlob } = await buildDrawingZipFromManifest(manifest);
  const overviewBlob = makeOverviewPdf(manifest);
  const rfq = findTopLevelRFQPdf();
  const rfqBlob = rfq ? new Blob([await rfq.entry.async('arraybuffer')], {type:'application/pdf'}) : null;

  // Names & sizes
  const id = (extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName) || '').trim();
  const safeTitle = (extractedTitle||'').replace(/[\\/:*?"<>|]+/g, ' ').trim();
  const subjectRaw = `RFQ${id}${safeTitle? ' - '+safeTitle : ''}`.trim();
  const subject = rfc2047(subjectRaw);
  const drawingName = `${manifest.root}.zip`;
  const overviewName = 'File overview.pdf';
  const rfqName = rfq ? rfq.path.split('/').pop() : 'RFQ.pdf';

  const drawingSize = drawingBlob.size || '';
  const overviewSize = overviewBlob.size || '';
  const rfqSize = rfqBlob ? (rfqBlob.size||'') : '';

  // Body
  const bodyText = await loadRfqTemplate();

  // Base64 payloads
  const b64zip = wrap76(await b64FromBlob(drawingBlob));
  const b64ov  = wrap76(await b64FromBlob(overviewBlob));
  const b64rfq = rfqBlob ? wrap76(await b64FromBlob(rfqBlob)) : '';

  const boundary = '=_rfq_' + Math.random().toString(36).slice(2);
  const parts = [];
  parts.push(
`MIME-Version: 1.0
Date: ${new Date().toUTCString()}
Subject: ${subject}
X-Unsent: 1
Content-Type: multipart/mixed; boundary="${boundary}"

--${boundary}
Content-Type: text/plain; charset="UTF-8"
Content-Transfer-Encoding: 7bit

${bodyText}

--${boundary}
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="${drawingName}"${drawingSize? `; size=${drawingSize}`:''}

${b64zip}

--${boundary}
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="${overviewName}"${overviewSize? `; size=${overviewSize}`:''}

${b64ov}
`);
  if (rfqBlob){
    parts.push(
`--${boundary}
Content-Type: application/octet-stream
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="${rfqName}"${rfqSize? `; size=${rfqSize}`:''}

${b64rfq}
`);
  }
  parts.push(`--${boundary}--`);
  const eml = parts.join('\r\n');
  return new Blob([eml], {type:'message/rfc822'});
}
async function onDownloadEml(){
  try{
    const mode = $('#mode').value;
    if(mode !== 'rfq'){ return; }
    const root = ($('#rootName').value || '').trim() || autoRootForMode(mode);
    const manifest = buildManifest(mode, root);
    if(manifest.items.length === 0){ $('#status').textContent = 'No files found in DRW folders.'; return; }
    $('#status').textContent = 'Building .eml…';
    const emlBlob = await buildRfqEml(manifest);
    const id = extractIdCandidate(detectedRoot) || extractIdCandidate(lastFileName) || '';
    const safeTitle = (extractedTitle||'').replace(/[\\/:*?"<>|]+/g, ' ').trim();
    const name = `RFQ${id}${safeTitle? ' - '+safeTitle : ''}.eml`.trim();
    download(name, emlBlob);
    $('#status').textContent = `Downloaded: ${name}`;
  }catch(e){
    $('#status').textContent = e.message || 'EML build failed';
  }
}
function refreshEmlBtn(){
  const btn = document.querySelector('#btnEml');
  const ok = !!(inZip) && document.querySelector('#mode').value === 'rfq';
  btn.disabled = !ok;
}

(function init(){
  $('#mode').addEventListener('change', ()=>{ $('#rootName').value = autoRootForMode($('#mode').value); autoPreview(); refreshEmailBtn(); refreshEmlBtn(); });
  $('#rootName').addEventListener('input', ()=> autoPreview());

  const drop = $('#drop'); const input = $('#file');
  drop.addEventListener('dragover', e=>{ e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', ()=> drop.classList.remove('dragover'));
  drop.addEventListener('drop', e=>{ e.preventDefault(); drop.classList.remove('dragover'); const f = e.dataTransfer.files[0]; if(!f) return; if(!/\.zip$/i.test(f.name)){ $('#status').textContent = 'Please select a .zip export.'; return; } $('#status').textContent='Uploading…'; loadZip(f).catch(err=>{ $('#status').textContent = err.message; }); });
  drop.addEventListener('click', ()=> input.click());
  input.addEventListener('change', ()=>{ const f = input.files[0]; if(!f) return; if(!/\.zip$/i.test(f.name)){ $('#status').textContent = 'Please select a .zip export.'; return; } $('#status').textContent='Uploading…'; loadZip(f).catch(err=>{ $('#status').textContent = err.message; }); });

  $('#btnGo').addEventListener('click', ()=>{ if(inZip) onDownload().catch(err=>{ $('#status').textContent=err.message; }); });
  $('#btnEmail').addEventListener('click', ()=> onEmail());
  $('#btnEml').addEventListener('click', ()=> onDownloadEml());
})();