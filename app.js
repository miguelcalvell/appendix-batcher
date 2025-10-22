(function(){
  const VERSION = 'v6.7';
  const { PDFDocument } = PDFLib; // only for pre-counting; worker uses full API

  // ---------- State & DOM helpers ----------
  const $ = (id)=>document.getElementById(id);
  const state = { queue: [], outputs: [], queuedCount: 0, running:false };
  const logBuf = [];
  function log(msg){
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    logBuf.push(line);
    if(logBuf.length>500) logBuf.splice(0, logBuf.length-500);
    $('log').textContent = logBuf.join('\n');
    $('log').scrollTop = $('log').scrollHeight;
  }
  function clearLog(){ logBuf.length=0; $('log').textContent=''; }
  function setProgress(pct, text){
    $('progressBar').style.width = `${Math.max(0,Math.min(100,pct))}%`;
    if(text) $('progressText').textContent = text;
  }
  function showSaving(on, label){
    const wrap = $('savingWrap');
    const lbl = $('savingLabel');
    if(on){
      if(label) lbl.textContent = label;
      wrap.classList.remove('hidden');
      wrap.offsetWidth; // force reflow
    } else {
      wrap.classList.add('hidden');
    }
  }

  function updateDropzoneBadge(){
    const badge = document.querySelector('#dropzone .dz-title');
    if(!badge) return;
    badge.textContent = state.queuedCount>0 ? `Queued: ${state.queuedCount} file(s)` : 'Drag & drop files here';
  }
  function addToQueue(files){
    let added = 0;
    for(const f of files){
      state.queue.push(f);
      added++;
      state.queuedCount++;
    }
    log(`Queued ${added} file(s). Total: ${state.queuedCount}`);
    updateDropzoneBadge();
  }

  // ---------- Constants ----------
  const LETTER_W = 612, LETTER_H = 792;
  const MARGIN = 36;          // 0.5"
  const HEADER_BAND = 54;     // 0.75"
  const FONT_SIZE = 10;
  const APP_RE = /^appendix[\s_\-]*([A-Z]+|\d+)(?:[^\d(]*?(\d+))?(?:.*?\((\d+)\s*of\s*(\d+)\))?/i;

  // ---------- Utils ----------
  function lettersToOrder(s){
    s = s.toUpperCase();
    let n = 0;
    for(let i=0;i<s.length;i++){
      const c = s.charCodeAt(i);
      if(c<65 || c>90) return null;
      n = n*26 + (c-64);
    }
    return n;
  }

  function parseName(name){
    const base = name.replace(/\.[^/.]+$/, '').trim();
    const m = base.match(APP_RE);
    if(!m) return null;
    const label = m[1];
    const isAlpha = /[A-Za-z]/.test(label);
    const appendix_order = isAlpha ? lettersToOrder(label) : parseInt(label,10);
    if(appendix_order==null || Number.isNaN(appendix_order)) return null;

    const bareAfter = m[2] ? parseInt(m[2],10) : null;
    const y = m[3] ? parseInt(m[3],10) : null;
    const Y = m[4] ? parseInt(m[4],10) : null;
    const part_y = (y!=null) ? y : (bareAfter!=null ? bareAfter : null);

    return { appendix_label: label.toString(), appendix_order, part_y, part_Y: Y };
  }

  function sortIntoGroups(files){
    const items = [];
    for(const f of files){
      const key = parseName(f.name);
      if(!key){ log(`WARN: skipping (bad name): ${f.name}`); continue; }
      items.push({file:f, name:f.name, key, isPDF:/\.pdf$/i.test(f.name)});
    }
    items.sort((a,b)=>{
      if(a.key.appendix_order!==b.key.appendix_order) return a.key.appendix_order - b.key.appendix_order;
      const ay=a.key.part_y, by=b.key.part_y;
      if(ay==null && by!=null) return 1;
      if(ay!=null && by==null) return -1;
      if(ay!=null && by!=null && ay!==by) return ay-by;
      return a.name.localeCompare(b.name);
    });
    const groups=[]; let cur=null;
    for(const it of items){
      if(!cur || cur.appendix_order!==it.key.appendix_order){
        cur = { appendix_order: it.key.appendix_order, label: it.key.appendix_label, items: [] };
        groups.push(cur);
      }
      cur.items.push(it);
    }
    return groups;
  }

  async function bytes(file){ return new Uint8Array(await file.arrayBuffer()); }
  async function countPdfPagesFromBytes(u8){
    try{ const src = await PDFDocument.load(u8, { ignoreEncryption: true }); return src.getPageCount(); }
    catch(e){ return 0; }
  }
  function headerLabelFor(nameWithExt, header){
    const base = nameWithExt.replace(/\.[^/.]+$/, '');
    return (header && header.trim().length>0) ? `${base} – ${header}` : base;
  }

  // ---------- Drag & drop ----------
  const dz = $('dropzone');
  const filesInput = $('files');
  ['dragenter','dragover'].forEach(ev=>{
    window.addEventListener(ev,(e)=>{ e.preventDefault(); dz.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(ev=>{
    window.addEventListener(ev,(e)=>{ if(ev!=='drop') dz.classList.remove('dragover'); });
  });
  dz.addEventListener('drop',(e)=>{
    e.preventDefault(); dz.classList.remove('dragover');
    const files = Array.from(e.dataTransfer?.files || []);
    if(files.length) addToQueue(files);
  });
  $('pickBtn').addEventListener('click', ()=> filesInput.click());
  filesInput.addEventListener('change', ()=>{
    const files = Array.from(filesInput.files || []);
    if(files.length) addToQueue(files);
    filesInput.value = '';
  });

  // ---------- Worker ----------
  const worker = new Worker('saveWorker.js');

  function saveBatchWithWorker(batchIdx, jobs, header){
    return new Promise((resolve, reject)=>{
      const transfers = [];
      // Collect transferable ArrayBuffers
      for(const j of jobs){
        if(j.bytes && j.bytes.buffer) transfers.push(j.bytes.buffer);
      }
      worker.onmessage = (e)=>{
        const { ok, pdfBytes, error, pages } = e.data || {};
        if(ok){
          resolve({ pdfBytes: new Uint8Array(pdfBytes), pages });
        } else {
          reject(new Error(error || 'Worker failed'));
        }
      };
      worker.onerror = (err)=> reject(err);
      worker.postMessage({
        type: 'save',
        batchIdx,
        header,
        jobs,
        LETTER_W, LETTER_H, MARGIN, HEADER_BAND, FONT_SIZE
      }, transfers);
    });
  }

  // ---------- Main run ----------
  async function run(){
    if(state.running) return;
    state.running = true;
    $('runBtn').disabled = true;
    try{
      if(!$('keepHistory').checked){ $('links').innerHTML=''; state.outputs.length=0; }
      clearLog(); setProgress(0,'Scanning files…');

      const header = $('headerText').value || '';
      const allFiles = state.queue.slice();
      if(allFiles.length===0){ alert('Select or drop files first'); return; }

      const groups = sortIntoGroups(allFiles);
      log(`Found ${groups.length} appendices`);

      // Pre-count planned pages and pre-read bytes, build job items container per input
      let totalPlannedPages = 0;
      const prepared = []; // [{name,isPDF,bytes,pageCount}]
      for(const g of groups){
        for(const it of g.items){
          const u8 = await bytes(it.file);
          const pageCount = it.isPDF ? (await countPdfPagesFromBytes(u8) || 0) : 1;
          totalPlannedPages += pageCount;
          prepared.push({ groupOrder:g.appendix_order, label:g.label, name:it.name, isPDF:it.isPDF, bytes:u8, pageCount });
        }
      }
      if(totalPlannedPages===0){ log('ERROR: No renderable pages (bad PDFs or no images).'); return; }

      // Build batches without splitting an appendix
      const targetPages = Math.max(1, parseInt(($('targetPages').value||'50'),10));
      let batches = [];
      let cur = { jobs: [], pages: 0, startLabel: null };
      let lastAppendixOrder = null;

      // Reconstruct the prepared list grouped by appendix order
      const byAppendix = new Map();
      for(const p of prepared){
        const key = p.groupOrder;
        if(!byAppendix.has(key)) byAppendix.set(key, []);
        byAppendix.get(key).push(p);
      }
      const appendixOrders = Array.from(byAppendix.keys()).sort((a,b)=>a-b);

      for(const aord of appendixOrders){
        const arr = byAppendix.get(aord);
        const appendixPages = arr.reduce((s,x)=>s+x.pageCount,0);
        if(cur.pages>0 && (cur.pages + appendixPages) > targetPages){
          batches.push(cur);
          cur = { jobs: [], pages: 0, startLabel: null };
        }
        if(!cur.startLabel) cur.startLabel = arr[0].label;
        for(const p of arr){
          cur.jobs.push({ kind: p.isPDF ? 'pdf' : 'img', name: p.name, bytes: p.bytes });
          cur.pages += p.pageCount;
        }
        lastAppendixOrder = aord;
      }
      if(cur.pages>0) batches.push(cur);

      let donePages = 0;
      for(let i=0;i<batches.length;i++){
        const b = batches[i];
        log(`Rendering batch ${String(i+1).padStart(3,'0')} from Appendix ${b.startLabel} (~${b.pages} pages)`);
        setProgress((donePages/totalPlannedPages)*100, `Rendering… ${donePages}/${totalPlannedPages} pages`);

        // Offload to worker
        showSaving(true, `Saving batch ${String(i+1).padStart(3,'0')}…`);
        const { pdfBytes, pages } = await saveBatchWithWorker(i+1, b.jobs, header).finally(()=> showSaving(false));

        // Link
        const blob = new Blob([pdfBytes], {type:'application/pdf'});
        const name = `batch_${String(i+1).padStart(3,'0')}.pdf`;
        state.outputs.push({ name, blob });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = name;
        a.textContent = `⬇ ${name}`;
        $('links').appendChild(a);
        log(`Saved ${name} (${pages} pages)`);

        donePages += pages;
        setProgress((donePages/totalPlannedPages)*100, `Rendering… ${donePages}/${totalPlannedPages} pages`);
      }

      $('zipBtn').disabled = state.outputs.length === 0;
      $('zipBtn').onclick = async ()=>{
        const zip = new JSZip();
        for(const o of state.outputs) zip.file(o.name, o.blob);
        const z = await zip.generateAsync({type:'blob'});
        const link = document.createElement('a');
        link.href = URL.createObjectURL(z);
        link.download = 'appendix_batches.zip';
        link.click();
      };

      setProgress(100, 'Done.');
      log('Done.');
    } finally {
      state.running = false;
      $('runBtn').disabled = false;
      const ver = document.getElementById('version'); if(ver){ ver.textContent = VERSION; }
    }
  }

  // ---------- Buttons ----------
  $('runBtn').addEventListener('click', ()=>{
    if(!$('keepHistory').checked){ $('links').innerHTML=''; state.outputs.length=0; }
    $('progressBar').style.width='0%';
    $('progressText').textContent='Starting…';
    run().catch(e=>{ console.error(e); log('FATAL: '+(e.message||e)); setProgress(0,'Error'); });
  });
  $('clearBtn').addEventListener('click', ()=>{
    state.queue.length=0; state.queuedCount=0; updateDropzoneBadge();
    $('headerText').value=''; $('targetPages').value=50; $('links').innerHTML='';
    clearLog(); setProgress(0,'Idle');
  });
  $('zipBtn').disabled = true;

  // init DZ label
  updateDropzoneBadge();
})();