(function(){
  const VERSION = 'v6.9';
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  // ---------- State & DOM helpers ----------
  const $ = (id)=>document.getElementById(id);
  const state = { queue: [], outputs: [], queuedCount: 0, running:false };
  const logBuf = [];
  function log(msg){
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    logBuf.push(line);
    if(logBuf.length>1500) logBuf.splice(0, logBuf.length-1500);
    $('log').textContent = logBuf.join('\\n');
    $('log').scrollTop = $('log').scrollHeight;
  }
  function clearLog(){ logBuf.length=0; $('log').textContent=''; }
  function setProgress(pct, text){
    $('progressBar').style.width = `${Math.max(0,Math.min(100,pct))}%`;
    if(text) $('progressText').textContent = text;
  }
  function showSavingSpinner(on){
    const el = $('savingSpinner');
    if(on){ el.classList.remove('hidden'); }
    else { el.classList.add('hidden'); }
  }
  const raf = () => new Promise(res=>requestAnimationFrame(()=>res()));
  const delay = (ms) => new Promise(res=>setTimeout(res, ms));

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
  async function imageDims(file){
    const url = URL.createObjectURL(file);
    try{
      const img = new Image();
      const p = new Promise((res,rej)=>{ img.onload=()=>res({w:img.naturalWidth,h:img.naturalHeight}); img.onerror=rej; });
      img.src = url;
      return await p;
    } finally { URL.revokeObjectURL(url); }
  }
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

      // Pre-count planned pages
      let totalPlannedPages = 0;
      const prepared = [];
      for(const g of groups){
        for(const it of g.items){
          const u8 = await bytes(it.file);
          const pageCount = it.isPDF ? (await countPdfPagesFromBytes(u8) || 0) : 1;
          totalPlannedPages += pageCount;
          prepared.push({ groupOrder:g.appendix_order, label:g.label, file:it.file, name:it.name, isPDF:it.isPDF, bytes:u8, pageCount });
        }
      }
      if(totalPlannedPages===0){ log('ERROR: No renderable pages (bad PDFs or no images).'); return; }

      let batchIdx = 1;
      let doc = await PDFDocument.create();
      let font = await doc.embedFont(StandardFonts.Helvetica);
      let curPages = 0;
      let donePages = 0;

      async function finalize(){
        if(curPages===0) return;
        showSavingSpinner(true);
        // Let the spinner paint before heavy work
        await raf(); await delay(30);
        try{
          const pdfBytes = await doc.save({ updateFieldAppearances:false });
          const blob = new Blob([pdfBytes], {type:'application/pdf'});
          const name = `batch_${String(batchIdx).padStart(3,'0')}.pdf`;
          state.outputs.push({name, blob});
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = name;
          a.textContent = `⬇ ${name}`;
          $('links').appendChild(a);
          log(`Saved ${name} (${curPages} pages)`);
        } finally {
          showSavingSpinner(false);
        }
        batchIdx += 1;
        doc = await PDFDocument.create();
        font = await doc.embedFont(StandardFonts.Helvetica);
        curPages = 0;
      }

      // Build by appendix
      const byAppendix = new Map();
      for(const p of prepared){
        const key = p.groupOrder;
        if(!byAppendix.has(key)) byAppendix.set(key, []);
        byAppendix.get(key).push(p);
      }
      const appendixOrders = Array.from(byAppendix.keys()).sort((a,b)=>a-b);
      const targetPages = Math.max(1, parseInt(($('targetPages').value||'50'),10));

      for(const aord of appendixOrders){
        const arr = byAppendix.get(aord);
        const appendixPages = arr.reduce((s,x)=>s+x.pageCount,0);
        if(curPages>0 && (curPages + appendixPages) > targetPages){
          await finalize();
        }
        log(`Rendering Appendix ${arr[0].label} (~${appendixPages} pages)`);

        for(const it of arr){
          const label = headerLabelFor(it.name, header);
          log(`  • ${it.name}`);

          if(it.isPDF){
            try{
              const src = await PDFDocument.load(it.bytes, { ignoreEncryption:true });
              const total = src.getPageCount();
              log(`    PDF pages: ${total}`);
              for(let i=0;i<total;i++){
                let embedded;
                try{
                  const arr = await doc.embedPdf(src, [i]);
                  embedded = arr && arr[0];
                }catch(e){
                  log(`    ERROR: embed page ${i+1}/${total} failed (${e.message||e}).`);
                  continue;
                }
                if(!embedded){
                  log(`    ERROR: missing embedded page ${i+1}/${total}.`);
                  continue;
                }
                const page = doc.addPage([LETTER_W, LETTER_H]);
                const right = LETTER_W - MARGIN;
                const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
                page.drawText(label, { x: right - textWidth, y: LETTER_H - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

                const dstW = LETTER_W - 2*MARGIN;
                const dstH = LETTER_H - (MARGIN + HEADER_BAND) - MARGIN;
                const { width:sW, height:sH } = src.getPage(i).getSize();
                const scale = Math.min(dstW/sW, dstH/sH);
                const drawW = sW*scale, drawH = sH*scale;
                const cx = MARGIN + (dstW - drawW)/2;
                const cy = MARGIN + (dstH - drawH)/2;
                page.drawPage(embedded, { x: cx, y: cy, width: drawW, height: drawH });

                curPages += 1; donePages += 1;
                setProgress((donePages/totalPlannedPages)*100, `Rendering… ${donePages}/${totalPlannedPages} pages`);
              }
            }catch(e){
              log(`    ERROR: PDF load failed: ${it.name}. (${e.message||e})`);
            }
          } else {
            try{
              const dims = await imageDims(it.file);
              const landscape = dims.w >= dims.h;
              const width = landscape ? LETTER_H : LETTER_W;
              const height = landscape ? LETTER_W : LETTER_H;
              const page = doc.addPage([width, height]);
              const right = width - MARGIN;
              const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
              page.drawText(label, { x: right - textWidth, y: height - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

              let img;
              try{ img = await doc.embedPng(it.bytes); }catch{}
              if(!img){
                try{ img = await doc.embedJpg(it.bytes); }catch{}
              }
              if(!img){ log(`    ERROR: image embed failed: ${it.name}`); continue; }

              const dstW = width - 2*MARGIN;
              const dstH = height - (MARGIN + HEADER_BAND) - MARGIN;
              const sW = img.width, sH = img.height;
              const scale = Math.min(dstW/sW, dstH/sH);
              const drawW = sW*scale, drawH = sH*scale;
              const cx = MARGIN + (dstW - drawW)/2;
              const cy = MARGIN + (dstH - drawH)/2;
              page.drawImage(img, { x: cx, y: cy, width: drawW, height: drawH });

              curPages += 1; donePages += 1;
              setProgress((donePages/totalPlannedPages)*100, `Rendering… ${donePages}/${totalPlannedPages} pages`);
            }catch(e){
              log(`    ERROR: image failed: ${it.name}. (${e.message||e})`);
            }
          }
        }
      }

      await finalize();

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

  // init DZ label + version
  updateDropzoneBadge();
  const ver = document.getElementById('version'); if(ver){ ver.textContent = VERSION; }
})();