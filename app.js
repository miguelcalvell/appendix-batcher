(function(){
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  // ---------- State & DOM helpers ----------
  const $ = (id)=>document.getElementById(id);
  const state = { queue: [], outputs: [], queuedCount: 0 };
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
  const APP_RE = /^appendix[\s_\-]*([0-9]+)(?:[^\d]*(\d+))?(?:.*?\((\d+)\s*of\s*(\d+)\))?/i;

  // ---------- Utils ----------
  function parseName(name){
    const base = name.replace(/\.[^/.]+$/, '').trim();
    const m = base.match(APP_RE);
    if(!m) return null;
    const appendix_num = parseInt(m[1],10);
    const bare = m[2] ? parseInt(m[2],10) : null;
    const y = m[3] ? parseInt(m[3],10) : null;
    const Y = m[4] ? parseInt(m[4],10) : null;
    const part_y = (y!=null) ? y : (bare!=null ? bare : null);
    return { appendix_num, part_y, part_Y: Y };
  }

  function sortIntoGroups(files){
    const items = [];
    for(const f of files){
      const key = parseName(f.name);
      if(!key){ log(`WARN: skipping (bad name): ${f.name}`); continue; }
      items.push({file:f, name:f.name, key, isPDF:/\.pdf$/i.test(f.name)});
    }
    items.sort((a,b)=>{
      if(a.key.appendix_num!==b.key.appendix_num) return a.key.appendix_num - b.key.appendix_num;
      const ay=a.key.part_y, by=b.key.part_y;
      if(ay==null && by!=null) return 1;
      if(ay!=null && by==null) return -1;
      if(ay!=null && by!=null && ay!==by) return ay-by;
      return a.name.localeCompare(b.name);
    });
    const groups=[]; let cur=null;
    for(const it of items){
      if(!cur || cur.appendix_num!==it.key.appendix_num){
        cur = { appendix_num: it.key.appendix_num, items: [] };
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
    if(!$('keepHistory').checked){ $('links').innerHTML=''; state.outputs.length=0; }
    clearLog(); setProgress(0,'Scanning files…');

    const header = $('headerText').value || '';
    const allFiles = state.queue.slice();
    if(allFiles.length===0){ alert('Select or drop files first'); return; }

    const groups = sortIntoGroups(allFiles);
    log(`Found ${groups.length} appendices`);

    // Pre-count planned pages
    let totalPlannedPages = 0;
    for(const g of groups){
      for(const it of g.items){
        if(it.isPDF) totalPlannedPages += (await countPdfPagesFromBytes(await bytes(it.file))) || 0;
        else totalPlannedPages += 1;
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
      batchIdx += 1;
      doc = await PDFDocument.create();
      font = await doc.embedFont(StandardFonts.Helvetica);
      curPages = 0;
    }

    const manifest = [["input_name","appendix_num","part_y","is_pdf","pages_in_item","batch","batch_page_start","batch_page_end"]];

    for(const group of groups){
      // Count this appendix
      let appendixPages = 0;
      const appendixPageCounts = [];
      for(const it of group.items){
        if(it.isPDF){
          const pc = (await countPdfPagesFromBytes(await bytes(it.file))) || 0;
          appendixPages += pc; appendixPageCounts.push(pc);
        } else { appendixPages += 1; appendixPageCounts.push(1); }
      }
      const targetPages = Math.max(1, parseInt(($('targetPages').value||'50'),10));
      if(curPages>0 && (curPages + appendixPages) > targetPages){
        await finalize();
      }

      log(`Rendering Appendix ${group.appendix_num} (~${appendixPages} pages)`);
      let batchStartPage = curPages + 1;

      for(const it of group.items){
        const label = headerLabelFor(it.name, header);

        if(it.isPDF){
          try{
            const u8 = await bytes(it.file);
            const src = await PDFDocument.load(u8, { ignoreEncryption:true });
            const total = src.getPageCount();

            for(let i=0;i<total;i++){
              let embedded;
              try {
                // Use loaded source doc rather than raw bytes (more reliable)
                const arr = await doc.embedPdf(src, [i]);
                embedded = arr && arr[0];
              } catch (e) {
                log(`ERROR: Embed page ${i+1}/${total} failed: ${it.name}. (${e.message||e})`);
                continue;
              }
              if(!embedded){
                log(`ERROR: Missing embedded page ${i+1}/${total}: ${it.name}.`);
                continue;
              }

              // Build a new Letter page and draw header + scaled content box
              const page = doc.addPage([LETTER_W, LETTER_H]);
              const right = LETTER_W - MARGIN;
              const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
              page.drawText(label, { x: right - textWidth, y: LETTER_H - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

              const dstW = LETTER_W - 2*MARGIN;
              const dstH = LETTER_H - (MARGIN + HEADER_BAND) - MARGIN;

              // Source page size comes from src.getPage(i)
              const { width:sW, height:sH } = src.getPage(i).getSize();
              const scale = Math.min(dstW/sW, dstH/sH);
              const drawW = sW*scale, drawH = sH*scale;
              const cx = MARGIN + (dstW - drawW)/2;
              const cy = MARGIN + (dstH - drawH)/2;

              page.drawPage(embedded, { x: cx, y: cy, width: drawW, height: drawH });

              curPages += 1; donePages += 1;
              setProgress((donePages/totalPlannedPages)*100, `Rendering… ${donePages}/${totalPlannedPages} pages`);
            }

            manifest.push([it.name, group.appendix_num, it.key.part_y ?? "", "true", total, batchIdx, batchStartPage, batchStartPage+total-1]);
            batchStartPage += total;
          }catch(e){
            log(`ERROR: PDF load failed: ${it.name}. Skipped. (${e.message||e})`);
          }
        } else {
          try{
            const u8 = await bytes(it.file);
            const dims = await imageDims(it.file);
            const landscape = dims.w >= dims.h;
            const width = landscape ? LETTER_H : LETTER_W;
            const height = landscape ? LETTER_W : LETTER_H;
            const page = doc.addPage([width, height]);

            const right = width - MARGIN;
            const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
            page.drawText(label, { x: right - textWidth, y: height - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

            let img;
            if(/\.png$/i.test(it.name)) img = await doc.embedPng(u8); else img = await doc.embedJpg(u8);
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
            manifest.push([it.name, group.appendix_num, it.key.part_y ?? "", "false", 1, batchIdx, batchStartPage, batchStartPage]);
            batchStartPage += 1;
          }catch(e){
            log(`ERROR: Image failed: ${it.name}. Skipped. (${e.message||e})`);
          }
        }
      }
    }

    await finalize();

    // manifest.csv
    const csv = [["input_name","appendix_num","part_y","is_pdf","pages_in_item","batch","batch_page_start","batch_page_end"]]
      .concat(manifest.slice(1))
      .map(r => r.map(x => `"${String(x).replace(/"/g,'""')}"`).join(','))
      .join('\\n');
    const blob = new Blob([csv], {type:'text/csv'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'manifest.csv';
    a.textContent = '⬇ manifest.csv';
    $('links').appendChild(a);

    // ZIP from in-memory blobs
    $('zipBtn').disabled = state.outputs.length === 0;
    $('zipBtn').onclick = async ()=>{
      const zip = new JSZip();
      for(const o of state.outputs) zip.file(o.name, o.blob);
      zip.file('manifest.csv', blob);
      const z = await zip.generateAsync({type:'blob'});
      const link = document.createElement('a');
      link.href = URL.createObjectURL(z);
      link.download = 'appendix_batches.zip';
      link.click();
    };

    setProgress(100, 'Done.');
    log('Done.');
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