(function(){
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  // ---------- DOM helpers ----------
  const $ = (id)=>document.getElementById(id);
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

  // ---------- Constants ----------
  const LETTER_W = 612, LETTER_H = 792;
  const MARGIN = 36;          // 0.5"
  const HEADER_BAND = 54;     // 0.75" safe zone for header
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
      if(ay!=null && ay!==by) return ay-by;
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

  function fitWithin(sw, sh, dw, dh){
    const s = Math.min(dw/sw, dh/sh);
    return { w: sw*s, h: sh*s };
  }

  async function countPdfPages(u8){
    try{
      const src = await PDFDocument.load(u8, { ignoreEncryption: true });
      return src.getPageCount();
    }catch(e){ return 0; }
  }

  function headerLabelFor(nameWithExt, header){
    const base = nameWithExt.replace(/\.[^/.]+$/, '');
    return (header && header.trim().length>0) ? `${base} – ${header}` : base;
  }

  // ---------- Drag & drop ----------
  const dz = $('dropzone');
  const filesInput = $('files');
  function addFiles(fileList){
    const dt = new DataTransfer();
    for(const f of filesInput.files) dt.items.add(f);
    for(const f of fileList) dt.items.add(f);
    filesInput.files = dt.files;
    log(`Queued ${fileList.length} file(s). Total: ${filesInput.files.length}`);
  }
  ['dragenter','dragover'].forEach(ev=>{
    window.addEventListener(ev,(e)=>{ e.preventDefault(); dz.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(ev=>{
    window.addEventListener(ev,(e)=>{ if(ev!=='drop') dz.classList.remove('dragover'); });
  });
  dz.addEventListener('drop',(e)=>{
    e.preventDefault(); dz.classList.remove('dragover');
    const files = Array.from(e.dataTransfer.files||[]);
    if(files.length) addFiles(files);
  });
  $('pickBtn').addEventListener('click', ()=> filesInput.click());

  // ---------- Main run ----------
  async function run(){
    if(!$('keepHistory').checked){ $('links').innerHTML=''; }
    clearLog(); setProgress(0,'Scanning files…');

    const header = $('headerText').value || '';
    const allFiles = Array.from(filesInput.files||[]);
    if(allFiles.length===0){ alert('Select or drop files first'); return; }

    const groups = sortIntoGroups(allFiles);
    log(`Found ${groups.length} appendices`);

    let totalPlannedPages = 0;
    for(const g of groups){
      for(const it of g.items){
        if(it.isPDF) totalPlannedPages += (await countPdfPages(await bytes(it.file))) || 0;
        else totalPlannedPages += 1;
      }
    }
    if(totalPlannedPages===0){ log('ERROR: No renderable pages (bad PDFs or no images).'); return; }

    const outputs = [];
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
      outputs.push({name, blob});
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
      let appendixPages = 0;
      for(const it of group.items){
        if(it.isPDF) appendixPages += (await countPdfPages(await bytes(it.file))) || 0;
        else appendixPages += 1;
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
            let embeddedPages;
            try{
              embeddedPages = await doc.embedPdf(u8);
            }catch(e){
              log(`ERROR: Embed PDF failed: ${it.name}. Skipping. (${e.message||e})`);
              continue;
            }
            for(let i=0;i<total;i++){
              const page = doc.addPage([LETTER_W, LETTER_H]);
              // header in safe top-right
              const right = LETTER_W - MARGIN;
              const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
              page.drawText(label, { x: right - textWidth, y: LETTER_H - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });
              // content
              const dstW = LETTER_W - 2*MARGIN;
              const dstH = LETTER_H - (MARGIN + HEADER_BAND) - MARGIN;
              const srcPage = src.getPage(i);
              const { width:sW, height:sH } = srcPage.getSize();
              const scale = Math.min(dstW/sW, dstH/sH);
              const drawW = sW*scale, drawH = sH*scale;
              const cx = MARGIN + (dstW - drawW)/2;
              const cy = MARGIN + (dstH - drawH)/2;
              const ep = embeddedPages[i];
              page.drawPage(ep, { x: cx, y: cy, width: drawW, height: drawH });

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
            // header in safe top-right
            const right = width - MARGIN;
            const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
            page.drawText(label, { x: right - textWidth, y: height - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });
            // image
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
    const csvRows = [["input_name","appendix_num","part_y","is_pdf","pages_in_item","batch","batch_page_start","batch_page_end"]];
    for(const r of csvRows.slice(1)){} // placeholder to keep style consistent
    const csvBody = [["input_name","appendix_num","part_y","is_pdf","pages_in_item","batch","batch_page_start","batch_page_end"]];
    // We already built manifest with header; re-use it directly:
    const csv = manifest.map(r => r.map(x => `"${String(x).replace(/"/g,'""')}"`).join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'manifest.csv';
    a.textContent = '⬇ manifest.csv';
    $('links').appendChild(a);

    $('zipBtn').disabled = false;
    $('zipBtn').onclick = async ()=>{
      const zip = new JSZip();
      const linkEls = Array.from($('links').querySelectorAll('a')).filter(a=>/batch_\d+\.pdf$/.test(a.download));
      for(const l of linkEls){
        const resp = await fetch(l.href);
        const b = await resp.blob();
        zip.file(l.download, b);
      }
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
    if(!$('keepHistory').checked){ $('links').innerHTML=''; }
    $('progressBar').style.width='0%';
    $('progressText').textContent='Starting…';
    run().catch(e=>{ console.error(e); log('FATAL: '+(e.message||e)); setProgress(0,'Error'); });
  });
  $('clearBtn').addEventListener('click', ()=>{
    $('files').value=''; $('headerText').value=''; $('targetPages').value=50; $('links').innerHTML='';
    clearLog(); setProgress(0,'Idle');
  });
  $('zipBtn').disabled = true;
})();