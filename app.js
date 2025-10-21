(function(){
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  const $ = (id)=>document.getElementById(id);
  const log = (msg)=>{ const l=$('log'); l.textContent += msg + '\n'; l.scrollTop = l.scrollHeight; };
  const clearLog=()=>{ $('log').textContent=''; };

  // Page constants (US Letter)
  const LETTER_W = 612, LETTER_H = 792;
  const MARGIN = 36;             // 0.5"
  const HEADER_BAND = 54;        // 0.75"
  const FONT_SIZE = 10;

  // Regex: Appendix 1, Appendix 1 (1 of 5), Appendix 10 2
  const APP_RE = /^appendix[\s_\-]*([0-9]+)(?:[^\d]*(\d+))?(?:.*?\((\d+)\s*of\s*(\d+)\))?/i;

  function parseName(name) {
    const base = name.replace(/\.[^/.]+$/, '').trim();
    const m = base.match(APP_RE);
    if(!m) return null;
    const appendix_num = parseInt(m[1], 10);
    const bare = m[2] ? parseInt(m[2], 10) : null;
    const y = m[3] ? parseInt(m[3], 10) : null;
    const Y = m[4] ? parseInt(m[4], 10) : null;
    const part_y = (y != null) ? y : (bare != null ? bare : null);
    return { appendix_num, part_y, part_Y: Y };
  }

  function sortIntoGroups(files){
    const items = [];
    for(const f of files){
      const key = parseName(f.name);
      if(!key){ log(`WARN: skipping (bad name): ${f.name}`); continue; }
      items.push({file:f, name:f.name, key, isPDF:/\.pdf$/i.test(f.name)});
    }
    // primary: appendix number; secondary: part_y (nulls last); tertiary: filename
    items.sort((a,b)=>{
      if(a.key.appendix_num !== b.key.appendix_num) return a.key.appendix_num - b.key.appendix_num;
      const ay=a.key.part_y, by=b.key.part_y;
      if(ay==null && by!=null) return 1;
      if(ay!=null && by==null) return -1;
      if(ay!=null && by!=null && ay!==by) return ay-by;
      return a.name.localeCompare(b.name);
    });
    // group by appendix
    const groups = [];
    let cur = null;
    for(const it of items){
      if(!cur || cur.appendix_num !== it.key.appendix_num){
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
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  function fitWithin(sw, sh, dw, dh){
    const s = Math.min(dw/sw, dh/sh);
    return { w: sw*s, h: sh*s };
  }

  async function countPdfPages(u8){
    const src = await PDFDocument.load(u8);
    return src.getPageCount();
  }

  async function run(){
    clearLog();
    const files = Array.from($('files').files||[]);
    if(files.length===0){ alert('Select files first'); return; }
    const header = $('headerText').value || '';
    const targetPages = Math.max(1, parseInt($('targetPages').value||'50',10));

    const groups = sortIntoGroups(files);
    log(`Found ${groups.length} appendices`);

    // Outputs to collect for "Download All as ZIP"
    const outputs = [];

    // Current batch
    let batchIdx = 1;
    let doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    let curPages = 0;
    let batchStartPage = 1;

    async function finalize(){
      if(curPages===0) return;
      const pdfBytes = await doc.save({ updateFieldAppearances:false });
      const blob = new Blob([pdfBytes], {type:'application/pdf'});
      const name = `batch_${String(batchIdx).padStart(3,'0')}.pdf`;
      outputs.push({name, blob});

      // Link
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.textContent = `⬇ ${name}`;
      $('links').appendChild(a);
      $('downloads').classList.remove('hidden');

      log(`Saved ${name} (${curPages} pages)`);
      batchIdx += 1;
      doc = await PDFDocument.create();
      curPages = 0;
      batchStartPage = 1;
    }

    // Manifest
    const manifest = [["input_name","appendix_num","part_y","is_pdf","pages_in_item","batch","batch_page_start","batch_page_end"]];

    for(const group of groups){
      // Dry-run count of pages for split decision
      let appendixPages = 0;
      for(const it of group.items){
        if(it.isPDF){
          appendixPages += await countPdfPages(await bytes(it.file));
        } else appendixPages += 1;
      }
      if(curPages>0 && (curPages + appendixPages) > targetPages){
        await finalize();
      }

      log(`Rendering Appendix ${group.appendix_num} (~${appendixPages} pages)`);

      for(const it of group.items){
        if(it.isPDF){
          const u8 = await bytes(it.file);
          const src = await PDFDocument.load(u8);
          const total = src.getPageCount();
          const embeddedPages = await doc.embedPdf(u8);
          for(let i=0;i<total;i++){
            const page = doc.addPage([LETTER_W, LETTER_H]);
            // Header text top-right
            const right = LETTER_W - MARGIN;
            const textWidth = font.widthOfTextAtSize(header, FONT_SIZE);
            page.drawText(header, { x: right - textWidth, y: LETTER_H - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

            // Content area
            const dstW = LETTER_W - 2*MARGIN;
            const dstH = LETTER_H - (MARGIN + HEADER_BAND) - MARGIN;
            const srcPage = src.getPage(i);
            const { width: sW, height: sH } = srcPage.getSize();
            const scale = Math.min(dstW/sW, dstH/sH);
            const drawW = sW * scale, drawH = sH * scale;
            const cx = MARGIN + (dstW - drawW)/2;
            const cy = MARGIN + (dstH - drawH)/2;

            const ep = embeddedPages[i];
            page.drawPage(ep, { x: cx, y: cy, width: drawW, height: drawH });
            curPages += 1;
          }
          manifest.push([it.name, group.appendix_num, it.key.part_y ?? "", "true", total, batchIdx, batchStartPage, batchStartPage+total-1]);
          batchStartPage += total;
        } else {
          const u8 = await bytes(it.file);
          const dims = await imageDims(it.file);
          const landscape = dims.w >= dims.h;
          const width = landscape ? LETTER_H : LETTER_W;
          const height = landscape ? LETTER_W : LETTER_H;
          const page = doc.addPage([width, height]);

          const right = width - MARGIN;
          const textWidth = font.widthOfTextAtSize(header, FONT_SIZE);
          page.drawText(header, { x: right - textWidth, y: height - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

          let img;
          if (/\.png$/i.test(it.name)) img = await doc.embedPng(u8);
          else img = await doc.embedJpg(u8);

          const dstW = width - 2*MARGIN;
          const dstH = height - (MARGIN + HEADER_BAND) - MARGIN;
          const sW = img.width, sH = img.height;
          const scale = Math.min(dstW/sW, dstH/sH);
          const drawW = sW * scale, drawH = sH * scale;
          const cx = MARGIN + (dstW - drawW)/2;
          const cy = MARGIN + (dstH - drawH)/2;
          page.drawImage(img, { x: cx, y: cy, width: drawW, height: drawH });
          curPages += 1;
          manifest.push([it.name, group.appendix_num, it.key.part_y ?? "", "false", 1, batchIdx, batchStartPage, batchStartPage]);
          batchStartPage += 1;
        }
      }
    }

    await finalize();

    // manifest.csv
    const csv = manifest.map(r => r.map(x => `"${String(x).replace(/"/g,'""')}"`).join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'manifest.csv';
    a.textContent = '⬇ manifest.csv';
    $('links').appendChild(a);

    // Enable "Download All as ZIP"
    $('zipBtn').disabled = outputs.length === 0;
    $('zipBtn').onclick = async ()=>{
      const zip = new JSZip();
      outputs.forEach(o => zip.file(o.name, o.blob));
      zip.file('manifest.csv', blob);
      const z = await zip.generateAsync({type:'blob'});
      const link = document.createElement('a');
      link.href = URL.createObjectURL(z);
      link.download = 'appendix_batches.zip';
      link.click();
    };

    log('Done.');
  }

  $('runBtn').addEventListener('click', ()=>{
    $('links').innerHTML=''; $('downloads').classList.add('hidden'); run().catch(e=>{ console.error(e); log('ERROR: '+e.message); });
  });
  $('clearBtn').addEventListener('click', ()=>{
    $('files').value=''; $('headerText').value=''; $('targetPages').value=50; $('links').innerHTML=''; $('downloads').classList.add('hidden'); clearLog();
  });
})();