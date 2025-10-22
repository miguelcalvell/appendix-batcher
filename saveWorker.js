/* saveWorker.js: builds and saves a batch PDF off the main thread */
// Load pdf-lib inside the worker
importScripts('https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js');

const { PDFDocument, StandardFonts, rgb } = PDFLib;

self.onmessage = async (e) => {
  const { type, batchIdx, header, jobs, LETTER_W, LETTER_H, MARGIN, HEADER_BAND, FONT_SIZE } = e.data || {};
  if(type !== 'save'){ return; }
  try{
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);

    let totalPages = 0;

    for(const job of jobs){
      const name = job.name || 'Page';
      const label = header && header.trim().length>0 ? `${name.replace(/\.[^/.]+$/, '')} – ${header}` : name.replace(/\.[^/.]+$/, '');

      if(job.kind === 'pdf'){
        const u8 = new Uint8Array(job.bytes);
        let src;
        try{
          src = await PDFDocument.load(u8, { ignoreEncryption: true });
        } catch(e){
          // skip this file
          continue;
        }
        const count = src.getPageCount();
        for(let i=0;i<count;i++){
          let embedded;
          try{
            const arr = await doc.embedPdf(src, [i]);
            embedded = arr && arr[0];
          }catch(err){
            continue;
          }
          if(!embedded) continue;

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
          totalPages++;
        }
      } else if(job.kind === 'img'){
        const u8 = new Uint8Array(job.bytes);
        // Attempt PNG then JPG
        let img, isPng = false;
        try{ img = await doc.embedPng(u8); isPng = true; }catch{}
        if(!img){
          try{ img = await doc.embedJpg(u8); }catch{}
        }
        if(!img) continue;

        const dims = { w: img.width, h: img.height };
        const landscape = dims.w >= dims.h;
        const width = landscape ? LETTER_H : LETTER_W;
        const height = landscape ? LETTER_W : LETTER_H;
        const page = doc.addPage([width, height]);

        const right = width - MARGIN;
        const textWidth = font.widthOfTextAtSize(label, FONT_SIZE);
        page.drawText(label, { x: right - textWidth, y: height - MARGIN - FONT_SIZE - 4, size: FONT_SIZE, font, color: rgb(0,0,0) });

        const dstW = width - 2*MARGIN;
        const dstH = height - (MARGIN + HEADER_BAND) - MARGIN;
        const sW = img.width, sH = img.height;
        const scale = Math.min(dstW/sW, dstH/sH);
        const drawW = sW*scale, drawH = sH*scale;
        const cx = MARGIN + (dstW - drawW)/2;
        const cy = MARGIN + (dstH - drawH)/2;
        page.drawImage(img, { x: cx, y: cy, width: drawW, height: drawH });
        totalPages++;
      }
    }

    const pdfBytes = await doc.save({ updateFieldAppearances: false });
    self.postMessage({ ok:true, pdfBytes, pages: totalPages }, [pdfBytes.buffer]);
  }catch(err){
    self.postMessage({ ok:false, error: (err && err.message) || String(err) });
  }
};
