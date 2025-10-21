# Appendix Batcher (Web App, Mobile-Optimized)

**What:** A 100% in-browser tool to sort mixed PNG/PDF files named like `Appendix 1 (1 of 5)`,
stamp a header, and output chunked PDFs (~50 pages) **without splitting an Appendix**.

- Works fully offline (PWA). No uploads.
- Mobile-friendly UI (bigger controls). "Download All as ZIP" included.
- Host on GitHub Pages (free).

## Use
1. Open the site (or `index.html` locally).
2. Tap **Input files** and select all PNGs/PDFs (iOS allows multi-select).
3. Enter header text. Set target pages (default 50). Tap **Run**.
4. Download `batch_001.pdf`, etc., or **Download All as ZIP**. Also get `manifest.csv`.

## Tech
- `pdf-lib` for PDF composition, `JSZip` for bundling outputs.
- US Letter pages; PDFs forced **portrait**; PNGs **auto orientation** (landscape if width ≥ height).
- Header top-right, inside a safe header band (0.75"). No cropping; scale to fit.

## Deploy (GitHub Pages)
- Create repo, upload these files at the **root**.
- **Settings → Pages** → Branch: `main`, Folder: `/`.
- Wait ~1 min; site appears at `https://<you>.github.io/<repo>/`.
