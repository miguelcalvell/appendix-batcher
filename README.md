# Appendix Batcher — v6

**Fixes & Improvements**
- Prints **`<filename without extension> – <your header text>`** in the **safe top-right band** on **every page**.
- Drag-and-drop works across desktop & iOS; no `DataTransfer()` dependency.
- Compact UI (sticky app bar, capped log, progress bar).
- ZIP built from in-memory blobs (no `fetch(blob:)` issues).
- Robust PDF handling: `ignoreEncryption:true`, graceful skip on errors.
- Service worker cache **v6** to force clients to pull the update.

## Deploy
Replace files in your repo root with these versions. Then hard-refresh the site.
