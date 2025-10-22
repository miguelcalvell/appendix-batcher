# Appendix Batcher — Fixed Header + UI

- Prints **`<filename without extension> – <your header text>`** in the **safe top-right band** on **every page** (PDF pages and images).
- Visible **drag & drop** area + file picker.
- Compact **log** (last ~500 lines) and **progress bar**.
- PDF handling more robust (`ignoreEncryption`, graceful skip on errors).
- Chunking still respects **never split an appendix**.

## Deploy
Replace files in your repo root: `index.html`, `style.css`, `app.js`, `sw.js`, `manifest.json`.
Then hard-refresh your site. (Service worker cache bumped to **v5**.)
