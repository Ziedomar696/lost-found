# مفقودات مراسي — Lost & Found

App link: **https://ziedomar696.github.io/lost-found/**

- `apps-script/`: the Google Apps Script project (Code.gs, Index.html, appsscript.json).
  Data goes to the Google Sheet **مفقودات مراسي**; photos go to two private Drive folders.
- `index.html`, `manifest.webmanifest`, `sw.js`, `icons/`: the installable app (GitHub Pages).
  `APP_URL` in `index.html` points at the Apps Script `/exec` link.

Updating the code: paste the new Code.gs / Index.html into the Apps Script project (keep your `SETUP_CODE`),
then Deploy → Manage deployments → ✏️ → New version. The link stays the same.
