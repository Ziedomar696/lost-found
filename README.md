# مفقودات مراسي — Lost & Found

App link: **https://ziedomar696.github.io/lost-found/**

- `apps-script/`: the Google Apps Script project (Code.gs, Index.html, appsscript.json).
  New version (مخزن مفقودات مراسي): username/password accounts managed by the admin, roles
  (الإدارة، أمين مخزن، مشرف، الكنترول روم), exit permits, deliveries with signature, loss reports,
  control-room requests, and editable lists. Deploy as Execute as: Me / Who has access: Anyone.
- `index.html`, `manifest.webmanifest`, `sw.js`, `icons/`: the installable app (GitHub Pages).
  `APP_URL` in `index.html` points at the Apps Script `/exec` link.

Updating the code: paste the new Code.gs / Index.html into the Apps Script project (keep your `SETUP_CODE`),
then Deploy → Manage deployments → ✏️ → New version. The link stays the same.
