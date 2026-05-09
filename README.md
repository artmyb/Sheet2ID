# Sheet2ID

Sheet2ID is a local web app that turns spreadsheet rows into printable ID-card PDFs. It supports:

- `CSV`, `XLSX`, and `XLS` spreadsheet uploads
- photo matching based on values like `name` or `id` appearing in photo filenames
- background switching by spreadsheet field rules such as `duty`
- configurable card size, page size, grid layout, and text/photo placement
- a GUI editor for page settings, matching rules, and card elements
- a sticky live single-card preview plus a separate page-setup preview

## Run it

1. Install dependencies once:

```powershell
npm.cmd install
```

2. Start the app:

```powershell
node server.js
```

3. Open `http://127.0.0.1:3300`

## Deploy on Vercel

This repo is now set up for Vercel deployment.

Why it fits Vercel's Express model:

- the Express entry file is `server.js`, which matches Vercel's supported Express entry names
- the app exports the Express instance with `module.exports = app`
- static assets already live in `public/`, which is the location Vercel serves from its CDN
- `vercel.json` is included and sets the function timeout explicitly for `server.js`

Files added or relied on for deployment:

- `server.js`
- `public/**`
- `package.json`
- `vercel.json`

Dashboard deploy flow:

1. Push the latest commit to GitHub.
2. In Vercel, click `New Project`.
3. Import this repository.
4. On the project configuration screen:
   - Root Directory: leave as the repo root
   - Framework Preset: let Vercel auto-detect Express, or choose `Other` if it does not
   - Build Command: leave empty
   - Output Directory: leave empty
   - Install Command: leave empty
5. Confirm the Node version is `24.x`.
6. Click `Deploy`.

After the first deploy:

1. Open the deployment URL.
2. Confirm `/` loads the app shell.
3. Confirm `/api/default-config` returns JSON.
4. Try a very small spreadsheet upload first.

CLI alternative:

```powershell
npm install -g vercel
vercel
```

For a production deployment from the CLI:

```powershell
vercel --prod
```

Important limitation on Vercel:

- Vercel Functions currently have a `4.5 MB` request-body limit, and this app uploads the spreadsheet plus all selected photos/backgrounds in a single request.
- That means this deployment is safe on Vercel only for small demos or very small asset sets.
- For real-world batches of photos, a better deployment shape is:
  client uploads directly to storage first, or
  the PDF generation runs on a platform without the same body-size limit.

Runtime note:

- The PDF renderer now tries both Windows and common Linux font paths, so it is much more likely to work on Vercel's Linux runtime than before.
- On Vercel, files inside `public/` are served by the platform CDN. `express.static()` is still useful locally, but Vercel serves those assets separately.

## How it works

1. Upload a spreadsheet.
2. Optionally select a photo folder and a background folder.
3. Edit the layout in the GUI builder.
4. Watch the live ID preview and the page-setup preview update automatically.
5. Click `Generate Full PDF` to create the final ID-card sheet.

## Config notes

The GUI builder controls:

- `page`: page preset or custom size, margin, orientation, gap
- `layout`: how many cards per row and per column
- `card`: card size, border, default background, and rule-based backgrounds
- `photos.fields`: spreadsheet columns used to match image filenames
- `elements`: absolute-positioned `text` and `photo` blocks inside each card

There is also an advanced JSON panel for power users, but it is optional.

Text elements can use:

- `field`: direct spreadsheet column name
- `template`: string with placeholders like `{{name}}`
- `prefix` and `suffix`
- `font`, `fontSize`, `minFontSize`, `align`, `valign`, `lineHeight`

Photo elements can use:

- `fit`: `cover` or `contain`
- `xMm`, `yMm`, `widthMm`, `heightMm`
- optional placeholder styling
