# Sheet2ID

Sheet2ID is a local web app that turns spreadsheet rows into printable ID-card PDFs. It supports:

- `CSV`, `XLSX`, and `XLS` spreadsheet uploads
- photo matching based on values like `name` or `id` appearing in photo filenames
- background switching by spreadsheet field rules such as `duty`
- configurable card size, page size, grid layout, and text/photo placement
- a GUI editor for page settings, matching rules, and card elements
- previewing one card or one page before generating the full PDF

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

This project can be deployed to Vercel because:

- the Express entry file is [server.js](/c:/Users/mybas/Desktop/Sheet2ID/server.js), which matches Vercel's supported Express entry names
- static files already live in `public/`
- the app exports the Express instance with `module.exports = app`

Basic deploy flow:

1. Push this folder to GitHub, GitLab, or Bitbucket.
2. Import the repo into Vercel.
3. Keep the framework preset as `Other` or let Vercel auto-detect Express.
4. Leave the build command empty.
5. Set the Node version to `24.x` if Vercel does not pick it automatically.
6. Deploy.

CLI alternative:

```powershell
npm install -g vercel
vercel
```

Important limitation:

- Vercel Functions currently have a `4.5 MB` request-body limit, and this app uploads the spreadsheet plus all selected photos/backgrounds in a single request.
- That means the current architecture is only safe on Vercel for small demos or very small asset sets.
- For real-world batches of photos, a better deployment shape is:
  client uploads directly to storage first, or
  the PDF generation runs on a platform without the same body-size limit.

Font note:

- The PDF renderer now tries both Windows and common Linux font paths, so it is much more likely to work on Vercel's Linux runtime than before.

## How it works

1. Upload a spreadsheet.
2. Optionally select a photo folder and a background folder.
3. Edit the layout in the GUI builder.
4. Click `Preview` to inspect matching plus either a single-card preview or a one-page preview.
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
