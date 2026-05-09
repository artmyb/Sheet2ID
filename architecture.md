# Sheet2ID Architecture

## 1. Product Definition

Sheet2ID is a local-first web application that converts spreadsheet rows into printable ID-card PDFs.

Each spreadsheet row becomes one card. The app optionally matches:

- a person photo by comparing spreadsheet values such as `name` and `id` against uploaded photo filenames
- a background image by evaluating rule-based conditions such as `duty == Manager`

The user can control:

- source spreadsheet sheet selection
- page size, orientation, page margin, and gap
- grid layout such as `2 columns x 5 rows`
- card width, height, fill, and border
- photo matching fields
- background mapping rules
- absolutely positioned text and photo elements inside the card
- separate live previews for one isolated card and for page setup before full export

The app is designed to run locally with a browser UI and an Express server. It does not require a database and does not persist uploads on disk.

## 2. Primary Goals

The application is built around these goals:

1. Convert spreadsheet rows into cards with very little user setup.
2. Let non-technical users edit layout visually instead of writing raw JSON.
3. Keep the full workflow local and simple.
4. Support multilingual names and titles, including Turkish characters such as `S`, `C`, `G`, `I`, `O`, and `U` with diacritics.
5. Support a quick live-preview cycle before generating the final batch PDF.
6. Remain configurable enough to support many card templates without changing code.

## 3. Non-Goals

Current non-goals or deliberate simplifications:

- no database
- no user accounts
- no saved projects on the server
- no drag-and-drop WYSIWYG canvas editor
- no vector template system
- no per-row custom element visibility rules beyond `hideWhenEmpty`
- no OCR or face detection
- no remote storage integration

## 4. Runtime Stack

### Backend

- Node.js
- Express 5
- Multer with in-memory uploads
- `xlsx` for CSV/XLS/XLSX parsing
- `pdf-lib` for PDF creation
- `@pdf-lib/fontkit` for embedding Unicode-capable fonts

### Frontend

- plain HTML
- plain CSS
- plain browser JavaScript
- no framework

### Process model

- one Express process
- static assets served from `public/`
- JSON/config and uploads travel in multipart form requests
- generated PDFs are returned directly in the HTTP response

## 5. Project Layout

```text
Sheet2ID/
  server.js
  package.json
  architecture.md
  architecture.mmd
  README.md
  public/
    index.html
    styles.css
    app.js
  src/
    default-config.js
    lib/
      config.js
      spreadsheet.js
      assets.js
      pdf.js
      utils.js
```

### File responsibilities

`server.js`
- creates the Express app
- configures Multer upload handling
- exposes the HTTP API
- constructs request context objects
- builds preview-document configuration variants for single-card and page PDF previews

`public/index.html`
- defines the page structure
- contains file inputs, a sticky live-preview panel, the config panel, the summary table, the generated-PDF iframe, and datalists such as font suggestions

`public/styles.css`
- handles the application layout
- styles the builder UI
- keeps the live-preview panel visible with sticky positioning on wide screens
- implements the collapsible section and collapsible card presentation

`public/app.js`
- loads the default config
- stores the working config in localStorage
- renders the GUI builder
- synchronizes GUI state with raw JSON
- refreshes the summary preview automatically
- refreshes the live single-card preview PDF automatically
- renders the page-setup rectangle preview in the browser
- loads returned PDFs into the appropriate iframe

`src/default-config.js`
- contains the initial configuration object used by both backend and frontend

`src/lib/config.js`
- normalizes and validates config data on the server

`src/lib/spreadsheet.js`
- loads the spreadsheet
- chooses a sheet
- converts rows into normalized row objects with lookup metadata

`src/lib/assets.js`
- indexes image uploads
- matches backgrounds and photos
- builds card models
- builds preview summary JSON

`src/lib/pdf.js`
- renders cards into PDF pages
- resolves fonts
- embeds images
- fits text into boxes

`src/lib/utils.js`
- shared helpers for unit conversion, normalization, lookup, templating, color parsing, and merging

## 6. End-to-End User Flow

1. The browser opens `/`.
2. The frontend requests `/api/default-config`.
3. The frontend normalizes that config for the GUI and merges any saved local config from `localStorage`.
4. The user selects:
   - one spreadsheet file
   - an optional photos folder
   - an optional backgrounds folder
5. The user edits settings in the collapsible GUI builder.
6. The frontend redraws the page-setup preview immediately in the browser from the current config.
7. When a spreadsheet is available, the frontend also refreshes the summary JSON and the single-card preview PDF automatically.
8. The user can change `previewStart` to inspect a different source row.
9. The live single-card preview stays visible in a sticky panel while the rest of the page scrolls.
10. When satisfied, the user clicks `Generate Full PDF`.
11. The frontend sends the multipart payload to `/api/generate`.
12. The final PDF loads into the generated-document iframe and becomes downloadable.

## 7. HTTP API

### `GET /api/default-config`

Purpose:
- return the shared default config used to bootstrap the GUI

Response:
- JSON object equal to `src/default-config.js`

### `POST /api/preview`

Purpose:
- parse inputs
- match photos and backgrounds
- return lightweight preview data for the first 50 rows

Multipart fields:
- `spreadsheet`: required, single file
- `photos`: optional, many files
- `backgrounds`: optional, many files
- `config`: required JSON string in practice, but the server tolerates empty input and falls back to defaults
- `previewMode`: ignored by this route
- `previewStart`: ignored by this route

Response fields:
- `sourceSheet`
- `availableSheets`
- `totalRows`
- `matchedPhotos`
- `missingPhotos`
- `matchedBackgrounds`
- `headers`
- `displayFields`
- `cards`

Important behavior:
- `cards` is intentionally capped to the first 50 rows for the table preview
- this route does not return a PDF

### `POST /api/preview-document`

Purpose:
- render a preview PDF without generating the entire batch
- power the browser's live single-card preview

Multipart fields:
- same fields as `/api/preview`
- `previewMode` is used
- `previewStart` is used

Preview modes:

`single-card`
- isolates exactly one card
- creates a temporary derived config with:
  - page size `Custom`
  - one column
  - one row
  - page width = `card.widthMm + 2 * previewMargin`
  - page height = `card.heightMm + 2 * previewMargin`
  - page gap = `0`
  - layout centered
- orientation is chosen from the card proportions:
  - landscape if card width >= card height
  - portrait otherwise

`page`
- renders one page worth of cards using the real current layout
- card count = `layout.columns * layout.rows`
- cards start at the 1-based row selected by `previewStart`

Current browser usage:

- the frontend always requests `single-card`
- page placement is previewed client-side as rectangles instead of by requesting a page PDF

### `POST /api/generate`

Purpose:
- render the full final PDF

Multipart fields:
- same as `/api/preview`

Response:
- binary PDF
- content disposition is `attachment`

## 8. Upload Handling

Uploads are configured in `server.js` with Multer memory storage.

Limits:

- `fileSize`: `25 * 1024 * 1024`
- `files`: `1000`

Accepted upload groups:

- `spreadsheet`: max `1`
- `photos`: max `800`
- `backgrounds`: max `200`

Why memory storage is used:

- avoids temporary file cleanup complexity
- makes the app simple for local execution
- allows direct handoff into parsing and PDF generation

Tradeoff:

- large jobs increase memory usage because files are stored in RAM during the request

## 9. Shared Data Models

### 9.1 Config Object

The full working config shape is:

```js
{
  spreadsheet: {
    sheetName: "",
    skipEmptyRows: true
  },
  page: {
    size: "A4",
    orientation: "portrait",
    widthMm: 210,
    heightMm: 297,
    marginMm: 10,
    gapMm: 6
  },
  layout: {
    columns: 2,
    rows: 5,
    centerOnPage: true
  },
  card: {
    widthMm: 85.6,
    heightMm: 54,
    fillColor: "#FFFFFF",
    borderColor: "#D6D0C4",
    borderWidthMm: 0.35,
    background: {
      default: "",
      rules: [
        {
          field: "duty",
          equals: "Manager",
          background: "manager"
        },
        {
          field: "duty",
          equals: "Coach",
          background: "coach"
        }
      ]
    }
  },
  photos: {
    fields: ["name", "id"],
    placeholderLabel: "NO PHOTO"
  },
  elements: [
    {
      id: "photo",
      type: "photo",
      xMm: 6,
      yMm: 8,
      widthMm: 24,
      heightMm: 30,
      fit: "cover",
      borderColor: "#FFFFFF",
      borderWidthMm: 0.5,
      fillColor: "#F3F0EA"
    },
    {
      id: "name",
      type: "text",
      field: "name",
      xMm: 34,
      yMm: 11,
      widthMm: 46,
      heightMm: 12,
      font: "Helvetica-Bold",
      fontSize: 14,
      minFontSize: 10,
      color: "#111111",
      align: "left",
      valign: "top",
      lineHeight: 1.1,
      uppercase: true
    }
  ]
}
```

### 9.2 Normalized Spreadsheet Row

Each parsed row becomes an object that includes both original spreadsheet keys and metadata:

```js
{
  Name: "Alice Stone",
  Duty: "Manager",
  ID: "A-001",
  __lookup: {
    name: "Alice Stone",
    duty: "Manager",
    id: "A-001"
  },
  __rowNumber: 2,
  __isEmpty: false
}
```

Important rules:

- `__rowNumber` starts at `2` because row `1` is the header row
- `__lookup` stores normalized header keys so field lookups can be case-insensitive and punctuation-insensitive

### 9.3 Asset Model

Each uploaded image file is normalized into:

```js
{
  id: "0-photo_001.jpg",
  originalname: "photo_001.jpg",
  relativeName: "staff/photo_001.jpg",
  basename: "photo_001",
  extension: ".jpg",
  normalizedBase: "photo 001",
  compactBase: "photo001",
  buffer: <Buffer>
}
```

### 9.4 Card Model

Each row becomes:

```js
{
  index: 0,
  row: <normalized spreadsheet row>,
  backgroundAsset: <asset or null>,
  photoAsset: <asset or null>
}
```

## 10. Config Normalization Rules

Server normalization is performed in `src/lib/config.js`.

Important guarantees:

- missing nested objects are created
- numeric fields are coerced
- minimums are enforced
- layout rows and columns are at least `1`
- card dimensions are at least `1`
- border widths, padding, page margin, and page gap are clamped to `>= 0`
- `elements` is always an array of normalized element objects
- if all elements are removed, the default elements are restored
- orientation is normalized to `portrait` or `landscape`

Client normalization in `public/app.js` does similar work for UI stability:

- color inputs are forced to valid 6-digit hex values
- missing arrays become empty arrays
- rule mode is inferred from whether a rule uses `equals`, `contains`, or `regex`
- element IDs are auto-generated when missing

The frontend and backend both normalize because:

- the frontend needs predictable fields to render the form
- the backend must not trust browser state

## 11. Spreadsheet Parsing

Implemented in `src/lib/spreadsheet.js`.

### Supported formats

- `.csv`
- `.xlsx`
- `.xls`

### Parsing procedure

1. Read the uploaded spreadsheet buffer with `XLSX.read(..., { type: "buffer" })`.
2. Determine the requested sheet name from `config.spreadsheet.sheetName`.
3. If the requested sheet is not present, fall back to the workbook's first sheet.
4. Throw an error if the workbook has no sheets.
5. Parse row objects with:
   - `defval: ""`
   - `raw: false`
   - `blankrows: false`
6. Parse the raw header row separately with `header: 1`.
7. Convert each row using `decorateRow(row, index + 2)`.
8. If `skipEmptyRows` is true, discard rows where every value is empty.

### Header and field matching behavior

Field lookups are not limited to exact original header spelling.

Example:

- spreadsheet header: `Employee ID`
- config field: `employee_id`
- config field: `Employee-ID`

All of these can resolve because the lookup normalizes keys by:

- Unicode decomposition
- removal of combining marks
- lowercase conversion
- replacement of non-alphanumeric runs with spaces
- trimming

## 12. Asset Indexing

Implemented in `src/lib/assets.js`.

Supported image extensions:

- `.png`
- `.jpg`
- `.jpeg`

Indexing procedure:

1. Ignore files without buffers.
2. Ignore files whose extension is not supported.
3. Parse filename pieces with `path.parse`.
4. Store:
   - base filename
   - extension
   - normalized and compacted basename
   - original and relative names
   - raw buffer

The app does not inspect image pixels for matching. Matching is filename-driven only.

## 13. Background Matching

Background matching is deterministic and rule-ordered.

Algorithm:

1. Iterate through `config.card.background.rules` in array order.
2. For each rule, evaluate `matchesRule(row, rule)`.
3. If a rule matches, try to find a background asset by `rule.background`.
4. Return the first matched asset.
5. If no rule matches or the referenced asset is missing, try `config.card.background.default`.
6. If that also fails, return `null`.

### Supported rule styles

Exact match:

```js
{ field: "duty", equals: "Manager", background: "manager" }
```

Contains match:

```js
{ field: "department", contains: "sales", background: "sales" }
```

Regex match:

```js
{ field: "title", regex: "^senior", flags: "i", background: "senior" }
```

### Rule evaluation details

`equals`
- compares `normalizeCompact(value)` against `normalizeCompact(ruleValue)`

`contains`
- compares `normalizeKey(value).includes(normalizeKey(rule.contains))`

`regex`
- creates `new RegExp(rule.regex, rule.flags || "i")`
- invalid regex patterns fail safely and simply do not match

### Asset reference matching

A background reference such as `"manager"` is matched against uploaded files using this order:

1. exact `compactBase`
2. exact `normalizedBase`
3. partial `compactBase` inclusion
4. partial `normalizedBase` inclusion

This lets `"manager"` match files like:

- `manager.png`
- `manager blue.jpg`
- `staff-manager-card.jpeg`

## 14. Photo Matching

Photo matching is heuristic and score-based.

Implemented in `resolvePhotoAsset`, `buildCandidates`, and `scoreAsset`.

### Candidate creation

The app uses `config.photos.fields` to generate matching candidates from each row.

For each configured field:

1. Look up the row value.
2. Skip empty values.
3. Build:
   - `normalized`
   - `compact`
   - `tokens`
4. Assign a weight based on field order:
   - first field weight = `5`
   - second field weight = `4`
   - third field weight = `3`
   - fourth field weight = `2`
   - fifth and later field weight = `1`

The app also builds one combined candidate from all configured fields joined with spaces, with weight `2`.

Tokens are filtered so that:

- tokens with at least 2 characters are allowed
- purely numeric tokens are allowed even if shorter

### Scoring formula

For each asset and each candidate:

- exact compact match: `+140`
- exact normalized match: `+120`
- one compact string contains the other: `+75`
- each matched token: `+14`
- if all tokens match and there is more than one token: `+35`
- if asset compact base starts with candidate compact: `+10`

The candidate score is then multiplied by the candidate weight.

The best candidate score for that asset becomes the asset score.
The highest scoring asset wins.

If the best score is `0`, the row is treated as having no matched photo.

### Why this works

This approach handles common filename patterns such as:

- `alice-stone.jpg`
- `A001_alice.png`
- `stone_alice_id_A001.jpeg`

without requiring a perfect filename convention.

## 15. Preview Summary JSON

Implemented in `buildPreview`.

Purpose:

- give the user a quick table to inspect data matching before export
- refresh continuously during the live-preview editing loop

Preview summary fields:

- `totalRows`
- `matchedPhotos`
- `missingPhotos`
- `matchedBackgrounds`
- `headers`
- `displayFields`
- `cards`

`displayFields` is derived from the union of:

- `config.photos.fields`
- all `text` element fields
- all background rule fields

Duplicates are removed while preserving only truthy values.

The table row payload is limited to the first 50 cards and includes:

- original spreadsheet row number
- selected display values
- matched photo filename
- matched background filename

## 16. Preview Document Rendering

Preview rendering is split across separate browser surfaces.

Reason:

- the user wanted page placement and card-content preview to be separate
- the ID preview needed to remain visible while the rest of the page scrolls
- page placement does not require full PDF rendering

### Live single-card preview

Use case:
- inspect typography, spacing, background choice, and photo crop for one person

Behavior:
- one row only
- one card only
- rendered through `/api/preview-document`
- temporary custom page dimensions based on the card dimensions
- keeps the actual element layout identical to final rendering

### Page-setup preview

Use case:
- inspect the real sheet layout and spacing

Behavior:
- rendered directly in the browser, not on the server
- draws only page and card-slot rectangles
- uses the real page size, margin, orientation, gap, card size, and grid layout
- highlights when the configured slots extend past the page margins

### Server-side page preview capability

The backend still supports a page PDF preview through `previewMode = page`.

This is currently retained as route capability, but the browser UI no longer uses it because the lightweight rectangle preview is faster and better suited for page-placement tweaking.

### Preview start row

`previewStart` is 1-based in the UI.

Backend handling:

- parse integer
- clamp to at least `1`
- convert to zero-based index
- clamp to `context.cards.length - 1`

## 17. PDF Rendering Pipeline

Implemented in `src/lib/pdf.js`.

### Overall procedure

1. Create a new `PDFDocument`.
2. Register `fontkit`.
3. Embed the fonts needed by text elements plus default regular/bold fonts.
4. Create an image cache map.
5. Resolve page size.
6. Resolve grid geometry.
7. Determine cards per page.
8. Add pages as needed.
9. For each card:
   - resolve the target card rectangle
   - draw base fill
   - draw background image
   - draw card border
   - render each configured element in order
10. Save the PDF as bytes.

### Coordinate system

`pdf-lib` uses points, not millimeters.

Conversion:

```text
pt = mm * 72 / 25.4
```

All user-facing geometry values are stored in millimeters and converted during rendering.

### Page size resolution

If `config.page.size` matches a named `pdf-lib` preset such as `A4` or `Letter`, use that preset.

Otherwise:

- width = `mmToPt(config.page.widthMm)`
- height = `mmToPt(config.page.heightMm)`

Orientation handling:

- landscape => wider dimension becomes width
- portrait => taller dimension becomes height

### Grid resolution

Inputs:

- page margin
- page gap
- card width
- card height
- number of columns
- number of rows
- whether to center on page

Formulas:

```text
usableWidth  = pageWidth  - 2 * margin
usableHeight = pageHeight - 2 * margin
gridWidth    = columns * cardWidth  + (columns - 1) * gap
gridHeight   = rows    * cardHeight + (rows    - 1) * gap
```

Offsets:

- if centering is enabled, remaining space is split on both sides
- otherwise the grid begins at the top-left margin area

### Card rectangle resolution

For a given `slot`:

```text
column = slot % columns
row    = floor(slot / columns)
```

Card origin:

```text
x = offsetX + column * (cardWidth + gap)
y = pageHeight - offsetY - cardHeight - row * (cardHeight + gap)
```

This keeps rows flowing top-to-bottom visually while respecting PDF bottom-left coordinates.

## 18. Element Rendering

Elements are rendered strictly in config array order.

That means later elements can visually cover earlier ones.

### 18.1 Photo elements

Photo element flow:

1. Resolve element box in points.
2. If `fillColor` exists, draw the box fill first.
3. If a photo asset exists:
   - embed or reuse cached image
   - draw it using `cover` or `contain`
4. If no photo asset exists:
   - draw placeholder text centered in the box
   - font comes from the embedded default bold font or default regular font
   - label comes from `config.photos.placeholderLabel`
5. If `borderWidthMm > 0`, draw the element border

### 18.2 Text elements

Text element flow:

1. Build text using `template`, `field`, `prefix`, `suffix`, and `value`.
2. Return early if the text is empty and `hideWhenEmpty` is true.
3. Uppercase the final text if `uppercase` is enabled.
4. Resolve the font from the font cache.
5. Resolve the text box and inner padding.
6. Fit the text into the box by reducing size from `fontSize` down to `minFontSize` in steps of `0.5`.
7. Wrap the text into lines based on measured width.
8. If text still overflows at minimum size, truncate the final visible line with `...`.
9. Place the text according to horizontal and vertical alignment.
10. Draw each line separately.

### Text source priority

Priority order:

1. `template`
2. `field`
3. `value`

Examples:

Field-based:

```js
{ field: "name" }
```

Templated:

```js
{ template: "{{name}} / {{id}}" }
```

Static:

```js
{ value: "VISITOR" }
```

Prefix and suffix are applied in field or static mode.

## 19. Text Fitting and Wrapping

The text-fitting behavior is important because the layout is absolute and card space is limited.

### Fitting loop

For sizes from `fontSize` down to `minFontSize`:

1. compute `lineHeight = size * element.lineHeight`
2. compute how many lines can fit vertically
3. wrap text to the available width
4. if wrapped line count fits, accept that size

If no size fits:

- use `minFontSize`
- wrap anyway
- trim to max visible lines
- truncate the last line with ellipsis

### Word wrapping

Paragraphs split on newlines.

Within each paragraph:

- words are assembled into a line while the measured width still fits
- if a single word is too long, it is broken character by character

This avoids total failure for long IDs or unbroken strings.

## 20. Image Rendering Modes

Two image fit modes are supported:

### `contain`

- scale image to fit entirely inside the box
- preserve aspect ratio
- center it within the box

### `cover`

- scale image until the box is fully covered
- preserve aspect ratio
- center the image
- clip overflow using a rectangular clip path

This is used for photo crops and full-card backgrounds.

## 21. Font System

The app originally relied on built-in WinAnsi fonts, which fail for many non-English characters.
The current system embeds real fonts with `@pdf-lib/fontkit`.

### Font embedding strategy

1. Collect the font names requested by text elements.
2. Normalize each font name to an internal key.
3. For each key, try a list of candidate font files in order.
4. If no candidate is readable, fall back to a `pdf-lib` standard font.

### Why candidate lists exist

The app can run on:

- Windows local development
- Linux-like environments

So each font family tries:

- Windows system font paths first when applicable
- then common Linux font paths such as DejaVu, Liberation, or Noto

### Supported families

The renderer recognizes a broad set of names, including:

- Helvetica family
- Arial family
- Arial Narrow family
- Segoe UI family
- Calibri family
- Verdana family
- Tahoma family
- Trebuchet MS family
- Century Gothic family
- Franklin Gothic Medium
- Candara family
- Corbel family
- Times / Times New Roman family
- Georgia family
- Palatino Linotype family
- Book Antiqua family
- Cambria family
- Constantia family
- Garamond family
- Rockwell family
- Courier / Courier New family
- Consolas family
- Lucida Console
- Comic Sans MS
- Impact
- DejaVu Sans / Serif / Sans Mono
- Liberation Sans / Serif / Mono
- Noto Sans / Serif / Sans Mono

### Alias handling

Aliases map common user-facing names to internal renderer keys.

Examples:

- `Times New Roman` -> `times-roman`
- `Courier New Bold` -> `courier-bold`
- `DejaVu Sans` -> `helvetica`
- `Noto Serif Bold` -> `times-bold`
- `monospace` -> `courier`

### Unicode behavior

Because the app embeds actual font files when available, it can render names such as:

- `Sule Cetin`
- `Ilker Dogan`
- and the original Turkish-character versions those spellings came from

without the WinAnsi encoding crash.

## 22. Frontend Structure

The frontend is intentionally framework-free.

Main page areas:

1. Hero header
2. Assets
3. Live Previews
4. Visual Builder
5. Match Summary
6. Generated PDF

### Assets panel

Contains:

- spreadsheet file input
- photos folder input with `webkitdirectory`
- backgrounds folder input with `webkitdirectory`
- full generate button
- status message

### Live Previews

Contains:

- preview start row numeric input
- sticky iframe for the live single-card PDF preview
- download link for the current single-card preview
- browser-rendered page-setup rectangle preview
- page-setup summary text

### Visual Builder

Contains collapsible sections for:

- Spreadsheet
- Page
- Grid Layout
- Card
- Photo Matching
- Background Rules
- Card Elements

It also contains a collapsible `Advanced JSON` area for power users.

### Match Summary

Contains:

- three summary metric cards
- a table showing the first 50 preview rows

### Generated PDF

Contains:

- download link
- iframe used to display the final generated PDF

## 23. Frontend State Model

`public/app.js` keeps a single `state` object with:

- `defaultConfig`
- `currentConfig`
- `availableSheets`
- `activePreviewPdfUrl`
- `activeGeneratedPdfUrl`
- `uiOpen`
- `livePreviewTimer`
- `livePreviewRequestId`
- `livePreviewController`

### Local storage

Config is saved in:

```text
sheet2id-config-v2
```

This means:

- page refresh keeps the last edited config
- uploaded files are not persisted
- collapsible openness is reset only when explicitly requested by code paths such as default reset

### UI openness tracking

The app remembers open/closed state for:

- top-level builder sections
- background rule cards
- element cards

This is implemented with `<details>` elements and `data-ui-group` / `data-ui-key` attributes.

Why this matters:

- long parameter lists no longer force the page to stay fully expanded
- the user can focus on one configuration area at a time

## 24. Frontend Config Rendering

The config GUI is not static HTML.
It is rendered dynamically from `state.currentConfig`.

Key renderer responsibilities:

- render inputs for primitive config fields
- render and normalize background rules
- render and normalize card elements
- keep raw JSON synchronized with GUI state
- re-render the page-setup diagram after config edits
- trigger a debounced live-preview refresh after config edits
- intercept numeric mouse-wheel adjustments so page scrolling does not happen at the same time
- rebuild sections after structural edits such as:
  - add rule
  - remove rule
  - add text element
  - add photo element
  - duplicate element
  - remove element

This dynamic approach keeps the UI consistent with the current config schema.

## 25. Preview and Generate Submission Flow in the Browser

### Live preview refresh

The browser no longer has a manual `Preview` button.

Instead, the live preview flow runs automatically when:

- spreadsheet, photo, or background uploads change
- `previewStart` changes
- GUI config fields change
- advanced JSON is applied
- the config is reset to defaults

Browser flow:

1. redraw the page-setup rectangle preview immediately from `state.currentConfig`
2. if no spreadsheet is present, show empty-state preview placeholders and stop
3. debounce rapid edits
4. submit multipart payload to `/api/preview`
5. render summary metrics and table from the JSON response
6. if rows exist, submit a fresh multipart payload to `/api/preview-document` with `previewMode = single-card`
7. convert the preview PDF response to a blob
8. assign a blob URL to the sticky preview iframe
9. enable the single-card preview download link

Two requests are used because:

- summary JSON and binary PDF have different response types
- keeping them separate simplifies both backend routes

### Generate button

When the user clicks `Generate Full PDF`:

1. verify spreadsheet exists
2. build multipart payload
3. submit to `/api/generate`
4. load returned PDF into the generated-document iframe
5. update the download link filename to the final export name

## 26. Error Handling

All major server routes use `try/catch` and pass failures into a common Express error middleware.

The middleware returns:

```js
{ error: "message" }
```

Default status code:

- `400` unless a thrown error provides `statusCode`

Typical error cases:

- missing spreadsheet
- invalid config JSON
- workbook with no sheets
- impossible layout with zero room for cards
- no rows available for preview

Frontend behavior:

- catches failed fetch responses
- tries to read `{ error }` JSON
- shows the error string in the status banner

## 27. Performance Characteristics

This app is efficient enough for typical local batch jobs but is not optimized for huge workloads.

Important characteristics:

- uploads are held in memory
- images are embedded lazily and cached per request
- fonts are embedded only for actually requested font keys
- photo matching compares every row candidate against every photo asset, so very large photo sets will scale linearly

Current complexity notes:

- spreadsheet parsing is roughly linear in row count
- asset indexing is linear in file count
- photo matching is approximately `rowCount * photoCount * candidateCount`

For local desktop usage, this is acceptable for normal team ID batches.

## 28. Constraints and Tradeoffs

### Strengths

- very easy local setup
- good configurability
- no database overhead
- Unicode-friendly PDF output
- always-on live previews speed up iteration

### Limitations

- no persistent project save format beyond exported JSON and browser localStorage
- no background job queue
- no true visual drag positioning
- memory-only uploads
- no server-side font upload feature
- summary preview table always shows the first 50 rows, not the current preview slice

## 29. Rebuild Blueprint

This section is the shortest path to reproducing the app from scratch.

### Step 1: Create the Node project

Install:

- `express`
- `multer`
- `xlsx`
- `pdf-lib`
- `@pdf-lib/fontkit`

Use CommonJS to match the current implementation.

### Step 2: Build the Express server

Implement:

- static hosting from `public/`
- Multer memory storage
- upload field groups for spreadsheet, photos, backgrounds
- routes:
  - `GET /api/default-config`
  - `POST /api/preview`
  - `POST /api/preview-document`
  - `POST /api/generate`

### Step 3: Define a shared default config

Include:

- spreadsheet settings
- page settings
- grid layout settings
- card fill and border settings
- background rules
- photo matching fields
- default card elements

### Step 4: Implement config normalization

Guarantee:

- all nested objects exist
- values are coerced to expected types
- dimensions are clamped
- element arrays are valid

### Step 5: Implement spreadsheet parsing

With `xlsx`:

- select requested sheet or first sheet
- parse object rows
- parse headers
- decorate rows with normalized lookup data and row numbers
- optionally drop empty rows

### Step 6: Implement image indexing and matching

Index uploaded images by:

- basename
- normalized basename
- compact basename

Implement:

- ordered background rule evaluation
- score-based photo matching

### Step 7: Implement PDF rendering

Build:

- millimeter-to-point conversion
- page size resolution
- grid layout calculation
- card rectangle calculation
- image placement with `contain` and `cover`
- text wrapping and shrink-to-fit
- font embedding through `fontkit`

### Step 8: Implement preview derivation

Add:

- JSON summary preview route
- preview-document PDF route
- single-card preview config rewriting
- page preview card slicing
- client-side page-setup rectangle rendering

### Step 9: Build the browser UI

Create:

- file upload area
- sticky live single-card preview
- page-setup rectangle preview
- preview row control
- status messages
- summary metrics and table
- iframe for generated PDF output
- download link
- dynamic builder UI for config editing
- optional advanced JSON editor

### Step 10: Add local state and collapsible sections

Persist the config in localStorage.
Use collapsible sections so large configurations remain manageable.

### Step 11: Verify with multilingual data

Test with:

- non-ASCII names
- missing photo rows
- rules that map to uploaded background files
- live one-card preview
- page-setup rectangle preview
- full batch generation

## 30. Current Architecture Summary

The app uses a very direct pipeline:

```text
browser inputs
  -> multipart upload
  -> server parses config + spreadsheet + image assets
  -> rows become card models
  -> matches determine photo/background assets
  -> preview summary JSON and single-card PDF are generated
  -> browser displays summary table, sticky ID preview, and client-side page-setup rectangles
```

The overall design is intentionally simple:

- no database
- no build system
- no framework
- configuration-driven rendering
- live-preview-first workflow

That simplicity is the reason the app is easy to run locally, easy to modify, and easy to replicate.
