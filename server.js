const path = require("path");
const express = require("express");
const multer = require("multer");
const { defaultConfig } = require("./src/default-config");
const { buildCardModels, buildPreview, indexImageFiles } = require("./src/lib/assets");
const { normalizeConfig } = require("./src/lib/config");
const { generatePdf } = require("./src/lib/pdf");
const { loadSpreadsheet } = require("./src/lib/spreadsheet");

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
    files: 1000,
  },
});

const uploadFields = upload.fields([
  { name: "spreadsheet", maxCount: 1 },
  { name: "photos", maxCount: 800 },
  { name: "backgrounds", maxCount: 200 },
]);

app.use(express.static(path.join(__dirname, "public")));

app.get("/api/default-config", (_req, res) => {
  res.json(defaultConfig);
});

app.post(
  "/api/preview",
  uploadFields,
  async (req, res, next) => {
    try {
      const context = await buildRequestContext(req);
      res.json({
        sourceSheet: context.spreadsheet.sheetName,
        availableSheets: context.spreadsheet.workbookSheetNames,
        ...buildPreview(context.cards, context.config, context.spreadsheet.headers),
      });
    } catch (error) {
      next(error);
    }
  }
);

app.post(
  "/api/preview-document",
  uploadFields,
  async (req, res, next) => {
    try {
      const context = await buildRequestContext(req);
      const preview = buildPreviewDocument(context, req.body);
      const pdfBytes = await generatePdf(preview.cards, preview.config);

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `inline; filename="${preview.filename}"`);
      res.send(Buffer.from(pdfBytes));
    } catch (error) {
      next(error);
    }
  }
);

app.post(
  "/api/generate",
  uploadFields,
  async (req, res, next) => {
    try {
      const context = await buildRequestContext(req);
      const pdfBytes = await generatePdf(context.cards, context.config);

      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", 'attachment; filename="sheet2id-cards.pdf"');
      res.send(Buffer.from(pdfBytes));
    } catch (error) {
      next(error);
    }
  }
);

app.use((error, _req, res, _next) => {
  const message = error && error.message ? error.message : "Unexpected server error.";
  const statusCode = error && error.statusCode ? error.statusCode : 400;
  res.status(statusCode).json({ error: message });
});

async function buildRequestContext(req) {
  const config = normalizeConfig(parseConfig(req.body.config));
  const spreadsheetFile = req.files && req.files.spreadsheet ? req.files.spreadsheet[0] : null;
  const photoFiles = req.files && req.files.photos ? req.files.photos : [];
  const backgroundFiles = req.files && req.files.backgrounds ? req.files.backgrounds : [];
  const spreadsheet = loadSpreadsheet(spreadsheetFile, config);
  const photoAssets = indexImageFiles(photoFiles);
  const backgroundAssets = indexImageFiles(backgroundFiles);
  const cards = buildCardModels(spreadsheet.rows, photoAssets, backgroundAssets, config);

  return {
    config,
    spreadsheet,
    cards,
  };
}

function parseConfig(rawConfig) {
  if (!rawConfig) {
    return {};
  }

  try {
    return JSON.parse(rawConfig);
  } catch {
    throw new Error("The configuration JSON is not valid.");
  }
}

function buildPreviewDocument(context, body = {}) {
  if (!context.cards.length) {
    throw new Error("There are no rows to preview.");
  }

  const previewMode = String(body.previewMode || "single-card").toLowerCase() === "page" ? "page" : "single-card";
  const previewStart = Math.max(1, Number.parseInt(body.previewStart, 10) || 1);
  const startIndex = Math.min(context.cards.length - 1, previewStart - 1);

  if (previewMode === "page") {
    const cardsPerPage = Math.max(1, context.config.layout.columns * context.config.layout.rows);
    return {
      cards: context.cards.slice(startIndex, startIndex + cardsPerPage),
      config: context.config,
      filename: "sheet2id-preview-page.pdf",
    };
  }

  const previewMargin = Math.max(6, Number(context.config.page.marginMm) || 10);
  const singleCardConfig = {
    ...context.config,
    page: {
      ...context.config.page,
      size: "Custom",
      orientation: context.config.card.widthMm >= context.config.card.heightMm ? "landscape" : "portrait",
      widthMm: context.config.card.widthMm + previewMargin * 2,
      heightMm: context.config.card.heightMm + previewMargin * 2,
      marginMm: previewMargin,
      gapMm: 0,
    },
    layout: {
      ...context.config.layout,
      columns: 1,
      rows: 1,
      centerOnPage: true,
    },
  };

  return {
    cards: context.cards.slice(startIndex, startIndex + 1),
    config: singleCardConfig,
    filename: "sheet2id-preview-card.pdf",
  };
}

const PORT = Number(process.env.PORT) || 3300;

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Sheet2ID running at http://127.0.0.1:${PORT}`);
  });
}

module.exports = app;
