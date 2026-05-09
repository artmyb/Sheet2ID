const { defaultConfig } = require("../default-config");
const { deepMerge, integerOr, numberOr } = require("./utils");

function normalizeConfig(inputConfig = {}) {
  const config = deepMerge(defaultConfig, inputConfig);

  config.spreadsheet = config.spreadsheet || {};
  config.page = config.page || {};
  config.layout = config.layout || {};
  config.card = config.card || {};
  config.card.background = config.card.background || {};
  config.photos = config.photos || {};
  config.elements = Array.isArray(config.elements) ? config.elements : defaultConfig.elements;

  config.spreadsheet.sheetName = String(config.spreadsheet.sheetName || "");
  config.spreadsheet.skipEmptyRows = config.spreadsheet.skipEmptyRows !== false;

  config.page.size = String(config.page.size || "A4");
  config.page.orientation = String(config.page.orientation || "portrait").toLowerCase() === "landscape" ? "landscape" : "portrait";
  config.page.widthMm = numberOr(config.page.widthMm, defaultConfig.page.widthMm);
  config.page.heightMm = numberOr(config.page.heightMm, defaultConfig.page.heightMm);
  config.page.marginMm = Math.max(0, numberOr(config.page.marginMm, defaultConfig.page.marginMm));
  config.page.gapMm = Math.max(0, numberOr(config.page.gapMm, defaultConfig.page.gapMm));

  config.layout.columns = Math.max(1, integerOr(config.layout.columns, defaultConfig.layout.columns));
  config.layout.rows = Math.max(1, integerOr(config.layout.rows, defaultConfig.layout.rows));
  config.layout.centerOnPage = config.layout.centerOnPage !== false;

  config.card.widthMm = Math.max(1, numberOr(config.card.widthMm, defaultConfig.card.widthMm));
  config.card.heightMm = Math.max(1, numberOr(config.card.heightMm, defaultConfig.card.heightMm));
  config.card.fillColor = String(config.card.fillColor || defaultConfig.card.fillColor);
  config.card.borderColor = String(config.card.borderColor || defaultConfig.card.borderColor);
  config.card.borderWidthMm = Math.max(0, numberOr(config.card.borderWidthMm, defaultConfig.card.borderWidthMm));
  config.card.background.default = String(config.card.background.default || "");
  config.card.background.rules = Array.isArray(config.card.background.rules) ? config.card.background.rules : [];

  config.photos.fields = Array.isArray(config.photos.fields) ? config.photos.fields.map(String).filter(Boolean) : defaultConfig.photos.fields;
  config.photos.placeholderLabel = String(config.photos.placeholderLabel || defaultConfig.photos.placeholderLabel);

  config.elements = config.elements
    .filter((element) => element && typeof element === "object")
    .map((element, index) => normalizeElement(element, index));

  if (config.elements.length === 0) {
    config.elements = defaultConfig.elements.map((element, index) => normalizeElement(element, index));
  }

  return config;
}

function normalizeElement(element, index) {
  return {
    id: String(element.id || `element-${index + 1}`),
    type: String(element.type || "text").toLowerCase(),
    field: element.field ? String(element.field) : "",
    template: element.template ? String(element.template) : "",
    value: element.value !== undefined ? String(element.value) : "",
    prefix: element.prefix ? String(element.prefix) : "",
    suffix: element.suffix ? String(element.suffix) : "",
    xMm: numberOr(element.xMm, 0),
    yMm: numberOr(element.yMm, 0),
    widthMm: Math.max(0, numberOr(element.widthMm, 0)),
    heightMm: Math.max(0, numberOr(element.heightMm, 0)),
    paddingMm: Math.max(0, numberOr(element.paddingMm, 0)),
    borderColor: String(element.borderColor || ""),
    borderWidthMm: Math.max(0, numberOr(element.borderWidthMm, 0)),
    fillColor: String(element.fillColor || ""),
    fit: String(element.fit || "cover").toLowerCase(),
    font: String(element.font || "Helvetica"),
    fontSize: Math.max(1, numberOr(element.fontSize, 11)),
    minFontSize: Math.max(1, numberOr(element.minFontSize, 7)),
    lineHeight: Math.max(0.9, numberOr(element.lineHeight, 1.15)),
    color: String(element.color || "#111111"),
    align: String(element.align || "left").toLowerCase(),
    valign: String(element.valign || "top").toLowerCase(),
    uppercase: Boolean(element.uppercase),
    maxLines: Math.max(1, integerOr(element.maxLines, 99)),
    hideWhenEmpty: Boolean(element.hideWhenEmpty),
  };
}

module.exports = {
  normalizeConfig,
};

