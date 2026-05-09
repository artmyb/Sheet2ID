const { rgb } = require("pdf-lib");

const MM_TO_PT = 72 / 25.4;

function mmToPt(value) {
  return Number(value || 0) * MM_TO_PT;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function normalizeKey(value) {
  return transliterateForMatching(value)
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function transliterateForMatching(value) {
  const turkishAsciiMap = {
    "Ç": "C",
    "ç": "c",
    "Ğ": "G",
    "ğ": "g",
    "İ": "I",
    "ı": "i",
    "Ö": "O",
    "ö": "o",
    "Ş": "S",
    "ş": "s",
    "Ü": "U",
    "ü": "u",
  };

  return String(value ?? "").replace(/[ÇçĞğİıÖöŞşÜü]/g, (character) => turkishAsciiMap[character] || character);
}

function normalizeCompact(value) {
  return normalizeKey(value).replace(/\s+/g, "");
}

function tokenize(value) {
  return normalizeKey(value)
    .split(" ")
    .map((part) => part.trim())
    .filter(Boolean);
}

function coerceText(value) {
  return String(value ?? "").trim();
}

function getRowValue(row, field) {
  if (!field) {
    return "";
  }

  if (Object.prototype.hasOwnProperty.call(row, field)) {
    return coerceText(row[field]);
  }

  const normalized = normalizeKey(field);
  if (row.__lookup && Object.prototype.hasOwnProperty.call(row.__lookup, normalized)) {
    return coerceText(row.__lookup[normalized]);
  }

  return "";
}

function renderTemplate(template, row) {
  return String(template ?? "").replace(/{{\s*([^}]+?)\s*}}/g, (_, fieldName) => getRowValue(row, fieldName));
}

function hexToRgb(hex, fallback = "#111111") {
  const cleaned = String(hex || fallback).trim().replace(/^#/, "");
  const normalized =
    cleaned.length === 3
      ? cleaned
          .split("")
          .map((char) => char + char)
          .join("")
      : cleaned;

  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return rgb(0.066, 0.066, 0.066);
  }

  const red = parseInt(normalized.slice(0, 2), 16) / 255;
  const green = parseInt(normalized.slice(2, 4), 16) / 255;
  const blue = parseInt(normalized.slice(4, 6), 16) / 255;
  return rgb(red, green, blue);
}

function numberOr(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function integerOr(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function deepMerge(base, override) {
  if (Array.isArray(base) || Array.isArray(override)) {
    return Array.isArray(override) ? override : Array.isArray(base) ? base : [];
  }

  if (!isPlainObject(base) || !isPlainObject(override)) {
    return override === undefined ? base : override;
  }

  const merged = { ...base };
  for (const [key, value] of Object.entries(override)) {
    merged[key] = key in base ? deepMerge(base[key], value) : value;
  }
  return merged;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

module.exports = {
  clamp,
  coerceText,
  deepMerge,
  getRowValue,
  hexToRgb,
  integerOr,
  isPlainObject,
  mmToPt,
  normalizeCompact,
  normalizeKey,
  numberOr,
  renderTemplate,
  tokenize,
  unique,
};
