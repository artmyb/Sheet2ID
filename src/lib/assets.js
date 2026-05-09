const path = require("path");
const { getRowValue, normalizeCompact, normalizeKey, tokenize, unique } = require("./utils");

const SUPPORTED_IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg"]);

function indexImageFiles(files = []) {
  return files
    .filter((file) => file && file.buffer && isSupportedImage(file.originalname))
    .map((file, index) => {
      const parsed = path.parse(file.originalname);
      return {
        id: `${index}-${parsed.base}`,
        originalname: parsed.base,
        relativeName: file.originalname,
        basename: parsed.name,
        extension: parsed.ext.toLowerCase(),
        normalizedBase: normalizeKey(parsed.name),
        compactBase: normalizeCompact(parsed.name),
        buffer: file.buffer,
      };
    });
}

function buildCardModels(rows, photoAssets, backgroundAssets, config) {
  return rows.map((row, index) => {
    const backgroundAsset = resolveBackgroundAsset(row, backgroundAssets, config);
    const photoAsset = resolvePhotoAsset(row, photoAssets, config);
    return {
      index,
      row,
      backgroundAsset,
      photoAsset,
    };
  });
}

function buildPreview(cards, config, headers) {
  const displayFields = unique([
    ...config.photos.fields,
    ...config.elements.filter((element) => element.type === "text" && element.field).map((element) => element.field),
    ...config.card.background.rules.map((rule) => rule.field).filter(Boolean),
  ]);

  return {
    totalRows: cards.length,
    matchedPhotos: cards.filter((card) => card.photoAsset).length,
    missingPhotos: cards.filter((card) => !card.photoAsset).length,
    matchedBackgrounds: cards.filter((card) => card.backgroundAsset).length,
    headers,
    displayFields,
    cards: cards.slice(0, 50).map((card) => ({
      rowNumber: card.row.__rowNumber,
      photo: card.photoAsset ? card.photoAsset.originalname : "",
      background: card.backgroundAsset ? card.backgroundAsset.originalname : "",
      values: Object.fromEntries(displayFields.map((field) => [field, getRowValue(card.row, field)])),
    })),
  };
}

function resolveBackgroundAsset(row, backgroundAssets, config) {
  for (const rule of config.card.background.rules) {
    if (matchesRule(row, rule)) {
      const matchedAsset = findAssetByReference(backgroundAssets, rule.background);
      if (matchedAsset) {
        return matchedAsset;
      }
    }
  }

  return findAssetByReference(backgroundAssets, config.card.background.default);
}

function resolvePhotoAsset(row, photoAssets, config) {
  if (!photoAssets.length) {
    return null;
  }

  const candidates = buildCandidates(row, config.photos.fields);
  if (!candidates.length) {
    return null;
  }

  let bestAsset = null;
  let bestScore = 0;

  for (const asset of photoAssets) {
    const score = scoreAsset(asset, candidates);
    if (score > bestScore) {
      bestScore = score;
      bestAsset = asset;
    }
  }

  return bestScore > 0 ? bestAsset : null;
}

function buildCandidates(row, fields) {
  const candidates = [];

  fields.forEach((field, index) => {
    const rawValue = getRowValue(row, field);
    if (!rawValue) {
      return;
    }

    candidates.push({
      weight: Math.max(1, 5 - index),
      normalized: normalizeKey(rawValue),
      compact: normalizeCompact(rawValue),
      tokens: tokenize(rawValue).filter((token) => token.length >= 2 || /^\d+$/.test(token)),
    });
  });

  const combined = fields.map((field) => getRowValue(row, field)).filter(Boolean).join(" ");
  if (combined) {
    candidates.push({
      weight: 2,
      normalized: normalizeKey(combined),
      compact: normalizeCompact(combined),
      tokens: tokenize(combined).filter((token) => token.length >= 2 || /^\d+$/.test(token)),
    });
  }

  return candidates.filter((candidate) => candidate.compact);
}

function scoreAsset(asset, candidates) {
  let bestCandidateScore = 0;

  for (const candidate of candidates) {
    let candidateScore = 0;

    if (asset.compactBase === candidate.compact) {
      candidateScore += 140;
    } else if (asset.normalizedBase === candidate.normalized) {
      candidateScore += 120;
    } else if (asset.compactBase.includes(candidate.compact) || candidate.compact.includes(asset.compactBase)) {
      candidateScore += 75;
    }

    const matchedTokens = candidate.tokens.filter((token) => asset.normalizedBase.includes(token));
    candidateScore += matchedTokens.length * 14;

    if (candidate.tokens.length > 1 && matchedTokens.length === candidate.tokens.length) {
      candidateScore += 35;
    }

    if (candidate.compact && asset.compactBase.startsWith(candidate.compact)) {
      candidateScore += 10;
    }

    candidateScore *= candidate.weight;
    bestCandidateScore = Math.max(bestCandidateScore, candidateScore);
  }

  return bestCandidateScore;
}

function matchesRule(row, rule) {
  const value = getRowValue(row, rule.field);
  if (!value) {
    return false;
  }

  if (rule.equals !== undefined || rule.value !== undefined) {
    return normalizeCompact(value) === normalizeCompact(rule.equals ?? rule.value);
  }

  if (rule.contains !== undefined) {
    return normalizeKey(value).includes(normalizeKey(rule.contains));
  }

  if (rule.regex) {
    try {
      return new RegExp(rule.regex, rule.flags || "i").test(value);
    } catch {
      return false;
    }
  }

  return false;
}

function findAssetByReference(assets, reference) {
  if (!reference) {
    return null;
  }

  const parsedReference = path.parse(String(reference));
  const rawReference = String(reference);
  const candidateReferences = unique([
    rawReference,
    parsedReference.base,
    parsedReference.name,
  ]);
  const compactReferences = candidateReferences.map((value) => normalizeCompact(value)).filter(Boolean);
  const normalizedReferences = candidateReferences.map((value) => normalizeKey(value)).filter(Boolean);

  const exactAsset = assets.find(
    (asset) =>
      candidateReferences.includes(asset.originalname) ||
      candidateReferences.includes(asset.relativeName) ||
      compactReferences.includes(asset.compactBase) ||
      normalizedReferences.includes(asset.normalizedBase)
  );

  if (exactAsset) {
    return exactAsset;
  }

  return assets.find(
    (asset) =>
      compactReferences.some((candidate) => asset.compactBase.includes(candidate) || candidate.includes(asset.compactBase)) ||
      normalizedReferences.some((candidate) => asset.normalizedBase.includes(candidate) || candidate.includes(asset.normalizedBase))
  ) || null;
}

function isSupportedImage(filename) {
  return SUPPORTED_IMAGE_EXTENSIONS.has(path.extname(filename || "").toLowerCase());
}

module.exports = {
  buildCardModels,
  buildPreview,
  indexImageFiles,
};
