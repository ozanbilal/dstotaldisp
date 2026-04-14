const SUMMARY_KIND_ORDER = new Map([
  ["pair", 10],
  ["single", 20],
  ["db_pair", 30],
  ["db_single", 40],
  ["other", 90],
]);

function toSummarySlug(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "summary";
}

function cleanNumberList(values) {
  return (Array.isArray(values) ? values : []).map((value) => {
    const num = Number(value);
    return Number.isFinite(num) ? num : null;
  });
}

function normalizeVariant(variant, index = 0) {
  const depths = cleanNumberList(variant?.depths || []);
  const values = cleanNumberList(variant?.values || []);
  const count = Math.min(depths.length, values.length);
  const points = [];
  for (let idx = 0; idx < count; idx += 1) {
    const x = values[idx];
    const y = depths[idx];
    if (x == null || y == null) continue;
    points.push({ x, y });
  }
  return {
    variantKey: String(variant?.variantKey || `variant-${index}`),
    displayLabel: String(variant?.displayLabel || variant?.variantKey || `Variant ${index + 1}`),
    methodClass: String(variant?.methodClass || "derived"),
    confidenceRank: Number.isFinite(Number(variant?.confidenceRank)) ? Number(variant.confidenceRank) : 999,
    confidenceLabel: String(variant?.confidenceLabel || ""),
    valid: !!variant?.valid && points.length > 0,
    rawValid: !!variant?.valid,
    reason: String(variant?.reason || ""),
    sourceRefs: (Array.isArray(variant?.sourceRefs) ? variant.sourceRefs : []).map((item) => String(item || "")).filter(Boolean),
    depths: depths.slice(0, count),
    values: values.slice(0, count),
    points,
  };
}

function normalizeSummaryCatalog(summaryCatalog) {
  return (Array.isArray(summaryCatalog) ? summaryCatalog : [])
    .map((entry, index) => {
      const variants = (Array.isArray(entry?.variants) ? entry.variants : [])
        .map((variant, variantIndex) => normalizeVariant(variant, variantIndex))
        .sort((left, right) => left.confidenceRank - right.confidenceRank || left.displayLabel.localeCompare(right.displayLabel));
      if (!variants.length) return null;
      const summaryKind = String(entry?.summaryKind || "other");
      const validVariantCount = variants.filter((variant) => variant.valid).length;
      const preferredVariantKey = String(entry?.preferredVariantKey || "");
      return {
        summaryId: String(entry?.summaryId || `summary-${index}-${toSummarySlug(entry?.summaryLabel || entry?.pairKey || "")}`),
        summaryLabel: String(entry?.summaryLabel || entry?.pairKey || `Summary ${index + 1}`),
        sourceSystem: String(entry?.sourceSystem || "deepsoil"),
        summaryKind,
        axis: String(entry?.axis || ""),
        pairKey: String(entry?.pairKey || ""),
        inputKind: String(entry?.inputKind || ""),
        preferredVariantKey,
        validVariantCount: validVariantCount || Number(entry?.validVariantCount || 0),
        detailSourceIds: (Array.isArray(entry?.detailSourceIds) ? entry.detailSourceIds : []).map((item) => String(item || "")).filter(Boolean),
        artifactPairKeys: (Array.isArray(entry?.artifactPairKeys) ? entry.artifactPairKeys : []).map((item) => String(item || "")).filter(Boolean),
        warnings: (Array.isArray(entry?.warnings) ? entry.warnings : []).map((item) => String(item || "")).filter(Boolean),
        coverage: {
          availableLayerCount: Number(entry?.coverage?.availableLayerCount || 0),
          profileLayerCount: Number(entry?.coverage?.profileLayerCount || 0),
          limitedData: !!entry?.coverage?.limitedData,
          label: String(entry?.coverage?.label || ""),
        },
        variants,
      };
    })
    .filter(Boolean)
    .sort((left, right) => {
      const orderLeft = SUMMARY_KIND_ORDER.get(left.summaryKind) || 90;
      const orderRight = SUMMARY_KIND_ORDER.get(right.summaryKind) || 90;
      return orderLeft - orderRight || left.summaryLabel.localeCompare(right.summaryLabel);
    });
}

function buildSummaryViewerScene(summaryCatalog, selection = {}) {
  const summaries = normalizeSummaryCatalog(summaryCatalog);
  const variantVisibilityMap = selection?.variantVisibilityMap && typeof selection.variantVisibilityMap === "object"
    ? selection.variantVisibilityMap
    : {};
  const primaryVariantMap = selection?.primaryVariantMap && typeof selection.primaryVariantMap === "object"
    ? selection.primaryVariantMap
    : {};
  if (!summaries.length) {
    return {
      summaries,
      summaryCount: 0,
      activeSummary: null,
      activeSummaryId: "",
      activeSummaryIndex: -1,
      validVariants: [],
      invalidVariants: [],
      visibleVariants: [],
      primaryVariant: null,
      variantVisibilityMap,
      primaryVariantMap,
    };
  }

  const requestedSummaryId = String(selection?.activeSummaryId || "").trim();
  const activeSummary = summaries.find((item) => item.summaryId === requestedSummaryId) || summaries[0];
  const activeSummaryIndex = summaries.findIndex((item) => item.summaryId === activeSummary.summaryId);
  const validVariants = activeSummary.variants.filter((variant) => variant.valid);
  const invalidVariants = activeSummary.variants.filter((variant) => !variant.valid);
  const visibleVariants = validVariants.filter((variant) => {
    const key = `${activeSummary.summaryId}::${variant.variantKey}`;
    return variantVisibilityMap[key] !== false;
  });
  const requestedPrimaryVariantKey = String(primaryVariantMap?.[activeSummary.summaryId] || "").trim();
  const primaryVariant =
    visibleVariants.find((variant) => variant.variantKey === requestedPrimaryVariantKey) ||
    visibleVariants.find((variant) => variant.variantKey === activeSummary.preferredVariantKey) ||
    visibleVariants[0] ||
    validVariants.find((variant) => variant.variantKey === requestedPrimaryVariantKey) ||
    validVariants.find((variant) => variant.variantKey === activeSummary.preferredVariantKey) ||
    validVariants[0] ||
    null;

  return {
    summaries,
    summaryCount: summaries.length,
    activeSummary,
    activeSummaryId: activeSummary.summaryId,
    activeSummaryIndex,
    validVariants,
    invalidVariants,
    visibleVariants,
    primaryVariant,
    variantVisibilityMap,
    primaryVariantMap,
  };
}

export {
  buildSummaryViewerScene,
  normalizeSummaryCatalog,
};
