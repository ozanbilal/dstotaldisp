function cleanToken(value) {
  return String(value || "").trim();
}

function pluralize(count, singular, plural = `${singular}s`) {
  const safeCount = Number.isFinite(Number(count)) ? Number(count) : 0;
  return `${safeCount} ${safeCount === 1 ? singular : plural}`;
}

export function buildSpectrumScopeKey(scene = {}) {
  const activeSourceId = cleanToken(scene.activeSourceId);
  const activeFamilyKey = cleanToken(scene.activeFamilyKey);
  const activeChartKey = cleanToken(scene.activeChartKey);
  if (!activeSourceId || !activeFamilyKey || !activeChartKey) return "";
  const activeLayerIndex = Number(scene.activeLayerIndex);
  return [
    activeSourceId,
    activeFamilyKey,
    activeChartKey,
    Number.isFinite(activeLayerIndex) ? Math.max(0, Math.floor(activeLayerIndex)) : 0,
  ].join("::");
}

export function getScopedSpectrumPeriodMax(spectrumPeriodMaxMap = {}, scene = {}) {
  const scopeKey = buildSpectrumScopeKey(scene);
  if (!scopeKey) return "";
  return cleanToken(spectrumPeriodMaxMap?.[scopeKey]);
}

export function summarizeViewerSourceKinds(sources = []) {
  const sourceKinds = {};
  (Array.isArray(sources) ? sources : []).forEach((source) => {
    const key = cleanToken(source?.sourceKind || "source") || "source";
    sourceKinds[key] = (sourceKinds[key] || 0) + 1;
  });

  return {
    sourceKinds,
    pairViewCount: Number(sourceKinds.pair || 0) + Number(sourceKinds.db_pair || 0),
    singleDirectionCount: Number(sourceKinds.single || 0) + Number(sourceKinds.db_single || 0),
    aggregateViewCount: Number(sourceKinds.method3_aggregate || 0),
  };
}

export function formatViewerSourceMeta({ sources = [], sourceCount = null, fileCount = 0 } = {}) {
  const safeSourceCount = Number.isFinite(Number(sourceCount))
    ? Number(sourceCount)
    : Array.isArray(sources)
      ? sources.length
      : 0;
  const { pairViewCount, singleDirectionCount, aggregateViewCount } = summarizeViewerSourceKinds(sources);
  const parts = [pluralize(safeSourceCount, "viewer source")];

  if (Number(fileCount) > 0) parts.push(pluralize(Number(fileCount), "input file"));
  if (pairViewCount > 0) parts.push(pluralize(pairViewCount, "pair view"));
  if (singleDirectionCount > 0) {
    parts.push(pluralize(singleDirectionCount, "single-direction source", "single-direction sources"));
  }
  if (aggregateViewCount > 0) parts.push(pluralize(aggregateViewCount, "aggregate view"));

  return parts.join(" · ");
}
