function cleanToken(value) {
  return String(value || "").trim();
}

function normalizeShellModeToken(value) {
  const normalized = cleanToken(value).toLowerCase();
  return normalized === "legacy" ? "legacy" : normalized === "viewer" ? "viewer" : "";
}

export function parseViewerUrlState(search = "") {
  const raw = String(search || "");
  const params = new URLSearchParams(raw.startsWith("?") ? raw.slice(1) : raw);
  const shellMode = normalizeShellModeToken(params.get("shell"));
  const activeSummaryId = cleanToken(params.get("summary"));
  const activeSourceId = cleanToken(params.get("source"));
  const activeFamilyKey = cleanToken(params.get("family"));
  const activeChartKey = cleanToken(params.get("chart"));
  const rawLayer = cleanToken(params.get("layer"));
  const parsedLayer = rawLayer ? Number(rawLayer) : null;
  const activeLayerIndex =
    Number.isFinite(parsedLayer) && parsedLayer >= 0 ? Math.max(0, Math.floor(parsedLayer)) : null;
  const hasViewerRoute = !!(activeSourceId || activeFamilyKey || activeChartKey || rawLayer);

  return {
    hasAny:
      params.has("shell") ||
      params.has("summary") ||
      params.has("source") ||
      params.has("family") ||
      params.has("chart") ||
      params.has("layer"),
    shellMode: shellMode || null,
    activeSummaryId: activeSummaryId || "",
    viewerPrefs: {
      ...(hasViewerRoute ? { chartMode: "focus", compareSourceIds: [], seriesVisibilityMap: {} } : {}),
      ...(activeSummaryId ? { activeSummaryId } : {}),
      ...(activeSourceId ? { activeSourceId } : {}),
      ...(activeFamilyKey ? { activeFamilyKey } : {}),
      ...(activeChartKey ? { activeChartKey } : {}),
      ...(activeLayerIndex != null ? { activeLayerIndex } : {}),
    },
  };
}

export function buildViewerUrlSearch(state = {}) {
  const params = new URLSearchParams();
  const shellMode = normalizeShellModeToken(state.shellMode) || "viewer";
  params.set("shell", shellMode);

  const activeSummaryId = cleanToken(state.activeSummaryId);
  const activeSourceId = cleanToken(state.activeSourceId);
  const activeFamilyKey = cleanToken(state.activeFamilyKey);
  const activeChartKey = cleanToken(state.activeChartKey);
  const layerIndex = Number(state.activeLayerIndex);

  if (activeSummaryId) params.set("summary", activeSummaryId);
  if (activeSourceId) params.set("source", activeSourceId);
  if (activeFamilyKey) params.set("family", activeFamilyKey);
  if (activeChartKey) params.set("chart", activeChartKey);
  if (activeSourceId && Number.isFinite(layerIndex) && layerIndex >= 0) params.set("layer", String(Math.max(0, Math.floor(layerIndex))));

  const query = params.toString();
  return query ? `?${query}` : "";
}

export function resolveSummaryPlotState(activeSummary, visibleVariants, plotlyReady) {
  const visibleCount = Array.isArray(visibleVariants) ? visibleVariants.length : 0;
  const isEmpty = !activeSummary || visibleCount === 0 || !plotlyReady;
  return {
    isEmpty,
    message: activeSummary ? "Bu kayıtta görünür geçerli yöntem yok." : "Özet grafik burada görünecek.",
  };
}

export function prefersReducedMotion(matchMediaImpl = globalThis.matchMedia) {
  if (typeof matchMediaImpl !== "function") return false;
  try {
    return !!matchMediaImpl("(prefers-reduced-motion: reduce)")?.matches;
  } catch {
    return false;
  }
}

export function resolveScrollBehavior(reducedMotion) {
  return reducedMotion ? "auto" : "smooth";
}

export function openDetailDrawer(panel) {
  if (!panel) return false;
  panel.open = true;
  return true;
}
