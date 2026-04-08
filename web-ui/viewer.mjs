const SHELL_MODES = new Set(["viewer", "legacy"]);
const GROUP_ORDER = new Map([
  ["Primary Outputs", 10],
  ["Method-2", 20],
  ["Method-3 Aggregate", 30],
  ["DB Direct", 40],
  ["Other Outputs", 90],
]);
const VIEWER_CHART_MODES = new Set(["focus", "compare"]);

function toViewerSlug(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 64) || "item";
}

function getViewerResultFocusId(result, index = 0) {
  const groupPart = toViewerSlug(result?.viewerGroup || "output");
  const filePart = toViewerSlug(result?.outputFileName || result?.pairKey || result?.viewerKind || "result");
  const indexPart = Number.isFinite(Number(result?.viewerIndex)) ? Number(result.viewerIndex) : Number(index) || 0;
  return `viewer-result-${groupPart}-${filePart}-${indexPart}`;
}

function getRenderablePreviewCharts(result) {
  return (Array.isArray(result?.previewCharts) ? result.previewCharts : []).filter(
    (chart) => Array.isArray(chart?.series) && chart.series.some((series) => Array.isArray(series?.points) && series.points.length)
  );
}

function normalizeShellMode(mode) {
  const value = String(mode || "").trim().toLowerCase();
  return SHELL_MODES.has(value) ? value : "viewer";
}

function normalizeViewerChartMode(mode) {
  const value = String(mode || "").trim().toLowerCase();
  if (value === "stack") return "focus";
  return VIEWER_CHART_MODES.has(value) ? value : "focus";
}

function deriveViewerGroup(result) {
  const metrics = result?.metrics || {};
  const explicitGroup = String(result?.viewerGroup || "").trim();
  const explicitKind = String(result?.viewerKind || "").trim();
  const explicitOrder = Number(result?.viewerGroupOrder);

  const mode = String(metrics.mode || "").trim().toLowerCase();
  const axis = String(metrics.axis || "").trim().toUpperCase();

  if (explicitGroup) {
    return {
      viewerGroup: explicitGroup,
      viewerKind: explicitKind || explicitGroup,
      viewerGroupOrder: Number.isFinite(explicitOrder) ? explicitOrder : (GROUP_ORDER.get(explicitGroup) || 90),
    };
  }

  if (mode.startsWith("db_")) {
    let viewerKind = "DB Output";
    if (mode === "db_pair") viewerKind = "DB Pair";
    else if (mode === "db_single") viewerKind = "DB Single";
    else if (mode === "db_method2_single") viewerKind = "DB Method-2";
    else if (mode === "db_method3" || mode === "db_method3_aggregate") viewerKind = "DB Method-3";
    return {
      viewerGroup: "DB Direct",
      viewerKind,
      viewerGroupOrder: mode === "db_method2_single" ? 41 : mode === "db_method3_aggregate" ? 42 : 40,
    };
  }

  if (mode === "method2_single") {
    return {
      viewerGroup: "Method-2",
      viewerKind: axis ? `Method-2 ${axis}` : "Method-2",
      viewerGroupOrder: 20,
    };
  }

  if (mode === "method3" || mode === "method3_aggregate") {
    return {
      viewerGroup: "Method-3 Aggregate",
      viewerKind: "Method-3",
      viewerGroupOrder: 30,
    };
  }

  if (mode === "pair") {
    return {
      viewerGroup: "Primary Outputs",
      viewerKind: "Pair",
      viewerGroupOrder: 10,
    };
  }

  if (mode === "single") {
    return {
      viewerGroup: "Primary Outputs",
      viewerKind: "Single",
      viewerGroupOrder: 11,
    };
  }

  return {
    viewerGroup: "Other Outputs",
    viewerKind: mode || "Output",
    viewerGroupOrder: 90,
  };
}

function classifyViewerResult(result, index = 0) {
  const group = deriveViewerGroup(result);
  return {
    ...(result || {}),
    viewerGroup: group.viewerGroup,
    viewerKind: group.viewerKind,
    viewerGroupOrder: group.viewerGroupOrder,
    viewerIndex: index,
  };
}

function groupViewerResults(results) {
  const grouped = new Map();
  (Array.isArray(results) ? results : []).forEach((result, index) => {
    const item = classifyViewerResult(result, index);
    const key = item.viewerGroup || "Other Outputs";
    if (!grouped.has(key)) {
      grouped.set(key, {
        key,
        label: key,
        order: item.viewerGroupOrder ?? GROUP_ORDER.get(key) ?? 90,
        chartCount: 0,
        results: [],
      });
    }

    const bucket = grouped.get(key);
    bucket.results.push(item);
    bucket.chartCount += Array.isArray(item.previewCharts)
      ? item.previewCharts.filter((chart) => Array.isArray(chart?.series) && chart.series.length).length
      : 0;
  });

  return [...grouped.values()]
    .sort((a, b) => a.order - b.order || a.label.localeCompare(b.label))
    .map((group) => ({
      ...group,
      results: [...group.results].sort((a, b) => {
        const left = String(a.outputFileName || a.pairKey || "");
        const right = String(b.outputFileName || b.pairKey || "");
        return left.localeCompare(right);
      }),
    }));
}

function summarizeViewerResults(results) {
  const groups = groupViewerResults(results);
  const totalCharts = groups.reduce((sum, group) => sum + group.chartCount, 0);
  const totalResults = groups.reduce((sum, group) => sum + group.results.length, 0);
  return { groups, totalCharts, totalResults };
}

function normalizeSourceCatalog(sourceCatalog) {
  return (Array.isArray(sourceCatalog) ? sourceCatalog : [])
    .map((entry, index) => {
      const sourceId = String(entry?.sourceId || `source-${index}`);
      const sourceLabel = String(entry?.sourceLabel || sourceId);
      const families = (Array.isArray(entry?.families) ? entry.families : [])
        .map((family) => {
          const charts = (Array.isArray(family?.charts) ? family.charts : []).filter((chart) => {
            if (!chart || typeof chart !== "object") return false;
            const series = Array.isArray(chart.series) ? chart.series.filter((item) => Array.isArray(item?.points) && item.points.length) : [];
            const layerViews = Array.isArray(chart.layerViews)
              ? chart.layerViews
                  .map((view) => ({
                    ...view,
                    series: (Array.isArray(view?.series) ? view.series : []).filter(
                      (item) => Array.isArray(item?.points) && item.points.length
                    ),
                  }))
                  .filter((view) => view.series.length)
              : [];
            return series.length > 0 || layerViews.length > 0;
          });
          return charts.length
            ? {
                familyKey: String(family?.familyKey || `${sourceId}-family`),
                familyLabel: String(family?.familyLabel || family?.familyKey || "Family"),
                chartType: String(family?.chartType || "time"),
                supportsOverlay: !!family?.supportsOverlay,
                supportsLayerSelection: !!family?.supportsLayerSelection,
                defaultVisibleSeries: Array.isArray(family?.defaultVisibleSeries)
                  ? family.defaultVisibleSeries.map((item) => String(item || "")).filter(Boolean)
                  : [],
                layers: Array.isArray(family?.layers) ? family.layers : [],
                charts,
              }
            : null;
        })
        .filter(Boolean);
      return families.length
        ? {
            sourceId,
            sourceLabel,
            sourceKind: String(entry?.sourceKind || "source"),
            axis: String(entry?.axis || ""),
            pairKey: String(entry?.pairKey || ""),
            artifactPairKeys: Array.isArray(entry?.artifactPairKeys) ? entry.artifactPairKeys : [],
            families,
          }
        : null;
    })
    .filter(Boolean);
}

function summarizeSourceCatalog(sourceCatalog) {
  const sources = normalizeSourceCatalog(sourceCatalog);
  const familyCount = sources.reduce((sum, source) => sum + source.families.length, 0);
  const chartCount = sources.reduce(
    (sum, source) => sum + source.families.reduce((familySum, family) => familySum + family.charts.length, 0),
    0
  );
  return {
    sources,
    sourceCount: sources.length,
    familyCount,
    chartCount,
  };
}

function findFamily(source, familyKey) {
  return (source?.families || []).find((family) => family.familyKey === familyKey) || null;
}

function findChart(source, familyKey, chartKey) {
  const family = findFamily(source, familyKey);
  return family?.charts?.find((chart) => chart.chartKey === chartKey) || null;
}

function getChartSeries(chart, layerIndex = 0) {
  if (!chart) return { layerIndex: 0, layerLabel: "", layerCount: 0, series: [] };
  if (Array.isArray(chart.layerViews) && chart.layerViews.length) {
    const safeIndex = Math.max(0, Math.min(Number(layerIndex) || 0, chart.layerViews.length - 1));
    const layerView = chart.layerViews[safeIndex];
    return {
      layerIndex: safeIndex,
      layerLabel: String(layerView?.layerLabel || `Layer ${safeIndex + 1}`),
      layerCount: chart.layerViews.length,
      series: Array.isArray(layerView?.series) ? layerView.series : [],
    };
  }
  return {
    layerIndex: 0,
    layerLabel: "",
    layerCount: 0,
    series: Array.isArray(chart.series) ? chart.series : [],
  };
}

function buildSeriesEntry({
  source,
  family,
  chart,
  layerIndex,
  layerLabel,
  series,
  compareMode,
  isActiveSource,
  seriesVisibilityMap,
}) {
  const entryKey = `${source.sourceId}::${chart.chartKey}::${layerIndex}::${series.seriesKey || series.name}`;
  const explicitVisible = Object.prototype.hasOwnProperty.call(seriesVisibilityMap || {}, entryKey)
    ? !!seriesVisibilityMap[entryKey]
    : null;
  const visible = explicitVisible == null ? true : explicitVisible;
  return {
    entryKey,
    sourceId: source.sourceId,
    sourceLabel: source.sourceLabel,
    familyKey: family.familyKey,
    chartKey: chart.chartKey,
    layerIndex,
    layerLabel,
    seriesKey: String(series.seriesKey || series.name || "series"),
    seriesLabel: String(series.name || "Series"),
    displayLabel: compareMode && !isActiveSource ? `${source.sourceLabel} · ${series.name}` : String(series.name || "Series"),
    points: Array.isArray(series.points) ? series.points : [],
    visible,
  };
}

function buildPlotSeriesEntries(activeSource, activeFamily, activeChart, activeLayerIndex, compareSources, chartMode, seriesVisibilityMap) {
  if (!activeSource || !activeFamily || !activeChart) return [];

  const allSources = [activeSource, ...(chartMode === "compare" ? compareSources : [])];
  const compareMode = chartMode === "compare" && compareSources.length > 0;
  const entries = [];

  allSources.forEach((source, sourceIndex) => {
    const family = source.sourceId === activeSource.sourceId ? activeFamily : findFamily(source, activeFamily.familyKey);
    const chart = source.sourceId === activeSource.sourceId ? activeChart : findChart(source, activeFamily.familyKey, activeChart.chartKey);
    if (!family || !chart) return;
    const chartState = getChartSeries(chart, activeLayerIndex);
    (chartState.series || []).forEach((series) => {
      entries.push(
        buildSeriesEntry({
          source,
          family,
          chart,
          layerIndex: chartState.layerIndex,
          layerLabel: chartState.layerLabel,
          series,
          compareMode,
          isActiveSource: sourceIndex === 0,
          seriesVisibilityMap,
        })
      );
    });
  });

  return entries;
}

function buildSourceViewerScene(sourceCatalog, selection = {}) {
  const summary = summarizeSourceCatalog(sourceCatalog);
  const sources = summary.sources;
  const chartMode = normalizeViewerChartMode(selection.chartMode);
  const seriesVisibilityMap = selection?.seriesVisibilityMap && typeof selection.seriesVisibilityMap === "object"
    ? selection.seriesVisibilityMap
    : {};

  if (!sources.length) {
    return {
      ...summary,
      chartMode,
      activeSource: null,
      activeSourceId: "",
      activeSourceIndex: -1,
      activeFamily: null,
      activeFamilyKey: "",
      activeChart: null,
      activeChartKey: "",
      activeLayerIndex: 0,
      activeLayerLabel: "",
      activeLayerCount: 0,
      compareCandidates: [],
      compareSources: [],
      compareSourceIds: [],
      plotSeriesEntries: [],
      seriesVisibilityMap,
    };
  }

  const requestedSourceId = String(selection?.activeSourceId || "").trim();
  const activeSource = sources.find((source) => source.sourceId === requestedSourceId) || sources[0];
  const activeSourceIndex = sources.findIndex((source) => source.sourceId === activeSource.sourceId);

  const requestedFamilyKey = String(selection?.activeFamilyKey || "").trim();
  const requestedFamily = findFamily(activeSource, requestedFamilyKey);
  const activeFamily = requestedFamily || activeSource.families[0] || null;
  const familyFallback = !!requestedFamilyKey && !requestedFamily;
  const requestedChartKey = String(selection?.activeChartKey || "").trim();
  const requestedChart = activeFamily?.charts?.find((chart) => chart.chartKey === requestedChartKey) || null;
  const activeChart = requestedChart || activeFamily?.charts?.[0] || null;
  const chartFallback = !!requestedChartKey && !requestedChart;
  const chartState = getChartSeries(activeChart, selection?.activeLayerIndex);

  const compareCandidates = sources.filter((source) => {
    if (source.sourceId === activeSource.sourceId) return false;
    const family = findFamily(source, activeFamily?.familyKey || "");
    const chart = findChart(source, activeFamily?.familyKey || "", activeChart?.chartKey || "");
    return !!family && !!chart;
  });

  const compareSourceIds = Array.isArray(selection?.compareSourceIds)
    ? selection.compareSourceIds.map((value) => String(value || "")).filter((value) => compareCandidates.some((item) => item.sourceId === value))
    : [];
  const compareSources = compareCandidates.filter((source) => compareSourceIds.includes(source.sourceId));
  const plotSeriesEntries = buildPlotSeriesEntries(
    activeSource,
    activeFamily,
    activeChart,
    chartState.layerIndex,
    compareSources,
    chartMode,
    seriesVisibilityMap
  );

  return {
    ...summary,
    chartMode,
    activeSource,
    activeSourceId: activeSource.sourceId,
    activeSourceIndex,
    activeFamily,
    activeFamilyKey: activeFamily?.familyKey || "",
    familyFallback,
    activeChart,
    activeChartKey: activeChart?.chartKey || "",
    chartFallback,
    activeLayerIndex: chartState.layerIndex,
    activeLayerLabel: chartState.layerLabel,
    activeLayerCount: chartState.layerCount,
    compareCandidates,
    compareSources,
    compareSourceIds,
    plotSeriesEntries,
    seriesVisibilityMap,
  };
}

function cycleSceneItem(items, currentId, delta, getId = (item) => item) {
  const list = Array.isArray(items) ? items : [];
  if (!list.length) return null;
  const currentIndex = Math.max(
    0,
    list.findIndex((item) => getId(item) === currentId)
  );
  const nextIndex = (currentIndex + delta + list.length) % list.length;
  return list[nextIndex];
}

export {
  buildSourceViewerScene,
  classifyViewerResult,
  cycleSceneItem,
  getViewerResultFocusId,
  groupViewerResults,
  normalizeShellMode,
  normalizeSourceCatalog,
  normalizeViewerChartMode,
  summarizeSourceCatalog,
  summarizeViewerResults,
};
