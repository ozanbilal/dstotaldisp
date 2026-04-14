import assert from "node:assert/strict";
import { buildSourceViewerScene } from "../web-ui/viewer.mjs";
import { buildSummaryViewerScene } from "../web-ui/summary_viewer.mjs";
import {
  buildViewerUrlSearch,
  openDetailDrawer,
  parseViewerUrlState,
  resolveScrollBehavior,
  resolveSummaryPlotState,
} from "../web-ui/viewer_shell_state.mjs";

const sourceCatalog = [
  {
    sourceId: "source-a",
    sourceLabel: "Input A",
    sourceKind: "single",
    axis: "X",
    pairKey: "PAIR|A",
    artifactPairKeys: ["PAIR|A"],
    families: [
      {
        familyKey: "input-motion",
        familyLabel: "Input Motion",
        chartType: "time",
        supportsOverlay: true,
        supportsLayerSelection: false,
        defaultVisibleSeries: ["trace-a"],
        layers: [],
        charts: [
          {
            chartKey: "input-acceleration",
            chartLabel: "Input Acceleration",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Acceleration (g)",
            invertY: false,
            series: [{ seriesKey: "trace-a", name: "Trace A", points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] }],
          },
        ],
      },
      {
        familyKey: "layer-series",
        familyLabel: "Layer Series",
        chartType: "time",
        supportsOverlay: false,
        supportsLayerSelection: true,
        defaultVisibleSeries: ["layer-a"],
        layers: [{ layerIndex: 0, layerLabel: "Layer 1", depth: 0 }],
        charts: [
          {
            chartKey: "layer-displacement",
            chartLabel: "Layer Displacement",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Displacement (m)",
            invertY: false,
            layerViews: [
              {
                layerIndex: 0,
                layerLabel: "Layer 1",
                depth: 0,
                series: [{ seriesKey: "layer-a", name: "Layer A", points: [{ x: 0, y: 0 }, { x: 1, y: 0.2 }] }],
              },
            ],
          },
        ],
      },
    ],
  },
];

const summaryCatalog = [
  {
    summaryId: "pair-1",
    summaryLabel: "Pair 1",
    summaryKind: "pair",
    axis: "XY",
    pairKey: "PAIR|A",
    inputKind: "xlsx",
    preferredVariantKey: "strain_input_total",
    detailSourceIds: ["source-a"],
    artifactPairKeys: ["PAIR|A"],
    warnings: [],
    coverage: { availableLayerCount: 1, profileLayerCount: 1, limitedData: false, label: "Tam veri" },
    variants: [
      {
        variantKey: "strain_input_total",
        displayLabel: "Strain + Input Proxy",
        methodClass: "computed",
        valid: true,
        confidenceRank: 20,
        confidenceLabel: "Primary",
        reason: "best",
        depths: [0.5],
        values: [0.1],
      },
    ],
  },
];

const parsed = parseViewerUrlState(
  "?shell=legacy&summary=pair-1&source=source-a&family=layer-series&chart=layer-displacement&layer=2"
);
assert.equal(parsed.hasAny, true);
assert.equal(parsed.shellMode, "legacy");
assert.equal(parsed.viewerPrefs.activeSummaryId, "pair-1");
assert.equal(parsed.viewerPrefs.activeSourceId, "source-a");
assert.equal(parsed.viewerPrefs.activeFamilyKey, "layer-series");
assert.equal(parsed.viewerPrefs.activeChartKey, "layer-displacement");
assert.equal(parsed.viewerPrefs.activeLayerIndex, 2);
assert.equal(parsed.viewerPrefs.chartMode, "focus");
assert.deepEqual(parsed.viewerPrefs.compareSourceIds, []);

const rebuilt = buildViewerUrlSearch({
  shellMode: parsed.shellMode,
  activeSummaryId: parsed.viewerPrefs.activeSummaryId,
  activeSourceId: parsed.viewerPrefs.activeSourceId,
  activeFamilyKey: parsed.viewerPrefs.activeFamilyKey,
  activeChartKey: parsed.viewerPrefs.activeChartKey,
  activeLayerIndex: parsed.viewerPrefs.activeLayerIndex,
});
assert.equal(
  rebuilt,
  "?shell=legacy&summary=pair-1&source=source-a&family=layer-series&chart=layer-displacement&layer=2"
);

const invalid = parseViewerUrlState("?shell=viewer&summary=missing&source=missing&family=bad&chart=bad&layer=-2");
const fallbackScene = buildSourceViewerScene(sourceCatalog, invalid.viewerPrefs);
assert.equal(fallbackScene.activeSourceId, "source-a");
assert.equal(fallbackScene.activeFamilyKey, "input-motion");
assert.equal(fallbackScene.activeChartKey, "input-acceleration");
assert.equal(fallbackScene.activeLayerIndex, 0);

const fallbackSummaryScene = buildSummaryViewerScene(summaryCatalog, invalid.viewerPrefs);
assert.equal(fallbackSummaryScene.activeSummaryId, "pair-1");

const emptyPlotState = resolveSummaryPlotState(null, [], true);
assert.equal(emptyPlotState.isEmpty, true);
assert.equal(emptyPlotState.message, "Özet grafik burada görünecek.");

const fullPlotState = resolveSummaryPlotState(summaryCatalog[0], summaryCatalog[0].variants, true);
assert.equal(fullPlotState.isEmpty, false);

assert.equal(resolveScrollBehavior(true), "auto");
assert.equal(resolveScrollBehavior(false), "smooth");

const detailPanel = { open: false };
assert.equal(openDetailDrawer(detailPanel), true);
assert.equal(detailPanel.open, true);

console.log("viewer shell state test passed");
