import assert from "node:assert/strict";
import {
  buildSourceViewerScene,
  cycleSceneItem,
  groupViewerResults,
  normalizeShellMode,
  normalizeViewerChartMode,
  summarizeViewerResults,
} from "../web-ui/viewer.mjs";

const sampleResults = [
  {
    outputFileName: "pair.xlsx",
    metrics: { mode: "pair" },
    previewCharts: [{ title: "Pair", sheetName: "Sheet1", series: [{ name: "X", points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] }] }],
  },
  {
    outputFileName: "method2_x.xlsx",
    metrics: { mode: "method2_single", axis: "X" },
    previewCharts: [{ title: "Method2", sheetName: "Sheet2", series: [{ name: "X", points: [{ x: 0, y: 1 }, { x: 1, y: 2 }] }] }],
  },
  {
    outputFileName: "method3.xlsx",
    metrics: { mode: "method3_aggregate" },
    previewCharts: [{ title: "Method3", sheetName: "Sheet3", series: [{ name: "RSS", points: [{ x: 0, y: 2 }, { x: 1, y: 3 }] }] }],
  },
];

const sampleSourceCatalog = [
  {
    sourceId: "source-a",
    sourceLabel: "Input A",
    sourceKind: "single",
    axis: "X",
    pairKey: "METHOD2|A",
    artifactPairKeys: ["METHOD2|A"],
    families: [
      {
        familyKey: "input-motion",
        familyLabel: "Input Motion",
        chartType: "time",
        supportsOverlay: true,
        supportsLayerSelection: false,
        defaultVisibleSeries: ["input-a"],
        layers: [],
        charts: [
          {
            chartKey: "input-acceleration",
            chartLabel: "Input Acceleration",
            sheetName: "Input Motion",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Acceleration (g)",
            invertY: false,
            series: [{ seriesKey: "input-a", name: "Input A", points: [{ x: 0, y: 0 }, { x: 1, y: 1 }] }],
          },
        ],
      },
      {
        familyKey: "layer-series",
        familyLabel: "Layer Series",
        chartType: "time",
        supportsOverlay: false,
        supportsLayerSelection: true,
        defaultVisibleSeries: ["layer-a-1"],
        layers: [{ layerIndex: 0, layerLabel: "Layer 1", depth: 0.0 }, { layerIndex: 1, layerLabel: "Layer 2", depth: 1.0 }],
        charts: [
          {
            chartKey: "layer-displacement",
            chartLabel: "Layer Displacement",
            sheetName: "Layer",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Displacement (m)",
            invertY: false,
            layerViews: [
              { layerIndex: 0, layerLabel: "Layer 1", depth: 0.0, series: [{ seriesKey: "layer-a-1", name: "A-L1", points: [{ x: 0, y: 0 }, { x: 1, y: 0.2 }] }] },
              { layerIndex: 1, layerLabel: "Layer 2", depth: 1.0, series: [{ seriesKey: "layer-a-2", name: "A-L2", points: [{ x: 0, y: 0 }, { x: 1, y: 0.4 }] }] },
            ],
          },
        ],
      },
    ],
  },
  {
    sourceId: "source-b",
    sourceLabel: "Input B",
    sourceKind: "single",
    axis: "Y",
    pairKey: "METHOD2|B",
    artifactPairKeys: ["METHOD2|B"],
    families: [
      {
        familyKey: "input-motion",
        familyLabel: "Input Motion",
        chartType: "time",
        supportsOverlay: true,
        supportsLayerSelection: false,
        defaultVisibleSeries: ["input-b"],
        layers: [],
        charts: [
          {
            chartKey: "input-acceleration",
            chartLabel: "Input Acceleration",
            sheetName: "Input Motion",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Acceleration (g)",
            invertY: false,
            series: [{ seriesKey: "input-b", name: "Input B", points: [{ x: 0, y: 1 }, { x: 1, y: 0.5 }] }],
          },
        ],
      },
      {
        familyKey: "layer-series",
        familyLabel: "Layer Series",
        chartType: "time",
        supportsOverlay: false,
        supportsLayerSelection: true,
        defaultVisibleSeries: ["layer-b-1"],
        layers: [{ layerIndex: 0, layerLabel: "Layer 1", depth: 0.0 }, { layerIndex: 1, layerLabel: "Layer 2", depth: 1.0 }],
        charts: [
          {
            chartKey: "layer-displacement",
            chartLabel: "Layer Displacement",
            sheetName: "Layer",
            chartType: "time",
            xLabel: "Time (s)",
            yLabel: "Displacement (m)",
            invertY: false,
            layerViews: [
              { layerIndex: 0, layerLabel: "Layer 1", depth: 0.0, series: [{ seriesKey: "layer-b-1", name: "B-L1", points: [{ x: 0, y: 0 }, { x: 1, y: 0.3 }] }] },
              { layerIndex: 1, layerLabel: "Layer 2", depth: 1.0, series: [{ seriesKey: "layer-b-2", name: "B-L2", points: [{ x: 0, y: 0 }, { x: 1, y: 0.5 }] }] },
            ],
          },
        ],
      },
    ],
  },
];

assert.equal(normalizeShellMode("viewer"), "viewer");
assert.equal(normalizeShellMode("legacy"), "legacy");
assert.equal(normalizeShellMode("junk"), "viewer");

assert.equal(normalizeViewerChartMode("focus"), "focus");
assert.equal(normalizeViewerChartMode("compare"), "compare");
assert.equal(normalizeViewerChartMode("stack"), "focus");
assert.equal(normalizeViewerChartMode("junk"), "focus");

const grouped = groupViewerResults(sampleResults);
assert.deepEqual(grouped.map((group) => group.label), ["Primary Outputs", "Method-2", "Method-3 Aggregate"]);

const resultSummary = summarizeViewerResults(sampleResults);
assert.equal(resultSummary.totalResults, 3);
assert.equal(resultSummary.totalCharts, 3);

const defaultScene = buildSourceViewerScene(sampleSourceCatalog, {});
assert.equal(defaultScene.chartMode, "focus");
assert.equal(defaultScene.activeSourceId, "source-a");
assert.equal(defaultScene.activeFamilyKey, "input-motion");
assert.equal(defaultScene.activeChartKey, "input-acceleration");
assert.equal(defaultScene.plotSeriesEntries.length, 1);

const compareScene = buildSourceViewerScene(sampleSourceCatalog, {
  chartMode: "compare",
  activeSourceId: "source-a",
  activeFamilyKey: "input-motion",
  activeChartKey: "input-acceleration",
  compareSourceIds: ["source-b"],
});
assert.equal(compareScene.chartMode, "compare");
assert.deepEqual(compareScene.compareSourceIds, ["source-b"]);
assert.equal(compareScene.plotSeriesEntries.length, 2);

const layerScene = buildSourceViewerScene(sampleSourceCatalog, {
  activeSourceId: "source-a",
  activeFamilyKey: "layer-series",
  activeChartKey: "layer-displacement",
  activeLayerIndex: 99,
});
assert.equal(layerScene.activeLayerIndex, 1);
assert.equal(layerScene.activeLayerCount, 2);
assert.equal(layerScene.activeLayerLabel, "Layer 2");
assert.equal(layerScene.plotSeriesEntries.length, 1);

const nextSource = cycleSceneItem(defaultScene.sources, "source-a", 1, (source) => source.sourceId);
assert.equal(nextSource.sourceId, "source-b");

console.log("viewer model test passed");
