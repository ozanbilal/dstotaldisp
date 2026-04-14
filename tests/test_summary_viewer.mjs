import assert from "node:assert/strict";
import { buildSummaryViewerScene } from "../web-ui/summary_viewer.mjs";

const summaryCatalog = [
  {
    summaryId: "pair-1",
    summaryLabel: "Pair 1",
    summaryKind: "pair",
    axis: "XY",
    pairKey: "PAIR|1",
    inputKind: "xlsx",
    preferredVariantKey: "strain_input_total",
    detailSourceIds: ["source-pair-1"],
    artifactPairKeys: ["PAIR|1"],
    warnings: [],
    coverage: { availableLayerCount: 3, profileLayerCount: 3, limitedData: false, label: "Tam veri" },
    variants: [
      {
        variantKey: "strain_input_total",
        displayLabel: "Strain + Input Proxy",
        methodClass: "computed",
        valid: true,
        confidenceRank: 20,
        confidenceLabel: "Primary",
        reason: "best",
        depths: [0.5, 1.5, 2.5],
        values: [0.1, 0.2, 0.3],
      },
      {
        variantKey: "profile_offset_total",
        displayLabel: "Profile Offset",
        methodClass: "approximate",
        valid: true,
        confidenceRank: 40,
        confidenceLabel: "Support",
        reason: "approx",
        depths: [0.5, 1.5, 2.5],
        values: [0.12, 0.18, 0.28],
      },
      {
        variantKey: "time_history_total",
        displayLabel: "Time History",
        methodClass: "indirect",
        valid: false,
        confidenceRank: 50,
        confidenceLabel: "Support",
        reason: "missing layers",
        depths: [0.5, 1.5, 2.5],
        values: [0.13, 0.19, 0.29],
      },
    ],
  },
  {
    summaryId: "single-1",
    summaryLabel: "Single 1",
    summaryKind: "single",
    axis: "X",
    pairKey: "SINGLE|1",
    inputKind: "xlsx",
    preferredVariantKey: "profile_offset_total",
    detailSourceIds: ["source-single-1"],
    artifactPairKeys: ["SINGLE|1"],
    warnings: ["limited"],
    coverage: { availableLayerCount: 2, profileLayerCount: 3, limitedData: true, label: "Sinirli veri" },
    variants: [
      {
        variantKey: "profile_offset_total",
        displayLabel: "Profile Offset",
        methodClass: "approximate",
        valid: true,
        confidenceRank: 40,
        confidenceLabel: "Support",
        reason: "available",
        depths: [0.5, 1.5],
        values: [0.05, 0.08],
      },
    ],
  },
];

const defaultScene = buildSummaryViewerScene(summaryCatalog, {});
assert.equal(defaultScene.summaryCount, 2);
assert.equal(defaultScene.activeSummaryId, "pair-1");
assert.equal(defaultScene.primaryVariant.variantKey, "strain_input_total");
assert.equal(defaultScene.validVariants.length, 2);
assert.equal(defaultScene.invalidVariants.length, 1);
assert.equal(defaultScene.visibleVariants.length, 2);

const filteredScene = buildSummaryViewerScene(summaryCatalog, {
  activeSummaryId: "pair-1",
  variantVisibilityMap: { "pair-1::profile_offset_total": false },
});
assert.equal(filteredScene.visibleVariants.length, 1);
assert.equal(filteredScene.visibleVariants[0].variantKey, "strain_input_total");

const singleScene = buildSummaryViewerScene(summaryCatalog, { activeSummaryId: "single-1" });
assert.equal(singleScene.activeSummary.summaryKind, "single");
assert.equal(singleScene.primaryVariant.variantKey, "profile_offset_total");
assert.equal(singleScene.activeSummary.coverage.limitedData, true);

const forcedPrimaryScene = buildSummaryViewerScene(summaryCatalog, {
  activeSummaryId: "pair-1",
  primaryVariantMap: { "pair-1": "profile_offset_total" },
});
assert.equal(forcedPrimaryScene.primaryVariant.variantKey, "profile_offset_total");

console.log("summary viewer test passed");
