import assert from "node:assert/strict";
import {
  buildSpectrumScopeKey,
  formatViewerSourceMeta,
  getScopedSpectrumPeriodMax,
  summarizeViewerSourceKinds,
} from "../web-ui/viewer_controls_state.mjs";

const spectrumScene = {
  activeSourceId: "source-a",
  activeFamilyKey: "input-motion",
  activeChartKey: "response-spectrum",
  activeLayerIndex: 2,
};

assert.equal(
  buildSpectrumScopeKey(spectrumScene),
  "source-a::input-motion::response-spectrum::2"
);
assert.equal(
  getScopedSpectrumPeriodMax(
    {
      "source-a::input-motion::response-spectrum::2": "5.5",
      "source-b::input-motion::response-spectrum::0": "2.0",
    },
    spectrumScene
  ),
  "5.5"
);
assert.equal(
  getScopedSpectrumPeriodMax(
    {
      "source-b::input-motion::response-spectrum::0": "2.0",
    },
    spectrumScene
  ),
  ""
);

const summary = summarizeViewerSourceKinds([
  { sourceKind: "pair" },
  { sourceKind: "single" },
  { sourceKind: "single" },
  { sourceKind: "db_pair" },
  { sourceKind: "method3_aggregate" },
]);
assert.equal(summary.pairViewCount, 2);
assert.equal(summary.singleDirectionCount, 2);
assert.equal(summary.aggregateViewCount, 1);

assert.equal(
  formatViewerSourceMeta({
    sources: [
      { sourceKind: "pair" },
      { sourceKind: "single" },
      { sourceKind: "single" },
      { sourceKind: "db_pair" },
      { sourceKind: "method3_aggregate" },
    ],
    sourceCount: 5,
    fileCount: 4,
  }),
  "5 viewer sources · 4 input files · 2 pair views · 2 single-direction sources · 1 aggregate view"
);

console.log("viewer controls state test passed");
