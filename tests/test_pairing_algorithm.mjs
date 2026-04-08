import assert from "node:assert/strict";
import { detectDirectionInfo, resolveDeepsoilPairingCandidates } from "../web-ui/pairing.mjs";

const sampleFiles = [
  { name: "Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN1111.xlsx" },
  { name: "Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN1111.xlsx" },
  { name: "Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN1633.xlsx" },
  { name: "Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN1633.xlsx" },
];

const resolved = resolveDeepsoilPairingCandidates(sampleFiles);
const pairKeys = resolved.pairs
  .map(([xFile, yFile]) => `${xFile.name}|${yFile.name}`)
  .sort();

assert.deepEqual(pairKeys, [
  "Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN1111.xlsx|Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN1111.xlsx",
  "Results_profile_Profile 1_motion_Yalova_Horizontal_1_Matched_RSN1633.xlsx|Results_profile_Profile 1_motion_Yalova_Horiontal_2_Matched_RSN1633.xlsx",
]);
assert.equal(resolved.leftovers.length, 0);
assert.equal(detectDirectionInfo(sampleFiles[0].name).side, "X");
assert.equal(detectDirectionInfo(sampleFiles[1].name).side, "Y");
assert.equal(detectDirectionInfo("Results_profile_1_motion_RSN801_LOMAP_SJTE225.xlsx").side, "X");
assert.equal(detectDirectionInfo("Results_profile_1_motion_RSN801_LOMAP_SJTE315.xlsx").side, "Y");

console.log("pairing algorithm test passed");
