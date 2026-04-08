const PAIR_SCORE_AUTO_THRESHOLD = 105;
const PAIR_SCORE_SUGGEST_THRESHOLD = 40;

const PAIRING_NOISE_TOKENS = new Set([
  "RESULTS",
  "RESULT",
  "PROFILE",
  "MOTION",
  "OUTPUT",
  "OUT",
  "ACC",
  "ACCEL",
  "ACCELERATION",
  "VEL",
  "DISP",
  "SA",
  "SV",
  "SD",
  "TH",
  "TXT",
  "CSV",
  "XLSX",
  "XLS",
  "DD1",
  "DD2",
  "DD3",
  "DD4",
]);

const PAIRING_SCALE_MARKERS = {
  SCALED: "scaled",
  SCALE: "scaled",
  MATCH: "matched",
  MATCHED: "matched",
  ADJ: "adjusted",
  ADJUSTED: "adjusted",
  BENZESTIRILMIS: "adjusted",
};

const PAIRING_AXIS_TOKENS = {
  X: new Set(["X", "HORIZONTAL1", "H1", "HN1", "HNE", "EW", "E", "W", "000", "180", "270", "360", "225", "210"]),
  Y: new Set(["Y", "HORIZONTAL2", "H2", "HN2", "HNN", "NS", "N", "S", "090", "045", "135", "315", "300"]),
  V: new Set(["V", "Z", "VERTICAL", "UD", "UPDOWN", "HNZ"]),
};

const PAIRING_TRAILING_COMPONENT_SUFFIXES = new Set([
  "DN",
  "UP",
  "LT",
  "RT",
  "LF",
  "RG",
  "LR",
  "TB",
  "L1",
  "L2",
  "T1",
  "T2",
  "X1",
  "X2",
  "Y1",
  "Y2",
  "Z1",
  "Z2",
  "V1",
  "V2",
  "E1",
  "E2",
  "W1",
  "W2",
  "N1",
  "N2",
  "S1",
  "S2",
]);

function getPairingLeafName(name) {
  const value = String(name || "").replace(/\\/g, "/");
  const parts = value.split("/").filter(Boolean);
  return parts.length ? parts[parts.length - 1] : "";
}

function isPairingTrailingComponentSuffix(token) {
  const upper = String(token || "").trim().toUpperCase();
  if (!upper) return false;
  if (PAIRING_TRAILING_COMPONENT_SUFFIXES.has(upper)) return true;
  if (/^[A-Z]\d{1,2}$/.test(upper)) return true;
  return false;
}

function stripPairingTrailingComponentSuffixes(name) {
  const text = String(name || "").replace(/[_.-]+$/g, "");
  if (!text) return "";
  const parts = text.split(/[_.-]+/).filter(Boolean);
  while (parts.length > 1 && isPairingTrailingComponentSuffix(parts[parts.length - 1])) {
    parts.pop();
  }
  return parts.join("_");
}

function canonicalizePairingLeaf(name) {
  const leaf = getPairingLeafName(name).replace(/\.[^.]+$/, "");
  return leaf
    .toUpperCase()
    .replace(/HORIONTAL/g, "HORIZONTAL")
    .replace(/BENZE[SZ]TIRILM[Iİ]S/gi, "BENZESTIRILMIS")
    .replace(/HORIZONTAL[_-]?1/g, "HORIZONTAL1")
    .replace(/HORIZONTAL[_-]?2/g, "HORIZONTAL2")
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeAzimuth(deg) {
  if (!Number.isFinite(deg)) return NaN;
  const value = deg % 360;
  return value < 0 ? value + 360 : value;
}

function getPairingAzimuthDelta(a, b) {
  if (!Number.isFinite(a) || !Number.isFinite(b)) return NaN;
  const left = normalizeAzimuth(a);
  const right = normalizeAzimuth(b);
  if (!Number.isFinite(left) || !Number.isFinite(right)) return NaN;
  const diff = Math.abs(left - right);
  return Math.min(diff, 360 - diff);
}

function isOrthogonalPairingAzimuth(a, b) {
  const delta = getPairingAzimuthDelta(a, b);
  if (!Number.isFinite(delta)) return false;
  return Math.abs(delta - 90) < 15 || Math.abs(delta - 270) < 15;
}

function stripOutputSuffix(stem) {
  return String(stem || "").replace(/[_.-](ACC|VEL|DISP|SA|SV|SD|TH)$/i, "");
}

function detectDirection(name) {
  const stem = stripOutputSuffix(String(name || "").replace(/\.[^.]+$/, ""));
  const f = stem.toUpperCase();
  if (/[_.-]X([_.-]|$)/.test(f) || f.endsWith("_X") || f.startsWith("X_")) return "X";
  if (/[_.-]Y([_.-]|$)/.test(f) || f.endsWith("_Y") || f.startsWith("Y_")) return "Y";
  if (
    /(HN1|H1|HNE|EW|000|180|270|360|225|210)$/.test(f) ||
    /[_.-](HN1|H1|HNE|E|W|EW|000|180|270|360|225|210)(?=[_.-]|$)/.test(f)
  ) {
    return "X";
  }
  if (
    /(HN2|H2|HNN|NS|090|045|135|315|300)$/.test(f) ||
    /[_.-](HN2|H2|HNN|N|S|NS|090|045|135|315|300)(?=[_.-]|$)/.test(f)
  ) {
    return "Y";
  }
  if (/\d(E|EW|W|X)$/.test(f)) return "X";
  if (/\d(N|NS|S|Y)$/.test(f)) return "Y";
  if (/[A-Za-z]X\d+$/u.test(stem)) return "X";
  if (/[A-Za-z]Y\d+$/u.test(stem)) return "Y";
  return "SINGLE";
}

function parseGenericAzimuthSuffixToken(value) {
  const upper = String(value || "").trim().toUpperCase();
  if (!upper) return null;
  const match = upper.match(/^(.*?)([NS]?)(\d{1,3})([EW]?)$/);
  if (!match) return null;
  const core = String(match[1] || "").replace(/[_-]+$/g, "");
  const azimuth = Number(match[3]);
  if (!core || !Number.isFinite(azimuth)) return null;
  const edge = String(match[4] || "");
  const prefix = String(match[2] || "");
  let side = null;
  if (edge === "E" || edge === "W" || edge === "EW") side = "X";
  else if (prefix === "N" || prefix === "S" || prefix === "NS") side = "Y";
  if (!side) return null;
  return {
    core,
    side,
    azimuth,
    padded: String(azimuth).padStart(3, "0"),
  };
}

function detectEmbeddedPairingAxisToken(token) {
  const value = String(token || "").trim().toUpperCase();
  if (!value) return null;
  const patterns = [
    { regex: /^(.*\d)(HORIZONTAL1)$/i, side: "X", hint: "HORIZONTAL1" },
    { regex: /^(.*\d)(HORIZONTAL2)$/i, side: "Y", hint: "HORIZONTAL2" },
    { regex: /^(.*\d)(HN1|H1|HNE|EW|000|180|270|360|225|210|X|E|W)$/i, side: "X" },
    { regex: /^(.*\d)(HN2|H2|HNN|NS|090|045|135|315|300|Y|N|S)$/i, side: "Y" },
    { regex: /^(.*\d)(HNZ|UD|UPDOWN|VERTICAL|V|Z)$/i, side: "V" },
  ];
  for (const pattern of patterns) {
    const match = value.match(pattern.regex);
    if (!match) continue;
    const core = String(match[1] || "").replace(/[_-]+$/g, "");
    if (!core) continue;
    const suffix = String(pattern.hint || match[2] || "").toUpperCase();
    return {
      side: pattern.side,
      hint: suffix,
      confidence: /^(X|Y|V|HORIZONTAL1|HORIZONTAL2)$/.test(suffix) ? 3 : 2,
      core,
    };
  }
  const genericAzimuth = parseGenericAzimuthSuffixToken(value);
  if (genericAzimuth?.core && genericAzimuth?.side) {
    return {
      side: genericAzimuth.side,
      hint: `AZ${genericAzimuth.padded}`,
      confidence: 1,
      core: genericAzimuth.core,
      azimuth: genericAzimuth.azimuth,
    };
  }
  return null;
}

function detectDirectionInfo(name) {
  const normalized = canonicalizePairingLeaf(name);
  const tokens = normalized ? normalized.split("_").filter(Boolean) : [];
  let side = null;
  let componentHint = "";
  let confidence = 0;
  let azimuthDeg = null;

  tokens.forEach((token) => {
    if (PAIRING_AXIS_TOKENS.X.has(token) && confidence < 3) {
      side = "X";
      componentHint = token;
      confidence = token === "X" || token === "HORIZONTAL1" ? 3 : 2;
    }
    if (PAIRING_AXIS_TOKENS.Y.has(token) && confidence < 3) {
      side = "Y";
      componentHint = token;
      confidence = token === "Y" || token === "HORIZONTAL2" ? 3 : 2;
    }
    if (PAIRING_AXIS_TOKENS.V.has(token) && confidence < 3) {
      side = "V";
      componentHint = token;
      confidence = 3;
    }
    if (!side) {
      const embedded = detectEmbeddedPairingAxisToken(token);
      if (embedded && embedded.confidence >= confidence) {
        side = embedded.side;
        componentHint = embedded.hint;
        confidence = embedded.confidence;
        if (Number.isFinite(embedded.azimuth)) azimuthDeg = embedded.azimuth;
      }
    }
  });

  if (!side && /(^|_)X\d+($|_)/.test(normalized)) {
    side = "X";
    componentHint = "XN";
    confidence = 2;
  }
  if (!side && /(^|_)Y\d+($|_)/.test(normalized)) {
    side = "Y";
    componentHint = "YN";
    confidence = 2;
  }
  if (!side && /\d(E|EW|W|X)$/.test(normalized)) {
    side = "X";
    componentHint = "AZIM_X";
    confidence = 1;
  }
  if (!side && /\d(N|NS|S|Y)$/.test(normalized)) {
    side = "Y";
    componentHint = "AZIM_Y";
    confidence = 1;
  }

  const fallback = detectDirection(name);
  return {
    side: side || fallback,
    explicit: confidence >= 2,
    componentHint,
    confidence,
    azimuthDeg,
  };
}

function tokenizePairingText(name) {
  const normalized = canonicalizePairingLeaf(name);
  if (!normalized) return [];
  return normalized.split("_").filter(Boolean);
}

function isPairingVerticalProfile(profile) {
  if (!profile) return false;
  if (profile.side === "V") return true;
  return ["V", "Z", "VERTICAL", "UD", "UPDOWN", "HNZ"].includes(String(profile.componentHint || "").toUpperCase());
}

function buildPairingClusterKey(profile) {
  const tokens = Array.isArray(profile?.tokens) ? profile.tokens.slice() : [];
  const filteredTokens = tokens.filter((token) => {
    const upper = String(token || "").toUpperCase();
    if (!upper) return false;
    if (PAIRING_NOISE_TOKENS.has(upper)) return false;
    if (PAIRING_SCALE_MARKERS[upper]) return false;
    if (PAIRING_AXIS_TOKENS.X.has(upper) || PAIRING_AXIS_TOKENS.Y.has(upper) || PAIRING_AXIS_TOKENS.V.has(upper)) return false;
    return true;
  });
  while (filteredTokens.length > 0) {
    const tail = String(filteredTokens[filteredTokens.length - 1] || "").toUpperCase();
    if (/^[A-Z]$/.test(tail) || isPairingTrailingComponentSuffix(tail)) {
      filteredTokens.pop();
      continue;
    }
    break;
  }
  const stemKey = filteredTokens.join("_");
  const parts = [
    profile?.recordId || "",
    stemKey || profile?.groupKey || profile?.base || "",
  ]
    .map((part) => String(part || "").trim())
    .filter(Boolean);
  const key = parts.join("||");
  if (key) return key;
  return recordBaseName(profile?.file?.name || "");
}

function recordBaseName(filename) {
  let name = String(filename || "").replace(/\.[^.]+$/, "");
  name = name.replace(/^dd[1-4][_-]?/i, "");
  name = name.replace(/^[xy][_-]/i, "");
  name = name.replace(/[_.-]?(h1|h2|hn1|hn2|hne|hnn|hnz)$/i, "");
  name = name.replace(/[_.-]?(ns|ew)$/i, "");
  name = name.replace(/[_.-]?([ns]\d{1,3}[ew])$/i, "");
  name = name.replace(/[_.-]?(000|090|180|270|360|045|135|225|315|210|300)$/i, "");
  name = name.replace(/[_.-]([enws])$/i, "");
  name = name.replace(/(\d)(e|w|n|s|x|y)$/i, "$1");
  name = name.replace(/([A-Za-z])([XY])(\d+)$/u, "$1$3");
  name = name.replace(/[_.-]{2,}/g, "_").replace(/^[_-]+|[_-]+$/g, "");
  return name || String(filename || "").replace(/\.[^.]+$/, "");
}

function deepsoilBaseName(filename) {
  let name = String(filename || "").replace(/\.[^.]+$/, "");
  name = name.replace(/^results_profile_\d+_motion_/i, "");
  name = name.replace(/^results_profile_[^_]+_motion_/i, "");
  name = name.replace(/^results_profile[^_]*_motion_/i, "");
  name = name.replace(/^results_profile_motion_/i, "");
  name = name.replace(/^results_profile_/i, "");
  name = name.replace(/^results_/i, "");
  name = name.replace(/^dd[1-4][_-]?/i, "");
  name = name.replace(/^[xy][_-]/i, "");
  name = name.replace(/[_.-]acc[_.-]?(e|w|n|s|ew|ns)$/i, "");
  name = name.replace(/[_.-](x|y|ew|ns|h1|h2|hn1|hn2|hne|hnn|hnz)$/i, "");
  name = name.replace(/[_.-](000|090|180|270|360|045|135|225|315|210|300)$/i, "");
  name = name.replace(/(\d)(e|w|n|s|x|y)$/i, "$1");
  name = name.replace(/([A-Za-z])([XY])(\d+)$/u, "$1$3");
  name = name.replace(/[_.-]{2,}/g, "_").replace(/^[_-]+|[_-]+$/g, "");
  return name;
}

function deepsoilBaseNameLoose(filename) {
  let name = deepsoilBaseName(filename);
  name = name.replace(/([NS]\d{1,3}[EW])$/i, "");
  name = name.replace(/(\d{1,5}(EW|NS))$/i, "");
  name = name.replace(/(\d{1,3}[EWNS])$/i, "");
  name = name.replace(/([NS]\d{1,3})$/i, "");
  name = name.replace(/(000|090|180|270|360|045|135|225|315|210|300)$/i, "");
  name = name.replace(/[_.-](\d{2,3})$/i, "");
  name = name.replace(/[_.-]{2,}/g, "_").replace(/^[_-]+|[_-]+$/g, "");
  return name;
}

function normalizePairingItem(item) {
  if (!item) return null;
  if (typeof item === "string") {
    const name = String(item).trim();
    return name ? { name } : null;
  }
  if (typeof item === "object") {
    const name = String(item.name || item.file?.name || item.path || "").trim();
    if (!name) return null;
    return Object.prototype.hasOwnProperty.call(item, "name") ? item : { ...item, name };
  }
  return null;
}

function normalizePairingItems(items) {
  const byName = new Map();
  (Array.isArray(items) ? items : []).forEach((item) => {
    const normalized = normalizePairingItem(item);
    if (!normalized) return;
    const key = String(normalized.name || "").trim().toLowerCase();
    if (!key || byName.has(key)) return;
    byName.set(key, normalized);
  });
  return [...byName.values()];
}

function buildPairingProfile(item) {
  const file = normalizePairingItem(item);
  const direction = detectDirectionInfo(file?.name || "");
  const rawTokens = tokenizePairingText(file?.name || "");
  let scaleState = "raw";
  const qualityFlags = [];
  const filteredTokens = [];

  rawTokens.forEach((token) => {
    if (PAIRING_SCALE_MARKERS[token]) {
      scaleState = PAIRING_SCALE_MARKERS[token];
      return;
    }
    if (PAIRING_AXIS_TOKENS.X.has(token) || PAIRING_AXIS_TOKENS.Y.has(token) || PAIRING_AXIS_TOKENS.V.has(token)) return;
    if (PAIRING_NOISE_TOKENS.has(token)) return;
    const embedded = detectEmbeddedPairingAxisToken(token);
    if (embedded) {
      filteredTokens.push(embedded.core);
      return;
    }
    if (/^X\d+$/.test(token) || /^Y\d+$/.test(token)) {
      filteredTokens.push(token.slice(1));
      return;
    }
    filteredTokens.push(token);
  });

  const recordId = rawTokens.find((token) => /^RSN\d+$/i.test(token)) || "";
  const numericTokens = filteredTokens.filter((token) => /^\d+$/.test(token));
  const coreTokens = filteredTokens.filter((token) => !/^\d+$/.test(token));
  const eventTokens = coreTokens.filter((token) => token !== recordId);
  const eventKey = eventTokens.join("_");
  const groupKey = [recordId, eventKey, numericTokens.join("_")].filter(Boolean).join("_") || recordBaseName(file?.name || "");
  if (!direction.explicit) qualityFlags.push("weak-axis");
  if (!groupKey) qualityFlags.push("weak-group");
  return {
    file,
    side: direction.side,
    explicitSide: direction.explicit ? direction.side : null,
    axisConfidence: direction.confidence,
    componentHint: direction.componentHint,
    azimuthDeg: direction.azimuthDeg,
    scaleState,
    recordId,
    base: recordBaseName(file?.name || ""),
    groupKey,
    eventKey,
    tokens: filteredTokens,
    numericTokens,
    qualityFlags,
  };
}

function getSharedPairingValues(left, right) {
  const sharedTokens = Array.from(new Set((left.tokens || []).filter((token) => right.tokens?.includes(token))));
  const sharedNumbers = Array.from(new Set((left.numericTokens || []).filter((token) => right.numericTokens?.includes(token))));
  return { sharedTokens, sharedNumbers };
}

function scorePairingRecords(xProfile, yProfile) {
  if (!xProfile || !yProfile || xProfile.file?.name === yProfile.file?.name) return null;
  if (xProfile.side !== "X" || yProfile.side !== "Y") return null;
  const { sharedTokens, sharedNumbers } = getSharedPairingValues(xProfile, yProfile);
  let score = 0;
  const reasons = [];
  if (xProfile.base && yProfile.base && xProfile.base === yProfile.base) {
    score += 135;
    reasons.push("base ayni");
  }
  if (xProfile.groupKey && yProfile.groupKey && xProfile.groupKey === yProfile.groupKey) {
    score += 120;
    reasons.push("groupKey ayni");
  } else if (xProfile.eventKey && yProfile.eventKey && xProfile.eventKey === yProfile.eventKey) {
    score += 55;
    reasons.push("event ayni");
  }
  if (xProfile.recordId && yProfile.recordId && xProfile.recordId === yProfile.recordId) {
    score += 90;
    reasons.push(xProfile.recordId);
  } else if (xProfile.recordId && yProfile.recordId && xProfile.recordId !== yProfile.recordId) {
    score -= 90;
  }
  if (xProfile.scaleState === yProfile.scaleState) {
    score += 16;
    reasons.push(`${xProfile.scaleState} etiketi ortak`);
  } else if (xProfile.scaleState !== "raw" || yProfile.scaleState !== "raw") {
    score -= 35;
  }
  score += sharedTokens.length * 10;
  score += sharedNumbers.length * 18;
  if (sharedNumbers.length > 0) reasons.push(`${sharedNumbers.join(",")} sayisal eslesme`);
  if (xProfile.axisConfidence > 0) score += xProfile.axisConfidence * 6;
  if (yProfile.axisConfidence > 0) score += yProfile.axisConfidence * 6;
  if (Number.isFinite(xProfile.azimuthDeg) && Number.isFinite(yProfile.azimuthDeg)) {
    if (isOrthogonalPairingAzimuth(xProfile.azimuthDeg, yProfile.azimuthDeg)) {
      score += 72;
      reasons.push(`${String(xProfile.azimuthDeg).padStart(3, "0")}/${String(yProfile.azimuthDeg).padStart(3, "0")} dik aci`);
    } else {
      const delta = getPairingAzimuthDelta(xProfile.azimuthDeg, yProfile.azimuthDeg);
      if (Number.isFinite(delta) && (delta < 15 || Math.abs(delta - 180) < 15)) {
        score -= 55;
      }
    }
  }
  if (xProfile.componentHint && yProfile.componentHint) {
    reasons.push(`${xProfile.componentHint}/${yProfile.componentHint}`);
  }
  if (!sharedTokens.length && !sharedNumbers.length && xProfile.groupKey !== yProfile.groupKey) score -= 28;
  return {
    x: xProfile.file,
    y: yProfile.file,
    base: xProfile.base === yProfile.base ? xProfile.base : recordBaseName(xProfile.file?.name || yProfile.file?.name || ""),
    score,
    reason: reasons.filter(Boolean).join(" + "),
  };
}

function scorePairingProfilesAny(left, right) {
  if (!left || !right || left.file?.name === right.file?.name) return null;
  const oriented = orientPairingProfiles(left, right);
  return scorePairingRecords(oriented.x, oriented.y);
}

function orientPairingProfiles(left, right) {
  const a = left ? { ...left } : null;
  const b = right ? { ...right } : null;
  if (!a || !b) return { x: a, y: b };
  const aIsX = a.side === "X";
  const aIsY = a.side === "Y";
  const bIsX = b.side === "X";
  const bIsY = b.side === "Y";
  if (aIsX && bIsY) return { x: a, y: b };
  if (aIsY && bIsX) return { x: b, y: a };
  if (Number.isFinite(a.azimuthDeg) && Number.isFinite(b.azimuthDeg)) {
    if (a.azimuthDeg <= b.azimuthDeg) {
      a.side = "X";
      b.side = "Y";
      return { x: a, y: b };
    }
    b.side = "X";
    a.side = "Y";
    return { x: b, y: a };
  }
  if (aIsX || bIsY) {
    a.side = "X";
    b.side = "Y";
    return { x: a, y: b };
  }
  if (aIsY || bIsX) {
    b.side = "X";
    a.side = "Y";
    return { x: b, y: a };
  }
  const aName = String(a.file?.name || "");
  const bName = String(b.file?.name || "");
  if (aName.localeCompare(bName) <= 0) {
    a.side = "X";
    b.side = "Y";
    return { x: a, y: b };
  }
  b.side = "X";
  a.side = "Y";
  return { x: b, y: a };
}

function computeMaximumWeightMatching(weights) {
  const rowCount = Array.isArray(weights) ? weights.length : 0;
  const colCount = rowCount > 0 && Array.isArray(weights[0]) ? weights[0].length : 0;
  if (!rowCount || !colCount) return [];
  const size = Math.max(rowCount, colCount);
  let maxWeight = 0;
  const matrix = Array.from({ length: size + 1 }, () => Array(size + 1).fill(0));
  for (let i = 1; i <= rowCount; i += 1) {
    for (let j = 1; j <= colCount; j += 1) {
      const value = Math.max(0, Number(weights[i - 1][j - 1]) || 0);
      matrix[i][j] = value;
      if (value > maxWeight) maxWeight = value;
    }
  }
  const u = Array(size + 1).fill(0);
  const v = Array(size + 1).fill(0);
  const p = Array(size + 1).fill(0);
  const way = Array(size + 1).fill(0);
  for (let i = 1; i <= size; i += 1) {
    p[0] = i;
    let j0 = 0;
    const minv = Array(size + 1).fill(Number.POSITIVE_INFINITY);
    const used = Array(size + 1).fill(false);
    do {
      used[j0] = true;
      const i0 = p[j0];
      let delta = Number.POSITIVE_INFINITY;
      let j1 = 0;
      for (let j = 1; j <= size; j += 1) {
        if (used[j]) continue;
        const cur = (maxWeight - matrix[i0][j]) - u[i0] - v[j];
        if (cur < minv[j]) {
          minv[j] = cur;
          way[j] = j0;
        }
        if (minv[j] < delta) {
          delta = minv[j];
          j1 = j;
        }
      }
      for (let j = 0; j <= size; j += 1) {
        if (used[j]) {
          u[p[j]] += delta;
          v[j] -= delta;
        } else {
          minv[j] -= delta;
        }
      }
      j0 = j1;
    } while (p[j0] !== 0);
    do {
      const j1 = way[j0];
      p[j0] = p[j1];
      j0 = j1;
    } while (j0 !== 0);
  }
  const result = [];
  for (let j = 1; j <= size; j += 1) {
    if (p[j] > 0 && p[j] <= rowCount && j <= colCount) {
      result.push([p[j] - 1, j - 1]);
    }
  }
  return result;
}

function pairFilesByScoredFuzzy(items) {
  const src = normalizePairingItems(items);
  if (src.length < 2) return { pairs: [], leftovers: src, suggestions: [] };
  const profiles = src.map((file) => buildPairingProfile(file));
  const clusterMap = new Map();
  profiles.forEach((profile) => {
    const key = buildPairingClusterKey(profile);
    if (!clusterMap.has(key)) clusterMap.set(key, []);
    clusterMap.get(key).push(profile);
  });

  const autoPairs = [];
  const suggestions = [];
  const usedNames = new Set();
  const unresolvedProfiles = [];

  const clusterEntries = Array.from(clusterMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));
  clusterEntries.forEach(([, cluster]) => {
    if (!cluster || cluster.length < 2) {
      if (Array.isArray(cluster) && cluster.length === 1) unresolvedProfiles.push(cluster[0]);
      return;
    }

    const vProfile = cluster.find((profile) => isPairingVerticalProfile(profile)) || null;
    const baseProfiles = cluster.filter((profile) => profile !== vProfile);

    if (cluster.length === 2 || (cluster.length === 3 && vProfile)) {
      const candidateProfiles = baseProfiles.slice();
      let bestCandidate = null;
      for (let i = 0; i < candidateProfiles.length; i += 1) {
        for (let j = i + 1; j < candidateProfiles.length; j += 1) {
          const candidate = scorePairingProfilesAny(candidateProfiles[i], candidateProfiles[j]);
          if (!candidate || candidate.score < PAIR_SCORE_SUGGEST_THRESHOLD) continue;
          if (!bestCandidate || candidate.score > bestCandidate.score) bestCandidate = candidate;
        }
      }
      if (!bestCandidate) {
        unresolvedProfiles.push(...cluster);
        return;
      }
      const tuple = [bestCandidate.x, bestCandidate.y, bestCandidate.base];
      if (vProfile) tuple.push(vProfile.file);
      const target = bestCandidate.score >= PAIR_SCORE_AUTO_THRESHOLD ? autoPairs : suggestions;
      if (target === autoPairs) {
        autoPairs.push(tuple);
      } else {
        suggestions.push({
          id: `${bestCandidate.x?.name || ""}||${bestCandidate.y?.name || ""}`,
          base: bestCandidate.base,
          x: bestCandidate.x,
          y: bestCandidate.y,
          v: vProfile?.file || null,
          score: bestCandidate.score,
          reason: bestCandidate.reason,
        });
      }
      [bestCandidate.x?.name, bestCandidate.y?.name, vProfile?.file?.name].filter(Boolean).forEach((name) => usedNames.add(name));
      return;
    }

    unresolvedProfiles.push(...cluster);
  });

  const unresolvedFiles = unresolvedProfiles.map((profile) => profile.file).filter((file) => file && !usedNames.has(file.name));
  const xProfiles = unresolvedProfiles.filter((profile) => profile.side === "X" && !usedNames.has(profile.file?.name));
  const yProfiles = unresolvedProfiles.filter((profile) => profile.side === "Y" && !usedNames.has(profile.file?.name));

  if (xProfiles.length && yProfiles.length) {
    const weights = xProfiles.map((xProfile) =>
      yProfiles.map((yProfile) => {
        const candidate = scorePairingRecords(xProfile, yProfile);
        return candidate ? candidate.score : 0;
      })
    );
    const matchIndexes = computeMaximumWeightMatching(weights);
    const matchedX = new Set();
    const matchedY = new Set();

    matchIndexes.forEach(([xIdx, yIdx]) => {
      const candidate = scorePairingRecords(xProfiles[xIdx], yProfiles[yIdx]);
      if (!candidate || candidate.score < PAIR_SCORE_SUGGEST_THRESHOLD) return;
      matchedX.add(candidate.x?.name);
      matchedY.add(candidate.y?.name);
      const tuple = [candidate.x, candidate.y, candidate.base];
      const target = candidate.score >= PAIR_SCORE_AUTO_THRESHOLD ? autoPairs : suggestions;
      if (target === autoPairs) {
        autoPairs.push(tuple);
      } else {
        suggestions.push({
          id: `${candidate.x?.name || ""}||${candidate.y?.name || ""}`,
          base: candidate.base,
          x: candidate.x,
          y: candidate.y,
          score: candidate.score,
          reason: candidate.reason,
        });
      }
    });

    const usedForSuggestions = new Set([...matchedX, ...matchedY]);
    const remainingCandidates = [];
    xProfiles.forEach((xProfile) => {
      if (usedForSuggestions.has(xProfile.file?.name)) return;
      yProfiles.forEach((yProfile) => {
        if (usedForSuggestions.has(yProfile.file?.name)) return;
        const candidate = scorePairingRecords(xProfile, yProfile);
        if (!candidate || candidate.score < PAIR_SCORE_SUGGEST_THRESHOLD) return;
        remainingCandidates.push(candidate);
      });
    });
    remainingCandidates
      .sort(
        (a, b) =>
          b.score - a.score ||
          String(a.base || "").localeCompare(String(b.base || "")) ||
          String(a.x?.name || "").localeCompare(String(b.x?.name || "")) ||
          String(a.y?.name || "").localeCompare(String(b.y?.name || ""))
      )
      .forEach((candidate) => {
        const xName = candidate.x?.name;
        const yName = candidate.y?.name;
        if (!xName || !yName || usedForSuggestions.has(xName) || usedForSuggestions.has(yName)) return;
        suggestions.push({
          id: `${xName}||${yName}`,
          base: candidate.base,
          x: candidate.x,
          y: candidate.y,
          score: candidate.score,
          reason: candidate.reason,
        });
        usedForSuggestions.add(xName);
        usedForSuggestions.add(yName);
      });
  } else if (unresolvedFiles.length) {
    unresolvedFiles.forEach((file) => {
      if (file?.name) usedNames.add(file.name);
    });
  }

  const usedAuto = new Set(autoPairs.flatMap(([xFile, yFile, , vFile]) => [xFile?.name, yFile?.name, vFile?.name]).filter(Boolean));
  const leftovers = src.filter((file) => !usedAuto.has(file.name));
  return {
    pairs: autoPairs,
    leftovers,
    suggestions,
  };
}

function pairByKey(items, keyFn) {
  const groups = new Map();
  normalizePairingItems(items).forEach((item) => {
    const key = String(keyFn(item) || "").trim() || "__missing__";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  });

  const pairs = [];
  const leftovers = [];
  const suggestions = [];

  Array.from(groups.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([, group]) => {
      if (group.length < 2) {
        leftovers.push(...group);
        return;
      }
      const resolved = pairFilesByScoredFuzzy(group);
      pairs.push(...resolved.pairs);
      suggestions.push(...resolved.suggestions);
      leftovers.push(...resolved.leftovers);
    });

  return { pairs, leftovers, suggestions };
}

function pairFilesByDeepsoil(items) {
  const src = normalizePairingItems(items);
  if (src.length < 2) return { pairs: [], leftovers: src, suggestions: [] };

  const strict = pairByKey(src, (item) => deepsoilBaseName(item.name) || recordBaseName(item.name));
  const loose = pairByKey(strict.leftovers, (item) => deepsoilBaseNameLoose(item.name) || deepsoilBaseName(item.name) || recordBaseName(item.name));
  const fuzzy = pairFilesByScoredFuzzy(loose.leftovers);

  return {
    pairs: [...strict.pairs, ...loose.pairs, ...fuzzy.pairs],
    leftovers: fuzzy.leftovers,
    suggestions: [...strict.suggestions, ...loose.suggestions, ...fuzzy.suggestions],
  };
}

function normalizePairEntry(entry) {
  if (!entry) return null;
  if (Array.isArray(entry) && entry.length >= 2) {
    const xName = String(entry[0] || "").trim();
    const yName = String(entry[1] || "").trim();
    return xName && yName ? { xName, yName } : null;
  }
  if (typeof entry === "object") {
    const xName = String(entry.xName || entry.x?.name || entry.x || "").trim();
    const yName = String(entry.yName || entry.y?.name || entry.y || "").trim();
    return xName && yName ? { xName, yName } : null;
  }
  return null;
}

function collectExcludedNames(existingPairs) {
  const excluded = new Set();
  (Array.isArray(existingPairs) ? existingPairs : []).forEach((entry) => {
    const pair = normalizePairEntry(entry);
    if (!pair) return;
    excluded.add(pair.xName.toLowerCase());
    excluded.add(pair.yName.toLowerCase());
  });
  return excluded;
}

function resolveDeepsoilPairingCandidates(items, options = {}) {
  const existingPairs = options?.existingPairs || [];
  const excluded = collectExcludedNames(existingPairs);
  const available = normalizePairingItems(items).filter((item) => !excluded.has(item.name.toLowerCase()));
  if (available.length < 2) {
    return {
      pairs: [],
      leftovers: available,
      suggestions: [],
    };
  }
  return pairFilesByDeepsoil(available);
}

export {
  PAIR_SCORE_AUTO_THRESHOLD,
  PAIR_SCORE_SUGGEST_THRESHOLD,
  canonicalizePairingLeaf,
  deepsoilBaseName,
  deepsoilBaseNameLoose,
  detectDirectionInfo,
  pairFilesByDeepsoil,
  pairFilesByScoredFuzzy,
  resolveDeepsoilPairingCandidates,
};
