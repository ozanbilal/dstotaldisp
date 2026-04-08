import assert from "node:assert/strict";
import { buildZipBlobFromResults } from "../web-ui/zip.mjs";

function readU16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function decodeText(bytes) {
  return new TextDecoder().decode(bytes);
}

async function parseZip(blob) {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const localEntries = [];
  const centralEntries = [];
  let offset = 0;

  while (offset < bytes.length) {
    const signature = readU32(bytes, offset);
    if (signature === 0x04034b50) {
      const flags = readU16(bytes, offset + 6);
      const method = readU16(bytes, offset + 8);
      const crc = readU32(bytes, offset + 14);
      const compressedSize = readU32(bytes, offset + 18);
      const uncompressedSize = readU32(bytes, offset + 22);
      const nameLength = readU16(bytes, offset + 26);
      const extraLength = readU16(bytes, offset + 28);
      const name = decodeText(bytes.subarray(offset + 30, offset + 30 + nameLength));
      const dataStart = offset + 30 + nameLength + extraLength;
      const data = bytes.subarray(dataStart, dataStart + compressedSize);
      localEntries.push({ name, flags, method, crc, compressedSize, uncompressedSize, data, offset });
      offset = dataStart + compressedSize;
      continue;
    }
    if (signature === 0x02014b50 || signature === 0x06054b50) {
      break;
    }
    throw new Error(`Unexpected ZIP signature at offset ${offset}: 0x${signature.toString(16)}`);
  }

  while (offset < bytes.length) {
    const signature = readU32(bytes, offset);
    if (signature === 0x02014b50) {
      const method = readU16(bytes, offset + 10);
      const crc = readU32(bytes, offset + 16);
      const compressedSize = readU32(bytes, offset + 20);
      const uncompressedSize = readU32(bytes, offset + 24);
      const nameLength = readU16(bytes, offset + 28);
      const extraLength = readU16(bytes, offset + 30);
      const commentLength = readU16(bytes, offset + 32);
      const localOffset = readU32(bytes, offset + 42);
      const name = decodeText(bytes.subarray(offset + 46, offset + 46 + nameLength));
      centralEntries.push({ name, method, crc, compressedSize, uncompressedSize, localOffset, offset });
      offset += 46 + nameLength + extraLength + commentLength;
      continue;
    }
    if (signature === 0x06054b50) {
      const totalEntries = readU16(bytes, offset + 8);
      const centralSize = readU32(bytes, offset + 12);
      const centralOffset = readU32(bytes, offset + 16);
      const commentLength = readU16(bytes, offset + 20);
      const comment = decodeText(bytes.subarray(offset + 22, offset + 22 + commentLength));
      return { bytes, localEntries, centralEntries, eocd: { totalEntries, centralSize, centralOffset, comment } };
    }
    throw new Error(`Unexpected central directory signature at offset ${offset}: 0x${signature.toString(16)}`);
  }

  throw new Error("ZIP EOCD record was not found.");
}

const results = [
  {
    outputFileName: "alpha.xlsx",
    outputBytesB64: Buffer.from("alpha").toString("base64"),
  },
  {
    outputFileName: "beta.xlsx",
    outputBytes: new Uint8Array([0, 1, 2, 3, 4]),
  },
];

const blob = await buildZipBlobFromResults(results, {
  yieldEvery: 1,
  timestamp: new Date("2026-03-30T12:00:00"),
});
const parsed = await parseZip(blob);

assert.equal(parsed.localEntries.length, 2);
assert.equal(parsed.centralEntries.length, 2);
assert.equal(parsed.eocd.totalEntries, 2);
assert.equal(parsed.eocd.comment, "");
assert.deepEqual(parsed.localEntries.map((entry) => entry.name), ["alpha.xlsx", "beta.xlsx"]);
assert.deepEqual(parsed.centralEntries.map((entry) => entry.name), ["alpha.xlsx", "beta.xlsx"]);
assert.equal(parsed.localEntries[0].method, 0);
assert.equal(parsed.localEntries[1].method, 0);
assert.ok(parsed.localEntries[0].flags & 0x0800);
assert.ok(parsed.localEntries[1].flags & 0x0800);
assert.deepEqual(Array.from(parsed.localEntries[0].data), Array.from(Buffer.from("alpha")));
assert.deepEqual(Array.from(parsed.localEntries[1].data), [0, 1, 2, 3, 4]);
assert.equal(parsed.centralEntries[0].localOffset, parsed.localEntries[0].offset);
assert.equal(parsed.centralEntries[1].localOffset, parsed.localEntries[1].offset);
assert.equal(parsed.eocd.centralOffset, parsed.centralEntries[0].offset);

console.log("zip builder test passed");
