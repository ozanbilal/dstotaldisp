const textEncoder = new TextEncoder();
const MAX_UINT32 = 0xffffffff;

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let crc = i;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) !== 0 ? (0xedb88320 ^ (crc >>> 1)) >>> 0 : crc >>> 1;
    }
    table[i] = crc >>> 0;
  }
  return table;
})();

function writeUint16LE(target, offset, value) {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32LE(target, offset, value) {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

function sleepFrame() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function decodeBase64(base64) {
  const value = String(base64 || "").replace(/\s+/g, "");
  if (!value) {
    return new Uint8Array();
  }

  if (typeof Buffer !== "undefined" && typeof Buffer.from === "function") {
    return Uint8Array.from(Buffer.from(value, "base64"));
  }

  if (typeof atob === "function") {
    const binary = atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  throw new Error("No base64 decoder available in this runtime.");
}

function normalizeBytesSource(value, label) {
  if (value instanceof Uint8Array) {
    return value;
  }
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value);
  }
  if (Array.isArray(value)) {
    return Uint8Array.from(value);
  }
  if (typeof value === "string") {
    return decodeBase64(value);
  }

  throw new Error(`Unsupported ZIP payload for ${label || "entry"}.`);
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }
  return (~crc) >>> 0;
}

function toDosDateTime(date) {
  const safe = date instanceof Date && !Number.isNaN(date.getTime()) ? date : new Date();
  const year = Math.min(Math.max(safe.getFullYear(), 1980), 2107);
  const dosDate = ((year - 1980) << 9) | ((safe.getMonth() + 1) << 5) | safe.getDate();
  const dosTime = (safe.getHours() << 11) | (safe.getMinutes() << 5) | Math.floor(safe.getSeconds() / 2);
  return { dosDate, dosTime };
}

function sanitizeEntryName(rawName, index, usedNames) {
  const fallback = `output_${index + 1}.xlsx`;
  const raw = String(rawName || fallback).replace(/\\/g, "/");
  const parts = raw.split("/").filter((part) => part && part !== "." && part !== "..");
  const lastPart = parts.length ? parts[parts.length - 1] : fallback;
  const cleaned = lastPart.replace(/[\x00-\x1f<>:"|?*]/g, "_").trim().replace(/\.+$/g, "");
  const base = cleaned || fallback;

  const key = base.toLowerCase();
  if (!usedNames.has(key)) {
    usedNames.add(key);
    return base;
  }

  const dot = base.lastIndexOf(".");
  let suffix = 1;
  while (true) {
    const next = dot >= 0 ? `${base.slice(0, dot)}_${suffix}${base.slice(dot)}` : `${base}_${suffix}`;
    const nextKey = next.toLowerCase();
    if (!usedNames.has(nextKey)) {
      usedNames.add(nextKey);
      return next;
    }
    suffix += 1;
  }
}

function createLocalFileHeader(nameBytes, crc, size, dosTime, dosDate) {
  const header = new Uint8Array(30 + nameBytes.length);
  writeUint32LE(header, 0, 0x04034b50);
  writeUint16LE(header, 4, 20);
  writeUint16LE(header, 6, 0x0800);
  writeUint16LE(header, 8, 0);
  writeUint16LE(header, 10, dosTime);
  writeUint16LE(header, 12, dosDate);
  writeUint32LE(header, 14, crc);
  writeUint32LE(header, 18, size);
  writeUint32LE(header, 22, size);
  writeUint16LE(header, 26, nameBytes.length);
  writeUint16LE(header, 28, 0);
  header.set(nameBytes, 30);
  return header;
}

function createCentralDirectoryHeader(nameBytes, crc, size, dosTime, dosDate, localOffset) {
  const header = new Uint8Array(46 + nameBytes.length);
  writeUint32LE(header, 0, 0x02014b50);
  writeUint16LE(header, 4, 20);
  writeUint16LE(header, 6, 20);
  writeUint16LE(header, 8, 0x0800);
  writeUint16LE(header, 10, 0);
  writeUint16LE(header, 12, dosTime);
  writeUint16LE(header, 14, dosDate);
  writeUint32LE(header, 16, crc);
  writeUint32LE(header, 20, size);
  writeUint32LE(header, 24, size);
  writeUint16LE(header, 28, nameBytes.length);
  writeUint16LE(header, 30, 0);
  writeUint16LE(header, 32, 0);
  writeUint16LE(header, 34, 0);
  writeUint16LE(header, 36, 0);
  writeUint32LE(header, 38, 0);
  writeUint32LE(header, 42, localOffset);
  header.set(nameBytes, 46);
  return header;
}

function createEndOfCentralDirectory(entryCount, centralSize, centralOffset) {
  const record = new Uint8Array(22);
  writeUint32LE(record, 0, 0x06054b50);
  writeUint16LE(record, 4, 0);
  writeUint16LE(record, 6, 0);
  writeUint16LE(record, 8, entryCount);
  writeUint16LE(record, 10, entryCount);
  writeUint32LE(record, 12, centralSize);
  writeUint32LE(record, 16, centralOffset);
  writeUint16LE(record, 20, 0);
  return record;
}

export async function buildZipBlobFromResults(results, options = {}) {
  if (!Array.isArray(results)) {
    throw new Error("results must be an array.");
  }

  const usedNames = new Set();
  const timestamp = options.timestamp instanceof Date ? options.timestamp : new Date();
  const { dosDate, dosTime } = toDosDateTime(timestamp);
  const onProgress = typeof options.onProgress === "function" ? options.onProgress : null;
  const yieldEvery = Math.max(1, Number(options.yieldEvery) || 1);
  const entries = [];

  for (let index = 0; index < results.length; index += 1) {
    const item = results[index];
    if (!item || typeof item !== "object") {
      continue;
    }

    const rawPayload = item.outputBytesB64 ?? item.outputBytes;
    if (rawPayload == null) {
      continue;
    }

    const rawName = item.outputFileName || item.fileName || `output_${index + 1}.xlsx`;
    const name = sanitizeEntryName(rawName, index, usedNames);
    const bytes = normalizeBytesSource(rawPayload, rawName);
    entries.push({ name, bytes });
  }

  if (!entries.length) {
    throw new Error("No ZIP entries were found in the provided results.");
  }

  const fileParts = [];
  const centralParts = [];
  let offset = 0;

  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    const nameBytes = textEncoder.encode(entry.name);
    const crc = crc32(entry.bytes);
    const size = entry.bytes.length;

    if (size > MAX_UINT32 || offset > MAX_UINT32) {
      throw new Error("ZIP64 is required for this output set; split the batch into smaller parts.");
    }

    const localHeader = createLocalFileHeader(nameBytes, crc, size, dosTime, dosDate);
    const centralHeader = createCentralDirectoryHeader(nameBytes, crc, size, dosTime, dosDate, offset);

    // XLSX files are already ZIP-compressed, so store them as-is.
    fileParts.push(localHeader, entry.bytes);
    centralParts.push(centralHeader);
    offset += localHeader.length + size;

    if (onProgress) {
      const completed = index + 1;
      const percent = Math.min(100, (completed / entries.length) * 100);
      onProgress({
        completed,
        total: entries.length,
        fileName: entry.name,
        percent,
      });
    }

    if ((index + 1) % yieldEvery === 0) {
      await sleepFrame();
    }
  }

  const centralOffset = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const eocd = createEndOfCentralDirectory(entries.length, centralSize, centralOffset);

  if (onProgress) {
    onProgress({
      completed: entries.length,
      total: entries.length,
      fileName: entries[entries.length - 1].name,
      percent: 100,
    });
  }

  return new Blob([...fileParts, ...centralParts, eocd], { type: "application/zip" });
}
