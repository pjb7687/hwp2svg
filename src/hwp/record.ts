/**
 * HWP 5.0 data record parser (browser-compatible, no Node.js Buffer).
 *
 * Record header is 32 bits:
 *   - TagID:  bits 0-9   (10 bits)
 *   - Level:  bits 10-19 (10 bits)
 *   - Size:   bits 20-31 (12 bits)
 *
 * If size == 0xFFF (4095), an additional DWORD follows with the actual size.
 */

export interface HwpRecord {
  tagId: number;
  level: number;
  size: number;
  data: Uint8Array;
  offset: number;
}

export function parseRecords(buf: Uint8Array): HwpRecord[] {
  const records: HwpRecord[] = [];
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  let pos = 0;

  while (pos < buf.length) {
    if (pos + 4 > buf.length) break;

    const header = view.getUint32(pos, true);
    const recordOffset = pos;
    pos += 4;

    const tagId = header & 0x3FF;
    const level = (header >> 10) & 0x3FF;
    let size = (header >> 20) & 0xFFF;

    if (size === 0xFFF) {
      if (pos + 4 > buf.length) break;
      size = view.getUint32(pos, true);
      pos += 4;
    }

    if (pos + size > buf.length) {
      const data = buf.subarray(pos, buf.length);
      records.push({ tagId, level, size: data.length, data, offset: recordOffset });
      break;
    }

    const data = buf.subarray(pos, pos + size);
    records.push({ tagId, level, size, data, offset: recordOffset });
    pos += size;
  }

  return records;
}

/** Read a UTF-16LE string from a Uint8Array. */
export function readWString(buf: Uint8Array, offset: number, charCount: number): string {
  const codes: number[] = [];
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  for (let i = 0; i < charCount; i++) {
    codes.push(view.getUint16(offset + i * 2, true));
  }
  return String.fromCharCode(...codes);
}

/** Read a length-prefixed WCHAR string (WORD length + WCHAR[]). Returns [string, bytesConsumed]. */
export function readLPWString(buf: Uint8Array, offset: number): [string, number] {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const len = view.getUint16(offset, true);
  const str = readWString(buf, offset + 2, len);
  return [str, 2 + len * 2];
}

/** Helper to create a DataView from a Uint8Array record data. */
export function dataView(data: Uint8Array): DataView {
  return new DataView(data.buffer, data.byteOffset, data.byteLength);
}
