import { XlsParseError } from '../errors';
import type { DecodedBitmap } from './types';

type DibInfo = {
  width: number;
  height: number;
  topDown: boolean;
  bitCount: number;
  compression: number;
  pixelOffset: number;
  palette: Array<[number, number, number, number]>;
  masks: [number, number, number, number];
};

function ensureRange(bytes: Uint8Array, offset: number, length: number) {
  if (
    !Number.isSafeInteger(offset) ||
    !Number.isSafeInteger(length) ||
    offset < 0 ||
    length < 0 ||
    offset + length > bytes.length
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB 数据范围越界');
  }
}

function readPalette(
  bytes: Uint8Array,
  offset: number,
  count: number,
  entrySize: 3 | 4,
) {
  ensureRange(bytes, offset, count * entrySize);
  const palette: DibInfo['palette'] = [];
  for (let index = 0; index < count; index += 1) {
    const base = offset + index * entrySize;
    palette.push([
      bytes[base + 2],
      bytes[base + 1],
      bytes[base],
      entrySize === 4 && bytes[base + 3] ? bytes[base + 3] : 255,
    ]);
  }
  return palette;
}

function parseCoreHeader(bytes: Uint8Array, view: DataView): DibInfo {
  const width = view.getUint16(4, true);
  const height = view.getUint16(6, true);
  const planes = view.getUint16(8, true);
  const bitCount = view.getUint16(10, true);
  if (!width || !height || planes !== 1) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB Core Header 无效');
  }
  const paletteCount = bitCount <= 8 ? 2 ** bitCount : 0;
  const pixelOffset = 12 + paletteCount * 3;
  return {
    width,
    height,
    topDown: false,
    bitCount,
    compression: 0,
    pixelOffset,
    palette: readPalette(bytes, 12, paletteCount, 3),
    masks: [0, 0, 0, 0],
  };
}

function parseInfoHeader(
  bytes: Uint8Array,
  view: DataView,
  headerSize: number,
): DibInfo {
  ensureRange(bytes, 0, headerSize);
  const width = view.getInt32(4, true);
  const signedHeight = view.getInt32(8, true);
  const planes = view.getUint16(12, true);
  const bitCount = view.getUint16(14, true);
  const compression = view.getUint32(16, true);
  const colorCount = view.getUint32(32, true);
  if (
    width <= 0 ||
    signedHeight === 0 ||
    planes !== 1 ||
    ![1, 4, 8, 16, 24, 32].includes(bitCount) ||
    ![0, 1, 2, 3, 6].includes(compression)
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB Info Header 无效');
  }
  if (
    (compression === 1 && bitCount !== 8) ||
    (compression === 2 && bitCount !== 4) ||
    (signedHeight < 0 && (compression === 1 || compression === 2)) ||
    ((compression === 3 || compression === 6) &&
      bitCount !== 16 &&
      bitCount !== 32)
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB 压缩方式与位深不匹配');
  }

  let paletteOffset = headerSize;
  let masks: DibInfo['masks'] =
    bitCount === 16
      ? [0x7c00, 0x03e0, 0x001f, 0]
      : bitCount === 32
      ? [0x00ff0000, 0x0000ff00, 0x000000ff, 0]
      : [0, 0, 0, 0];
  if (compression === 3 || compression === 6) {
    const masksOffset = headerSize >= 52 ? 40 : headerSize;
    const maskCount = compression === 6 || headerSize >= 56 ? 4 : 3;
    ensureRange(bytes, masksOffset, maskCount * 4);
    masks = [
      view.getUint32(masksOffset, true),
      view.getUint32(masksOffset + 4, true),
      view.getUint32(masksOffset + 8, true),
      maskCount === 4 ? view.getUint32(masksOffset + 12, true) : 0,
    ];
    if (headerSize === 40) paletteOffset += maskCount * 4;
  }
  const paletteCount = bitCount <= 8 ? colorCount || 2 ** bitCount : colorCount;
  const palette = readPalette(bytes, paletteOffset, paletteCount, 4);
  return {
    width,
    height: Math.abs(signedHeight),
    topDown: signedHeight < 0,
    bitCount,
    compression,
    pixelOffset: paletteOffset + paletteCount * 4,
    palette,
    masks,
  };
}

function parseDibInfo(bytes: Uint8Array) {
  ensureRange(bytes, 0, 12);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerSize = view.getUint32(0, true);
  if (headerSize === 12) return parseCoreHeader(bytes, view);
  if (![40, 52, 56, 108, 124].includes(headerSize)) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      `暂不支持 DIB Header 大小 ${headerSize}`,
    );
  }
  return parseInfoHeader(bytes, view, headerSize);
}

function extractMasked(raw: number, mask: number, fallback: number) {
  if (!mask) return fallback;
  let shift = 0;
  let shiftedMask = mask >>> 0;
  while ((shiftedMask & 1) === 0 && shift < 32) {
    shiftedMask >>>= 1;
    shift += 1;
  }
  const component = ((raw & mask) >>> shift) >>> 0;
  return Math.round((component / shiftedMask) * 255);
}

function writeRgba(
  rgba: Uint8ClampedArray,
  width: number,
  row: number,
  column: number,
  color: [number, number, number, number],
) {
  const offset = (row * width + column) * 4;
  rgba[offset] = color[0];
  rgba[offset + 1] = color[1];
  rgba[offset + 2] = color[2];
  rgba[offset + 3] = color[3];
}

function decodeUncompressed(
  bytes: Uint8Array,
  info: DibInfo,
  rgba: Uint8ClampedArray,
) {
  const stride = Math.floor((info.width * info.bitCount + 31) / 32) * 4;
  const required = stride * info.height;
  if (!Number.isSafeInteger(required)) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB 行跨度计算溢出');
  }
  ensureRange(bytes, info.pixelOffset, required);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let sourceRow = 0; sourceRow < info.height; sourceRow += 1) {
    const targetRow = info.topDown ? sourceRow : info.height - 1 - sourceRow;
    const rowOffset = info.pixelOffset + sourceRow * stride;
    for (let column = 0; column < info.width; column += 1) {
      let color: [number, number, number, number];
      if (info.bitCount <= 8) {
        const bitOffset = column * info.bitCount;
        const rawByte = bytes[rowOffset + Math.floor(bitOffset / 8)];
        const shift = 8 - info.bitCount - (bitOffset % 8);
        const paletteIndex = (rawByte >> shift) & (2 ** info.bitCount - 1);
        color = info.palette[paletteIndex] ?? [0, 0, 0, 255];
      } else if (info.bitCount === 16) {
        const raw = view.getUint16(rowOffset + column * 2, true);
        color = [
          extractMasked(raw, info.masks[0], 0),
          extractMasked(raw, info.masks[1], 0),
          extractMasked(raw, info.masks[2], 0),
          extractMasked(raw, info.masks[3], 255),
        ];
      } else if (info.bitCount === 24) {
        const offset = rowOffset + column * 3;
        color = [bytes[offset + 2], bytes[offset + 1], bytes[offset], 255];
      } else {
        const raw = view.getUint32(rowOffset + column * 4, true);
        color = [
          extractMasked(raw, info.masks[0], 0),
          extractMasked(raw, info.masks[1], 0),
          extractMasked(raw, info.masks[2], 0),
          extractMasked(raw, info.masks[3], 255),
        ];
      }
      writeRgba(rgba, info.width, targetRow, column, color);
    }
  }
}

function decodeRle(bytes: Uint8Array, info: DibInfo, rgba: Uint8ClampedArray) {
  let offset = info.pixelOffset;
  let x = 0;
  let y = 0;
  let ended = false;
  const writeIndex = (paletteIndex: number) => {
    if (x >= info.width || y >= info.height) {
      throw new XlsParseError('INVALID_RECORD_DATA', 'DIB RLE 写入越界');
    }
    writeRgba(
      rgba,
      info.width,
      info.height - 1 - y,
      x,
      info.palette[paletteIndex] ?? [0, 0, 0, 255],
    );
    x += 1;
  };

  while (offset + 2 <= bytes.length && !ended) {
    const count = bytes[offset++];
    const value = bytes[offset++];
    if (count) {
      for (let index = 0; index < count; index += 1) {
        writeIndex(
          info.compression === 1
            ? value
            : index % 2
            ? value & 0x0f
            : value >> 4,
        );
      }
      continue;
    }
    if (value === 0) {
      x = 0;
      y += 1;
      if (y > info.height) {
        throw new XlsParseError('INVALID_RECORD_DATA', 'DIB RLE 行号越界');
      }
    } else if (value === 1) {
      ended = true;
    } else if (value === 2) {
      ensureRange(bytes, offset, 2);
      x += bytes[offset++];
      y += bytes[offset++];
      if (x > info.width || y >= info.height) {
        throw new XlsParseError('INVALID_RECORD_DATA', 'DIB RLE Delta 越界');
      }
    } else {
      const pixelCount = value;
      const byteCount =
        info.compression === 1 ? pixelCount : Math.ceil(pixelCount / 2);
      ensureRange(bytes, offset, byteCount);
      for (let index = 0; index < pixelCount; index += 1) {
        const packed = bytes[offset + Math.floor(index / 2)];
        writeIndex(
          info.compression === 1
            ? bytes[offset + index]
            : index % 2
            ? packed & 0x0f
            : packed >> 4,
        );
      }
      offset += byteCount;
      if (byteCount % 2) offset += 1;
    }
  }
  if (!ended) {
    throw new XlsParseError('TRUNCATED_RECORD', 'DIB RLE 数据缺少结束标记');
  }
}

/** 将 DIB 解码为顶部起始的 RGBA 像素，不依赖 DOM 或 Canvas。 */
export function decodeDib(bytes: Uint8Array): DecodedBitmap {
  const info = parseDibInfo(bytes);
  const pixelCount = info.width * info.height;
  const byteLength = pixelCount * 4;
  if (
    !Number.isSafeInteger(pixelCount) ||
    !Number.isSafeInteger(byteLength) ||
    pixelCount <= 0
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB 像素数量溢出');
  }
  let rgba: Uint8ClampedArray;
  try {
    rgba = new Uint8ClampedArray(byteLength);
  } catch {
    throw new XlsParseError('INVALID_RECORD_DATA', 'DIB 像素缓冲区分配失败');
  }
  const background = info.palette[0] ?? [0, 0, 0, 0];
  for (let row = 0; row < info.height; row += 1) {
    for (let column = 0; column < info.width; column += 1) {
      writeRgba(rgba, info.width, row, column, background);
    }
  }
  if (info.compression === 1 || info.compression === 2) {
    decodeRle(bytes, info, rgba);
  } else {
    decodeUncompressed(bytes, info, rgba);
  }
  return { width: info.width, height: info.height, rgba };
}
