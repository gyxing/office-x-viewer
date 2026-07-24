import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { OFFICE_ART_RECORD } from './officeArtRecords';
import { parseOfficeArtRecords } from './parseOfficeArt';
import type {
  Biff8DrawingImageFormat,
  OfficeArtRecord,
  ParsedBlip,
} from './types';

function findSignature(bytes: Uint8Array, signature: number[]) {
  const limit = Math.min(bytes.length - signature.length, 96);
  for (let offset = 0; offset <= limit; offset += 1) {
    if (signature.every((value, index) => bytes[offset + index] === value)) {
      return offset;
    }
  }
  return -1;
}

function sniffRaster(bytes: Uint8Array) {
  const signatures: Array<{
    format: Biff8DrawingImageFormat;
    signature: number[];
  }> = [
    {
      format: 'png',
      signature: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
    },
    { format: 'jpeg', signature: [0xff, 0xd8, 0xff] },
    { format: 'gif', signature: [0x47, 0x49, 0x46, 0x38] },
  ];
  for (const item of signatures) {
    const offset = findSignature(bytes, item.signature);
    if (offset >= 0) {
      return { format: item.format, bytes: bytes.subarray(offset) };
    }
  }
  const dibHeaders = new Set([12, 40, 52, 56, 108, 124]);
  const limit = Math.min(bytes.length - 4, 96);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = 0; offset <= limit; offset += 1) {
    if (dibHeaders.has(view.getUint32(offset, true))) {
      return { format: 'dib' as const, bytes: bytes.subarray(offset) };
    }
  }
  return undefined;
}

function formatFromRecordType(type: number): Biff8DrawingImageFormat {
  if (type === OFFICE_ART_RECORD.BLIP_PNG) return 'png';
  if (type === OFFICE_ART_RECORD.BLIP_JPEG) return 'jpeg';
  if (type === OFFICE_ART_RECORD.BLIP_DIB) return 'dib';
  if (type === OFFICE_ART_RECORD.BLIP_WMF) return 'wmf';
  if (type === OFFICE_ART_RECORD.BLIP_EMF) return 'emf';
  if (type === OFFICE_ART_RECORD.BLIP_PICT) return 'pict';
  return 'unknown';
}

function extractMetafilePayload(
  bytes: Uint8Array,
  warnings: SpreadsheetWarning[],
) {
  for (const uidLength of [16, 32]) {
    const headerOffset = uidLength;
    if (headerOffset + 34 > bytes.length) continue;
    const view = new DataView(
      bytes.buffer,
      bytes.byteOffset + headerOffset,
      bytes.length - headerOffset,
    );
    const savedSize = view.getUint32(28, true);
    const compression = view.getUint8(32);
    const dataOffset = headerOffset + 34;
    if (savedSize > 0 && dataOffset + savedSize <= bytes.length) {
      if (compression !== 0x00 && compression !== 0xfe) {
        warnings.push({
          code: 'UNKNOWN_METAFILE_COMPRESSION',
          message: `暂不支持 metafile 压缩标记 0x${compression
            .toString(16)
            .toUpperCase()}`,
        });
        return undefined;
      }
      return {
        bytes: bytes.subarray(dataOffset, dataOffset + savedSize),
        compressed: compression === 0x00,
      };
    }
  }
  warnings.push({
    code: 'INVALID_METAFILE_BLIP',
    message: 'metafile BLIP Header 或数据长度无效',
  });
  return undefined;
}

function extractBlip(
  record: OfficeArtRecord,
  index: number,
  warnings: SpreadsheetWarning[],
): ParsedBlip | undefined {
  const declaredFormat = formatFromRecordType(record.type);
  const sniffed = sniffRaster(record.data);
  if (sniffed) {
    return {
      index,
      ...sniffed,
      warnings: [],
    };
  }
  if (declaredFormat === 'wmf' || declaredFormat === 'emf') {
    const payload = extractMetafilePayload(record.data, warnings);
    if (!payload) return undefined;
    return {
      index,
      format: declaredFormat,
      ...payload,
      warnings: [],
    };
  }
  if (declaredFormat === 'pict') {
    warnings.push({
      code: 'UNSUPPORTED_PICT',
      message: '暂不支持 PICT 图片，已跳过该对象',
    });
    return undefined;
  }
  warnings.push({
    code: 'UNKNOWN_BLIP_FORMAT',
    message: `无法识别 BLIP 图片格式 0x${record.type
      .toString(16)
      .toUpperCase()}`,
  });
  return undefined;
}

function collectRecords(
  records: OfficeArtRecord[],
  type: number,
  result: OfficeArtRecord[] = [],
) {
  for (const record of records) {
    if (record.type === type) result.push(record);
    if (record.children) collectRecords(record.children, type, result);
  }
  return result;
}

function blipRecordTypeFromBse(value: number) {
  if (value === 2) return OFFICE_ART_RECORD.BLIP_EMF;
  if (value === 3) return OFFICE_ART_RECORD.BLIP_WMF;
  if (value === 4) return OFFICE_ART_RECORD.BLIP_PICT;
  if (value === 5) return OFFICE_ART_RECORD.BLIP_JPEG;
  if (value === 6) return OFFICE_ART_RECORD.BLIP_PNG;
  if (value === 7) return OFFICE_ART_RECORD.BLIP_DIB;
  return 0;
}

/** 从 Dgg/BStore 中提取并编号可显示的 BLIP 数据。 */
export function parseBlips(
  records: OfficeArtRecord[],
  warnings: SpreadsheetWarning[],
) {
  const result: ParsedBlip[] = [];
  const bseRecords = collectRecords(records, OFFICE_ART_RECORD.BSE);
  bseRecords.forEach((bse, position) => {
    const index = position + 1;
    if (bse.data.length < 36) {
      warnings.push({
        code: 'INVALID_BSE',
        message: `BSE ${index} 记录长度不足`,
        offset: bse.offset,
      });
      return;
    }
    const nameLength = bse.data[33];
    const embeddedOffset = 36 + nameLength;
    if (embeddedOffset > bse.data.length) {
      warnings.push({
        code: 'INVALID_BSE',
        message: `BSE ${index} 名称长度越界`,
        offset: bse.offset,
      });
      return;
    }
    let embedded: OfficeArtRecord | undefined;
    try {
      embedded = parseOfficeArtRecords(
        bse.data.subarray(embeddedOffset),
        warnings,
      )[0];
    } catch (error) {
      warnings.push({
        code: 'INVALID_EMBEDDED_BLIP',
        message: `BSE ${index} 的嵌入 BLIP 结构无效：${
          error instanceof Error ? error.message : '未知错误'
        }`,
        offset: bse.offset,
      });
    }
    if (!embedded) {
      const sniffed = sniffRaster(bse.data.subarray(embeddedOffset));
      if (sniffed) {
        result.push({ index, ...sniffed, warnings: [] });
        return;
      }
      const fallbackType = blipRecordTypeFromBse(bse.data[0]);
      if (!fallbackType) return;
      embedded = {
        version: 0,
        instance: 0,
        type: fallbackType,
        length: bse.data.length - embeddedOffset,
        offset: bse.offset + 8 + embeddedOffset,
        data: bse.data.subarray(embeddedOffset),
      };
    }
    const parsed = extractBlip(embedded, index, warnings);
    if (parsed) result.push(parsed);
  });
  return result;
}
