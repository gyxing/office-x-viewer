import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { OFFICE_ART_RECORD } from './officeArtRecords';
import { parseClientAnchor } from './parseAnchors';
import { parseBlips } from './parseBlips';
import { parseOfficeArtRecords } from './parseOfficeArt';
import type {
  Biff8DrawingImage,
  Biff8DrawingShape,
  OfficeArtRecord,
} from './types';

type ShapeProperties = {
  blipIndex?: number;
  name?: string;
};

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

function parseShapeId(record: OfficeArtRecord | undefined) {
  if (!record || record.data.length < 4) return undefined;
  return new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  ).getUint32(0, true);
}

function decodeUtf16(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let value = '';
  for (let offset = 0; offset + 1 < bytes.length; offset += 2) {
    const character = view.getUint16(offset, true);
    if (!character) break;
    value += String.fromCharCode(character);
  }
  return value;
}

function parseShapeProperties(record: OfficeArtRecord | undefined) {
  const result: ShapeProperties = {};
  if (!record) return result;
  const propertyBytes = record.instance * 6;
  if (propertyBytes > record.data.length) return result;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  let complexOffset = propertyBytes;
  for (let index = 0; index < record.instance; index += 1) {
    const offset = index * 6;
    const operation = view.getUint16(offset, true);
    const propertyId = operation & 0x3fff;
    const value = view.getUint32(offset + 2, true);
    if (propertyId === 0x0104) result.blipIndex = value;
    if (operation & 0x8000) {
      if (complexOffset + value > record.data.length) break;
      if (propertyId === 0x0380) {
        result.name = decodeUtf16(
          record.data.subarray(complexOffset, complexOffset + value),
        );
      }
      complexOffset += value;
    }
  }
  return result;
}

/** 提取所有带 ClientAnchor 的形状，供图表与图片按工作簿顺序关联。 */
export function parseBiff8DrawingShapes(
  sheetBytes: Uint8Array,
  warnings: SpreadsheetWarning[] = [],
): Biff8DrawingShape[] {
  if (!sheetBytes.length) return [];
  const records = parseOfficeArtRecords(sheetBytes, warnings);
  const result: Biff8DrawingShape[] = [];
  for (const container of collectRecords(
    records,
    OFFICE_ART_RECORD.SP_CONTAINER,
  )) {
    const children = container.children ?? [];
    const fsp = children.find(
      (record) => record.type === OFFICE_ART_RECORD.FSP,
    );
    const anchor = children.find(
      (record) => record.type === OFFICE_ART_RECORD.CLIENT_ANCHOR,
    );
    if (!anchor) continue;
    const shapeId = parseShapeId(fsp);
    const properties = parseShapeProperties(
      children.find((record) => record.type === OFFICE_ART_RECORD.FOPT),
    );
    try {
      result.push({
        id: `xls-shape-${shapeId ?? result.length + 1}`,
        shapeId,
        shapeType: fsp?.instance,
        name: properties.name,
        blipIndex: properties.blipIndex,
        anchor: parseClientAnchor(anchor.data),
      });
    } catch {
      // 图片解析流程会产生带对象名称的锚点告警，这里只跳过无效候选。
    }
  }
  return result;
}

/** 关联工作簿 BLIP、工作表形状和 ClientAnchor，生成图片实例。 */
export function parseBiff8Drawings(
  groupBytes: Uint8Array,
  sheetBytes: Uint8Array,
  warnings: SpreadsheetWarning[] = [],
): Biff8DrawingImage[] {
  if (!groupBytes.length || !sheetBytes.length) return [];
  const groupRecords = parseOfficeArtRecords(groupBytes, warnings);
  const sheetRecords = parseOfficeArtRecords(sheetBytes, warnings);
  const blips = new Map(
    parseBlips(groupRecords, warnings).map((blip) => [blip.index, blip]),
  );
  const images: Biff8DrawingImage[] = [];

  for (const container of collectRecords(
    sheetRecords,
    OFFICE_ART_RECORD.SP_CONTAINER,
  )) {
    const children = container.children ?? [];
    const shapeId = parseShapeId(
      children.find((record) => record.type === OFFICE_ART_RECORD.FSP),
    );
    const properties = parseShapeProperties(
      children.find((record) => record.type === OFFICE_ART_RECORD.FOPT),
    );
    const anchorRecord = children.find(
      (record) => record.type === OFFICE_ART_RECORD.CLIENT_ANCHOR,
    );
    if (!properties.blipIndex || !anchorRecord) continue;
    const blip = blips.get(properties.blipIndex);
    if (!blip) {
      warnings.push({
        code: 'MISSING_BLIP',
        message: `形状 ${shapeId ?? 'unknown'} 引用的 BLIP ${
          properties.blipIndex
        } 不存在`,
        offset: container.offset,
      });
      continue;
    }
    try {
      images.push({
        id: `xls-image-${shapeId ?? images.length + 1}`,
        name: properties.name,
        alt: properties.name,
        format: blip.format,
        bytes: blip.bytes,
        compressed: blip.compressed,
        anchor: parseClientAnchor(anchorRecord.data),
        warnings: [...blip.warnings],
      });
    } catch (error) {
      warnings.push({
        code: 'INVALID_IMAGE_ANCHOR',
        message: `形状 ${shapeId ?? 'unknown'} 的锚点无效：${
          error instanceof Error ? error.message : '未知错误'
        }`,
        offset: anchorRecord.offset,
      });
    }
  }
  return images;
}
