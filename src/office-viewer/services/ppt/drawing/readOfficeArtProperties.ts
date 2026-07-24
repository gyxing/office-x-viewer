import type { OfficeArtRecord } from '../../../shared/officeart';
import type { PptOfficeArtProperty } from './types';

/** 读取 OfficeArt 属性表，并将复杂属性数据与属性项重新关联。 */
export function readPptOfficeArtProperties(record: OfficeArtRecord | undefined) {
  const properties = new Map<number, PptOfficeArtProperty>();
  if (!record) return properties;
  const count = record.instance;
  const tableLength = count * 6;
  if (tableLength > record.data.length) return properties;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  let complexOffset = tableLength;

  for (let index = 0; index < count; index += 1) {
    const offset = index * 6;
    const operation = view.getUint16(offset, true);
    const value = view.getUint32(offset + 2, true);
    const isComplex = Boolean(operation & 0x8000);
    const property: PptOfficeArtProperty = {
      id: operation & 0x3fff,
      value,
      isBlip: Boolean(operation & 0x4000),
    };
    if (isComplex && complexOffset + value <= record.data.length) {
      property.complexData = record.data.subarray(
        complexOffset,
        complexOffset + value,
      );
      complexOffset += value;
    }
    properties.set(property.id, property);
  }
  return properties;
}
