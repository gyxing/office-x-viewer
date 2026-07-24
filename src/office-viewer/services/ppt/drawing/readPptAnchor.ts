import type { OfficeArtRecord } from '../../../shared/officeart';
import type { PptShapeAnchor } from './types';

const MASTER_UNIT_TO_PX = 1 / 8;

/** 将 PPT ClientAnchor 的“上、左、右、下”主单位转换为像素。 */
export function readPptAnchor(
  record: OfficeArtRecord | undefined,
): PptShapeAnchor | undefined {
  if (!record || record.data.length < 8) return undefined;
  const view = new DataView(
    record.data.buffer,
    record.data.byteOffset,
    record.data.byteLength,
  );
  const top = view.getInt16(0, true) * MASTER_UNIT_TO_PX;
  const left = view.getInt16(2, true) * MASTER_UNIT_TO_PX;
  const right = view.getInt16(4, true) * MASTER_UNIT_TO_PX;
  const bottom = view.getInt16(6, true) * MASTER_UNIT_TO_PX;
  return {
    x: Math.min(left, right),
    y: Math.min(top, bottom),
    width: Math.abs(right - left),
    height: Math.abs(bottom - top),
  };
}
