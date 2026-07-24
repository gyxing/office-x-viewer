import { XlsParseError } from '../errors';
import type { Biff8Anchor } from './types';

/** 解析 OfficeArt ClientAnchor，并将偏移归一化为单元格比例。 */
export function parseClientAnchor(bytes: Uint8Array): Biff8Anchor {
  if (bytes.length < 18) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'ClientAnchor 记录长度不足');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const fromColumn = view.getUint16(2, true);
  const fromColumnOffset = view.getUint16(4, true);
  const fromRow = view.getUint16(6, true);
  const fromRowOffset = view.getUint16(8, true);
  const toColumn = view.getUint16(10, true);
  const toColumnOffset = view.getUint16(12, true);
  const toRow = view.getUint16(14, true);
  const toRowOffset = view.getUint16(16, true);
  const reversed =
    toRow < fromRow ||
    (toRow === fromRow && toColumn < fromColumn) ||
    fromColumnOffset > 1023 ||
    toColumnOffset > 1023 ||
    fromRowOffset > 255 ||
    toRowOffset > 255;
  if (reversed) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      'ClientAnchor 坐标或偏移无效',
    );
  }
  return {
    from: {
      row: fromRow,
      column: fromColumn,
      rowFraction: fromRowOffset / 256,
      columnFraction: fromColumnOffset / 1024,
    },
    to: {
      row: toRow,
      column: toColumn,
      rowFraction: toRowOffset / 256,
      columnFraction: toColumnOffset / 1024,
    },
  };
}
