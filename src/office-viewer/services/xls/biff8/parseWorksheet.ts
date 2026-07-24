import type { SpreadsheetWarning } from '../../spreadsheet/types';
import { XlsParseError } from '../errors';
import type {
  Biff8Cell,
  Biff8Merge,
  Biff8SheetDescriptor,
  Biff8WorkbookGlobals,
  Biff8Worksheet,
} from '../types';
import {
  Biff8Reader,
  Biff8RecordCursor,
  type ParseYieldState,
  yieldToBrowserIfNeeded,
} from './Biff8Reader';
import { BIFF8_RECORD, BIFF8_SUBSTREAM, BIFF8_VERSION } from './constants';
import { decodeBiff8Formula } from './formulas';
import { readBiff8UnicodeString } from './strings';

const ERROR_VALUES: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

function cellKey(row: number, column: number) {
  return `${row}:${column}`;
}

function readCellHeader(reader: Biff8Reader) {
  return {
    row: reader.readUint16(),
    column: reader.readUint16(),
    xfIndex: reader.readUint16(),
  };
}

function decodeRk(raw: number) {
  let value: number;
  if (raw & 0x02) {
    value = raw >> 2;
  } else {
    const bytes = new Uint8Array(8);
    const view = new DataView(bytes.buffer);
    view.setUint32(4, raw & 0xfffffffc, true);
    value = view.getFloat64(0, true);
  }
  return raw & 0x01 ? value / 100 : value;
}

function validateWorksheetBof(
  cursor: Biff8RecordCursor,
  descriptor: Biff8SheetDescriptor,
) {
  const record = cursor.next();
  if (!record || record.id !== BIFF8_RECORD.BOF || record.size < 4) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      `工作表“${descriptor.name}”缺少有效 BOF`,
      { offset: record?.offset, recordId: record?.id },
    );
  }
  const reader = new Biff8Reader(record.data);
  const version = reader.readUint16();
  const substreamType = reader.readUint16();
  if (
    version !== BIFF8_VERSION ||
    substreamType !== BIFF8_SUBSTREAM.WORKSHEET
  ) {
    throw new XlsParseError(
      version === BIFF8_VERSION
        ? 'INVALID_RECORD_DATA'
        : 'UNSUPPORTED_BIFF_VERSION',
      `工作表“${descriptor.name}”不是 BIFF8 Worksheet 子流`,
      { offset: record.offset, recordId: record.id },
    );
  }
}

function parseFormulaCell(
  data: Uint8Array,
  globals: Biff8WorkbookGlobals,
  warnings: SpreadsheetWarning[],
  sheetName: string,
  recordOffset: number,
) {
  const reader = new Biff8Reader(data);
  const header = readCellHeader(reader);
  const resultBytes = reader.readBytes(8);
  reader.readUint16();
  reader.readBytes(4);
  const tokenLength = reader.readUint16();
  const tokens = reader.readBytes(tokenLength);
  const decoded = decodeBiff8Formula(tokens, {
    row: header.row,
    column: header.column,
    definedNames: globals.definedNames,
    sheets: globals.sheets,
  });

  let value: Biff8Cell['value'];
  let cachedType: Biff8Cell['cachedType'];
  if (resultBytes[6] === 0xff && resultBytes[7] === 0xff) {
    const resultType = resultBytes[0];
    cachedType =
      resultType === 0
        ? 'string'
        : resultType === 1
        ? 'boolean'
        : resultType === 2
        ? 'error'
        : 'blank';
    value =
      resultType === 1
        ? Boolean(resultBytes[2])
        : resultType === 2
        ? ERROR_VALUES[resultBytes[2]] ?? '#ERROR!'
        : null;
  } else {
    const view = new DataView(
      resultBytes.buffer,
      resultBytes.byteOffset,
      resultBytes.byteLength,
    );
    value = view.getFloat64(0, true);
    cachedType = 'number';
  }
  if (decoded.unsupported) {
    warnings.push({
      code: 'UNSUPPORTED_FORMULA_TOKEN',
      message: '公式包含未识别 token，已保留缓存值和原始 token',
      sheetName,
      offset: recordOffset,
    });
  }
  return {
    ...header,
    value,
    cachedType,
    formula: decoded.formula,
    formulaTokens: decoded.formulaTokens,
  } satisfies Biff8Cell;
}

function rangesOverlap(left: Biff8Merge, right: Biff8Merge) {
  return !(
    left.endRow < right.startRow ||
    left.startRow > right.endRow ||
    left.endColumn < right.startColumn ||
    left.startColumn > right.endColumn
  );
}

function addMerge(
  merge: Biff8Merge,
  merges: Biff8Merge[],
  warnings: SpreadsheetWarning[],
  descriptor: Biff8SheetDescriptor,
  offset: number,
) {
  const invalid =
    merge.startRow > merge.endRow ||
    merge.startColumn > merge.endColumn ||
    merge.endRow > 0xffff ||
    merge.endColumn > 0xff;
  if (invalid || merges.some((existing) => rangesOverlap(existing, merge))) {
    warnings.push({
      code: invalid ? 'INVALID_MERGE' : 'OVERLAPPING_MERGE',
      message: invalid
        ? '已忽略坐标无效的合并单元格'
        : '合并区域发生重叠，已保留先出现的区域',
      sheetName: descriptor.name,
      offset,
    });
    return;
  }
  merges.push(merge);
}

/** 解析单个 BIFF8 Worksheet 子流，保留稀疏单元格和布局元数据。 */
export async function parseBiff8Worksheet(
  workbookStream: Uint8Array,
  descriptor: Biff8SheetDescriptor,
  globals: Biff8WorkbookGlobals,
  endOffset: number,
  yieldState: ParseYieldState,
): Promise<Biff8Worksheet> {
  const cursor = new Biff8RecordCursor(
    workbookStream,
    descriptor.streamOffset,
    endOffset,
  );
  validateWorksheetBof(cursor, descriptor);
  const cells = new Map<string, Biff8Cell>();
  const rows = new Map<number, Biff8Worksheet['rows'][number]>();
  const columns: Biff8Worksheet['columns'] = [];
  const merges: Biff8Merge[] = [];
  const warnings: SpreadsheetWarning[] = [];
  const drawingRecords: Biff8Worksheet['drawingRecords'] = [];
  const chartSubstreams: Biff8Worksheet['chartSubstreams'] = [];
  let dimensions: Biff8Worksheet['dimensions'];
  let defaultColumnWidth = 8.43;
  let defaultRowHeightTwips = 300;
  let hasDrawingRecords = false;
  let hasChartRecords = false;
  let pendingStringFormula: Biff8Cell | undefined;
  let reachedEof = false;
  let substreamDepth = 1;

  for (let record = cursor.next(); record; record = cursor.next()) {
    if (record.id === BIFF8_RECORD.BOF) {
      const nestedBof = new Biff8Reader(record.data);
      const version = nestedBof.readUint16();
      const substreamType = nestedBof.readUint16();
      if (
        version === BIFF8_VERSION &&
        substreamType === BIFF8_SUBSTREAM.CHART
      ) {
        hasChartRecords = true;
        const records = [record];
        let chartDepth = 1;
        while (chartDepth > 0) {
          const chartRecord = cursor.next();
          if (!chartRecord) {
            warnings.push({
              code: 'MISSING_CHART_EOF',
              message: '嵌入图表子流缺少 EOF，已忽略该图表',
              sheetName: descriptor.name,
              offset: record.offset,
            });
            records.length = 0;
            break;
          }
          records.push(chartRecord);
          if (chartRecord.id === BIFF8_RECORD.BOF) chartDepth += 1;
          if (chartRecord.id === BIFF8_RECORD.EOF) chartDepth -= 1;
        }
        if (records.length) {
          chartSubstreams.push({ offset: record.offset, records });
        }
        await yieldToBrowserIfNeeded(yieldState);
        continue;
      }
      substreamDepth += 1;
      await yieldToBrowserIfNeeded(yieldState);
      continue;
    }
    if (record.id === BIFF8_RECORD.EOF) {
      substreamDepth -= 1;
      if (substreamDepth === 0) {
        reachedEof = true;
        break;
      }
      await yieldToBrowserIfNeeded(yieldState);
      continue;
    }
    if (substreamDepth > 1) {
      await yieldToBrowserIfNeeded(yieldState);
      continue;
    }
    if (record.id !== BIFF8_RECORD.STRING) pendingStringFormula = undefined;
    const reader = new Biff8Reader(record.data);
    switch (record.id) {
      case BIFF8_RECORD.DIMENSIONS:
        dimensions = {
          firstRow: reader.readUint32(),
          lastRowExclusive: reader.readUint32(),
          firstColumn: reader.readUint16(),
          lastColumnExclusive: reader.readUint16(),
        };
        break;
      case BIFF8_RECORD.ROW: {
        const index = reader.readUint16();
        reader.readBytes(4);
        const heightTwips = reader.readUint16();
        reader.readBytes(4);
        const flags = reader.readUint32();
        rows.set(index, {
          index,
          heightTwips: flags & 0x0040 ? undefined : heightTwips,
          hidden: Boolean(flags & 0x0020),
          outlineLevel: flags & 0x0007,
        });
        break;
      }
      case BIFF8_RECORD.COLINFO: {
        const firstColumn = reader.readUint16();
        const lastColumn = reader.readUint16();
        const widthCharacters = reader.readUint16() / 256;
        const xfIndex = reader.readUint16();
        const flags = reader.readUint16();
        columns.push({
          firstColumn,
          lastColumn,
          widthCharacters,
          xfIndex,
          hidden: Boolean(flags & 0x0001),
          outlineLevel: (flags >> 8) & 0x07,
        });
        break;
      }
      case BIFF8_RECORD.DEFCOLWIDTH:
        defaultColumnWidth = reader.readUint16();
        break;
      case BIFF8_RECORD.STANDARDWIDTH:
        defaultColumnWidth = reader.readUint16() / 256;
        break;
      case BIFF8_RECORD.DEFAULTROWHEIGHT:
        reader.readUint16();
        defaultRowHeightTwips = reader.readUint16();
        break;
      case BIFF8_RECORD.NUMBER: {
        const header = readCellHeader(reader);
        cells.set(cellKey(header.row, header.column), {
          ...header,
          value: reader.readFloat64(),
          cachedType: 'number',
        });
        break;
      }
      case BIFF8_RECORD.RK: {
        const header = readCellHeader(reader);
        cells.set(cellKey(header.row, header.column), {
          ...header,
          value: decodeRk(reader.readUint32()),
          cachedType: 'number',
        });
        break;
      }
      case BIFF8_RECORD.MULRK: {
        const row = reader.readUint16();
        const firstColumn = reader.readUint16();
        let column = firstColumn;
        while (reader.remaining > 2) {
          const xfIndex = reader.readUint16();
          cells.set(cellKey(row, column), {
            row,
            column,
            xfIndex,
            value: decodeRk(reader.readUint32()),
            cachedType: 'number',
          });
          column += 1;
        }
        const lastColumn = reader.readUint16();
        if (lastColumn !== column - 1) {
          throw new XlsParseError(
            'INVALID_RECORD_DATA',
            'MulRK 末列与单元格数量不一致',
            { offset: record.offset, recordId: record.id },
          );
        }
        break;
      }
      case BIFF8_RECORD.LABELSST: {
        const header = readCellHeader(reader);
        const sharedStringIndex = reader.readUint32();
        const value = globals.sharedStrings[sharedStringIndex];
        if (value === undefined) {
          throw new XlsParseError(
            'INVALID_RECORD_DATA',
            `LabelSst 索引 ${sharedStringIndex} 超出 SST`,
            { offset: record.offset, recordId: record.id },
          );
        }
        cells.set(cellKey(header.row, header.column), {
          ...header,
          value,
          cachedType: 'string',
        });
        break;
      }
      case BIFF8_RECORD.BOOLERR: {
        const header = readCellHeader(reader);
        const rawValue = reader.readUint8();
        const isError = Boolean(reader.readUint8());
        cells.set(cellKey(header.row, header.column), {
          ...header,
          value: isError
            ? ERROR_VALUES[rawValue] ?? '#ERROR!'
            : Boolean(rawValue),
          cachedType: isError ? 'error' : 'boolean',
        });
        break;
      }
      case BIFF8_RECORD.BLANK: {
        const header = readCellHeader(reader);
        cells.set(cellKey(header.row, header.column), {
          ...header,
          value: null,
          cachedType: 'blank',
        });
        break;
      }
      case BIFF8_RECORD.MULBLANK: {
        const row = reader.readUint16();
        const firstColumn = reader.readUint16();
        let column = firstColumn;
        while (reader.remaining > 2) {
          cells.set(cellKey(row, column), {
            row,
            column,
            xfIndex: reader.readUint16(),
            value: null,
            cachedType: 'blank',
          });
          column += 1;
        }
        const lastColumn = reader.readUint16();
        if (lastColumn !== column - 1) {
          throw new XlsParseError(
            'INVALID_RECORD_DATA',
            'MulBlank 末列与单元格数量不一致',
            { offset: record.offset, recordId: record.id },
          );
        }
        break;
      }
      case BIFF8_RECORD.FORMULA: {
        const cell = parseFormulaCell(
          record.data,
          globals,
          warnings,
          descriptor.name,
          record.offset,
        );
        cells.set(cellKey(cell.row, cell.column), cell);
        if (cell.cachedType === 'string') pendingStringFormula = cell;
        break;
      }
      case BIFF8_RECORD.STRING:
        if (pendingStringFormula) {
          pendingStringFormula.value = readBiff8UnicodeString(reader).value;
          pendingStringFormula = undefined;
        } else {
          warnings.push({
            code: 'ORPHAN_FORMULA_STRING',
            message: '发现未紧跟字符串公式的 String 记录，已忽略',
            sheetName: descriptor.name,
            offset: record.offset,
          });
        }
        break;
      case BIFF8_RECORD.MERGECELLS: {
        const count = reader.readUint16();
        for (let index = 0; index < count; index += 1) {
          addMerge(
            {
              startRow: reader.readUint16(),
              endRow: reader.readUint16(),
              startColumn: reader.readUint16(),
              endColumn: reader.readUint16(),
            },
            merges,
            warnings,
            descriptor,
            record.offset,
          );
        }
        break;
      }
      case BIFF8_RECORD.MSODRAWING:
      case BIFF8_RECORD.TXO:
      case BIFF8_RECORD.IMDATA:
      case BIFF8_RECORD.NOTE:
      case BIFF8_RECORD.OBJ:
        hasDrawingRecords = true;
        {
          const chunks = [record.data];
          while (cursor.peek()?.id === BIFF8_RECORD.CONTINUE) {
            chunks.push(cursor.next()!.data);
          }
          drawingRecords.push({
            recordId: record.id,
            offset: record.offset,
            chunks,
          });
        }
        break;
      case BIFF8_RECORD.WINDOW2:
        if (record.size < 10) {
          throw new XlsParseError(
            'INVALID_RECORD_DATA',
            'Window2 记录长度无效',
            { offset: record.offset, recordId: record.id },
          );
        }
        break;
      default:
        break;
    }
    await yieldToBrowserIfNeeded(yieldState);
  }

  if (!reachedEof) {
    warnings.push({
      code: 'MISSING_SHEET_EOF',
      message: '工作表子流在下一个已验证边界处结束，未找到 EOF',
      sheetName: descriptor.name,
      offset: endOffset,
    });
  }
  if (pendingStringFormula) {
    warnings.push({
      code: 'MISSING_FORMULA_STRING',
      message: '字符串公式缺少紧随其后的 String 缓存记录',
      sheetName: descriptor.name,
    });
  }

  return {
    descriptor,
    cells: Array.from(cells.values()).sort(
      (left, right) => left.row - right.row || left.column - right.column,
    ),
    rows: Array.from(rows.values()).sort(
      (left, right) => left.index - right.index,
    ),
    columns: columns.sort(
      (left, right) => left.firstColumn - right.firstColumn,
    ),
    merges,
    defaultColumnWidth,
    defaultRowHeightTwips,
    dimensions,
    hasDrawingRecords,
    hasChartRecords,
    chartSubstreams,
    drawingRecords,
    warnings,
  };
}
