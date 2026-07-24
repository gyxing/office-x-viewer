import type {
  SpreadsheetCell,
  SpreadsheetCellStyle,
  SpreadsheetColumn,
  SpreadsheetMerge,
  SpreadsheetSheet,
  SpreadsheetWorkbook,
} from '../spreadsheet/types';
import type { PortableResource } from '../parsing/protocol/messages';
import { BIFF8_RECORD } from './biff8/constants';
import { formatBiff8Value } from './biff8/numberFormats';
import { parseBiff8Charts } from './chart/parseCharts';
import { createPortableImageResource } from './drawing/createPortableImageResource';
import {
  parseBiff8Drawings,
  parseBiff8DrawingShapes,
} from './drawing/parseDrawings';
import type {
  Biff8BorderStyle,
  Biff8Cell,
  Biff8CellFormat,
  Biff8SheetDescriptor,
  Biff8Workbook,
  Biff8Worksheet,
} from './types';

const CSS_DPI = 96;
const TWIPS_PER_INCH = 1440;
const POINTS_PER_INCH = 72;
const DEFAULT_COLUMN_PIXELS = 64;
const DEFAULT_ROW_PIXELS = 20;

function columnLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cellRef(row: number, column: number) {
  return `${columnLabel(column + 1)}${row + 1}`;
}

function twipsToPixels(value: number) {
  return (value / TWIPS_PER_INCH) * CSS_DPI;
}

function pointsToPixels(value: number) {
  return (value / POINTS_PER_INCH) * CSS_DPI;
}

function columnWidthToPixels(width: number) {
  if (!Number.isFinite(width) || width <= 0) return DEFAULT_COLUMN_PIXELS;
  return Math.max(1, Math.floor(((width * 256 + 18) / 256) * 7));
}

const SYSTEM_COLORS: Record<number, string> = {
  0: '#000000',
  1: '#ffffff',
  2: '#ff0000',
  3: '#00ff00',
  4: '#0000ff',
  5: '#ffff00',
  6: '#ff00ff',
  7: '#00ffff',
};

function resolveColor(index: number | undefined, palette: string[]) {
  if (index === undefined || index === 64 || index === 0x7fff) return undefined;
  return SYSTEM_COLORS[index] ?? palette[index - 8];
}

function borderToCss(border: Biff8BorderStyle | undefined, palette: string[]) {
  if (!border?.style) return undefined;
  const width = border.style === 2 ? 2 : border.style === 5 ? 3 : 1;
  const lineStyle =
    border.style === 3 || border.style === 8
      ? 'dashed'
      : border.style === 4 || border.style === 7
      ? 'dotted'
      : border.style === 6
      ? 'double'
      : 'solid';
  return `${width}px ${lineStyle} ${
    resolveColor(border.colorIndex, palette) ?? '#000000'
  }`;
}

function alignmentFromValue(value: number | undefined) {
  if (value === 1) return 'left' as const;
  if (value === 2 || value === 6) return 'center' as const;
  if (value === 3) return 'right' as const;
  if (value === 5 || value === 7) return 'justify' as const;
  return undefined;
}

function verticalAlignmentFromValue(value: number | undefined) {
  if (value === 0) return 'top' as const;
  if (value === 1) return 'middle' as const;
  if (value === 2) return 'bottom' as const;
  return undefined;
}

function resolveCellStyle(
  xf: Biff8CellFormat | undefined,
  workbook: Biff8Workbook,
): SpreadsheetCellStyle | undefined {
  if (!xf) return undefined;
  const { globals } = workbook;
  const font = globals.fonts[xf.fontIndex];
  const borderTop = borderToCss(xf.topBorder, globals.palette);
  const borderRight = borderToCss(xf.rightBorder, globals.palette);
  const borderBottom = borderToCss(xf.bottomBorder, globals.palette);
  const borderLeft = borderToCss(xf.leftBorder, globals.palette);
  const style: SpreadsheetCellStyle = {
    bold: font?.bold,
    italic: font?.italic,
    underline: font?.underline,
    color: resolveColor(font?.colorIndex, globals.palette),
    fontFamily: font?.name,
    fontSize: font ? pointsToPixels(font.heightTwips / 20) : undefined,
    backgroundColor: xf.fillPattern
      ? resolveColor(xf.fillForegroundColorIndex, globals.palette)
      : undefined,
    horizontalAlign: alignmentFromValue(xf.horizontalAlign),
    verticalAlign: verticalAlignmentFromValue(xf.verticalAlign),
    wrapText: xf.wrapText,
    border: Boolean(borderTop || borderRight || borderBottom || borderLeft),
    borderTop,
    borderRight,
    borderBottom,
    borderLeft,
  };
  return Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as SpreadsheetCellStyle;
}

function computeUsedRange(sheet: Biff8Worksheet) {
  let maxRow = Math.max(0, (sheet.dimensions?.lastRowExclusive ?? 0) - 1);
  let maxColumn = Math.max(0, (sheet.dimensions?.lastColumnExclusive ?? 0) - 1);
  for (const cell of sheet.cells) {
    maxRow = Math.max(maxRow, cell.row);
    maxColumn = Math.max(maxColumn, cell.column);
  }
  for (const row of sheet.rows) maxRow = Math.max(maxRow, row.index);
  for (const column of sheet.columns) {
    maxColumn = Math.max(maxColumn, column.lastColumn);
  }
  for (const merge of sheet.merges) {
    maxRow = Math.max(maxRow, merge.endRow);
    maxColumn = Math.max(maxColumn, merge.endColumn);
  }
  return { maxRow, maxColumn };
}

function findColumnInfo(sheet: Biff8Worksheet, column: number) {
  return sheet.columns.find(
    (item) => column >= item.firstColumn && column <= item.lastColumn,
  );
}

function adaptCell(
  source: Biff8Cell | undefined,
  row: number,
  column: number,
  fallbackXfIndex: number | undefined,
  workbook: Biff8Workbook,
): SpreadsheetCell {
  const xfIndex = source?.xfIndex ?? fallbackXfIndex;
  const xf =
    xfIndex === undefined ? undefined : workbook.globals.cellFormats[xfIndex];
  const rawValue = source?.value;
  const format = xf ? workbook.globals.formats.get(xf.formatIndex) : undefined;
  return {
    ref: cellRef(row, column),
    rowIndex: row + 1,
    columnIndex: column + 1,
    value: source
      ? formatBiff8Value(rawValue ?? null, format, workbook.globals.date1904)
      : '',
    rawValue: source && rawValue !== null ? String(rawValue) : undefined,
    type: source?.cachedType,
    styleId: xfIndex,
    style: resolveCellStyle(xf, workbook),
    formula: source?.formula,
    formulaTokens: source?.formulaTokens,
  };
}

function adaptMerges(
  sheet: Biff8Worksheet,
  cells: Map<string, SpreadsheetCell>,
) {
  const merges: SpreadsheetMerge[] = [];
  for (const source of sheet.merges) {
    const anchorRef = cellRef(source.startRow, source.startColumn);
    const anchor = cells.get(anchorRef)!;
    anchor.rowSpan = source.endRow - source.startRow + 1;
    anchor.colSpan = source.endColumn - source.startColumn + 1;
    for (let row = source.startRow; row <= source.endRow; row += 1) {
      for (
        let column = source.startColumn;
        column <= source.endColumn;
        column += 1
      ) {
        if (row !== source.startRow || column !== source.startColumn) {
          cells.get(cellRef(row, column))!.hiddenByMerge = true;
        }
      }
    }
    merges.push({
      ref: `${anchorRef}:${cellRef(source.endRow, source.endColumn)}`,
      startRow: source.startRow + 1,
      startColumn: source.startColumn + 1,
      endRow: source.endRow + 1,
      endColumn: source.endColumn + 1,
    });
  }
  return merges;
}

function adaptWorksheet(
  sheet: Biff8Worksheet,
  workbook: Biff8Workbook,
): SpreadsheetSheet {
  const { maxRow, maxColumn } = computeUsedRange(sheet);
  const sourceCells = new Map(
    sheet.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]),
  );
  const cells = new Map<string, SpreadsheetCell>();
  const columns: SpreadsheetColumn[] = Array.from(
    { length: maxColumn + 1 },
    (_, column) => {
      const info = findColumnInfo(sheet, column);
      return {
        index: column + 1,
        label: columnLabel(column + 1),
        width: columnWidthToPixels(
          info?.widthCharacters ?? sheet.defaultColumnWidth,
        ),
        hidden: info?.hidden,
      };
    },
  );

  const rowInfoByIndex = new Map(
    sheet.rows.map((rowInfo) => [rowInfo.index, rowInfo]),
  );
  const rows = Array.from({ length: maxRow + 1 }, (_, row) => {
    const rowInfo = rowInfoByIndex.get(row);
    const rowCells = Array.from({ length: maxColumn + 1 }, (_, column) => {
      const columnInfo = findColumnInfo(sheet, column);
      const cell = adaptCell(
        sourceCells.get(`${row}:${column}`),
        row,
        column,
        columnInfo?.xfIndex,
        workbook,
      );
      cells.set(cell.ref, cell);
      return cell;
    });
    return {
      index: row + 1,
      height: rowInfo?.heightTwips
        ? twipsToPixels(rowInfo.heightTwips)
        : sheet.defaultRowHeightTwips
        ? twipsToPixels(sheet.defaultRowHeightTwips)
        : DEFAULT_ROW_PIXELS,
      hidden: rowInfo?.hidden,
      cells: rowCells,
    };
  });
  const merges = adaptMerges(sheet, cells);
  const endRef = cellRef(maxRow, maxColumn);
  return {
    id: sheet.descriptor.id,
    name: sheet.descriptor.name,
    path: `/Workbook/${sheet.descriptor.name}`,
    kind: 'worksheet',
    range: endRef === 'A1' ? 'A1' : `A1:${endRef}`,
    rowCount: maxRow + 1,
    columnCount: maxColumn + 1,
    columns,
    rows,
    merges,
    images: [],
    charts: [],
  };
}

function createChartSheetPlaceholder(
  descriptor: Biff8SheetDescriptor,
): SpreadsheetSheet {
  return {
    id: descriptor.id,
    name: descriptor.name,
    path: `/Workbook/${descriptor.name}`,
    kind: 'chart',
    range: 'A1',
    rowCount: 1,
    columnCount: 1,
    columns: [{ index: 1, label: 'A', width: DEFAULT_COLUMN_PIXELS }],
    rows: [
      {
        index: 1,
        height: DEFAULT_ROW_PIXELS,
        cells: [{ ref: 'A1', rowIndex: 1, columnIndex: 1, value: '' }],
      },
    ],
    merges: [],
    images: [],
    charts: [],
  };
}

/** 将一个 BIFF8 工作表描述符适配为可独立传输的预览模型。 */
export function adaptBiff8Sheet(
  source: Biff8Workbook,
  descriptor: Biff8SheetDescriptor,
): SpreadsheetSheet | undefined {
  const worksheet = source.worksheets.find(
    (sheet) => sheet.descriptor.id === descriptor.id,
  );
  if (worksheet) return adaptWorksheet(worksheet, source);
  if (descriptor.type !== 'chart') return undefined;
  return createChartSheetPlaceholder(descriptor);
}

/** 将 BIFF8 中间模型适配为 XLSX 预览器复用的通用工作簿。 */
export function adaptBiff8Workbook(source: Biff8Workbook): SpreadsheetWorkbook {
  return {
    sheets: source.globals.sheets.flatMap((descriptor) => {
      const sheet = adaptBiff8Sheet(source, descriptor);
      return sheet ? [sheet] : [];
    }),
    warnings: source.warnings.length ? source.warnings : undefined,
  };
}

function concatenateChunks(chunks: Uint8Array[]) {
  const length = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  chunks.forEach((chunk) => {
    result.set(chunk, offset);
    offset += chunk.length;
  });
  return result;
}

function pointGeometry(
  sheet: SpreadsheetSheet,
  point: {
    row: number;
    column: number;
    rowFraction: number;
    columnFraction: number;
  },
) {
  const x = sheet.columns
    .slice(0, point.column)
    .reduce((sum, column) => sum + (column.hidden ? 0 : column.width), 0);
  const y = sheet.rows
    .slice(0, point.row)
    .reduce((sum, row) => sum + (row.hidden ? 0 : row.height), 0);
  const column = sheet.columns[point.column];
  const row = sheet.rows[point.row];
  const columnWidth = column?.hidden
    ? 0
    : column?.width ?? DEFAULT_COLUMN_PIXELS;
  const rowHeight = row?.hidden ? 0 : row?.height ?? DEFAULT_ROW_PIXELS;
  return {
    x: x + columnWidth * point.columnFraction,
    y: y + rowHeight * point.rowFraction,
    columnWidth,
    rowHeight,
  };
}

function ensureSheetBounds(
  sheet: SpreadsheetSheet,
  requiredRows: number,
  requiredColumns: number,
) {
  while (sheet.columns.length < requiredColumns) {
    const index = sheet.columns.length + 1;
    sheet.columns.push({
      index,
      label: columnLabel(index),
      width: DEFAULT_COLUMN_PIXELS,
    });
    sheet.rows.forEach((row) => {
      row.cells.push({
        ref: `${columnLabel(index)}${row.index}`,
        rowIndex: row.index,
        columnIndex: index,
        value: '',
      });
    });
  }
  while (sheet.rows.length < requiredRows) {
    const index = sheet.rows.length + 1;
    sheet.rows.push({
      index,
      height: DEFAULT_ROW_PIXELS,
      cells: sheet.columns.map((column) => ({
        ref: `${column.label}${index}`,
        rowIndex: index,
        columnIndex: column.index,
        value: '',
      })),
    });
  }
  sheet.rowCount = sheet.rows.length;
  sheet.columnCount = sheet.columns.length;
  const endRef = `${columnLabel(sheet.columnCount)}${sheet.rowCount}`;
  sheet.range = endRef === 'A1' ? 'A1' : `A1:${endRef}`;
}

export type XlsResourceCollector = {
  add(resource: PortableResource): Promise<string>;
};

/** 解析并附加 XLS 绘图图片，资源的实体化方式由运行环境注入。 */
export async function attachBiff8DrawingImages(
  target: SpreadsheetWorkbook,
  source: Biff8Workbook,
  resources: XlsResourceCollector,
) {
  const groupBytes = concatenateChunks(
    source.globals.drawingGroupRecords.flatMap((record) => record.chunks),
  );
  if (!groupBytes.length) return;
  const warnings = target.warnings ?? [];
  target.warnings = warnings;

  for (const sourceSheet of source.worksheets) {
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === sourceSheet.descriptor.id,
    );
    if (!targetSheet) continue;
    const sheetBytes = concatenateChunks(
      sourceSheet.drawingRecords
        .filter((record) => record.recordId === BIFF8_RECORD.MSODRAWING)
        .flatMap((record) => record.chunks),
    );
    if (!sheetBytes.length) continue;
    let images: ReturnType<typeof parseBiff8Drawings>;
    try {
      images = parseBiff8Drawings(groupBytes, sheetBytes, warnings);
    } catch (error) {
      warnings.push({
        code: 'INVALID_SHEET_DRAWING',
        message: `工作表“${sourceSheet.descriptor.name}”的绘图结构无效：${
          error instanceof Error ? error.message : '未知错误'
        }`,
        sheetName: sourceSheet.descriptor.name,
      });
      continue;
    }
    for (let imageIndex = 0; imageIndex < images.length; imageIndex += 1) {
      const image = images[imageIndex];
      ensureSheetBounds(
        targetSheet,
        image.anchor.to.row + 1,
        image.anchor.to.column + 1,
      );
      try {
        const resource = await createPortableImageResource(
          image,
          `xls:${sourceSheet.descriptor.id}:${image.id}:${imageIndex}`,
        );
        const src = await resources.add(resource.resource);
        warnings.push(
          ...resource.warnings.map((warning) => ({
            ...warning,
            sheetName: warning.sheetName ?? sourceSheet.descriptor.name,
          })),
        );
        const from = pointGeometry(targetSheet, image.anchor.from);
        const to = pointGeometry(targetSheet, image.anchor.to);
        targetSheet.images.push({
          id: image.id,
          name: image.name,
          alt: image.alt,
          src,
          from: {
            row: image.anchor.from.row + 1,
            column: image.anchor.from.column + 1,
            rowOffset: from.rowHeight * image.anchor.from.rowFraction,
            columnOffset: from.columnWidth * image.anchor.from.columnFraction,
          },
          to: {
            row: image.anchor.to.row + 1,
            column: image.anchor.to.column + 1,
            rowOffset: to.rowHeight * image.anchor.to.rowFraction,
            columnOffset: to.columnWidth * image.anchor.to.columnFraction,
          },
          x: from.x,
          y: from.y,
          width: Math.max(1, to.x - from.x),
          height: Math.max(1, to.y - from.y),
        });
      } catch (error) {
        warnings.push({
          code: 'IMAGE_RENDER_FAILED',
          message: `图片“${image.name ?? image.id}”转换失败：${
            error instanceof Error ? error.message : '未知错误'
          }`,
          sheetName: sourceSheet.descriptor.name,
        });
      }
    }
  }
}

/** 解析并附加内嵌图表与独立图表工作表。 */
export function attachBiff8Charts(
  target: SpreadsheetWorkbook,
  source: Biff8Workbook,
) {
  const warnings = target.warnings ?? [];
  target.warnings = warnings;
  for (const sourceSheet of source.worksheets) {
    if (!sourceSheet.chartSubstreams.length) continue;
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === sourceSheet.descriptor.id,
    );
    if (!targetSheet) continue;
    const sheetBytes = concatenateChunks(
      sourceSheet.drawingRecords
        .filter((record) => record.recordId === BIFF8_RECORD.MSODRAWING)
        .flatMap((record) => record.chunks),
    );
    const shapes = parseBiff8DrawingShapes(sheetBytes, warnings);
    const charts = parseBiff8Charts(
      source,
      sourceSheet.descriptor,
      sourceSheet.chartSubstreams,
      shapes,
      targetSheet.images,
      sourceSheet,
    );
    charts.forEach((item) => {
      ensureSheetBounds(
        targetSheet,
        item.anchor.to.row + 1,
        item.anchor.to.column + 1,
      );
      const from = pointGeometry(targetSheet, item.anchor.from);
      const to = pointGeometry(targetSheet, item.anchor.to);
      targetSheet.charts.push({
        id: item.id,
        title: item.title,
        chart: item.chart,
        from: {
          row: item.anchor.from.row + 1,
          column: item.anchor.from.column + 1,
          rowOffset: from.rowHeight * item.anchor.from.rowFraction,
          columnOffset: from.columnWidth * item.anchor.from.columnFraction,
        },
        to: {
          row: item.anchor.to.row + 1,
          column: item.anchor.to.column + 1,
          rowOffset: to.rowHeight * item.anchor.to.rowFraction,
          columnOffset: to.columnWidth * item.anchor.to.columnFraction,
        },
        x: from.x,
        y: from.y,
        width: Math.max(1, to.x - from.x),
        height: Math.max(1, to.y - from.y),
      });
      if (item.previewImageId) {
        targetSheet.images = targetSheet.images.filter(
          (image) => image.id !== item.previewImageId,
        );
      }
      warnings.push(...item.warnings);
    });
  }

  source.chartSheets.forEach((chartSheet) => {
    const targetSheet = target.sheets.find(
      (sheet) => sheet.id === chartSheet.descriptor.id,
    );
    if (!targetSheet) return;
    const chart = parseBiff8Charts(
      source,
      chartSheet.descriptor,
      [chartSheet.substream],
      [],
      [],
    )[0];
    if (!chart) return;
    targetSheet.charts.push({
      id: chart.id,
      title: chart.title,
      chart: chart.chart,
      from: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      to: { row: 1, column: 1, rowOffset: 0, columnOffset: 0 },
      x: 0,
      y: 0,
      width: 960,
      height: 600,
    });
    warnings.push(...chart.warnings);
  });
}
