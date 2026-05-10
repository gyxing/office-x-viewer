import { loadXlsxEntries } from './archive';
import type { OfficeEntryMap } from '../office/archive';
import { readXml } from '../office/archive';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../office/xml';
import { collectMedia, resolvePackageMediaRef, type OfficeRelationship } from '../office/media';
import { readRelationships } from '../office/relationships';
import { emuToPx } from '../office/units';
import { parseOfficeChartXml } from '../office/charts';
import { readOfficeTheme, type OfficeTheme } from '../office/theme';
import type {
  XlsxCell,
  XlsxCellStyle,
  XlsxColumn,
  XlsxChart,
  XlsxImage,
  XlsxMerge,
  XlsxRow,
  XlsxSheet,
  XlsxWorkbook,
} from './types';

type ParsedStyle = {
  fontId?: number;
  fillId?: number;
  borderId?: number;
  alignment?: XlsxCellStyle;
};

type StyleBook = {
  fonts: XlsxCellStyle[];
  fills: Array<Pick<XlsxCellStyle, 'backgroundColor'>>;
  borders: Array<Pick<XlsxCellStyle, 'border'>>;
  styles: ParsedStyle[];
};

type CellAddress = {
  row: number;
  column: number;
};

type XlsxPackageState = {
  entries: OfficeEntryMap;
  relationships: Record<string, Record<string, OfficeRelationship>>;
  mediaByPath: Record<string, string>;
  mediaByName: Record<string, string>;
  theme: OfficeTheme;
};

const DEFAULT_COLUMN_WIDTH = 88;
const DEFAULT_ROW_HEIGHT = 28;
const MAX_RENDERED_EMPTY_ROWS = 200;
const MAX_RENDERED_EMPTY_COLUMNS = 80;

function buildPackageState(entries: OfficeEntryMap): XlsxPackageState {
  const relationships: XlsxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = collectMedia(entries, 'xl/media/');

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
    theme: readOfficeTheme(readXml(entries, 'xl/theme/theme1.xml')),
  };
}

function decodeMojibake(value: string) {
  if (!/[ÃÂäåæçèé]|�|锟|鎬|宸|濮|韬|骞|涓|鍚|煡/.test(value)) {
    return value;
  }

  try {
    const bytes = new Uint8Array(Array.from(value, (char) => char.charCodeAt(0) & 0xff));
    const decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
    return decoded.includes('\uFFFD') ? value : decoded;
  } catch {
    return value;
  }
}

function readPlainText(node: Element | null | undefined) {
  if (!node) return '';
  return decodeMojibake(descendantsByLocalName(node, 't').map((item) => textContent(item)).join(''));
}

function readSharedStrings(xml: string) {
  if (!xml) return [];
  const doc = parseXml(xml);
  return childrenByLocalName(doc.documentElement, 'si').map(readPlainText);
}

function columnLabelToIndex(label: string) {
  return label.split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
}

function columnIndexToLabel(index: number) {
  let value = index;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function parseCellRef(ref: string): CellAddress {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { row: 1, column: 1 };
  return {
    row: Number(match[2]),
    column: columnLabelToIndex(match[1].toUpperCase()),
  };
}

function parseRange(range?: string) {
  if (!range) return undefined;
  const [start, end = start] = range.replace(/\$/g, '').split(':');
  const startCell = parseCellRef(start);
  const endCell = parseCellRef(end);
  return {
    startRow: startCell.row,
    startColumn: startCell.column,
    endRow: endCell.row,
    endColumn: endCell.column,
  };
}

function parseColor(node: Element | null | undefined) {
  const rgb = attr(node, 'rgb');
  if (!rgb) return undefined;
  const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
  return `#${hex}`;
}

function parseStyles(xml: string): StyleBook {
  if (!xml) return { fonts: [], fills: [], borders: [], styles: [] };
  const doc = parseXml(xml);
  const styleSheet = doc.documentElement;

  const fontsNode = childByLocalName(styleSheet, 'fonts');
  const fonts = childrenByLocalName(fontsNode, 'font').map((fontNode): XlsxCellStyle => ({
    bold: Boolean(childByLocalName(fontNode, 'b')),
    italic: Boolean(childByLocalName(fontNode, 'i')),
    underline: Boolean(childByLocalName(fontNode, 'u')),
    color: parseColor(childByLocalName(fontNode, 'color')),
  }));

  const fillsNode = childByLocalName(styleSheet, 'fills');
  const fills = childrenByLocalName(fillsNode, 'fill').map((fillNode) => {
    const pattern = childByLocalName(fillNode, 'patternFill');
    return {
      backgroundColor: parseColor(childByLocalName(pattern, 'fgColor')),
    };
  });

  const bordersNode = childByLocalName(styleSheet, 'borders');
  const borders = childrenByLocalName(bordersNode, 'border').map((borderNode) => ({
    border: Array.from(borderNode.children).some((side) => Boolean(attr(side, 'style'))),
  }));

  const cellXfs = childByLocalName(styleSheet, 'cellXfs');
  const styles = childrenByLocalName(cellXfs, 'xf').map((xfNode): ParsedStyle => {
    const alignment = childByLocalName(xfNode, 'alignment');
    const vertical = attr(alignment, 'vertical');
    const horizontal = attr(alignment, 'horizontal');
    return {
      fontId: Number(attr(xfNode, 'fontId') ?? 0),
      fillId: Number(attr(xfNode, 'fillId') ?? 0),
      borderId: Number(attr(xfNode, 'borderId') ?? 0),
      alignment: {
        horizontalAlign:
          horizontal === 'center' || horizontal === 'right' || horizontal === 'justify'
            ? horizontal
            : horizontal === 'left'
              ? 'left'
              : undefined,
        verticalAlign:
          vertical === 'top'
            ? 'top'
            : vertical === 'bottom'
              ? 'bottom'
              : vertical === 'center'
                ? 'middle'
                : undefined,
        wrapText: attr(alignment, 'wrapText') === '1',
      },
    };
  });

  return { fonts, fills, borders, styles };
}

function resolveStyle(styleId: number | undefined, styleBook: StyleBook): XlsxCellStyle | undefined {
  if (styleId === undefined) return undefined;
  const style = styleBook.styles[styleId];
  if (!style) return undefined;

  const resolved: XlsxCellStyle = {
    ...styleBook.fonts[style.fontId ?? 0],
    ...styleBook.fills[style.fillId ?? 0],
    ...styleBook.borders[style.borderId ?? 0],
    ...style.alignment,
  };

  return Object.fromEntries(
    Object.entries(resolved).filter(([, value]) => value !== undefined && value !== false),
  ) as XlsxCellStyle;
}

function excelWidthToPx(width?: number) {
  if (!width || !Number.isFinite(width)) return DEFAULT_COLUMN_WIDTH;
  return Math.max(40, Math.round(width * 7 + 5));
}

function pointToPx(point?: number) {
  if (!point || !Number.isFinite(point)) return DEFAULT_ROW_HEIGHT;
  return Math.max(22, Math.round(point * (96 / 72)));
}

function getColumnWidth(columns: XlsxColumn[], columnIndex: number) {
  return columns[columnIndex - 1]?.width ?? DEFAULT_COLUMN_WIDTH;
}

function getRowHeight(rowHeights: Map<number, number>, rowIndex: number) {
  return rowHeights.get(rowIndex) ?? DEFAULT_ROW_HEIGHT;
}

function anchorPosition(
  anchor: { row: number; column: number; rowOffset: number; columnOffset: number },
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
) {
  let x = 0;
  for (let column = 1; column < anchor.column; column += 1) {
    x += getColumnWidth(columns, column);
  }

  let y = 0;
  for (let row = 1; row < anchor.row; row += 1) {
    y += getRowHeight(rowHeights, row);
  }

  return {
    x: x + emuToPx(anchor.columnOffset),
    y: y + emuToPx(anchor.rowOffset),
  };
}

function readColumns(sheetNode: Element, maxColumn: number): XlsxColumn[] {
  const widths = new Map<number, XlsxColumn>();
  descendantsByLocalName(sheetNode, 'col').forEach((node) => {
    const min = Number(attr(node, 'min') ?? 1);
    const max = Math.min(Number(attr(node, 'max') ?? min), Math.max(maxColumn, MAX_RENDERED_EMPTY_COLUMNS));
    for (let index = min; index <= max; index += 1) {
      widths.set(index, {
        index,
        label: columnIndexToLabel(index),
        width: excelWidthToPx(Number(attr(node, 'width'))),
        hidden: attr(node, 'hidden') === '1',
      });
    }
  });

  return Array.from({ length: maxColumn }, (_, itemIndex) => {
    const index = itemIndex + 1;
    return widths.get(index) ?? {
      index,
      label: columnIndexToLabel(index),
      width: DEFAULT_COLUMN_WIDTH,
    };
  });
}

function readAnchorPoint(node: Element | null) {
  return {
    column: Number(textContent(childByLocalName(node, 'col'))) + 1,
    columnOffset: Number(textContent(childByLocalName(node, 'colOff')) || 0),
    row: Number(textContent(childByLocalName(node, 'row'))) + 1,
    rowOffset: Number(textContent(childByLocalName(node, 'rowOff')) || 0),
  };
}

function resolveMediaRef(target: string | undefined, packageState: XlsxPackageState) {
  return resolvePackageMediaRef(target, packageState.mediaByPath, packageState.mediaByName, 'xl');
}

function resolveXmlTarget(target: string | undefined, packageState: XlsxPackageState) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function readSheetCharts(
  sheetNode: Element,
  sheetPath: string,
  packageState: XlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
) {
  const drawing = descendantByLocalName(sheetNode, 'drawing');
  const drawingRelId = attr(drawing, 'r:id') ?? attr(drawing, 'id');
  if (!drawingRelId) return [];

  const sheetRelPath = sheetPath.replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/').concat('.rels');
  const drawingPath = packageState.relationships[sheetRelPath]?.[drawingRelId]?.target;
  if (!drawingPath) return [];

  const drawingXml = readXml(packageState.entries, drawingPath);
  if (!drawingXml) return [];

  const drawingRelPath = drawingPath.replace(/^xl\/drawings\//, 'xl/drawings/_rels/').concat('.rels');
  const drawingRels = packageState.relationships[drawingRelPath] ?? {};
  const drawingDoc = parseXml(drawingXml);

  return childrenByLocalName(drawingDoc.documentElement, 'twoCellAnchor')
    .map((anchorNode, index): XlsxChart | undefined => {
      const graphicFrame = childByLocalName(anchorNode, 'graphicFrame');
      const chartNode = descendantByLocalName(graphicFrame, 'chart');
      const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
      const target = relId ? drawingRels[relId]?.target : undefined;
      const chartPath = resolveXmlTarget(target, packageState);
      const xml = chartPath ? (packageState.entries.get(chartPath) as string | undefined) : undefined;
      if (!xml) return undefined;

      const startPoint = readAnchorPoint(childByLocalName(anchorNode, 'from'));
      const endPoint = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      const start = anchorPosition(startPoint, columns, rowHeights);
      const end = anchorPosition(endPoint, columns, rowHeights);
      const chart = parseOfficeChartXml(xml, packageState.theme);
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawingPath}-chart-${index + 1}`,
        title: name,
        chart,
        from: startPoint,
        to: endPoint,
        x: start.x,
        y: start.y,
        width: Math.max(1, end.x - start.x),
        height: Math.max(1, end.y - start.y),
      };
    })
    .filter(Boolean) as XlsxChart[];
}

function readSheetImages(
  sheetNode: Element,
  sheetPath: string,
  packageState: XlsxPackageState,
  columns: XlsxColumn[],
  rowHeights: Map<number, number>,
) {
  const drawing = descendantByLocalName(sheetNode, 'drawing');
  const drawingRelId = attr(drawing, 'r:id') ?? attr(drawing, 'id');
  if (!drawingRelId) return [];

  const sheetRelPath = sheetPath.replace(/^xl\/worksheets\//, 'xl/worksheets/_rels/').concat('.rels');
  const drawingPath = packageState.relationships[sheetRelPath]?.[drawingRelId]?.target;
  if (!drawingPath) return [];

  const drawingXml = readXml(packageState.entries, drawingPath);
  if (!drawingXml) return [];

  const drawingRelPath = drawingPath.replace(/^xl\/drawings\//, 'xl/drawings/_rels/').concat('.rels');
  const drawingRels = packageState.relationships[drawingRelPath] ?? {};
  const drawingDoc = parseXml(drawingXml);

  return childrenByLocalName(drawingDoc.documentElement, 'twoCellAnchor')
    .map((anchorNode, index): XlsxImage | undefined => {
      const from = readAnchorPoint(childByLocalName(anchorNode, 'from'));
      const to = readAnchorPoint(childByLocalName(anchorNode, 'to'));
      const blip = descendantByLocalName(anchorNode, 'blip');
      const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
      const target = embed ? drawingRels[embed]?.target : undefined;
      const src = resolveMediaRef(target, packageState);
      if (!src) return undefined;

      const start = anchorPosition(from, columns, rowHeights);
      const end = anchorPosition(to, columns, rowHeights);
      const name = attr(descendantByLocalName(anchorNode, 'cNvPr'), 'name');

      return {
        id: `${drawingPath}-${index + 1}`,
        name,
        alt: name,
        src,
        from,
        to,
        x: start.x,
        y: start.y,
        width: Math.max(1, end.x - start.x),
        height: Math.max(1, end.y - start.y),
      };
    })
    .filter(Boolean) as XlsxImage[];
}

function readCellValue(cellNode: Element, sharedStrings: string[]) {
  const type = attr(cellNode, 't');
  const valueNode = childByLocalName(cellNode, 'v');
  const rawValue = textContent(valueNode);

  if (type === 's') {
    return {
      rawValue,
      value: sharedStrings[Number(rawValue)] ?? '',
    };
  }

  if (type === 'inlineStr') {
    return {
      rawValue,
      value: readPlainText(childByLocalName(cellNode, 'is')),
    };
  }

  if (type === 'b') {
    return {
      rawValue,
      value: rawValue === '1' ? 'TRUE' : 'FALSE',
    };
  }

  return {
    rawValue,
    value: rawValue,
  };
}

function readMerges(sheetNode: Element) {
  const mergeCells = descendantByLocalName(sheetNode, 'mergeCells');
  return childrenByLocalName(mergeCells, 'mergeCell')
    .map((node) => {
      const ref = attr(node, 'ref') ?? '';
      const range = parseRange(ref);
      return range ? { ref, ...range } : undefined;
    })
    .filter(Boolean) as XlsxMerge[];
}

function applyMerges(cells: Map<string, XlsxCell>, merges: XlsxMerge[]) {
  merges.forEach((merge) => {
    const startRef = `${columnIndexToLabel(merge.startColumn)}${merge.startRow}`;
    const root = cells.get(startRef);
    if (root) {
      root.colSpan = merge.endColumn - merge.startColumn + 1;
      root.rowSpan = merge.endRow - merge.startRow + 1;
    }

    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
        if (row === merge.startRow && column === merge.startColumn) continue;
        const ref = `${columnIndexToLabel(column)}${row}`;
        const cell = cells.get(ref);
        if (cell) {
          cell.hiddenByMerge = true;
        } else {
          cells.set(ref, {
            ref,
            rowIndex: row,
            columnIndex: column,
            value: '',
            hiddenByMerge: true,
          });
        }
      }
    }
  });
}

function readSheet(
  xml: string,
  sheetInfo: Pick<XlsxSheet, 'id' | 'name' | 'path'>,
  sharedStrings: string[],
  styleBook: StyleBook,
  packageState: XlsxPackageState,
): XlsxSheet {
  const doc = parseXml(xml);
  const sheetNode = doc.documentElement;
  const range = attr(childByLocalName(sheetNode, 'dimension'), 'ref');
  const parsedRange = parseRange(range);
  const cells = new Map<string, XlsxCell>();
  let maxRow = parsedRange?.endRow ?? 0;
  let maxColumn = parsedRange?.endColumn ?? 0;

  descendantsByLocalName(sheetNode, 'c').forEach((cellNode) => {
    const ref = attr(cellNode, 'r') ?? 'A1';
    const address = parseCellRef(ref);
    const styleId = attr(cellNode, 's') ? Number(attr(cellNode, 's')) : undefined;
    const value = readCellValue(cellNode, sharedStrings);
    const cell: XlsxCell = {
      ref,
      rowIndex: address.row,
      columnIndex: address.column,
      value: value.value,
      rawValue: value.rawValue,
      type: attr(cellNode, 't'),
      styleId,
      style: resolveStyle(styleId, styleBook),
    };
    cells.set(ref, cell);
    maxRow = Math.max(maxRow, address.row);
    maxColumn = Math.max(maxColumn, address.column);
  });

  maxRow = Math.min(Math.max(maxRow, 1), MAX_RENDERED_EMPTY_ROWS);
  maxColumn = Math.min(Math.max(maxColumn, 1), MAX_RENDERED_EMPTY_COLUMNS);

  const merges = readMerges(sheetNode);
  applyMerges(cells, merges);

  const rowHeights = new Map<number, number>();
  descendantsByLocalName(sheetNode, 'row').forEach((rowNode) => {
    const rowIndex = Number(attr(rowNode, 'r') ?? 0);
    if (rowIndex) {
      rowHeights.set(rowIndex, pointToPx(Number(attr(rowNode, 'ht'))));
    }
  });
  const columns = readColumns(sheetNode, maxColumn);

  const rows: XlsxRow[] = Array.from({ length: maxRow }, (_, rowOffset) => {
    const rowIndex = rowOffset + 1;
    return {
      index: rowIndex,
      height: rowHeights.get(rowIndex) ?? DEFAULT_ROW_HEIGHT,
      cells: Array.from({ length: maxColumn }, (_, columnOffset) => {
        const columnIndex = columnOffset + 1;
        const ref = `${columnIndexToLabel(columnIndex)}${rowIndex}`;
        return cells.get(ref) ?? {
          ref,
          rowIndex,
          columnIndex,
          value: '',
        };
      }),
    };
  });

  return {
    ...sheetInfo,
    range,
    rowCount: maxRow,
    columnCount: maxColumn,
    columns,
    rows,
    merges,
    images: readSheetImages(sheetNode, sheetInfo.path, packageState, columns, rowHeights),
    charts: readSheetCharts(sheetNode, sheetInfo.path, packageState, columns, rowHeights),
  };
}

export async function parseXlsx(file: File): Promise<XlsxWorkbook> {
  const entries = await loadXlsxEntries(file);
  const packageState = buildPackageState(entries);
  const workbookXml = readXml(entries, 'xl/workbook.xml');
  const workbookRels = packageState.relationships['xl/_rels/workbook.xml.rels'] ?? {};
  const sharedStrings = readSharedStrings(readXml(entries, 'xl/sharedStrings.xml'));
  const styleBook = parseStyles(readXml(entries, 'xl/styles.xml'));
  const workbookDoc = parseXml(workbookXml);
  const sheets = childrenByLocalName(
    childByLocalName(workbookDoc.documentElement, 'sheets'),
    'sheet',
  ).map((sheetNode, index) => {
    const relId = attr(sheetNode, 'r:id') ?? attr(sheetNode, 'id') ?? '';
    const rel = workbookRels[relId];
    const path = rel?.target ?? `xl/worksheets/sheet${index + 1}.xml`;
    return readSheet(
      readXml(entries, path),
      {
        id: attr(sheetNode, 'sheetId') ?? String(index + 1),
        name: decodeMojibake(attr(sheetNode, 'name') ?? `Sheet ${index + 1}`),
        path,
      },
      sharedStrings,
      styleBook,
      packageState,
    );
  });

  return { sheets };
}
