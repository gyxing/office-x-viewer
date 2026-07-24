import { Biff8Reader, type Biff8Record } from '../biff8/Biff8Reader';
import { BIFF8_RECORD } from '../biff8/constants';
import { decodeBiff8Formula } from '../biff8/formulas';
import { readBiff8UnicodeString } from '../biff8/strings';
import type { Biff8Cell } from '../types';
import { collectChartNodes } from './ChartRecordTree';
import { CHART_CACHE_KIND } from './constants';
import type {
  Biff8ChartCache,
  Biff8ChartContext,
  Biff8ChartRecordNode,
  Biff8ChartSeries,
} from './types';

function setCacheValue(
  cache: Biff8ChartCache,
  kind: number,
  series: number,
  point: number,
  value: string | number | boolean | null,
) {
  const seriesMap = cache.get(kind) ?? new Map();
  const pointMap = seriesMap.get(series) ?? new Map();
  pointMap.set(point, value);
  seriesMap.set(series, pointMap);
  cache.set(kind, seriesMap);
}

/** 读取 SERIESDATA 区域，按数据种类、系列列和点行保留空洞。 */
export function parseChartCache(records: Biff8Record[]) {
  const cache: Biff8ChartCache = new Map();
  let kind = 0;
  for (const record of records) {
    if (record.id === BIFF8_RECORD.SIINDEX && record.data.length >= 2) {
      kind = new Biff8Reader(record.data).readUint16();
      continue;
    }
    if (!kind) continue;
    if (
      ![
        BIFF8_RECORD.NUMBER,
        BIFF8_RECORD.LABEL,
        BIFF8_RECORD.BOOLERR,
        BIFF8_RECORD.BLANK,
      ].includes(record.id as never) ||
      record.data.length < 6
    ) {
      continue;
    }
    const reader = new Biff8Reader(record.data);
    const point = reader.readUint16();
    const series = reader.readUint16();
    reader.readUint16();
    let value: string | number | boolean | null = null;
    if (record.id === BIFF8_RECORD.NUMBER) value = reader.readFloat64();
    if (record.id === BIFF8_RECORD.LABEL) {
      value = readBiff8UnicodeString(reader).value;
    }
    if (record.id === BIFF8_RECORD.BOOLERR) {
      const raw = reader.readUint8();
      value = reader.readUint8() ? null : Boolean(raw);
    }
    setCacheValue(cache, kind, series, point, value);
  }
  return cache;
}

function cacheColumn(cache: Biff8ChartCache, kind: number, series: number) {
  const points = cache.get(kind)?.get(series);
  if (!points?.size) return [];
  const length = Math.max(...points.keys()) + 1;
  return Array.from({ length }, (_, index) => points.get(index) ?? null);
}

function columnIndex(label: string) {
  return Array.from(label.toUpperCase()).reduce(
    (value, character) => value * 26 + character.charCodeAt(0) - 64,
    0,
  );
}

function resolveCells(
  formula: string | undefined,
  context: Biff8ChartContext,
): Biff8Cell[] {
  if (!formula) return [];
  const match = formula.match(
    /^=(?:'((?:''|[^'])+)'|([^!]+))?!?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/,
  );
  if (!match) return [];
  const sheetName = (match[1]?.replace(/''/g, "'") ?? match[2])?.trim();
  const sheet =
    (sheetName
      ? context.workbook.worksheets.find(
          (item) => item.descriptor.name === sheetName,
        )
      : context.sourceSheet) ?? context.sourceSheet;
  if (!sheet) return [];
  const firstColumn = columnIndex(match[3]) - 1;
  const firstRow = Number(match[4]) - 1;
  const lastColumn = match[5] ? columnIndex(match[5]) - 1 : firstColumn;
  const lastRow = match[6] ? Number(match[6]) - 1 : firstRow;
  const cells = new Map(
    sheet.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]),
  );
  const result: Biff8Cell[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) {
    for (let column = firstColumn; column <= lastColumn; column += 1) {
      result.push(
        cells.get(`${row}:${column}`) ?? {
          row,
          column,
          xfIndex: 0,
          value: null,
          cachedType: 'blank',
        },
      );
    }
  }
  return result;
}

function parseAiFormulas(
  node: Biff8ChartRecordNode,
  context: Biff8ChartContext,
) {
  const formulas = new Map<number, string | undefined>();
  for (const ai of collectChartNodes(node.children, BIFF8_RECORD.AI)) {
    if (ai.data.length < 8) continue;
    const reader = new Biff8Reader(ai.data);
    const linkId = reader.readUint8();
    reader.readBytes(5);
    const tokenLength = reader.readUint16();
    if (tokenLength > reader.remaining) continue;
    formulas.set(
      linkId,
      decodeBiff8Formula(reader.readBytes(tokenLength), {
        row: 0,
        column: 0,
        definedNames: context.workbook.globals.definedNames,
        sheets: context.workbook.globals.sheets,
      }).formula,
    );
  }
  return formulas;
}

function parseSeriesName(node: Biff8ChartRecordNode) {
  const text = collectChartNodes(node.children, BIFF8_RECORD.SERTXT)[0];
  if (!text || text.data.length < 4) return undefined;
  const reader = new Biff8Reader(text.data);
  reader.readUint16();
  const length = reader.readUint8();
  const unicode = Boolean(reader.readUint8() & 0x01);
  const bytes = reader.readBytes(length * (unicode ? 2 : 1));
  if (!unicode) {
    return Array.from(bytes, (value) => String.fromCharCode(value)).join('');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return Array.from({ length }, (_, index) =>
    String.fromCharCode(view.getUint16(index * 2, true)),
  ).join('');
}

/** 解析 Series、AI 引用和三类图表缓存。 */
export function parseBiff8ChartSeries(
  nodes: Biff8ChartRecordNode[],
  context: Biff8ChartContext,
): Biff8ChartSeries[] {
  const cache = parseChartCache(context.substream.records);
  return collectChartNodes(nodes, BIFF8_RECORD.SERIES).map((node, index) => {
    const formulas = parseAiFormulas(node, context);
    const groupNode = collectChartNodes(
      node.children,
      BIFF8_RECORD.SERTOCRT,
    )[0];
    const groupIndex =
      groupNode?.data.length >= 2
        ? new Biff8Reader(groupNode.data).readUint16()
        : 0;
    const cacheCategories = cacheColumn(
      cache,
      CHART_CACHE_KIND.CATEGORIES,
      index,
    );
    const cacheValues = cacheColumn(cache, CHART_CACHE_KIND.VALUES, index);
    const cacheBubbles = cacheColumn(cache, CHART_CACHE_KIND.BUBBLES, index);
    const referenceName = resolveCells(formulas.get(0), context)[0]?.value;
    return {
      name:
        parseSeriesName(node) ??
        (referenceName == null ? `系列 ${index + 1}` : String(referenceName)),
      groupIndex,
      categories: (cacheCategories.length
        ? cacheCategories
        : resolveCells(formulas.get(2), context).map((cell) => cell.value)
      ).map((value) =>
        typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value,
      ),
      values: (cacheValues.length
        ? cacheValues
        : resolveCells(formulas.get(1), context).map((cell) => cell.value)
      ).map((value) => (typeof value === 'number' ? value : null)),
      bubbleSizes: (cacheBubbles.length
        ? cacheBubbles
        : resolveCells(formulas.get(3), context).map((cell) => cell.value)
      ).map((value) => (typeof value === 'number' ? value : null)),
    };
  });
}
