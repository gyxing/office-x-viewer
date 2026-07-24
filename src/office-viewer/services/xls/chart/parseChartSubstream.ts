import { Biff8Reader } from '../biff8/Biff8Reader';
import { BIFF8_RECORD } from '../biff8/constants';
import { buildChartRecordTree, collectChartNodes } from './ChartRecordTree';
import { CHART_TYPE_RECORDS, FALLBACK_CHART_ANCHOR } from './constants';
import { parseChartFormatting } from './parseChartFormatting';
import { parseBiff8ChartSeries } from './parseChartSeries';
import type {
  Biff8ChartContext,
  Biff8ChartModel,
  Biff8ChartRecordNode,
} from './types';

function chartType(nodes: Biff8ChartRecordNode[]) {
  for (const [recordId, type] of CHART_TYPE_RECORDS) {
    const node = collectChartNodes(nodes, recordId)[0];
    if (!node) continue;
    if (recordId === BIFF8_RECORD.BAR && node.data.length >= 6) {
      const reader = new Biff8Reader(node.data);
      const overlap = reader.readInt16();
      const gapWidth = reader.readUint16();
      const flags = reader.readUint16();
      return {
        type: flags & 0x0001 ? 'bar' : 'column',
        overlap,
        gapWidth,
        stacking:
          flags & 0x0004
            ? ('percentStacked' as const)
            : flags & 0x0002
            ? ('stacked' as const)
            : undefined,
      };
    }
    if (recordId === BIFF8_RECORD.PIE && node.data.length >= 4) {
      const reader = new Biff8Reader(node.data);
      reader.readUint16();
      const holeSize = reader.readUint16();
      return { type: holeSize ? 'doughnut' : 'pie', holeSize };
    }
    if (recordId === BIFF8_RECORD.SCATTER && node.data.length >= 6) {
      return {
        type: new Biff8Reader(node.data).readBytes(4)[0]
          ? 'scatter'
          : 'scatter',
        bubble: Boolean(
          new DataView(
            node.data.buffer,
            node.data.byteOffset,
            node.data.byteLength,
          ).getUint16(4, true) & 0x0001,
        ),
      };
    }
    if (
      (recordId === BIFF8_RECORD.LINE || recordId === BIFF8_RECORD.AREA) &&
      node.data.length >= 2
    ) {
      const flags = new Biff8Reader(node.data).readUint16();
      return {
        type,
        stacking:
          flags & 0x0002
            ? ('percentStacked' as const)
            : flags & 0x0001
            ? ('stacked' as const)
            : undefined,
      };
    }
    return { type };
  }
  return { type: 'unknown' };
}

function chartGroupTypes(nodes: Biff8ChartRecordNode[]) {
  const groups = collectChartNodes(nodes, BIFF8_RECORD.CHARTFORMAT);
  if (!groups.length) return [chartType(nodes)];
  return groups.map((group) => chartType(group.children));
}

function warnUnknownRecords(
  nodes: Biff8ChartRecordNode[],
  context: Biff8ChartContext,
) {
  const warnings: Biff8ChartModel['warnings'] = [];
  const known = new Set<number>(Object.values(BIFF8_RECORD));
  const unknown = new Set<number>();
  const visit = (items: Biff8ChartRecordNode[]) => {
    items.forEach((node) => {
      if (node.id >= 0x1000 && node.id <= 0x10ff && !known.has(node.id)) {
        unknown.add(node.id);
      }
      visit(node.children);
    });
  };
  visit(nodes);
  unknown.forEach((recordId) =>
    warnings.push({
      code: 'UNKNOWN_CHART_RECORD',
      message: `已跳过未知图表记录 0x${recordId.toString(16).toUpperCase()}`,
      sheetName: context.descriptor.name,
    }),
  );
  return warnings;
}

/** 将一个完整 Chart 子流解析为与渲染层无关的模型。 */
export function parseChartSubstream(
  context: Biff8ChartContext,
): Biff8ChartModel {
  const nodes = buildChartRecordTree(context.substream.records);
  const groupTypes = chartGroupTypes(nodes);
  const type = groupTypes[0] ?? chartType(nodes);
  const series = parseBiff8ChartSeries(nodes, context);
  const formatting = parseChartFormatting(nodes, series);
  series.forEach((item) => {
    item.stacking = type.stacking;
  });
  const warnings = warnUnknownRecords(nodes, context);
  if (!series.some((item) => item.values.some((value) => value !== null))) {
    warnings.push({
      code: 'EMPTY_CHART_DATA',
      message: '图表没有可用的缓存或工作表引用数据',
      sheetName: context.descriptor.name,
    });
  }
  const categories = series.find((item) => item.categories.length)?.categories;
  const hasSecondaryAxis = collectChartNodes(
    nodes,
    BIFF8_RECORD.AXISPARENT,
  ).some(
    (node) =>
      node.data.length >= 2 && new Biff8Reader(node.data).readUint16() === 1,
  );
  const isStock =
    type.type === 'line' &&
    collectChartNodes(nodes, BIFF8_RECORD.DROPBAR).length > 0;
  return {
    id: `xls-chart-${context.descriptor.id}-${context.chartIndex + 1}`,
    sourceType: isStock ? 'stock' : type.bubble ? 'bubble' : type.type,
    groupTypes: groupTypes.map((group) =>
      group.bubble ? 'bubble' : group.type,
    ),
    is3d: collectChartNodes(nodes, BIFF8_RECORD.CHART3D).length > 0,
    hasSecondaryAxis,
    title: formatting.title,
    categories: (categories ?? []).map((value, index) =>
      value == null ? String(index + 1) : String(value),
    ),
    series,
    showLegend: formatting.showLegend,
    legendPosition: formatting.legendPosition,
    gapWidth: type.gapWidth,
    overlap: type.overlap,
    holeSize: type.holeSize,
    anchor: context.anchor ?? FALLBACK_CHART_ANCHOR,
    previewImageSrc: context.previewImageSrc,
    warnings,
  };
}
