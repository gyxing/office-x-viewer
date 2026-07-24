import { Biff8Reader } from '../biff8/Biff8Reader';
import { BIFF8_RECORD } from '../biff8/constants';
import { collectChartNodes } from './ChartRecordTree';
import type { Biff8ChartRecordNode, Biff8ChartSeries } from './types';

export type ParsedChartFormatting = {
  title?: string;
  showLegend: boolean;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
};

function colorRef(bytes: Uint8Array) {
  if (bytes.length < 4) return undefined;
  return `#${bytes[0].toString(16).padStart(2, '0')}${bytes[1]
    .toString(16)
    .padStart(2, '0')}${bytes[2].toString(16).padStart(2, '0')}`;
}

function readSerTxt(node: Biff8ChartRecordNode) {
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

function chartTitle(nodes: Biff8ChartRecordNode[]) {
  for (const text of collectChartNodes(nodes, BIFF8_RECORD.TEXT)) {
    const link = collectChartNodes(text.children, BIFF8_RECORD.OBJECTLINK)[0];
    if (
      link?.data.length >= 2 &&
      new Biff8Reader(link.data).readUint16() === 1
    ) {
      return readSerTxt(text);
    }
  }
  return undefined;
}

function applySeriesColors(
  nodes: Biff8ChartRecordNode[],
  series: Biff8ChartSeries[],
) {
  for (const format of collectChartNodes(nodes, BIFF8_RECORD.DATAFORMAT)) {
    if (format.data.length < 4) continue;
    const reader = new Biff8Reader(format.data);
    reader.readUint16();
    const seriesIndex = reader.readUint16();
    const target = series[seriesIndex];
    if (!target) continue;
    const area = collectChartNodes(format.children, BIFF8_RECORD.AREAFORMAT)[0];
    const line = collectChartNodes(format.children, BIFF8_RECORD.LINEFORMAT)[0];
    target.color = colorRef(area?.data ?? line?.data ?? new Uint8Array());
    if (line?.data.length && line.data.length >= 10) {
      target.lineWidth = line.data[8] === 0 ? 2 : 1;
    }
    const marker = collectChartNodes(
      format.children,
      BIFF8_RECORD.MARKERFORMAT,
    )[0];
    if (marker?.data.length && marker.data.length >= 14) {
      const markerType = new DataView(
        marker.data.buffer,
        marker.data.byteOffset,
        marker.data.byteLength,
      ).getUint16(12, true);
      target.marker = {
        symbol:
          markerType === 0 ? 'none' : markerType === 1 ? 'square' : 'circle',
        size: 7,
      };
    }
  }
}

/** 提取标题、图例和系列的基础颜色/线型。 */
export function parseChartFormatting(
  nodes: Biff8ChartRecordNode[],
  series: Biff8ChartSeries[],
): ParsedChartFormatting {
  applySeriesColors(nodes, series);
  const legend = collectChartNodes(nodes, BIFF8_RECORD.LEGEND)[0];
  let legendPosition: ParsedChartFormatting['legendPosition'];
  if (legend?.data.length && legend.data.length >= 17) {
    legendPosition =
      legend.data[16] === 0
        ? 'bottom'
        : legend.data[16] === 2
        ? 'top'
        : legend.data[16] === 3
        ? 'right'
        : legend.data[16] === 4
        ? 'left'
        : undefined;
  }
  return {
    title: chartTitle(nodes),
    showLegend: Boolean(legend),
    legendPosition,
  };
}
