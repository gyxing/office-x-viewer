import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from './xml';
import type { EChartsOption } from 'echarts';
import { DEFAULT_OFFICE_THEME, resolveOfficeThemeColor, type OfficeTheme } from './theme';

export type OfficeChartType =
  | 'line'
  | 'bar'
  | 'column'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'scatter'
  | 'bubble'
  | 'radar'
  | 'map'
  | 'unknown';

export type OfficeChartSeries = {
  name: string;
  values: number[];
  type?: OfficeChartType;
  stacking?: 'stacked' | 'percentStacked';
  stackGroup?: string;
  gapWidth?: number;
  overlap?: number;
  color?: string;
  pointColors?: string[];
  pointLabels?: string[];
  pointStyles?: Array<{
    color?: OfficeChartColor;
    borderColor?: string;
    borderWidth?: number;
  }>;
  dataLabels?: OfficeDataLabels;
  smooth?: boolean;
  lineWidth?: number;
  marker?: {
    symbol?: string;
    size?: number;
  };
};

export type OfficeDataLabels = {
  delete?: boolean;
  position?: string;
  separator?: string;
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLeaderLines?: boolean;
};

export type OfficeChartModel = {
  type: OfficeChartType;
  title?: string;
  categories: string[];
  series: OfficeChartSeries[];
  dataLabels?: OfficeDataLabels;
  showLegend?: boolean;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  legendStyle?: {
    itemWidth?: number;
    itemHeight?: number;
    textStyle?: {
      color?: string;
      fontFamily?: string;
      fontSize?: number;
      fontStyle?: string;
      fontWeight?: string | number;
    };
  };
  showDataLabels?: boolean;
  holeSize?: number;
  startAngle?: number;
  ofPieType?: 'bar' | 'pie';
  ofPieSecondPlotCount?: number;
  secondPieSize?: number;
  gapWidth?: number;
  overlap?: number;
  roseType?: 'radius' | 'area';
  radius?: [string, string];
  radarStyle?: string;
  radarRadius?: string;
  radarStartAngle?: number;
  radarSplitNumber?: number;
  radarIndicators?: Array<{
    name: string;
    max: number;
  }>;
  mapSeriesName?: string;
  mapRegion?: string;
  mapName?: string;
  mapGeoJsonUrl?: string;
  snapshotSrc?: string;
};

const DEFAULT_COLORS = ['#5470c6', '#91cc75', '#fac858', '#ee6666', '#73c0de', '#3ba272', '#fc8452'];
const OFFICE_FONT_FAMILY = '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';
const OFFICE_TEXT_STYLE = {
  color: '#334155',
  fontFamily: OFFICE_FONT_FAMILY,
};

type OfficeChartColorStop = {
  offset: number;
  color: string;
};

type OfficeChartColor =
  | string
  | {
      type: 'linear';
      x: number;
      y: number;
      x2: number;
      y2: number;
      colorStops: OfficeChartColorStop[];
      global?: boolean;
    };

const CHART_NODE_TO_TYPE: Record<string, OfficeChartType> = {
  linechart: 'line',
  barchart: 'column',
  piechart: 'pie',
  doughnutchart: 'doughnut',
  areachart: 'area',
  scatterchart: 'scatter',
  bubblechart: 'bubble',
  radarchart: 'radar',
  ofpiechart: 'pie',
};

export function decodeMojibake(value: string) {
  if (!/[脙脗盲氓忙莽猫茅]|锟|鍥|绯|绫|诲|埆|垪|棰/.test(value)) {
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

function firstText(node: Element | null | undefined) {
  const value = textContent(descendantByLocalName(node, 't')) || textContent(descendantByLocalName(node, 'v'));
  return decodeMojibake(value.trim());
}

function readCacheValues(node: Element | null | undefined, date1904 = false) {
  const strCache = descendantByLocalName(node, 'strCache');
  if (strCache) {
    return descendantsByLocalName(strCache, 'pt')
      .sort((a, b) => Number(attr(a, 'idx') ?? 0) - Number(attr(b, 'idx') ?? 0))
      .map((point) => decodeMojibake(textContent(childByLocalName(point, 'v')).trim()));
  }

  const numCache = descendantByLocalName(node, 'numCache');
  if (!numCache) return [];

  const cacheFormatCode = decodeMojibake(textContent(childByLocalName(numCache, 'formatCode')).trim());
  return descendantsByLocalName(numCache, 'pt')
    .sort((a, b) => Number(attr(a, 'idx') ?? 0) - Number(attr(b, 'idx') ?? 0))
    .map((point) => {
      const value = decodeMojibake(textContent(childByLocalName(point, 'v')).trim());
      const formatCode = decodeMojibake(attr(point, 'formatCode') ?? cacheFormatCode);
      return formatCacheValue(value, formatCode, date1904);
    });
}

function readNumericValues(node: Element | null | undefined) {
  return readCacheValues(node)
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
}

function normalizeType(chartNode: Element | null): OfficeChartType {
  if (!chartNode) return 'unknown';
  const localName = (chartNode.localName.split(':').pop() ?? chartNode.localName).toLowerCase();
  if (localName === 'barchart') {
    const barDir = attr(childByLocalName(chartNode, 'barDir'), 'val');
    return barDir === 'bar' ? 'bar' : 'column';
  }
  return CHART_NODE_TO_TYPE[localName] ?? 'unknown';
}

function readSeriesColorWithTheme(seriesNode: Element, theme: OfficeTheme) {
  const spPr = childByLocalName(seriesNode, 'spPr');
  const fillNode = childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill');
  const lineNode = childByLocalName(spPr, 'ln');
  const color = readFillValue(fillNode, theme) ?? readFillColor(lineNode, theme);
  if (typeof color === 'string') return color;

  const fallbackFill = childByLocalName(childByLocalName(seriesNode, 'spPr'), 'solidFill');
  const fallbackLine = childByLocalName(childByLocalName(seriesNode, 'spPr'), 'ln');
  const fallback = readFillValue(fallbackFill, DEFAULT_OFFICE_THEME) ?? readFillColor(fallbackLine, DEFAULT_OFFICE_THEME);
  return typeof fallback === 'string' ? fallback : undefined;
}

function readFillColor(node: Element | null | undefined, theme: OfficeTheme) {
  const fillNode = childByLocalName(node, 'solidFill') ?? childByLocalName(node, 'gradFill');
  const color = readFillValue(fillNode, theme);
  return typeof color === 'string' ? color : undefined;
}

function readPointStyles(seriesNode: Element, theme: OfficeTheme) {
  const styles: OfficeChartSeries['pointStyles'] = [];
  childrenByLocalName(seriesNode, 'dPt').forEach((pointNode) => {
    const index = Number(attr(childByLocalName(pointNode, 'idx'), 'val'));
    if (!Number.isFinite(index)) return;
    const spPr = childByLocalName(pointNode, 'spPr');
    if (!spPr) return;
    const fillNode = childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill');
    const lineNode = childByLocalName(spPr, 'ln');
    const color = readFillValue(fillNode, theme);
    const borderColor = readFillColor(lineNode, theme);
    const borderWidth = readLineWidth(lineNode);
    const hasVisibleBorder = Boolean(lineNode && !childByLocalName(lineNode, 'noFill') && (borderColor || borderWidth !== undefined));
    if (color || hasVisibleBorder) {
      styles[index] = {
        color,
        borderColor,
        borderWidth: borderWidth ?? (hasVisibleBorder ? 1 : undefined),
      };
    }
  });
  return styles;
}

function readPointColors(seriesNode: Element, theme: OfficeTheme) {
  const colors: string[] = [];
  childrenByLocalName(seriesNode, 'dPt').forEach((pointNode) => {
    const index = Number(attr(childByLocalName(pointNode, 'idx'), 'val'));
    if (!Number.isFinite(index)) return;
    const color = readFillColor(childByLocalName(pointNode, 'spPr'), theme);
    if (color) colors[index] = color;
  });
  return colors;
}

function localName(node: Element | null | undefined) {
  return (node?.localName.split(':').pop() ?? node?.localName ?? '').toLowerCase();
}

function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

function clamp255(value: number) {
  return Math.max(0, Math.min(255, value));
}

function normalizeHex(value?: string) {
  if (!value) return undefined;
  if (/^#?[0-9a-f]{6}$/i.test(value)) {
    return value.startsWith('#') ? value : `#${value}`;
  }
  return undefined;
}

function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  const value = Number.parseInt(normalized, 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

function rgbToHex(r: number, g: number, b: number) {
  return `#${[r, g, b]
    .map((value) => clamp255(value).toString(16).padStart(2, '0'))
    .join('')}`;
}

function rgbToHsl(r: number, g: number, b: number) {
  const red = r / 255;
  const green = g / 255;
  const blue = b / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const lightness = (max + min) / 2;
  if (max === min) {
    return { h: 0, s: 0, l: lightness };
  }
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue = 0;
  switch (max) {
    case red:
      hue = (green - blue) / delta + (green < blue ? 6 : 0);
      break;
    case green:
      hue = (blue - red) / delta + 2;
      break;
    default:
      hue = (red - green) / delta + 4;
      break;
  }
  return { h: hue * 60, s: saturation, l: lightness };
}

function hslToRgb(h: number, s: number, l: number) {
  const hue = ((h % 360) + 360) % 360 / 360;
  if (s === 0) {
    const value = Math.round(l * 255);
    return { r: value, g: value, b: value };
  }

  const hue2rgb = (p: number, q: number, t: number) => {
    let temp = t;
    if (temp < 0) temp += 1;
    if (temp > 1) temp -= 1;
    if (temp < 1 / 6) return p + (q - p) * 6 * temp;
    if (temp < 1 / 2) return q;
    if (temp < 2 / 3) return p + (q - p) * (2 / 3 - temp) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: Math.round(hue2rgb(p, q, hue + 1 / 3) * 255),
    g: Math.round(hue2rgb(p, q, hue) * 255),
    b: Math.round(hue2rgb(p, q, hue - 1 / 3) * 255),
  };
}

function readColorNode(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  const kind = localName(node);
  const base =
    kind === 'srgbclr'
      ? normalizeHex(attr(node, 'val'))
      : kind === 'schemeclr'
        ? normalizeHex(resolveOfficeThemeColor(attr(node, 'val'), theme))
        : kind === 'sysclr'
          ? normalizeHex(attr(node, 'lastClr') ?? attr(node, 'val'))
          : kind === 'prstclr'
            ? normalizeHex(attr(node, 'val'))
            : normalizeHex(
                readColorNode(childByLocalName(node, 'srgbClr') ?? childByLocalName(node, 'schemeClr') ?? childByLocalName(node, 'sysClr') ?? childByLocalName(node, 'prstClr'), theme),
              );
  if (!base) return undefined;

  const transforms = Array.from(node.children)
    .map((child) => ({
      type: localName(child),
      val: Number(attr(child, 'val') ?? 0),
    }))
    .filter((item) =>
      ['tint', 'shade', 'lummod', 'lumoff', 'huemod', 'hueoff', 'satmod', 'satoff', 'alpha'].includes(item.type),
    );

  let alpha = 1;
  let { r, g, b } = hexToRgb(base);
  let hsl = rgbToHsl(r, g, b);

  transforms.forEach((transform) => {
    const ratio = transform.val / 100000;
    switch (transform.type) {
      case 'tint':
        hsl.l = clamp01(hsl.l + (1 - hsl.l) * ratio);
        break;
      case 'shade':
      case 'lummod':
        hsl.l = clamp01(hsl.l * ratio);
        break;
      case 'lumoff':
        hsl.l = clamp01(hsl.l + ratio);
        break;
      case 'huemod':
        hsl.h *= ratio;
        break;
      case 'hueoff':
        hsl.h += transform.val / 60000;
        break;
      case 'satmod':
        hsl.s = clamp01(hsl.s * ratio);
        break;
      case 'satoff':
        hsl.s = clamp01(hsl.s + ratio);
        break;
      case 'alpha':
        alpha = clamp01(ratio);
        break;
      default:
        break;
    }
  });

  const rgb = hslToRgb(hsl.h, hsl.s, hsl.l);
  return alpha < 1 ? `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})` : rgbToHex(rgb.r, rgb.g, rgb.b);
}

function readFillValue(node: Element | null | undefined, theme: OfficeTheme): OfficeChartColor | undefined {
  if (!node) return undefined;
  const kind = localName(node);
  if (kind === 'gradfill') {
    const stops = childrenByLocalName(childByLocalName(node, 'gsLst'), 'gs')
      .map((stop) => {
        const offset = clamp01(Number(attr(stop, 'pos') ?? 0) / 100000);
        const color =
          readColorNode(childByLocalName(stop, 'srgbClr') ?? childByLocalName(stop, 'schemeClr') ?? childByLocalName(stop, 'sysClr') ?? childByLocalName(stop, 'prstClr'), theme) ??
          undefined;
        return color ? { offset, color } : undefined;
      })
      .filter((stop): stop is OfficeChartColorStop => Boolean(stop))
      .sort((a, b) => a.offset - b.offset);

    if (!stops.length) return undefined;

    const angle = Number(attr(childByLocalName(node, 'lin'), 'ang') ?? 5400000) / 60000;
    const radians = (angle * Math.PI) / 180;
    const x = 0.5 - Math.cos(radians) / 2;
    const y = 0.5 - Math.sin(radians) / 2;
    const x2 = 0.5 + Math.cos(radians) / 2;
    const y2 = 0.5 + Math.sin(radians) / 2;

    return {
      type: 'linear',
      x,
      y,
      x2,
      y2,
      global: false,
      colorStops: stops,
    };
  }

  return readColorNode(node, theme);
}

function readLineWidth(node: Element | null | undefined) {
  const width = Number(attr(node, 'w'));
  if (!Number.isFinite(width) || width <= 0) return undefined;
  return width / 9525;
}

function readPositiveNumber(node: Element | null | undefined) {
  const value = Number(attr(node, 'val'));
  return Number.isFinite(value) && value > 0 ? value : undefined;
}

function readOfPieSecondPlotCount(chartNode: Element, pointCount: number) {
  if (pointCount <= 1) return 0;
  const splitPos = readPositiveNumber(childByLocalName(chartNode, 'splitPos'));
  const splitType = attr(childByLocalName(chartNode, 'splitType'), 'val');
  if ((splitType === 'pos' || !splitType) && splitPos) {
    return Math.max(1, Math.min(pointCount - 1, Math.round(splitPos)));
  }

  // Some WPS/Office files omit splitType/splitPos after saving. In that case
  // the extra dPt after the real data points styles the aggregate "Other"
  // slice, while the last two real points are expanded in the secondary plot.
  return Math.max(1, Math.min(pointCount - 1, 2));
}

function isPieLikeChart(type: OfficeChartType, ofPieType?: 'bar' | 'pie') {
  return type === 'pie' || type === 'doughnut' || Boolean(ofPieType);
}

function readShowDataLabels(chartNode: Element | null) {
  return descendantsByLocalName(chartNode, 'dLbls').some((labelsNode) => {
    const showVal = childByLocalName(labelsNode, 'showVal');
    return attr(showVal, 'val') === '1' || attr(showVal, 'val') === 'true';
  });
}

function readDataLabels(labelsNode: Element | null | undefined): OfficeDataLabels | undefined {
  if (!labelsNode) return undefined;
  const dataLabels: OfficeDataLabels = {};
  const readBool = (name: string) => {
    const value = attr(childByLocalName(labelsNode, name), 'val');
    if (value === undefined) return undefined;
    return value === '1' || value === 'true';
  };
  const deleted = attr(childByLocalName(labelsNode, 'delete'), 'val');
  const position = attr(childByLocalName(labelsNode, 'dLblPos'), 'val');
  const separator = textContent(childByLocalName(labelsNode, 'separator')).trim();

  if (deleted === '1' || deleted === 'true') dataLabels.delete = true;
  if (position) dataLabels.position = position;
  if (separator) dataLabels.separator = separator;
  const showLegendKey = readBool('showLegendKey');
  const showVal = readBool('showVal');
  const showCatName = readBool('showCatName');
  const showSerName = readBool('showSerName');
  const showPercent = readBool('showPercent');
  const showBubbleSize = readBool('showBubbleSize');
  const showLeaderLines = readBool('showLeaderLines');
  if (showLegendKey !== undefined) dataLabels.showLegendKey = showLegendKey;
  if (showVal !== undefined) dataLabels.showVal = showVal;
  if (showCatName !== undefined) dataLabels.showCatName = showCatName;
  if (showSerName !== undefined) dataLabels.showSerName = showSerName;
  if (showPercent !== undefined) dataLabels.showPercent = showPercent;
  if (showBubbleSize !== undefined) dataLabels.showBubbleSize = showBubbleSize;
  if (showLeaderLines !== undefined) dataLabels.showLeaderLines = showLeaderLines;
  return Object.keys(dataLabels).length ? dataLabels : undefined;
}

function readLegendPosition(chartNode: Element | null) {
  const value = attr(childByLocalName(chartNode, 'legendPos'), 'val');
  if (value === 'b') return 'bottom';
  if (value === 'l') return 'left';
  if (value === 'r') return 'right';
  return 'top';
}

function readFontFamily(node: Element | null | undefined) {
  const latin = attr(childByLocalName(node, 'latin'), 'typeface');
  const eastAsia = attr(childByLocalName(node, 'ea'), 'typeface');
  const complex = attr(childByLocalName(node, 'cs'), 'typeface');
  const value = [eastAsia, latin, complex].filter((item) => item && !item.startsWith('+')).join(', ');
  return value || undefined;
}

function readLegendVisible(chartNode: Element | null) {
  if (!chartNode) return false;
  const deleted = attr(childByLocalName(chartNode, 'delete'), 'val');
  return deleted !== '1' && deleted !== 'true';
}

function readLegendStyle(chartNode: Element | null, theme: OfficeTheme): OfficeChartModel['legendStyle'] | undefined {
  const runProps = descendantByLocalName(chartNode, 'defRPr');
  const fontSize = Number(attr(runProps, 'sz'));
  const color = readFillColor(runProps, theme);
  const fontFamily = readFontFamily(runProps);
  const bold = attr(runProps, 'b');
  const italic = attr(runProps, 'i');
  const textStyle = {
    color,
    fontFamily,
    fontSize: Number.isFinite(fontSize) && fontSize > 0 ? fontSize / 100 : undefined,
    fontWeight: bold === '1' || bold === 'true' ? 600 : undefined,
    fontStyle: italic === '1' || italic === 'true' ? 'italic' : undefined,
  };
  const normalizedTextStyle = Object.fromEntries(Object.entries(textStyle).filter(([, value]) => value !== undefined));
  return Object.keys(normalizedTextStyle).length ? { textStyle: normalizedTextStyle } : undefined;
}

function readSeriesMarker(seriesNode: Element) {
  const markerNode = childByLocalName(seriesNode, 'marker');
  const symbol = attr(childByLocalName(markerNode, 'symbol'), 'val');
  const size = Number(attr(childByLocalName(markerNode, 'size'), 'val'));
  return {
    symbol: symbol ? symbol.toLowerCase() : undefined,
    size: Number.isFinite(size) && size > 0 ? size : undefined,
  };
}

function looksLikeDateFormat(formatCode: string) {
  return /[ymdhs]/i.test(formatCode) && !/^general$/i.test(formatCode);
}

function formatDateFromSerial(serial: number, formatCode: string, date1904: boolean) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const date = new Date(epoch + serial * 24 * 60 * 60 * 1000);
  const year = String(date.getUTCFullYear());
  const month = String(date.getUTCMonth() + 1);
  const paddedMonth = month.padStart(2, '0');
  const day = String(date.getUTCDate());
  const paddedDay = day.padStart(2, '0');

  return formatCode
    .replace(/yyyy/g, year)
    .replace(/yy/g, year.slice(-2))
    .replace(/mm/g, paddedMonth)
    .replace(/m/g, month)
    .replace(/dd/g, paddedDay)
    .replace(/d/g, day);
}

function formatCacheValue(value: string, formatCode: string, date1904: boolean) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return value;
  }

  if (!formatCode || !looksLikeDateFormat(formatCode)) {
    return value;
  }

  return formatDateFromSerial(numeric, formatCode, date1904);
}

function readDate1904(chartSpace: Element | null) {
  const date1904 = childByLocalName(chartSpace, 'date1904');
  return attr(date1904, 'val') === '1' || attr(date1904, 'val') === 'true';
}

function findChartNodes(plotArea: Element | null) {
  return Array.from(plotArea?.children ?? []).filter((child) =>
    (child.localName.split(':').pop() ?? child.localName).toLowerCase().endsWith('chart'),
  );
}

function readChartPlot(chartNode: Element, theme: OfficeTheme, date1904: boolean) {
  const type = normalizeType(chartNode);
  const chartKind = localName(chartNode);
  const seriesNodes = childrenByLocalName(chartNode, 'ser');
  const firstSeries = seriesNodes[0];
  const plotDataLabels = readDataLabels(childrenByLocalName(chartNode, 'dLbls')[0]);
  const grouping = attr(childByLocalName(chartNode, 'grouping'), 'val');
  const stacking: OfficeChartSeries['stacking'] =
    grouping === 'stacked' || grouping === 'percentStacked' ? grouping : undefined;
  const stackGroup = stacking ? `office-chart-${type}` : undefined;
  const categories = readCacheValues(descendantByLocalName(firstSeries, 'cat'), date1904).map(decodeMojibake);
  const firstSliceAngle = Number(attr(childByLocalName(chartNode, 'firstSliceAng'), 'val'));
  const gapWidth = readPositiveNumber(childByLocalName(chartNode, 'gapWidth'));
  const overlapValue = Number(attr(childByLocalName(chartNode, 'overlap'), 'val'));
  const overlap = Number.isFinite(overlapValue) ? overlapValue : undefined;
  const series = seriesNodes.map((seriesNode, index) => ({
    name: firstText(descendantByLocalName(seriesNode, 'tx')) || `Series ${index + 1}`,
    type,
    stacking,
    stackGroup,
    gapWidth,
    overlap,
    values: readNumericValues(descendantByLocalName(seriesNode, 'val')),
    color: readSeriesColorWithTheme(seriesNode, theme),
    lineWidth: readLineWidth(childByLocalName(childByLocalName(seriesNode, 'spPr'), 'ln')),
    pointColors: readPointColors(seriesNode, theme),
    pointStyles: readPointStyles(seriesNode, theme),
    dataLabels: readDataLabels(childrenByLocalName(seriesNode, 'dLbls')[0]) ?? plotDataLabels,
    smooth: attr(childByLocalName(seriesNode, 'smooth'), 'val') === '1' || attr(childByLocalName(seriesNode, 'smooth'), 'val') === 'true',
    marker: readSeriesMarker(seriesNode),
  }));
  const firstValueCount = series[0]?.values.length ?? 0;
  const ofPieTypeValue = attr(childByLocalName(chartNode, 'ofPieType'), 'val');
  const ofPieType = chartKind === 'ofpiechart' && (ofPieTypeValue === 'bar' || ofPieTypeValue === 'pie') ? ofPieTypeValue : undefined;

  return {
    type,
    categories,
    series,
    dataLabels: plotDataLabels,
    holeSize: type === 'doughnut' ? Number(attr(childByLocalName(chartNode, 'holeSize'), 'val') ?? 0) || undefined : undefined,
    startAngle: isPieLikeChart(type, ofPieType) ? Number.isFinite(firstSliceAngle) ? firstSliceAngle : 0 : undefined,
    ofPieType,
    ofPieSecondPlotCount: ofPieType ? readOfPieSecondPlotCount(chartNode, firstValueCount) : undefined,
    secondPieSize: ofPieType ? readPositiveNumber(childByLocalName(chartNode, 'secondPieSize')) : undefined,
    gapWidth,
    overlap,
    radarStyle: type === 'radar' ? attr(childByLocalName(chartNode, 'radarStyle'), 'val') : undefined,
  };
}

function niceRadarMax(value: number) {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const step = normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

function buildRadarIndicators(categories: string[], series: OfficeChartSeries[]) {
  return categories.map((name, index) => {
    const maxValue = series.reduce((acc, item) => Math.max(acc, item.values[index] ?? 0), 0);
    return {
      name,
      max: niceRadarMax(maxValue),
    };
  });
}

export function parseOfficeChartXml(xml: string, theme: OfficeTheme = DEFAULT_OFFICE_THEME): OfficeChartModel {
  const doc = parseXml(xml);
  const chartSpace = doc.documentElement;
  const chart = descendantByLocalName(chartSpace, 'chart');
  const plotArea = descendantByLocalName(chart, 'plotArea');
  const date1904 = readDate1904(chartSpace);
  const plots = findChartNodes(plotArea).map((chartNode) => readChartPlot(chartNode, theme, date1904));
  const primaryPlot = plots[0];
  const categories = primaryPlot?.categories ?? [];
  const series = plots.flatMap((plot) => plot.series);
  const type = primaryPlot?.type ?? 'unknown';
  const rawTitle = firstText(childByLocalName(chart, 'title'));
  const autoTitleDeleted = attr(childByLocalName(chart, 'autoTitleDeleted'), 'val');
  const title =
    rawTitle ||
    (childByLocalName(chart, 'title') && autoTitleDeleted !== '1' && autoTitleDeleted !== 'true'
      ? series[0]?.name
      : undefined);

  return {
    type,
    title: title || undefined,
    categories,
    series,
    showLegend: readLegendVisible(childByLocalName(chart, 'legend')),
    legendPosition: readLegendPosition(childByLocalName(chart, 'legend')),
    legendStyle: readLegendStyle(childByLocalName(chart, 'legend'), theme),
    showDataLabels: readShowDataLabels(plotArea),
    radarIndicators: type === 'radar' ? buildRadarIndicators(categories, series) : undefined,
    holeSize: primaryPlot?.holeSize,
    startAngle: primaryPlot?.startAngle,
    ofPieType: primaryPlot?.ofPieType,
    ofPieSecondPlotCount: primaryPlot?.ofPieSecondPlotCount,
    secondPieSize: primaryPlot?.secondPieSize,
    gapWidth: primaryPlot?.gapWidth,
    overlap: primaryPlot?.overlap,
    radarStyle: primaryPlot?.radarStyle,
  };
}

function resolveSeriesColor(series: OfficeChartSeries, index: number) {
  return series.color ?? DEFAULT_COLORS[index % DEFAULT_COLORS.length];
}

function resolveCategories(chart: OfficeChartModel) {
  if (chart.categories.length) {
    return chart.categories;
  }

  const maxLength = Math.max(...chart.series.map((series) => series.values.length), 0);
  return Array.from({ length: maxLength }, (_, index) => String(index + 1));
}

function normalizeSeriesType(type: OfficeChartType) {
  if (type === 'column' || type === 'bar') return 'bar';
  if (type === 'area') return 'line';
  if (type === 'scatter' || type === 'bubble') return 'scatter';
  if (type === 'radar') return 'radar';
  if (type === 'pie' || type === 'doughnut') return 'pie';
  return 'line';
}

function sanitizeMapRegionName(name: string) {
  return name
    .replace(/特别行政区$|壮族自治区$|回族自治区$|维吾尔自治区$|自治区$|省$|市$/g, '')
    .trim();
}

function scaleRoseRadius(radius: [string, string] | undefined): [string, string] | undefined {
  if (!radius) return undefined;
  const inner = Number(radius[0].replace(/%$/, ''));
  const outer = Number(radius[1].replace(/%$/, ''));
  if (!Number.isFinite(inner) || !Number.isFinite(outer) || outer <= 0) return radius;
  const fittedOuter = Math.min(58, outer);
  const fittedInner = Math.max(0, Math.min(fittedOuter - 4, Math.round((inner / outer) * fittedOuter)));
  return [`${fittedInner}%`, `${fittedOuter}%`];
}

function buildLegend(chart: OfficeChartModel, itemCount = chart.series.length) {
  if (chart.showLegend === false || itemCount <= 0) return undefined;

  const base = {
    type: 'scroll' as const,
    itemWidth: chart.legendStyle?.itemWidth ?? 10,
    itemHeight: chart.legendStyle?.itemHeight ?? 10,
    textStyle: {
      ...OFFICE_TEXT_STYLE,
      ...chart.legendStyle?.textStyle,
    },
  };

  switch (chart.legendPosition) {
    case 'bottom':
      return {
        ...base,
        bottom: 4,
      };
    case 'left':
      return {
        ...base,
        left: 8,
        top: chart.title ? 32 : 8,
        orient: 'vertical' as const,
      };
    case 'right':
      return {
        ...base,
        right: 8,
        top: chart.title ? 32 : 8,
        orient: 'vertical' as const,
      };
    default:
      return {
        ...base,
        top: chart.title ? 30 : 8,
      };
  }
}

function buildChartGrid(chart: OfficeChartModel) {
  const isBottomLegend = chart.legendPosition === 'bottom';
  const isSideLegend = chart.legendPosition === 'left' || chart.legendPosition === 'right';
  return {
    left: isSideLegend ? 70 : 40,
    right: isSideLegend ? 70 : 24,
    top: chart.title ? 56 : chart.legendPosition === 'top' ? 40 : 24,
    bottom: isBottomLegend ? 56 : 32,
    containLabel: true,
  };
}

function buildOfficeTitle(chart: OfficeChartModel) {
  return chart.title
    ? {
        text: chart.title,
        left: 'center',
        top: 8,
        textStyle: {
          fontSize: 14,
          fontWeight: 600,
          color: '#111827',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      }
    : undefined;
}

function resolveOfficePieStartAngle(chart: OfficeChartModel) {
  const officeAngle = chart.startAngle ?? 0;
  return ((90 - officeAngle) % 360 + 360) % 360;
}

function resolveOfficeRadarStartAngle(chart: OfficeChartModel) {
  return chart.radarStartAngle ?? 90;
}

function resolveOfficeRadarRadius(chart: OfficeChartModel) {
  return chart.radarRadius ?? (chart.legendPosition === 'bottom' ? '62%' : '68%');
}

function reorderRadarAxes<T>(items: T[]) {
  if (items.length <= 2) return items.slice();
  return [items[0], ...items.slice(1).reverse()];
}

function resolveOfficeRadarCenter(chart: OfficeChartModel): [string, string] {
  if (chart.legendPosition === 'bottom') return ['50%', chart.title ? '52%' : '48%'];
  if (chart.legendPosition === 'top') return ['50%', chart.title ? '58%' : '56%'];
  return ['50%', chart.title ? '56%' : '52%'];
}

function resolveBarWidthFromGap(gapWidth: number | undefined, seriesCount: number, overlap?: number) {
  if (!Number.isFinite(gapWidth) || gapWidth === undefined) return undefined;
  const visibleSeriesCount = Math.max(1, seriesCount);
  const overlapRatio = Math.max(-1, Math.min(1, (overlap ?? 0) / 100));
  const effectiveSeriesCount = Math.max(1, visibleSeriesCount - Math.max(0, overlapRatio) * (visibleSeriesCount - 1));
  const categoryWidth = 72;
  const width = categoryWidth / (effectiveSeriesCount + gapWidth / 100);
  return Math.max(6, Math.min(46, Math.round(width)));
}

function readPieLabelPosition(labels?: OfficeDataLabels, fallback: 'outside' | 'inside' = 'outside') {
  const position = labels?.position?.toLowerCase();
  if (!position) return fallback;
  if (position.includes('in')) return 'inside';
  if (position.includes('out')) return 'outside';
  return fallback;
}

function readCartesianLabelPosition(labels?: OfficeDataLabels, horizontal = false) {
  const position = labels?.position?.toLowerCase();
  if (!position) return horizontal ? 'right' : 'top';
  if (position === 'ctr' || position === 'center') return 'inside';
  if (position === 'inbase') return horizontal ? 'insideLeft' : 'insideBottom';
  if (position === 'inend') return horizontal ? 'insideRight' : 'insideTop';
  if (position === 'outend') return horizontal ? 'right' : 'top';
  if (position.includes('base')) return horizontal ? 'left' : 'bottom';
  if (position.includes('end')) return horizontal ? 'right' : 'top';
  return horizontal ? 'right' : 'top';
}

function buildDataLabelFormatter(labels: OfficeDataLabels | undefined, categories: string[]) {
  const showValue = labels?.showVal ?? false;
  const showCategory = labels?.showCatName ?? false;
  const showSeries = labels?.showSerName ?? false;
  const showPercent = labels?.showPercent ?? false;
  const separator = labels?.separator ?? '\n';
  return (params: unknown) => {
    const item = params as { name?: string; value?: unknown; dataIndex?: number; percent?: number; seriesName?: string };
    const value = Array.isArray(item.value) ? item.value[item.value.length - 1] : item.value;
    const category = item.name ?? (item.dataIndex !== undefined ? categories[item.dataIndex] : undefined);
    const parts: string[] = [];
    if (showSeries && item.seriesName) parts.push(item.seriesName);
    if (showCategory && category) parts.push(category);
    if (showValue && value !== undefined) parts.push(String(value));
    if (showPercent && item.percent !== undefined) parts.push(`${item.percent}%`);
    return parts.join(separator).trim();
  };
}

function shouldShowDataLabels(labels: OfficeDataLabels | undefined, chartShowDataLabels?: boolean) {
  if (labels?.delete) return false;
  if (!labels) return Boolean(chartShowDataLabels);
  const explicitFlags = [labels.showVal, labels.showCatName, labels.showSerName, labels.showPercent].filter(
    (value) => value !== undefined,
  );
  if (explicitFlags.length) return explicitFlags.some(Boolean);
  return Boolean(chartShowDataLabels);
}

function buildCartesianDataLabelConfig(
  labels: OfficeDataLabels | undefined,
  chartShowDataLabels: boolean | undefined,
  categories: string[],
  horizontal = false,
) {
  const effectiveLabels = labels ?? (chartShowDataLabels ? { showVal: true } : undefined);
  return {
    show: shouldShowDataLabels(effectiveLabels, chartShowDataLabels),
    position: readCartesianLabelPosition(effectiveLabels, horizontal),
    formatter: buildDataLabelFormatter(effectiveLabels, categories),
    color: '#334155',
    fontFamily: OFFICE_FONT_FAMILY,
  };
}

function buildPieDataLabelConfig(labels: OfficeDataLabels | undefined, showDataLabels?: boolean) {
  const position = readPieLabelPosition(labels, 'outside');
  const showValue = labels?.showVal ?? showDataLabels;
  const showCategory = labels?.showCatName ?? false;
  const showSeries = labels?.showSerName ?? false;
  const showPercent = labels?.showPercent ?? false;
  const separator = labels?.separator ?? '\n';
  const formatter = (params: unknown) => {
    const item = params as { name?: string; data?: { name?: string }; value?: number; percent?: number; seriesName?: string };
    const parts: string[] = [];
    if (showSeries && item.seriesName) parts.push(item.seriesName);
    if (showCategory && (item.data?.name ?? item.name)) parts.push(item.data?.name ?? item.name ?? '');
    if (showValue && item.value !== undefined) parts.push(String(item.value));
    if (showPercent && item.percent !== undefined) parts.push(`${item.percent}%`);
    return parts.join(separator).trim();
  };

  return {
    show: !labels?.delete && Boolean(showValue || showCategory || showSeries || showPercent || showDataLabels),
    position,
    formatter,
  };
}

function buildOfPieChartOption(chart: OfficeChartModel, categories: string[], palette: string[]): EChartsOption {
  const sourceSeries = chart.series[0];
  const pieLabels = chart.dataLabels ?? sourceSeries?.dataLabels;
  const values = sourceSeries?.values ?? [];
  const secondCount = Math.max(1, Math.min(values.length - 1, chart.ofPieSecondPlotCount ?? Math.ceil(values.length / 3)));
  const splitIndex = Math.max(1, values.length - secondCount);
  const mainNames = categories.slice(0, splitIndex);
  const secondaryNames = categories.slice(splitIndex, values.length);
  const secondaryValues = values.slice(splitIndex);
  const secondaryTotal = secondaryValues.reduce((sum, value) => sum + value, 0);
  const otherName = '其他';
  const otherStyle = buildPieItemStyle(sourceSeries, categories.length, palette);
  const startAngle = resolveOfficePieStartAngle(chart);
  const total = values.reduce((sum, value) => sum + value, 0);
  const beforeOtherTotal = values.slice(0, splitIndex).reduce((sum, value) => sum + value, 0);
  const otherStart = total ? startAngle - (beforeOtherTotal / total) * 360 : startAngle;
  const otherEnd = total ? otherStart - (secondaryTotal / total) * 360 : startAngle;
  const otherMid = ((otherStart + otherEnd) / 2) * (Math.PI / 180);
  const connectorAnchorX = 34 + Math.cos(otherMid) * 23;
  const connectorAnchorY = (chart.title ? 58 : 52) - Math.sin(otherMid) * 23;
  const mainData = [
    ...mainNames.map((name, index) => ({
      name,
      value: values[index] ?? 0,
      itemStyle: buildPieItemStyle(sourceSeries, index, palette),
    })),
    {
      name: otherName,
      value: secondaryTotal,
      itemStyle: otherStyle,
      tooltip: {
        formatter: `${otherName}<br/>${sourceSeries?.name ?? ''}: ${secondaryTotal}`,
      },
    },
  ];
  const secondarySize = Math.max(28, Math.min(70, Math.round(52 * ((chart.secondPieSize ?? 75) / 75))));
  const legend = buildLegend(chart, categories.length) as Record<string, unknown> | undefined;
  const tooltip = {
    trigger: 'item' as const,
    confine: true,
    appendToBody: true,
    backgroundColor: 'rgba(15, 23, 42, 0.96)',
    borderColor: 'rgba(15, 23, 42, 0.96)',
    textStyle: {
      color: '#fff',
      fontFamily: OFFICE_FONT_FAMILY,
    },
    formatter: (params: unknown) => {
      const item = params as { componentSubType?: string; data?: { name?: string }; name?: string; seriesName?: string; value?: unknown };
      const value = typeof item.value === 'number' ? item.value : Array.isArray(item.value) ? item.value[0] : '';
      const name = item.componentSubType === 'bar' ? item.seriesName : item.data?.name ?? item.name ?? '';
      return `${name}<br/>${sourceSeries?.name ?? ''}: ${value}`;
    },
  };

  const secondarySeries =
    chart.ofPieType === 'pie'
      ? [
          {
            type: 'pie' as const,
            radius: ['0%', `${secondarySize}%`] as [string, string],
            center: ['72%', chart.title ? '58%' : '52%'] as [string, string],
            startAngle,
            avoidLabelOverlap: true,
            label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
            labelLayout: {
              hideOverlap: true,
            },
            labelLine: {
              length: 12,
              length2: 8,
              smooth: true,
            },
            data: secondaryNames.map((name, index) => ({
              name,
              value: secondaryValues[index] ?? 0,
              itemStyle: buildPieItemStyle(sourceSeries, splitIndex + index, palette),
            })),
          },
        ]
      : secondaryNames.map((name, index) => ({
          name,
          type: 'bar' as const,
          stack: 'office-of-pie-secondary',
          barWidth: Math.max(26, Math.min(46, Math.round(secondarySize * 0.72))),
          xAxisIndex: 0,
          yAxisIndex: 0,
          data: [
            {
              value: secondaryValues[index] ?? 0,
              itemStyle: buildPieItemStyle(sourceSeries, splitIndex + index, palette),
            },
          ],
          label: {
            show: chart.showDataLabels,
            position: 'inside' as const,
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
          emphasis: {
            itemStyle: {
              shadowBlur: 6,
              shadowColor: 'rgba(15, 23, 42, 0.18)',
            },
          },
        }));

  return {
    animation: false,
    backgroundColor: '#fff',
    color: palette,
    textStyle: OFFICE_TEXT_STYLE,
    title: buildOfficeTitle(chart),
    tooltip,
    legend: legend
      ? {
          ...legend,
          data: categories,
        }
      : undefined,
    grid:
      chart.ofPieType === 'bar'
        ? {
            left: '64%',
            top: chart.title ? '34%' : '26%',
            width: '16%',
            height: chart.title ? '46%' : '52%',
            containLabel: false,
          }
        : undefined,
    xAxis:
      chart.ofPieType === 'bar'
        ? {
            type: 'category' as const,
            data: [''],
            show: false,
          }
        : undefined,
    yAxis:
      chart.ofPieType === 'bar'
        ? {
            type: 'value' as const,
            min: 0,
            max: secondaryTotal || undefined,
            show: false,
          }
        : undefined,
    graphic:
      chart.ofPieType === 'bar'
        ? [
            {
              type: 'line',
              left: `${connectorAnchorX}%`,
              top: `${connectorAnchorY}%`,
              shape: { x1: 0, y1: 0, x2: 132, y2: chart.title ? -28 : -24 },
              style: { stroke: '#cbd5e1', lineWidth: 1 },
              silent: true,
            },
            {
              type: 'line',
              left: `${connectorAnchorX}%`,
              top: `${connectorAnchorY}%`,
              shape: { x1: 0, y1: 0, x2: 132, y2: chart.title ? 44 : 40 },
              style: { stroke: '#cbd5e1', lineWidth: 1 },
              silent: true,
            },
          ]
        : undefined,
    series: [
      {
        type: 'pie' as const,
        radius: ['0%', '46%'] as [string, string],
        center: ['34%', chart.title ? '58%' : '52%'] as [string, string],
        startAngle,
        avoidLabelOverlap: true,
        label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
        labelLayout: {
          hideOverlap: true,
        },
        labelLine: {
          length: 12,
          length2: 8,
          smooth: true,
        },
        emphasis: {
          scale: false,
          itemStyle: {
            shadowBlur: 8,
            shadowColor: 'rgba(15, 23, 42, 0.18)',
          },
        },
        data: mainData,
      },
      ...secondarySeries,
    ],
  };
}

export function buildOfficeChartOption(chart: OfficeChartModel): EChartsOption {
  const categories = resolveCategories(chart);
  const normalizedSeriesTypes = chart.series.map((item) => normalizeSeriesType(item.type ?? chart.type));
  const uniqueSeriesTypes = new Set(normalizedSeriesTypes);
  const isHorizontalBar = chart.type === 'bar';
  const isPie = chart.type === 'pie' || chart.type === 'doughnut';
  const isRadar = chart.type === 'radar';
  const isScatter = normalizedSeriesTypes.length > 0 && normalizedSeriesTypes.every((type) => type === 'scatter');
  const usesMixedSeriesTypes = uniqueSeriesTypes.size > 1;
  const palette = chart.series.map((series, index) => resolveSeriesColor(series, index));
  const hasSeries = chart.series.length > 0;

  if (!hasSeries) {
    return {
      animation: false,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
            },
          }
        : undefined,
    };
  }

  const radarIndicators = chart.radarIndicators ?? (isRadar ? buildRadarIndicators(categories, chart.series) : undefined);
  const radarDisplayIndicators = isRadar && radarIndicators?.length ? reorderRadarAxes(radarIndicators) : undefined;
  const radarCategories = radarDisplayIndicators?.map((indicator) => indicator.name) ?? categories;

  if (chart.type === 'map') {
    const sourceSeries = chart.series[0];
    const values = sourceSeries?.values ?? [];
    const tierNames = Array.from(new Set((sourceSeries?.pointLabels ?? []).filter(Boolean)));
    const tierColors = tierNames
      .map((tier) => {
        const index = sourceSeries?.pointLabels?.indexOf(tier) ?? -1;
        return index >= 0 ? sourceSeries?.pointColors?.[index] : undefined;
      })
      .filter((color): color is string => Boolean(color));
    const data = categories.map((name, index) => ({
      name,
      value: values[index] ?? 0,
      labelName: sanitizeMapRegionName(name),
      tierName: sourceSeries?.pointLabels?.[index],
      itemStyle: {
        areaColor: sourceSeries?.pointColors?.[index] ?? '#e5edf8',
        borderColor: '#ffffff',
        borderWidth: 1,
      },
    }));

    return {
      animation: false,
      backgroundColor: '#ffffff',
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            subtext: chart.mapRegion,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
            subtextStyle: {
              color: '#64748b',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
        formatter: (params: unknown) => {
          const item = params as { data?: { tierName?: string }; name?: string; value?: unknown };
          const value = typeof item.value === 'number' ? item.value : '';
          const tier = item.data?.tierName ? `<br/>${item.data.tierName}` : '';
          return `${item.name ?? ''}<br/>${sourceSeries?.name ?? chart.mapSeriesName ?? ''}: ${value}${tier}`;
        },
      },
      graphic: tierNames.length
        ? {
            type: 'group',
            left: 12,
            bottom: 12,
            children: tierNames.flatMap((name, index) => [
              {
                type: 'rect',
                shape: { x: 0, y: index * 20, width: 10, height: 10 },
                style: { fill: tierColors[index] ?? '#cbd5e1' },
              },
              {
                type: 'text',
                left: 16,
                top: index * 20 - 2,
                style: {
                  text: name,
                  fill: '#334155',
                  font: `12px ${OFFICE_FONT_FAMILY}`,
                },
              },
            ]),
          }
        : undefined,
      series: [
        {
          name: sourceSeries?.name ?? chart.mapSeriesName,
          type: 'map' as const,
          map: chart.mapName ?? 'china',
          roam: true,
          selectedMode: false,
          layoutCenter: ['50%', chart.title ? '56%' : '52%'],
          layoutSize: tierNames.length ? '88%' : '92%',
          zoom: 1.08,
          itemStyle: {
            areaColor: '#eef3f8',
            borderColor: '#f8fafc',
            borderWidth: 1,
          },
          emphasis: {
            label: {
              show: true,
              color: '#0f172a',
              fontFamily: OFFICE_FONT_FAMILY,
            },
            itemStyle: {
              areaColor: '#f59e0b',
              borderColor: '#ffffff',
              borderWidth: 1.2,
            },
          },
          label: {
            show: chart.showDataLabels ?? true,
            color: '#1f2937',
            fontFamily: OFFICE_FONT_FAMILY,
            fontSize: 9,
            formatter: (params: unknown) => {
              const item = params as { data?: { labelName?: string }; name?: string };
              return item.data?.labelName ?? sanitizeMapRegionName(item.name ?? '');
            },
          },
          data,
        },
      ],
    };
  }

  if (isRadar && radarDisplayIndicators?.length) {
    return {
      animation: false,
      backgroundColor: '#fff',
      color: palette,
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      },
      legend: buildLegend(chart),
      radar: {
        center: resolveOfficeRadarCenter(chart),
        radius: resolveOfficeRadarRadius(chart),
        startAngle: resolveOfficeRadarStartAngle(chart),
        indicator: radarDisplayIndicators,
        splitNumber: chart.radarSplitNumber ?? 5,
        axisName: {
          color: '#475569',
          fontFamily: OFFICE_FONT_FAMILY,
        },
        splitArea: {
          areaStyle: {
            color: ['rgba(255,255,255,0)', 'rgba(148,163,184,0.04)'],
          },
        },
        splitLine: {
          lineStyle: {
            color: '#e2e8f0',
          },
        },
        axisLine: {
          lineStyle: {
            color: '#cbd5e1',
          },
        },
      },
      series: chart.series.map((item, index) => {
        const color = resolveSeriesColor(item, index);
        return {
          name: item.name,
          type: 'radar',
          areaStyle: chart.radarStyle === 'filled' ? { opacity: 0.18 } : undefined,
          data: [
            {
              value: reorderRadarAxes(item.values.slice(0, radarDisplayIndicators.length)),
              name: item.name,
            },
          ],
          symbol: item.marker?.symbol ?? 'circle',
          symbolSize: item.marker?.size ?? 6,
          lineStyle: {
            color,
            width: item.lineWidth ?? 2,
          },
          itemStyle: {
            color,
          },
          label: {
            show: shouldShowDataLabels(item.dataLabels ?? chart.dataLabels, chart.showDataLabels),
            formatter: buildDataLabelFormatter(item.dataLabels ?? chart.dataLabels, radarCategories),
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
        };
      }),
    };
  }

  if (isPie && chart.ofPieType && categories.length > 1) {
    return buildOfPieChartOption(chart, categories, palette);
  }

  if (isPie && !usesMixedSeriesTypes) {
    const sourceSeries = chart.series[0];
    const pieLabels = chart.dataLabels ?? sourceSeries?.dataLabels;
    const data = categories.map((name, index) => ({
      name,
      value: sourceSeries?.values[index] ?? 0,
      itemStyle: buildPieItemStyle(sourceSeries, index, palette),
    }));
    const innerRadius = chart.type === 'doughnut' && chart.holeSize
      ? `${Math.max(8, Math.min(90, Math.round(68 * (chart.holeSize / 100))))}%`
      : '0%';
    const radius: [string, string] = chart.roseType
      ? scaleRoseRadius(chart.radius) ?? [innerRadius, '58%']
      : chart.radius ?? [innerRadius, '68%'];

    return {
      animation: false,
      backgroundColor: '#fff',
      color: palette,
      textStyle: OFFICE_TEXT_STYLE,
      title: chart.title
        ? {
            text: chart.title,
            left: 'center',
            top: 8,
            textStyle: {
              fontSize: 14,
              fontWeight: 600,
              color: '#111827',
              fontFamily: OFFICE_FONT_FAMILY,
            },
          }
        : undefined,
      tooltip: {
        trigger: 'item',
        confine: true,
        appendToBody: true,
        backgroundColor: 'rgba(15, 23, 42, 0.96)',
        borderColor: 'rgba(15, 23, 42, 0.96)',
        textStyle: {
          color: '#fff',
          fontFamily: OFFICE_FONT_FAMILY,
        },
      },
      legend: buildLegend(chart, categories.length),
      series: [
        {
          type: 'pie' as const,
          radius,
          roseType: chart.roseType,
          startAngle: resolveOfficePieStartAngle(chart),
          padAngle: 0,
          center: ['50%', chart.roseType ? '50%' : chart.title ? '58%' : '50%'],
          avoidLabelOverlap: true,
          label: buildPieDataLabelConfig(pieLabels, chart.showDataLabels),
          labelLayout: {
            hideOverlap: true,
          },
          emphasis: {
            scale: false,
            itemStyle: {
              borderColor: '#ffffff',
              borderWidth: 1,
              shadowBlur: 8,
              shadowColor: 'rgba(15, 23, 42, 0.18)',
            },
          },
          labelLine: {
            length: 12,
            length2: 8,
          },
          data,
        },
      ],
    };
  }

  const series = chart.series.map((item, index) => {
    const seriesType = normalizeSeriesType(item.type ?? chart.type) as 'line' | 'bar' | 'scatter' | 'radar' | 'pie';
    const color = resolveSeriesColor(item, index);
    const isBarSeries = seriesType === 'bar';
    const isLineSeries = seriesType === 'line';
    const markerSymbol = item.marker?.symbol && item.marker.symbol !== 'none' ? item.marker.symbol : undefined;
    const hideSymbol = item.marker?.symbol === 'none';
    const isBubbleSeries = item.type === 'bubble';
    const barSeriesCount = chart.series.filter((seriesItem) => normalizeSeriesType(seriesItem.type ?? chart.type) === 'bar' && !seriesItem.stackGroup).length;
    const barWidth = isBarSeries ? resolveBarWidthFromGap(item.gapWidth ?? chart.gapWidth, barSeriesCount, item.overlap ?? chart.overlap) : undefined;
    const labelConfig = buildCartesianDataLabelConfig(item.dataLabels ?? chart.dataLabels, chart.showDataLabels, categories, isHorizontalBar);

    return {
      name: item.name,
      type: seriesType,
      stack: item.stackGroup,
      data:
        isScatter || isBubbleSeries
          ? item.values.map((value, valueIndex) => [categories[valueIndex] ?? String(valueIndex + 1), value])
          : item.values,
      areaStyle: item.type === 'area' ? { opacity: 0.18 } : undefined,
      smooth: item.smooth ?? (item.type === 'line' || item.type === 'area'),
      itemStyle: {
        color,
        borderColor: isBarSeries ? '#fff' : color,
        borderWidth: isBarSeries ? 1 : 0,
      },
      lineStyle: {
        color,
        width: item.lineWidth ?? (isLineSeries || item.type === 'area' ? 2 : 1),
      },
      emphasis: {
        itemStyle: {
          color,
          borderColor: isBarSeries ? '#fff' : color,
          borderWidth: isBarSeries ? 1 : 0,
          shadowBlur: isBarSeries ? 6 : 0,
          shadowColor: 'rgba(15, 23, 42, 0.18)',
        },
        lineStyle: {
          color,
          width: item.lineWidth ? item.lineWidth + 1 : isLineSeries || item.type === 'area' ? 3 : 1,
        },
      },
      showSymbol: hideSymbol ? false : markerSymbol ? true : isLineSeries || item.type === 'area' || isBubbleSeries || seriesType === 'scatter',
      symbol: markerSymbol,
      symbolSize: item.marker?.size ?? (isBubbleSeries ? 14 : 8),
      label: labelConfig,
      barWidth,
      barGap: isBarSeries && item.overlap !== undefined ? `${-item.overlap}%` : undefined,
      barCategoryGap: isBarSeries && item.gapWidth !== undefined ? `${item.gapWidth}%` : undefined,
      barMaxWidth: isBarSeries && !barWidth ? 32 : undefined,
    };
  }) as EChartsOption['series'];

  return {
    animation: false,
    backgroundColor: '#fff',
    color: palette,
    textStyle: OFFICE_TEXT_STYLE,
    title: chart.title
      ? {
          text: chart.title,
          left: 'center',
          top: 8,
          textStyle: {
            fontSize: 14,
            fontWeight: 600,
            color: '#111827',
            fontFamily: OFFICE_FONT_FAMILY,
          },
        }
      : undefined,
    tooltip: {
      trigger: isScatter ? 'item' : 'axis',
      confine: true,
      appendToBody: true,
      axisPointer: isScatter
        ? undefined
        : {
            type: isHorizontalBar ? 'shadow' : 'line',
          },
      backgroundColor: 'rgba(15, 23, 42, 0.96)',
      borderColor: 'rgba(15, 23, 42, 0.96)',
      textStyle: {
        color: '#fff',
        fontFamily: OFFICE_FONT_FAMILY,
      },
    },
    legend: buildLegend(chart),
    grid: buildChartGrid(chart),
    xAxis: isHorizontalBar
      ? {
          type: 'value',
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          splitLine: {
            lineStyle: {
              color: '#eef2f7',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        }
      : {
          type: 'category',
          data: categories,
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        },
    yAxis: isHorizontalBar
      ? {
          type: 'category',
          data: categories,
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisLabel: {
            hideOverlap: true,
            color: '#475569',
          },
        }
      : {
          type: 'value',
          axisLine: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          axisTick: {
            lineStyle: {
              color: '#cbd5e1',
            },
          },
          splitLine: {
            lineStyle: {
              color: '#eef2f7',
            },
          },
          axisLabel: {
            color: '#475569',
          },
        },
    series,
  };
}

function buildPieItemStyle(series: OfficeChartSeries | undefined, index: number, palette: string[]) {
  const pointStyle = series?.pointStyles?.[index];
  const fallbackColor = series?.pointColors?.[index] ?? (series ? resolveSeriesColor(series, index) : palette[index % palette.length]);
  const color = pointStyle?.color ?? fallbackColor;
  const itemStyle: Record<string, unknown> = {
    color,
  };
  if (pointStyle?.borderColor !== undefined) {
    itemStyle.borderColor = pointStyle.borderColor;
  }
  if (pointStyle?.borderWidth !== undefined) {
    itemStyle.borderWidth = pointStyle.borderWidth;
  }
  return itemStyle;
}
