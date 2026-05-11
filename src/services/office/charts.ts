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
  color?: string;
  pointColors?: string[];
  pointLabels?: string[];
  pointStyles?: Array<{
    color?: OfficeChartColor;
    borderColor?: string;
    borderWidth?: number;
  }>;
  smooth?: boolean;
  lineWidth?: number;
  marker?: {
    symbol?: string;
    size?: number;
  };
};

export type OfficeChartModel = {
  type: OfficeChartType;
  title?: string;
  categories: string[];
  series: OfficeChartSeries[];
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
  roseType?: 'radius' | 'area';
  radius?: [string, string];
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

function readShowDataLabels(chartNode: Element | null) {
  return descendantsByLocalName(chartNode, 'dLbls').some((labelsNode) => {
    const showVal = childByLocalName(labelsNode, 'showVal');
    return attr(showVal, 'val') === '1' || attr(showVal, 'val') === 'true';
  });
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
  const seriesNodes = childrenByLocalName(chartNode, 'ser');
  const firstSeries = seriesNodes[0];
  const grouping = attr(childByLocalName(chartNode, 'grouping'), 'val');
  const stacking: OfficeChartSeries['stacking'] =
    grouping === 'stacked' || grouping === 'percentStacked' ? grouping : undefined;
  const stackGroup = stacking ? `office-chart-${type}` : undefined;
  const categories = readCacheValues(descendantByLocalName(firstSeries, 'cat'), date1904).map(decodeMojibake);
  const firstSliceAngle = Number(attr(childByLocalName(chartNode, 'firstSliceAng'), 'val'));
  const series = seriesNodes.map((seriesNode, index) => ({
    name: firstText(descendantByLocalName(seriesNode, 'tx')) || `Series ${index + 1}`,
    type,
    stacking,
    stackGroup,
    values: readNumericValues(descendantByLocalName(seriesNode, 'val')),
    color: readSeriesColorWithTheme(seriesNode, theme),
    lineWidth: readLineWidth(childByLocalName(childByLocalName(seriesNode, 'spPr'), 'ln')),
    pointColors: readPointColors(seriesNode, theme),
    pointStyles: readPointStyles(seriesNode, theme),
    smooth: attr(childByLocalName(seriesNode, 'smooth'), 'val') === '1' || attr(childByLocalName(seriesNode, 'smooth'), 'val') === 'true',
    marker: readSeriesMarker(seriesNode),
  }));

  return {
    type,
    categories,
    series,
    holeSize: type === 'doughnut' ? Number(attr(childByLocalName(chartNode, 'holeSize'), 'val') ?? 0) || undefined : undefined,
    startAngle: Number.isFinite(firstSliceAngle) ? firstSliceAngle : undefined,
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

  return {
    type,
    title: firstText(childByLocalName(chart, 'title')) || undefined,
    categories,
    series,
    showLegend: readLegendVisible(childByLocalName(chart, 'legend')),
    legendPosition: readLegendPosition(childByLocalName(chart, 'legend')),
    legendStyle: readLegendStyle(childByLocalName(chart, 'legend'), theme),
    showDataLabels: readShowDataLabels(plotArea),
    radarIndicators: type === 'radar' ? buildRadarIndicators(categories, series) : undefined,
    holeSize: primaryPlot?.holeSize,
    startAngle: primaryPlot?.startAngle,
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

  if (isRadar && radarIndicators?.length) {
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
        indicator: radarIndicators,
        splitNumber: 4,
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
          data: [
            {
              value: item.values.slice(0, radarIndicators.length),
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
            show: chart.showDataLabels,
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
        };
      }),
    };
  }

  if (isPie && !usesMixedSeriesTypes) {
    const sourceSeries = chart.series[0];
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
          startAngle: chart.startAngle ?? 90,
          padAngle: 0,
          center: ['50%', chart.roseType ? '50%' : chart.title ? '58%' : '50%'],
          avoidLabelOverlap: true,
          label: {
            show: chart.showDataLabels,
            color: '#334155',
            fontFamily: OFFICE_FONT_FAMILY,
          },
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
      label: {
        show: chart.showDataLabels,
        color: '#334155',
        fontFamily: OFFICE_FONT_FAMILY,
      },
      barMaxWidth: isBarSeries ? 32 : undefined,
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
