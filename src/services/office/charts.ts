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
  | 'area'
  | 'scatter'
  | 'bubble'
  | 'unknown';

export type OfficeChartSeries = {
  name: string;
  values: number[];
  color?: string;
  pointColors?: string[];
  pointStyles?: Array<{
    color?: OfficeChartColor;
    borderColor?: string;
    borderWidth?: number;
  }>;
};

export type OfficeChartModel = {
  type: OfficeChartType;
  title?: string;
  categories: string[];
  series: OfficeChartSeries[];
  showLegend?: boolean;
  showDataLabels?: boolean;
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
  doughnutchart: 'pie',
  areachart: 'area',
  scatterchart: 'scatter',
  bubblechart: 'bubble',
};

function decodeMojibake(value: string) {
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

function readCacheValues(node: Element | null | undefined) {
  const cache = descendantByLocalName(node, 'strCache') ?? descendantByLocalName(node, 'numCache');
  return descendantsByLocalName(cache, 'pt')
    .sort((a, b) => Number(attr(a, 'idx') ?? 0) - Number(attr(b, 'idx') ?? 0))
    .map((point) => decodeMojibake(textContent(childByLocalName(point, 'v')).trim()));
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

function readSeriesColor(seriesNode: Element) {
  const solidFill = descendantByLocalName(childByLocalName(seriesNode, 'spPr'), 'solidFill');
  const color = readColorNode(solidFill, DEFAULT_OFFICE_THEME);
  if (!color || typeof color !== 'string') return undefined;
  return color;
}

function readSeriesColorWithTheme(seriesNode: Element, theme: OfficeTheme) {
  const fillNode = childByLocalName(childByLocalName(seriesNode, 'spPr'), 'solidFill') ?? childByLocalName(childByLocalName(seriesNode, 'spPr'), 'gradFill');
  const color = readFillValue(fillNode, theme);
  return color ?? readSeriesColor(seriesNode);
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
    if (color || borderColor || borderWidth !== undefined) {
      styles[index] = {
        color,
        borderColor,
        borderWidth,
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

function findChartNode(plotArea: Element | null) {
  return (
    Array.from(plotArea?.getElementsByTagName('*') ?? []).find((child) =>
      child !== plotArea &&
      (child.localName.split(':').pop() ?? child.localName).toLowerCase().endsWith('chart'),
    ) ?? null
  );
}

export function parseOfficeChartXml(xml: string, theme: OfficeTheme = DEFAULT_OFFICE_THEME): OfficeChartModel {
  const doc = parseXml(xml);
  const chart = descendantByLocalName(doc.documentElement, 'chart');
  const plotArea = descendantByLocalName(chart, 'plotArea');
  const chartNode = findChartNode(plotArea);
  const seriesNodes = descendantsByLocalName(chartNode, 'ser');
  const firstSeries = seriesNodes[0];

  return {
    type: normalizeType(chartNode),
    title: firstText(childByLocalName(chart, 'title')) || undefined,
    categories: readCacheValues(descendantByLocalName(firstSeries, 'cat')).map(decodeMojibake),
    series: seriesNodes.map((seriesNode, index) => ({
      name: firstText(descendantByLocalName(seriesNode, 'tx')) || `Series ${index + 1}`,
      values: readNumericValues(descendantByLocalName(seriesNode, 'val')),
      color: readSeriesColorWithTheme(seriesNode, theme),
      pointColors: readPointColors(seriesNode, theme),
      pointStyles: readPointStyles(seriesNode, theme),
    })),
    showLegend: Boolean(childByLocalName(chart, 'legend')),
    showDataLabels: readShowDataLabels(chartNode),
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

export function buildOfficeChartOption(chart: OfficeChartModel): EChartsOption {
  const categories = resolveCategories(chart);
  const isHorizontalBar = chart.type === 'bar';
  const isPie = chart.type === 'pie';
  const isScatter = chart.type === 'scatter' || chart.type === 'bubble';
  const chartSeriesType =
    chart.type === 'column'
      ? 'bar'
      : chart.type === 'bar'
        ? 'bar'
        : chart.type === 'area'
          ? 'line'
          : isPie
            ? 'pie'
            : isScatter
              ? 'scatter'
              : 'line';

  const palette = chart.series.map((series, index) => resolveSeriesColor(series, index));
  const hasSeries = chart.series.length > 0;

  if (isPie) {
    const sourceSeries = chart.series[0];
    const data = categories.map((name, index) => ({
      name,
      value: sourceSeries?.values[index] ?? 0,
      itemStyle: buildPieItemStyle(sourceSeries, index, palette),
    }));

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
      legend:
        chart.showLegend !== false && chart.series.length > 1
          ? {
              top: 32,
              type: 'scroll',
              itemWidth: 10,
              itemHeight: 10,
              textStyle: OFFICE_TEXT_STYLE,
            }
          : undefined,
      series: [
        {
          type: 'pie',
          radius: '68%',
          padAngle: 1,
          center: ['50%', chart.title ? '58%' : '50%'],
          itemStyle: {
            borderColor: 'transparent',
            borderWidth: 0,
          },
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

  const series = chart.series.map((item, index) => {
    const color = resolveSeriesColor(item, index);
    return {
      name: item.name,
      type: chartSeriesType,
      data: isScatter ? item.values.map((value, valueIndex) => [categories[valueIndex] ?? String(valueIndex + 1), value]) : item.values,
      areaStyle: chart.type === 'area' ? { opacity: 0.18 } : undefined,
      itemStyle: {
        color,
        borderColor: chartSeriesType === 'bar' ? '#fff' : color,
        borderWidth: chartSeriesType === 'bar' ? 1 : 0,
      },
      lineStyle: {
        color,
        width: chart.type === 'area' || chart.type === 'line' ? 2 : 1,
      },
      emphasis: {
        itemStyle: {
          color,
          borderColor: chartSeriesType === 'bar' ? '#fff' : color,
          borderWidth: chartSeriesType === 'bar' ? 1 : 0,
          shadowBlur: chartSeriesType === 'bar' ? 6 : 0,
          shadowColor: 'rgba(15, 23, 42, 0.18)',
        },
        lineStyle: {
          color,
          width: chart.type === 'area' || chart.type === 'line' ? 3 : 1,
        },
      },
      showSymbol: chart.type === 'line' || chart.type === 'area' || chart.type === 'bubble',
      symbolSize: chart.type === 'bubble' ? 14 : 8,
      label: {
        show: chart.showDataLabels,
        color: '#334155',
        fontFamily: OFFICE_FONT_FAMILY,
      },
      barMaxWidth: 32,
    };
  });

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
    legend:
      chart.showLegend !== false && chart.series.length > 1
        ? {
            top: chart.title ? 30 : 8,
            type: 'scroll',
            itemWidth: 10,
            itemHeight: 10,
            textStyle: OFFICE_TEXT_STYLE,
          }
        : undefined,
    grid: {
      left: 40,
      right: 24,
      top: chart.title ? 56 : 24,
      bottom: 32,
      containLabel: true,
    },
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
