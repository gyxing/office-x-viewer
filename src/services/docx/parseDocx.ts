import { loadDocxEntries } from './archive';
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
import { decodeMojibake, parseOfficeChartXml } from '../office/charts';
import { readOfficeTheme, resolveOfficeThemeColor, type OfficeTheme } from '../office/theme';
import type {
  DocxBlock,
  DocxChartBlock,
  DocxDocument,
  DocxImage,
  DocxInline,
  DocxPage,
  DocxParagraphBlock,
  DocxShape,
  DocxShapeItem,
  DocxTableBlock,
  DocxTableCell,
  DocxTextStyle,
} from './types';

type DocxPackageState = {
  entries: OfficeEntryMap;
  relationships: Record<string, Record<string, OfficeRelationship>>;
  mediaByPath: Record<string, string>;
  mediaByName: Record<string, string>;
};

type ParseContext = {
  packageState: DocxPackageState;
  documentRels: Record<string, OfficeRelationship>;
  theme: OfficeTheme;
  styles: DocxStyleCatalog;
  images: DocxImage[];
  imageIndex: number;
  chartIndex: number;
  shapeIndex: number;
};

type DocxStyleDefinition = {
  kind: 'paragraph' | 'character' | 'table';
  basedOn?: string;
  style: DocxTextStyle;
};

type DocxStyleCatalog = {
  defaults: {
    paragraph?: DocxTextStyle;
    run?: DocxTextStyle;
    paragraphStyleId?: string;
    tableStyleId?: string;
  };
  styles: Record<string, DocxStyleDefinition>;
};

const DEFAULT_PAGE: DocxPage = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

const DEFAULT_DOCX_FONT_FAMILY = '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';

function buildPackageState(entries: OfficeEntryMap): DocxPackageState {
  const relationships: DocxPackageState['relationships'] = {};

  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = readRelationships(value, path);
    }
  }

  const media = collectMedia(entries, 'word/media/');

  return {
    entries,
    relationships,
    mediaByPath: media.byPath,
    mediaByName: media.byName,
  };
}

function twipToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return (numberValue / 1440) * 96;
}

function pointToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return numberValue * (96 / 72);
}

function positiveTwipToPx(value?: string | number) {
  const result = twipToPx(value);
  return result !== undefined && result >= 0 ? result : undefined;
}

function eighthPointToPx(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue) || numberValue <= 0) return undefined;
  return (numberValue / 8) * (96 / 72);
}

function pctToRatio(value?: string | number) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return numberValue / 100;
}

function vmlUnitToPx(value?: string | number) {
  const raw = String(value ?? '').trim();
  if (!raw) return undefined;
  const match = raw.match(/^(-?\d+(?:\.\d+)?)(pt|px|in|cm|mm)?$/i);
  if (!match) return undefined;
  const numberValue = Number(match[1]);
  if (!Number.isFinite(numberValue)) return undefined;
  const unit = match[2]?.toLowerCase();
  if (unit === 'px') return numberValue;
  if (unit === 'in') return numberValue * 96;
  if (unit === 'cm') return (numberValue / 2.54) * 96;
  if (unit === 'mm') return (numberValue / 25.4) * 96;
  if (unit === 'pt') return pointToPx(numberValue);
  return emuToPx(numberValue);
}

function readCssDeclaration(style: string | undefined, name: string) {
  if (!style) return undefined;
  const match = style.match(new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i'));
  return match?.[1]?.trim();
}

function readCssSize(style: string | undefined, name: string, scale?: number) {
  const raw = readCssDeclaration(style, name);
  if (!raw) return undefined;
  if (scale !== undefined && /^-?\d+(?:\.\d+)?$/.test(raw)) {
    return Number(raw) * scale;
  }
  return vmlUnitToPx(raw);
}

function readDocxLineHeight(spacingNode: Element | null | undefined) {
  const value = Number(attr(spacingNode, 'w:line') ?? attr(spacingNode, 'line'));
  if (!Number.isFinite(value)) return undefined;
  const rule = attr(spacingNode, 'w:lineRule') ?? attr(spacingNode, 'lineRule');
  if (rule === 'exact' || rule === 'atLeast') {
    return twipToPx(value);
  }
  return value / 240;
}

function halfPointToPx(value?: string) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return (numberValue / 2) * (96 / 72);
}

function parseHexColor(value?: string) {
  if (!value || value === 'auto' || value === 'none') return undefined;
  if (!/^#?[0-9a-f]{6}$/i.test(value)) return undefined;
  return value.startsWith('#') ? value : `#${value}`;
}

function normalizeCssColor(value?: string) {
  if (!value || value === 'auto' || value === 'none') return undefined;
  return value.startsWith('#') ? value : parseHexColor(value);
}

function readVal(node: Element | null | undefined) {
  return attr(node, 'w:val') ?? attr(node, 'val');
}

const WORD_HIGHLIGHT_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  yellow: '#ffff00',
  white: '#ffffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
};

function clamp255(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function tintHexColor(color: string | undefined, tint?: string) {
  if (!color || !tint) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const tintValue = Number.parseInt(tint, 16);
  if (!Number.isFinite(tintValue)) return color;
  const ratio = Math.max(0, Math.min(1, tintValue / 255));
  const rgb = Number.parseInt(normalized, 16);
  const r = (rgb >> 16) & 255;
  const g = (rgb >> 8) & 255;
  const b = rgb & 255;
  return `#${[r, g, b]
    .map((channel) => clamp255(channel + (255 - channel) * ratio).toString(16).padStart(2, '0'))
    .join('')}`;
}

function shadeHexColor(color: string | undefined, shade?: string) {
  if (!color || !shade) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const shadeValue = Number.parseInt(shade, 16);
  if (!Number.isFinite(shadeValue)) return color;
  const ratio = Math.max(0, Math.min(1, shadeValue / 255));
  const rgb = Number.parseInt(normalized, 16);
  const r = (rgb >> 16) & 255;
  const g = (rgb >> 8) & 255;
  const b = rgb & 255;
  return `#${[r, g, b].map((channel) => clamp255(channel * ratio).toString(16).padStart(2, '0')).join('')}`;
}

function resolveThemeFillColor(node: Element | null | undefined, theme: OfficeTheme) {
  const themeFill = attr(node, 'w:themeFill') ?? attr(node, 'themeFill');
  const themeColor = resolveOfficeThemeColor(themeFill, theme);
  return shadeHexColor(
    tintHexColor(themeColor, attr(node, 'w:themeFillTint') ?? attr(node, 'themeFillTint')),
    attr(node, 'w:themeFillShade') ?? attr(node, 'themeFillShade'),
  );
}

function readShading(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  const directFill = parseHexColor(attr(node, 'w:fill') ?? attr(node, 'fill'));
  return directFill ?? resolveThemeFillColor(node, theme);
}

function readHighlight(node: Element | null | undefined) {
  const value = readVal(node);
  if (!value || value === 'none') return undefined;
  return WORD_HIGHLIGHT_COLORS[value] ?? parseHexColor(value);
}

function readBorder(node: Element | null | undefined) {
  const value = readVal(node);
  if (!node || !value || value === 'none' || value === 'nil') return undefined;
  const color = parseHexColor(attr(node, 'w:color') ?? attr(node, 'color')) ?? '#000';
  const width = eighthPointToPx(attr(node, 'w:sz') ?? attr(node, 'sz')) ?? 1;
  const style = value === 'dashed' || value === 'dashSmallGap' ? 'dashed' : value === 'dotted' ? 'dotted' : 'solid';
  return `${width}px ${style} ${color}`;
}

function readParagraphBorders(pPr: Element | null | undefined) {
  const pBdr = childByLocalName(pPr, 'pBdr');
  return {
    borderTop: readBorder(childByLocalName(pBdr, 'top')),
    borderRight: readBorder(childByLocalName(pBdr, 'right')),
    borderBottom: readBorder(childByLocalName(pBdr, 'bottom')),
    borderLeft: readBorder(childByLocalName(pBdr, 'left')),
    paddingTop: pointToPx(attr(childByLocalName(pBdr, 'top'), 'w:space') ?? attr(childByLocalName(pBdr, 'top'), 'space')),
    paddingRight: pointToPx(attr(childByLocalName(pBdr, 'right'), 'w:space') ?? attr(childByLocalName(pBdr, 'right'), 'space')),
    paddingBottom: pointToPx(attr(childByLocalName(pBdr, 'bottom'), 'w:space') ?? attr(childByLocalName(pBdr, 'bottom'), 'space')),
    paddingLeft: pointToPx(attr(childByLocalName(pBdr, 'left'), 'w:space') ?? attr(childByLocalName(pBdr, 'left'), 'space')),
  };
}

function readParagraphPropertyStyle(
  pPr: Element | null | undefined,
  theme: OfficeTheme,
): DocxTextStyle | undefined {
  if (!pPr) return undefined;
  const spacing = childByLocalName(pPr, 'spacing');
  const ind = childByLocalName(pPr, 'ind');
  const style: DocxTextStyle = {
    align: mapAlignment(readVal(childByLocalName(pPr, 'jc'))),
    spacingBefore: positiveTwipToPx(attr(spacing, 'w:before') ?? attr(spacing, 'before')),
    spacingAfter: positiveTwipToPx(attr(spacing, 'w:after') ?? attr(spacing, 'after')),
    indentLeft: twipToPx(attr(ind, 'w:left') ?? attr(ind, 'left')),
    indentRight: twipToPx(attr(ind, 'w:right') ?? attr(ind, 'right')),
    firstLineIndent: twipToPx(attr(ind, 'w:firstLine') ?? attr(ind, 'firstLine')),
    lineHeight: readDocxLineHeight(spacing),
    backgroundColor: readShading(childByLocalName(pPr, 'shd'), theme),
    ...readParagraphBorders(pPr),
  };
  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function readOnOff(node: Element | null | undefined) {
  if (!node) return undefined;
  const value = attr(node, 'w:val') ?? attr(node, 'val');
  if (value === undefined) return true;
  return value !== '0' && value !== 'false' && value !== 'off';
}

function firstDefined<T>(...values: Array<T | undefined>) {
  return values.find((value) => value !== undefined);
}

function readThemeFont(rFonts: Element | null | undefined, theme: OfficeTheme) {
  const themeFont =
    attr(rFonts, 'w:eastAsiaTheme') ??
    attr(rFonts, 'eastAsiaTheme') ??
    attr(rFonts, 'w:asciiTheme') ??
    attr(rFonts, 'asciiTheme') ??
    attr(rFonts, 'w:hAnsiTheme') ??
    attr(rFonts, 'hAnsiTheme') ??
    attr(rFonts, 'w:cstheme') ??
    attr(rFonts, 'cstheme');
  if (!themeFont) return undefined;
  return themeFont.toLowerCase().includes('major') ? theme.fontScheme?.majorFont : theme.fontScheme?.minorFont;
}

function quoteFontFamily(value?: string) {
  if (!value) return undefined;
  return value
    .split(',')
    .map((font) => font.trim())
    .filter(Boolean)
    .map((font) => (/^["'].*["']$/.test(font) || /^[a-z-]+$/i.test(font) ? font : `"${font}"`))
    .join(', ');
}

function readFontFamily(rPr: Element | null | undefined, theme: OfficeTheme, allowFallback = false) {
  const rFonts = childByLocalName(rPr, 'rFonts');
  const ascii = attr(rFonts, 'w:ascii') ?? attr(rFonts, 'ascii');
  const eastAsia = attr(rFonts, 'w:eastAsia') ?? attr(rFonts, 'eastAsia');
  const hAnsi = attr(rFonts, 'w:hAnsi') ?? attr(rFonts, 'hAnsi');
  const cs = attr(rFonts, 'w:cs') ?? attr(rFonts, 'cs');
  const themeFonts = theme.fontScheme ?? {};
  const explicitFont = eastAsia ?? ascii ?? hAnsi ?? cs ?? readThemeFont(rFonts, theme);
  if (explicitFont || !allowFallback) return quoteFontFamily(explicitFont);
  return quoteFontFamily(
    themeFonts.minorFont ??
    themeFonts.majorFont ??
    DEFAULT_DOCX_FONT_FAMILY,
  );
}

function readDocxStyles(entries: OfficeEntryMap, theme: OfficeTheme): DocxStyleCatalog {
  const xml = readXml(entries, 'word/styles.xml');
  if (!xml) return { defaults: {}, styles: {} };

  const doc = parseXml(xml);
  const root = doc.documentElement;
  const styles: Record<string, DocxStyleDefinition> = {};
  const defaults: DocxStyleCatalog['defaults'] = {};

  const docDefaults = childByLocalName(root, 'docDefaults');
  const rPrDefault = childByLocalName(childByLocalName(docDefaults, 'rPrDefault'), 'rPr');
  const pPrDefault = childByLocalName(childByLocalName(docDefaults, 'pPrDefault'), 'pPr');
  defaults.run = readTextStyle(rPrDefault, theme, true);
  defaults.paragraph = mergeTextStyle(
    readParagraphPropertyStyle(pPrDefault, theme),
    readTextStyle(childByLocalName(pPrDefault, 'rPr'), theme, true),
  );

  childrenByLocalName(root, 'style').forEach((styleNode) => {
    const styleId = attr(styleNode, 'styleId');
    const kindAttr = attr(styleNode, 'type');
    if (!styleId) return;
    if (kindAttr === 'paragraph' && attr(styleNode, 'w:default') === '1') defaults.paragraphStyleId = styleId;
    if (kindAttr === 'table' && attr(styleNode, 'w:default') === '1') defaults.tableStyleId = styleId;

    const basedOn = attr(childByLocalName(styleNode, 'basedOn'), 'w:val') ?? attr(childByLocalName(styleNode, 'basedOn'), 'val') ?? undefined;
    const name = styleId;
    const pPr = childByLocalName(styleNode, 'pPr');
    const rPr = childByLocalName(styleNode, 'rPr');

    let style: DocxTextStyle | undefined;
    if (kindAttr === 'paragraph') {
      style = mergeTextStyle(readParagraphPropertyStyle(pPr, theme), readTextStyle(rPr, theme));
    } else if (kindAttr === 'table') {
      const tblPr = childByLocalName(styleNode, 'tblPr');
      style = readParagraphPropertyStyle(childByLocalName(tblPr, 'pPr'), theme);
    } else {
      style = readTextStyle(rPr, theme);
    }

    if (style) {
      styles[name] = {
        kind: kindAttr === 'paragraph' ? 'paragraph' : kindAttr === 'table' ? 'table' : 'character',
        basedOn,
        style,
      };
    }
  });

  return { defaults, styles };
}

function readUnderline(rPr: Element | null | undefined) {
  const underline = childByLocalName(rPr, 'u');
  if (!underline) return undefined;
  const value = attr(underline, 'w:val') ?? attr(underline, 'val');
  return value !== 'none' && value !== '0' && value !== 'false';
}

function readTextStyle(
  rPr: Element | null | undefined,
  theme: OfficeTheme,
  allowFontFallback = false,
): DocxTextStyle | undefined {
  if (!rPr) return undefined;

  const color = parseHexColor(attr(childByLocalName(rPr, 'color'), 'w:val') ?? attr(childByLocalName(rPr, 'color'), 'val'));
  const style: DocxTextStyle = {
    bold: firstDefined(readOnOff(childByLocalName(rPr, 'b')), readOnOff(childByLocalName(rPr, 'bCs'))),
    italic: firstDefined(readOnOff(childByLocalName(rPr, 'i')), readOnOff(childByLocalName(rPr, 'iCs'))),
    underline: readUnderline(rPr),
    strike: firstDefined(readOnOff(childByLocalName(rPr, 'strike')), readOnOff(childByLocalName(rPr, 'dstrike'))),
    smallCaps: readOnOff(childByLocalName(rPr, 'smallCaps')),
    allCaps: readOnOff(childByLocalName(rPr, 'caps')),
    color,
    backgroundColor: readHighlight(childByLocalName(rPr, 'highlight')) ?? readShading(childByLocalName(rPr, 'shd'), theme),
    fontSize: halfPointToPx(attr(childByLocalName(rPr, 'sz'), 'w:val') ?? attr(childByLocalName(rPr, 'sz'), 'val')),
    fontFamily: readFontFamily(rPr, theme, allowFontFallback),
  };

  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function mergeTwoTextStyles(base?: DocxTextStyle, next?: DocxTextStyle): DocxTextStyle {
  return {
    ...base,
    ...next,
    fontSize: next?.fontSize ?? base?.fontSize,
    fontFamily: next?.fontFamily ?? base?.fontFamily,
    color: next?.color ?? base?.color,
    lineHeight: next?.lineHeight ?? base?.lineHeight,
    spacingBefore: next?.spacingBefore ?? base?.spacingBefore,
    spacingAfter: next?.spacingAfter ?? base?.spacingAfter,
    indentLeft: next?.indentLeft ?? base?.indentLeft,
    indentRight: next?.indentRight ?? base?.indentRight,
    firstLineIndent: next?.firstLineIndent ?? base?.firstLineIndent,
    backgroundColor: next?.backgroundColor ?? base?.backgroundColor,
    borderTop: next?.borderTop ?? base?.borderTop,
    borderRight: next?.borderRight ?? base?.borderRight,
    borderBottom: next?.borderBottom ?? base?.borderBottom,
    borderLeft: next?.borderLeft ?? base?.borderLeft,
    paddingTop: next?.paddingTop ?? base?.paddingTop,
    paddingRight: next?.paddingRight ?? base?.paddingRight,
    paddingBottom: next?.paddingBottom ?? base?.paddingBottom,
    paddingLeft: next?.paddingLeft ?? base?.paddingLeft,
    align: next?.align ?? base?.align,
  };
}

function mergeTextStyle(...styles: Array<DocxTextStyle | undefined>): DocxTextStyle | undefined {
  const merged = styles.reduce<DocxTextStyle>((acc, style) => mergeTwoTextStyles(acc, style), {});
  return Object.keys(merged).length ? merged : undefined;
}

function resolveDocxStyle(
  styleId: string | undefined,
  catalog: DocxStyleCatalog,
  seen: Set<string> = new Set(),
): DocxTextStyle | undefined {
  if (!styleId || seen.has(styleId)) return undefined;
  const entry = catalog.styles[styleId];
  if (!entry) return undefined;
  seen.add(styleId);
  const base = resolveDocxStyle(entry.basedOn, catalog, seen);
  return mergeTextStyle(base, entry.style);
}

function resolveParagraphStyle(
  pPr: Element | null | undefined,
  catalog: DocxStyleCatalog,
  theme: OfficeTheme,
) {
  const styleId = attr(childByLocalName(pPr, 'pStyle'), 'w:val') ?? attr(childByLocalName(pPr, 'pStyle'), 'val');
  const baseStyle = resolveDocxStyle(catalog.defaults.paragraphStyleId, catalog);
  const namedStyle = resolveDocxStyle(styleId, catalog);
  const style = mergeTextStyle(catalog.defaults.paragraph, baseStyle, namedStyle);
  const directStyle = readParagraphPropertyStyle(pPr, theme);
  return {
    align: directStyle?.align ?? style?.align,
    spacingBefore: directStyle?.spacingBefore ?? style?.spacingBefore,
    spacingAfter: directStyle?.spacingAfter ?? style?.spacingAfter,
    indentLeft: directStyle?.indentLeft ?? style?.indentLeft,
    indentRight: directStyle?.indentRight ?? style?.indentRight,
    firstLineIndent: directStyle?.firstLineIndent ?? style?.firstLineIndent,
    lineHeight: directStyle?.lineHeight ?? style?.lineHeight,
    backgroundColor: directStyle?.backgroundColor ?? style?.backgroundColor,
    borderTop: directStyle?.borderTop ?? style?.borderTop,
    borderRight: directStyle?.borderRight ?? style?.borderRight,
    borderBottom: directStyle?.borderBottom ?? style?.borderBottom,
    borderLeft: directStyle?.borderLeft ?? style?.borderLeft,
    paddingTop: directStyle?.paddingTop ?? style?.paddingTop,
    paddingRight: directStyle?.paddingRight ?? style?.paddingRight,
    paddingBottom: directStyle?.paddingBottom ?? style?.paddingBottom,
    paddingLeft: directStyle?.paddingLeft ?? style?.paddingLeft,
    style: mergeTextStyle(style, directStyle, readTextStyle(childByLocalName(pPr, 'rPr'), theme)),
  };
}

function resolveRunStyle(
  rPr: Element | null | undefined,
  catalog: DocxStyleCatalog,
  theme: OfficeTheme,
) {
  const styleId = attr(childByLocalName(rPr, 'rStyle'), 'w:val') ?? attr(childByLocalName(rPr, 'rStyle'), 'val');
  return mergeTextStyle(catalog.defaults.run, resolveDocxStyle(styleId, catalog), readTextStyle(rPr, theme));
}

function inlineInheritedStyle(style: DocxTextStyle | undefined): DocxTextStyle | undefined {
  if (!style) return undefined;
  const {
    backgroundColor,
    borderTop,
    borderRight,
    borderBottom,
    borderLeft,
    paddingTop,
    paddingRight,
    paddingBottom,
    paddingLeft,
    spacingBefore,
    spacingAfter,
    indentLeft,
    indentRight,
    firstLineIndent,
    lineHeight,
    align,
    ...inlineStyle
  } = style;
  return Object.keys(inlineStyle).length ? inlineStyle : undefined;
}

function mapAlignment(value?: string): DocxTextStyle['align'] | undefined {
  if (value === 'left' || value === 'center' || value === 'right' || value === 'justify') return value;
  if (value === 'both') return 'justify';
  return undefined;
}

function resolveMediaRef(target: string | undefined, packageState: DocxPackageState) {
  return resolvePackageMediaRef(target, packageState.mediaByPath, packageState.mediaByName, 'word');
}

function resolveXmlTarget(target: string | undefined, packageState: DocxPackageState) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function parseChartElement(node: Element, context: ParseContext): DocxChartBlock | undefined {
  const chartNode = descendantByLocalName(node, 'chart');
  const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const chartPath = resolveXmlTarget(target, context.packageState);
  const xml = chartPath ? (context.packageState.entries.get(chartPath) as string | undefined) : undefined;
  if (!xml) return undefined;

  const chart = parseOfficeChartXml(xml, context.theme);
  const extent = descendantByLocalName(node, 'extent') ?? descendantByLocalName(node, 'xfrm');
  const width = Math.max(160, Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320));
  const height = Math.max(120, Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220));
  context.chartIndex += 1;
  return {
    id: `docx-chart-${context.chartIndex}`,
    type: 'chart',
    chart,
    width,
    height,
  };
}

function readWebExtensionProperties(webExtensionNode: Element) {
  const properties: Record<string, string> = {};
  descendantsByLocalName(webExtensionNode, 'property').forEach((propertyNode) => {
    const key = attr(propertyNode, 'key');
    const value = attr(propertyNode, 'value');
    if (key && value !== undefined) {
      properties[key] = value;
    }
  });
  return properties;
}

function decodeMojibakeDeep<T>(value: T): T {
  if (typeof value === 'string') {
    return decodeMojibake(value) as T;
  }
  if (Array.isArray(value)) {
    return value.map((item) => decodeMojibakeDeep(item)) as T;
  }
  if (value && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, decodeMojibakeDeep(item)]),
    ) as T;
  }
  return value;
}

function parseJsonProperty<T>(properties: Record<string, string>, key: string): T | undefined {
  const raw = properties[key];
  if (!raw) return undefined;
  try {
    return decodeMojibakeDeep(JSON.parse(raw)) as T;
  } catch {
    try {
      return decodeMojibakeDeep(JSON.parse(raw.replace(/[?�]quot;/g, '"'))) as T;
    } catch {
      return undefined;
    }
  }
}

function normalizeLegendPosition(value: unknown): DocxChartBlock['chart']['legendPosition'] | undefined {
  if (typeof value !== 'string') return undefined;
  const lower = value.toLowerCase();
  if (lower.includes('bottom')) return 'bottom';
  if (lower.includes('top')) return 'top';
  if (lower.includes('left')) return 'left';
  if (lower.includes('right')) return 'right';
  return undefined;
}

function normalizeChartColor(value: unknown) {
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object') {
    const color = (value as { color?: unknown; rgb?: unknown }).color ?? (value as { rgb?: unknown }).rgb;
    if (typeof color === 'string') return color;
  }
  return undefined;
}

function collectChartColors(style: Record<string, unknown> | undefined, fallbackKey: 'seriesThemeColor' | 'fill') {
  if (!style) return [];
  if (fallbackKey === 'seriesThemeColor' && Array.isArray(style.seriesThemeColor)) {
    return style.seriesThemeColor.map(normalizeChartColor).filter((color): color is string => Boolean(color));
  }
  const fill = style.fill as { props?: unknown[] } | undefined;
  if (fallbackKey === 'fill' && Array.isArray(fill?.props)) {
    return fill.props
      .map((item) => normalizeChartColor((item as { color?: unknown } | undefined)?.color))
      .filter((color): color is string => Boolean(color));
  }
  return [];
}

function readWpsLegendStyle(legend: unknown): DocxChartBlock['chart']['legendStyle'] | undefined {
  if (!legend || typeof legend !== 'object') return undefined;
  const legendObject = legend as Record<string, unknown>;
  const textStyle = legendObject.textStyle && typeof legendObject.textStyle === 'object'
    ? (legendObject.textStyle as Record<string, unknown>)
    : legendObject;
  const fontFamily = textStyle.fontFamily;
  const fontSize = Number(textStyle.fontSize);
  const fontStyle = textStyle.fontStyle;
  const fontWeight = textStyle.fontWeight;
  const color = normalizeChartColor(textStyle.color);
  const itemWidth = Number(legendObject.itemWidth);
  const itemHeight = Number(legendObject.itemHeight);
  const style = {
    itemWidth: Number.isFinite(itemWidth) && itemWidth > 0 ? itemWidth : undefined,
    itemHeight: Number.isFinite(itemHeight) && itemHeight > 0 ? itemHeight : undefined,
    textStyle: {
      color,
      fontFamily: typeof fontFamily === 'string' ? fontFamily : undefined,
      fontSize: Number.isFinite(fontSize) && fontSize > 0 ? fontSize : undefined,
      fontStyle: typeof fontStyle === 'string' ? fontStyle : undefined,
      fontWeight: typeof fontWeight === 'string' || typeof fontWeight === 'number' ? fontWeight : undefined,
    },
  };
  const normalizedTextStyle = Object.fromEntries(Object.entries(style.textStyle).filter(([, value]) => value !== undefined));
  return {
    ...(style.itemWidth !== undefined ? { itemWidth: style.itemWidth } : {}),
    ...(style.itemHeight !== undefined ? { itemHeight: style.itemHeight } : {}),
    ...(Object.keys(normalizedTextStyle).length ? { textStyle: normalizedTextStyle } : {}),
  };
}

function readPercent(value: unknown) {
  if (typeof value !== 'string') return undefined;
  const parsed = Number(value.replace(/%$/, ''));
  return Number.isFinite(parsed) ? parsed : undefined;
}

function readRadiusPair(value: unknown): [string, string] | undefined {
  if (!Array.isArray(value) || typeof value[0] !== 'string' || typeof value[1] !== 'string') return undefined;
  return [value[0], value[1]];
}

function readWpsSeriesStyle(style: Record<string, unknown> | undefined) {
  return Array.isArray(style?.series) && style.series[0] && typeof style.series[0] === 'object'
    ? (style.series[0] as Record<string, unknown>)
    : undefined;
}

function readWpsPiePointStyles(seriesStyle: Record<string, unknown> | undefined, count: number) {
  const itemStyle = seriesStyle?.itemStyle;
  if (!itemStyle || typeof itemStyle !== 'object') return undefined;
  const borderColor = normalizeChartColor((itemStyle as { borderColor?: unknown }).borderColor);
  const borderWidth = Number((itemStyle as { borderWidth?: unknown }).borderWidth);
  if (!borderColor && !Number.isFinite(borderWidth)) return undefined;
  return Array.from({ length: count }, () => ({
    borderColor,
    borderWidth: Number.isFinite(borderWidth) ? borderWidth : undefined,
  }));
}

function resolveWebExtensionSnapshot(doc: XMLDocument, webExtensionPath: string, context: ParseContext) {
  const snapshot = descendantByLocalName(doc.documentElement, 'snapshot');
  const embed = attr(snapshot, 'r:embed') ?? attr(snapshot, 'embed');
  const relsPath = webExtensionPath.replace(/^word\/webExtensions\//, 'word/webExtensions/_rels/').concat('.rels');
  const target = embed ? context.packageState.relationships[relsPath]?.[embed]?.target : undefined;
  return resolveMediaRef(target, context.packageState);
}

function parseWpsWebExtensionChart(node: Element, context: ParseContext): DocxChartBlock | undefined {
  const webExtensionRef = descendantByLocalName(node, 'webExtensionRef');
  const relId = attr(webExtensionRef, 'r:id') ?? attr(webExtensionRef, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const webExtensionPath = resolveXmlTarget(target, context.packageState);
  const xml = webExtensionPath ? (context.packageState.entries.get(webExtensionPath) as string | undefined) : undefined;
  if (!xml || !webExtensionPath) return undefined;

  const doc = parseXml(xml);
  const snapshotSrc = resolveWebExtensionSnapshot(doc, webExtensionPath, context);
  const properties = readWebExtensionProperties(doc.documentElement);
  const demoData = parseJsonProperty<Record<string, unknown>>(properties, 'demoData');
  const style = parseJsonProperty<Record<string, unknown>>(properties, 'style');
  const extStyle = parseJsonProperty<Record<string, unknown>>(properties, 'extStyle');
  const dschart = parseJsonProperty<Record<string, unknown>>(properties, 'dschart');
  const mapData = dschart && typeof dschart === 'object' ? (dschart as { json?: { data?: unknown[]; props?: Record<string, unknown> } }).json : undefined;
  const chartStyle = style ?? mapData?.props;
  const title = chartStyle?.title && typeof chartStyle.title === 'object' ? (chartStyle.title as { text?: unknown; show?: unknown }) : undefined;
  const showLegend = chartStyle?.legend && typeof chartStyle.legend === 'object' ? (chartStyle.legend as { show?: unknown }).show : undefined;
  const legendPosition = chartStyle?.legend && typeof chartStyle.legend === 'object'
    ? normalizeLegendPosition((chartStyle.legend as { position?: unknown }).position)
    : undefined;
  const legendStyle = readWpsLegendStyle(chartStyle?.legend);
  const labelStyle = chartStyle?.label && typeof chartStyle.label === 'object' ? (chartStyle.label as Record<string, unknown>) : undefined;
  const textLabelStyle = chartStyle?.label && typeof chartStyle.label === 'object' && chartStyle.label.textLabel && typeof chartStyle.label.textLabel === 'object'
    ? (chartStyle.label.textLabel as Record<string, unknown>)
    : undefined;
  const numberLabelStyle = chartStyle?.label && typeof chartStyle.label === 'object' && chartStyle.label.numberLabel && typeof chartStyle.label.numberLabel === 'object'
    ? (chartStyle.label.numberLabel as Record<string, unknown>)
    : undefined;
  const labelPosition = typeof labelStyle?.position === 'string'
    ? labelStyle.position
    : typeof textLabelStyle?.position === 'string'
      ? textLabelStyle.position
      : typeof numberLabelStyle?.position === 'string'
        ? numberLabelStyle.position
        : undefined;
  const labelSeparator = typeof labelStyle?.separator === 'string'
    ? labelStyle.separator
    : typeof textLabelStyle?.separator === 'string'
      ? textLabelStyle.separator
      : typeof numberLabelStyle?.separator === 'string'
        ? numberLabelStyle.separator
        : undefined;
  const labelShowVal = Boolean(labelStyle?.show ?? numberLabelStyle?.show);
  const labelShowCatName = Boolean(labelStyle?.showCategoryName ?? textLabelStyle?.show ?? numberLabelStyle?.showCatName);
  const labelShowSerName = Boolean(labelStyle?.showSeriesName ?? textLabelStyle?.showSerName ?? numberLabelStyle?.showSerName);
  const labelShowPercent = Boolean(labelStyle?.showPercent ?? numberLabelStyle?.showPercent);
  const labelShowLeaderLines = Boolean(labelStyle?.showLeaderLines ?? numberLabelStyle?.showLeaderLines);
  const showDataLabels = Boolean(
    (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { show?: unknown }).show) ||
      (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { numberLabel?: { show?: unknown }; textLabel?: { show?: unknown } }).numberLabel?.show) ||
      (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { numberLabel?: { show?: unknown }; textLabel?: { show?: unknown } }).textLabel?.show),
  );
  const titleText =
    typeof title?.show === 'boolean' && title.show !== false && typeof title?.text === 'string'
      ? decodeMojibake(title.text)
      : undefined;
  const extent = descendantByLocalName(node, 'extent') ?? descendantByLocalName(node, 'xfrm');
  const width = Math.max(160, Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320));
  const height = Math.max(120, Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220));

  if (demoData && Array.isArray(demoData.data) && Array.isArray(demoData.data[0])) {
    const rows = demoData.data.slice(1).filter((row): row is unknown[] => Array.isArray(row));
    const headers = demoData.data[0] as unknown[];

    const isPie = String(properties.type ?? '').toLowerCase().includes('pie');
    const pieType = typeof style?.pieType === 'string' ? style.pieType.toLowerCase() : '';
    const radius = readRadiusPair(style?.radius);
    const seriesStyle = readWpsSeriesStyle(extStyle) ?? readWpsSeriesStyle(style);
    const roseType =
      seriesStyle?.roseType === 'radius' || seriesStyle?.roseType === 'area'
        ? seriesStyle.roseType
        : undefined;
    const categories = rows.map((row) => decodeMojibake(String(row[0] ?? '').trim())).filter(Boolean);
    const seriesNames = headers.slice(1).map((header, index) =>
      decodeMojibake(String(header ?? `Series ${index + 1}`).trim()),
    );
    const palette = collectChartColors(chartStyle, 'seriesThemeColor');
    const chartType = isPie && (pieType.includes('doughnut') || radius || roseType) ? 'doughnut' : isPie ? 'pie' : 'line';
    const isPieChart = chartType === 'pie' || chartType === 'doughnut';
    const piePointStyles = chartType === 'doughnut' || chartType === 'pie'
      ? readWpsPiePointStyles(seriesStyle, categories.length)
      : undefined;
    const sourceSeries = seriesNames.length
      ? seriesNames.map((name, index) => ({
          name,
          values: rows.map((row) => Number(row[index + 1] ?? 0) || 0),
          type: isPieChart ? ('pie' as const) : ((style?.areaStyle && typeof style.areaStyle === 'object' && (style.areaStyle as { show?: unknown }).show) ? ('area' as const) : ('line' as const)),
          color: palette[index],
          smooth:
            Boolean(
              (style?.smooth as boolean | undefined) ??
                (Array.isArray(style?.series)
                  ? (style.series as Array<{ smooth?: unknown }>)[0]?.smooth
                  : undefined),
            ),
          marker:
            style?.symbol && Array.isArray(style.symbol) && (style.symbol[0] as { show?: unknown; size?: unknown })?.show === false
              ? { symbol: 'none' as const, size: Number((style.symbol[0] as { size?: unknown })?.size) || 6 }
              : { size: Number((style?.series && Array.isArray(style.series) ? (style.series[0] as { symbolSize?: unknown })?.symbolSize : undefined)) || 6 },
        }))
      : [];

    context.chartIndex += 1;
    return {
      id: `docx-chart-${context.chartIndex}`,
      type: 'chart',
      chart: {
        type: chartType,
        title: titleText,
        categories,
        series: sourceSeries.length
          ? sourceSeries.map((series, index) =>
              isPieChart
                ? {
                    ...series,
                    pointColors: palette.length ? palette : undefined,
                    pointStyles: piePointStyles,
                  }
                : series,
            )
              : [
              {
                name: 'Series 1',
                values: rows.map((row) => Number(row[1] ?? 0) || 0),
                type: isPieChart ? ('pie' as const) : ('line' as const),
                pointColors: palette.length ? palette : undefined,
                pointStyles: piePointStyles,
              },
            ],
        showLegend: showLegend !== false,
        legendPosition,
        legendStyle,
        showDataLabels,
        dataLabels: {
          position: labelPosition,
          separator: labelSeparator,
          showVal: labelShowVal,
          showCatName: labelShowCatName,
          showSerName: labelShowSerName,
          showPercent: labelShowPercent,
          showLeaderLines: labelShowLeaderLines,
        },
        holeSize:
          chartType === 'doughnut'
            ? (() => {
                const parsed = readPercent(radius?.[0]);
                return Number.isFinite(parsed) ? parsed : undefined;
              })()
            : undefined,
        radius: roseType ? radius : undefined,
        roseType,
        startAngle:
          isPieChart && Number.isFinite(Number(style?.startAngle))
            ? Number(style?.startAngle)
            : isPieChart
              ? 0
              : undefined,
      },
      width,
      height,
    };
  }

  if (dschart && typeof dschart === 'object' && typeof properties.dschart_type === 'string' && properties.dschart_type.toLowerCase() === 'map') {
    const chartJson = mapData && typeof mapData === 'object' ? mapData : undefined;
    const table = Array.isArray(chartJson?.data) && Array.isArray(chartJson.data[0]) ? chartJson.data[0] : undefined;
    if (!table || table.length < 2) return undefined;

    const rows = table.slice(1).filter((row): row is unknown[] => Array.isArray(row));
    const categories = rows.map((row) => decodeMojibake(String(row[0] ?? '').trim())).filter(Boolean);
    const valueIndex = 1;
    const header = table[0];
    const tierIndex = Array.isArray(header) && header.length > 2 ? 2 : undefined;
    const seriesName = Array.isArray(header) && typeof header[valueIndex] === 'string'
      ? decodeMojibake(header[valueIndex])
      : 'Series 1';
    const colors = collectChartColors(chartStyle, 'fill');
    const tiers = tierIndex !== undefined
      ? rows.map((row) => decodeMojibake(String(row[tierIndex] ?? '').trim()))
      : [];
    const tierNames = Array.from(new Set(tiers.filter(Boolean)));
    const pointColors = tierNames.length
      ? tiers.map((tier) => colors[tierNames.indexOf(tier)]).filter((color): color is string => Boolean(color))
      : colors;

    context.chartIndex += 1;
    return {
      id: `docx-chart-${context.chartIndex}`,
      type: 'chart',
      chart: {
        type: 'map',
        title: titleText,
        categories,
        series: [
          {
            name: seriesName,
            values: rows.map((row) => Number(row[valueIndex] ?? 0) || 0),
            type: 'map' as const,
            pointColors: pointColors.length ? pointColors : undefined,
            pointLabels: tiers.length ? tiers : undefined,
          },
        ],
        showLegend: showLegend !== false,
        legendPosition,
        legendStyle,
        showDataLabels,
        mapSeriesName: seriesName,
        mapName: 'china',
        mapGeoJsonUrl: 'https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json',
        mapRegion:
          chartStyle?.mapRegion && typeof chartStyle.mapRegion === 'object'
            ? decodeMojibake(
                String(
                  (chartStyle.mapRegion as { country?: unknown; province?: unknown; city?: unknown }).city ||
                    (chartStyle.mapRegion as { country?: unknown; province?: unknown; city?: unknown }).province ||
                    (chartStyle.mapRegion as { country?: unknown; province?: unknown; city?: unknown }).country ||
                    '',
                ),
              )
            : undefined,
        snapshotSrc,
      },
      width,
      height,
    };
  }

  return undefined;
}

function parseDrawingImage(drawingNode: Element, context: ParseContext): DocxImage | undefined {
  const blip = descendantByLocalName(drawingNode, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  const src = resolveMediaRef(target, context.packageState);
  if (!src) return undefined;

  const extent = descendantByLocalName(drawingNode, 'extent');
  const docPr = descendantByLocalName(drawingNode, 'docPr');
  const width = Math.max(1, Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0))));
  const height = Math.max(1, Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0))));
  const name = attr(docPr, 'name');
  const image: DocxImage = {
    id: `docx-image-${context.imageIndex + 1}`,
    name,
    alt: attr(docPr, 'descr') ?? name,
    src,
    width,
    height,
  };
  context.imageIndex += 1;
  context.images.push(image);
  return image;
}

function readDrawingColor(node: Element | null | undefined, theme: OfficeTheme) {
  if (!node) return undefined;
  const solidFill = childByLocalName(node, 'solidFill') ?? (matchesLocalName(node, 'solidFill') ? node : null);
  if (!solidFill) return undefined;
  const srgb = childByLocalName(solidFill, 'srgbClr');
  const scheme = childByLocalName(solidFill, 'schemeClr');
  const sys = childByLocalName(solidFill, 'sysClr');
  return (
    parseHexColor(attr(srgb, 'val')) ??
    resolveOfficeThemeColor(attr(scheme, 'val'), theme) ??
    parseHexColor(attr(sys, 'lastClr') ?? attr(sys, 'val'))
  );
}

function readDrawingNoFill(node: Element | null | undefined) {
  return Boolean(childByLocalName(node, 'noFill'));
}

function parseDrawingLineStyle(spPr: Element | null | undefined, theme: OfficeTheme) {
  const line = childByLocalName(spPr, 'ln');
  if (!line || readDrawingNoFill(line)) return {};
  const width = emuToPx(Number(attr(line, 'w') ?? 0)) || 1;
  const color = readDrawingColor(line, theme) ?? '#000';
  const dash = attr(childByLocalName(line, 'prstDash'), 'val');
  const strokeDasharray = dash && dash !== 'solid' ? `${width * 3} ${width}` : undefined;
  return {
    border: `${width}px solid ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray,
  };
}

function parseDrawingFillColor(spPr: Element | null | undefined, theme: OfficeTheme) {
  if (readDrawingNoFill(spPr)) return undefined;
  return readDrawingColor(spPr, theme);
}

function parseDrawingXfrm(node: Element | null | undefined, scale?: { x?: number; y?: number }) {
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const rawLeft = Number(attr(off, 'x') ?? 0);
  const rawTop = Number(attr(off, 'y') ?? 0);
  const rawWidth = Number(attr(ext, 'cx') ?? 0);
  const rawHeight = Number(attr(ext, 'cy') ?? 0);
  return {
    left: Number.isFinite(rawLeft) ? rawLeft * (scale?.x ?? 1) : 0,
    top: Number.isFinite(rawTop) ? rawTop * (scale?.y ?? 1) : 0,
    width: Number.isFinite(rawWidth) ? rawWidth * (scale?.x ?? 1) : 0,
    height: Number.isFinite(rawHeight) ? rawHeight * (scale?.y ?? 1) : 0,
  };
}

function readWpgScale(groupNode: Element, width: number, height: number) {
  const xfrm = descendantByLocalName(childByLocalName(groupNode, 'grpSpPr'), 'xfrm');
  const chExt = childByLocalName(xfrm, 'chExt');
  const rawWidth = Number(attr(chExt, 'cx') ?? 0);
  const rawHeight = Number(attr(chExt, 'cy') ?? 0);
  return {
    x: Number.isFinite(rawWidth) && rawWidth > 0 ? width / rawWidth : 1,
    y: Number.isFinite(rawHeight) && rawHeight > 0 ? height / rawHeight : 1,
  };
}

function readDrawingShapeKind(spPr: Element | null | undefined): DocxShapeItem['kind'] {
  const geometry = childByLocalName(spPr, 'prstGeom');
  const preset = attr(geometry, 'prst');
  if (preset === 'line') return 'line';
  if (preset === 'ellipse') return 'ellipse';
  if (childByLocalName(spPr, 'custGeom')) return 'path';
  return 'rect';
}

function readDrawingTextAnchor(shapeNode: Element): DocxShapeItem['textVerticalAlign'] {
  const anchor = attr(childByLocalName(shapeNode, 'bodyPr'), 'anchor');
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b') return 'bottom';
  return 'top';
}

function readDrawingBodyPadding(shapeNode: Element) {
  const bodyPr = childByLocalName(shapeNode, 'bodyPr');
  return {
    paddingTop: vmlUnitToPx(attr(bodyPr, 'tIns')),
    paddingRight: vmlUnitToPx(attr(bodyPr, 'rIns')),
    paddingBottom: vmlUnitToPx(attr(bodyPr, 'bIns')),
    paddingLeft: vmlUnitToPx(attr(bodyPr, 'lIns')),
  };
}

function convertDrawingCustomGeometry(spPr: Element | null | undefined, width: number, height: number) {
  const pathNode = descendantByLocalName(childByLocalName(spPr, 'custGeom'), 'path');
  if (!pathNode) return undefined;
  const pathWidth = Number(attr(pathNode, 'w') ?? 0);
  const pathHeight = Number(attr(pathNode, 'h') ?? 0);
  const scaleX = Number.isFinite(pathWidth) && pathWidth > 0 ? width / pathWidth : 1;
  const scaleY = Number.isFinite(pathHeight) && pathHeight > 0 ? height / pathHeight : 1;
  const commands: string[] = [];

  Array.from(pathNode.children).forEach((child) => {
    if (matchesLocalName(child, 'close')) {
      commands.push('Z');
      return;
    }
    const points = descendantsByLocalName(child, 'pt').map((point) => {
      const x = Number(attr(point, 'x') ?? 0);
      const y = Number(attr(point, 'y') ?? 0);
      return Number.isFinite(x) && Number.isFinite(y) ? `${x * scaleX} ${y * scaleY}` : undefined;
    });
    if (matchesLocalName(child, 'moveTo') && points[0]) {
      commands.push(`M ${points[0]}`);
    }
    if (matchesLocalName(child, 'lnTo') && points[0]) {
      commands.push(`L ${points[0]}`);
    }
    if (matchesLocalName(child, 'cubicBezTo') && points.length >= 3 && points.every(Boolean)) {
      commands.push(`C ${points.join(' ')}`);
    }
  });

  return commands.length ? commands.join(' ') : undefined;
}

function parseWpgShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  scale: { x: number; y: number },
): DocxShapeItem | undefined {
  const spPr = childByLocalName(shapeNode, 'spPr');
  const size = parseDrawingXfrm(spPr, scale);
  const kind = readDrawingShapeKind(spPr);
  const isLine = kind === 'line';
  if (!isLine && (!size.width || !size.height)) return undefined;
  if (isLine && !size.width && !size.height) return undefined;

  const id = `wpg-item-${context.shapeIndex + 1}-${index + 1}`;
  const fillColor = parseDrawingFillColor(spPr, context.theme);
  const stroke = parseDrawingLineStyle(spPr, context.theme);
  const paragraphs = parseVmlTextBoxParagraphs(shapeNode, context, id).filter(
    (paragraph) => paragraph.text || paragraph.inlines.length,
  );
  const path = kind === 'path'
    ? convertDrawingCustomGeometry(spPr, size.width, size.height)
    : isLine
      ? `M 0 0 L ${size.width || 0} ${size.height || 0}`
      : undefined;

  return {
    id,
    kind,
    ...size,
    height: isLine && size.height === 0 ? 1 : size.height,
    ...readDrawingBodyPadding(shapeNode),
    path,
    viewBox: path ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}` : undefined,
    fillColor,
    ...stroke,
    borderRadius: kind === 'ellipse' ? '50%' : undefined,
    textVerticalAlign: readDrawingTextAnchor(shapeNode),
    paragraphs: paragraphs.length ? paragraphs : undefined,
  };
}

function parseWpgShape(node: Element, context: ParseContext): DocxShape | undefined {
  const group = descendantByLocalName(node, 'wgp');
  if (!group) return undefined;
  const extent = descendantByLocalName(node, 'extent');
  const width = Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)));
  const height = Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)));
  if (!width || !height) return undefined;

  const scale = readWpgScale(group, width, height);
  const items = childrenByLocalName(group, 'wsp')
    .map((shapeNode, index) => parseWpgShapeItem(shapeNode, index, context, scale))
    .filter((item): item is DocxShapeItem => Boolean(item));

  if (!items.length) return undefined;
  context.shapeIndex += 1;
  return {
    id: `docx-shape-${context.shapeIndex}`,
    width,
    height,
    items,
  };
}

function parseAlternateContentShape(node: Element, context: ParseContext): DocxShape | undefined {
  const choice = childByLocalName(node, 'Choice');
  const choiceDrawing = descendantByLocalName(choice, 'drawing');
  const choiceShape = choiceDrawing ? parseWpgShape(choiceDrawing, context) : undefined;
  if (choiceShape) return choiceShape;

  const fallback = childByLocalName(node, 'Fallback');
  const fallbackPict = descendantByLocalName(fallback, 'pict');
  return fallbackPict ? parseVmlShape(fallbackPict, context) : undefined;
}

function parseVmlCoordSize(node: Element, renderedWidth: number, renderedHeight: number) {
  const coordsize = attr(node, 'coordsize');
  const [coordWidth, coordHeight] = (coordsize ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    x: Number.isFinite(coordWidth) && coordWidth > 0 ? renderedWidth / coordWidth : undefined,
    y: Number.isFinite(coordHeight) && coordHeight > 0 ? renderedHeight / coordHeight : undefined,
  };
}

function readVmlCoordSize(node: Element) {
  const [width, height] = (attr(node, 'coordsize') ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    width: Number.isFinite(width) && width > 0 ? width : undefined,
    height: Number.isFinite(height) && height > 0 ? height : undefined,
  };
}

function parseVmlShapeSize(node: Element, scale?: { x?: number; y?: number }) {
  const style = attr(node, 'style');
  return {
    left: readCssSize(style, 'left', scale?.x) ?? 0,
    top: readCssSize(style, 'top', scale?.y) ?? 0,
    width: readCssSize(style, 'width', scale?.x) ?? 0,
    height: readCssSize(style, 'height', scale?.y) ?? 0,
  };
}

function vmlOnOff(value: string | undefined, fallback = true) {
  if (value === undefined) return fallback;
  return value !== 'f' && value !== 'false' && value !== '0' && value !== 'off';
}

function parseVmlStroke(shapeNode: Element) {
  const stroke = childByLocalName(shapeNode, 'stroke');
  if (!vmlOnOff(attr(shapeNode, 'stroked'), true) || !vmlOnOff(attr(stroke, 'on'), true)) {
    return {};
  }
  const color = normalizeCssColor(attr(stroke, 'color') ?? attr(shapeNode, 'strokecolor')) ?? '#000';
  const width = vmlUnitToPx(attr(stroke, 'weight')) ?? 1;
  const dashstyle = attr(stroke, 'dashstyle');
  const strokeDasharray = dashstyle
    ?.split(/\s+/)
    .map((item) => Number(item))
    .filter((item) => Number.isFinite(item) && item > 0)
    .map((item) => item * width)
    .join(' ');
  return {
    border: `${width}px solid ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray: strokeDasharray || undefined,
  };
}

function parseVmlFillColor(shapeNode: Element) {
  const fill = childByLocalName(shapeNode, 'fill');
  if (!vmlOnOff(attr(shapeNode, 'filled'), true) || !vmlOnOff(attr(fill, 'on'), true)) {
    return undefined;
  }
  return normalizeCssColor(attr(fill, 'color') ?? attr(shapeNode, 'fillcolor'));
}

function readVmlTextAnchor(shapeNode: Element): DocxShapeItem['textVerticalAlign'] {
  const anchor = readCssDeclaration(attr(shapeNode, 'style'), 'v-text-anchor');
  if (anchor === 'middle') return 'middle';
  if (anchor === 'bottom') return 'bottom';
  return 'top';
}

function convertVmlPathToSvgPath(path: string | undefined, width: number, height: number, node: Element) {
  if (!path) return undefined;
  const coordSize = readVmlCoordSize(node);
  const scaleX = coordSize.width ? width / coordSize.width : 1;
  const scaleY = coordSize.height ? height / coordSize.height : 1;
  const tokens = path.match(/[a-z]|-?\d+(?:\.\d+)?/gi) ?? [];
  const commands: string[] = [];
  let index = 0;
  let command = '';

  const readPoint = () => {
    const x = Number(tokens[index++]);
    const y = Number(tokens[index++]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return undefined;
    return `${x * scaleX} ${y * scaleY}`;
  };

  while (index < tokens.length) {
    const token = tokens[index++];
    if (/^[a-z]$/i.test(token)) {
      command = token.toLowerCase();
    } else {
      index -= 1;
    }

    if (command === 'm') {
      const point = readPoint();
      if (point) commands.push(`M ${point}`);
    } else if (command === 'l') {
      const point = readPoint();
      if (point) commands.push(`L ${point}`);
    } else if (command === 'c') {
      const points = [readPoint(), readPoint(), readPoint()];
      if (points.every(Boolean)) commands.push(`C ${points.join(' ')}`);
    } else if (command === 'x') {
      commands.push('Z');
    } else if (command === 'e') {
      break;
    } else {
      break;
    }
  }

  return commands.length ? commands.join(' ') : undefined;
}

function parseVmlTextBoxParagraphs(shapeNode: Element, context: ParseContext, id: string) {
  const textBox = descendantByLocalName(shapeNode, 'txbxContent');
  return childrenByLocalName(textBox, 'p').map((pNode, index) =>
    parseParagraph(pNode, `${id}-p-${index + 1}`, context),
  );
}

function hasVmlTextBox(shapeNode: Element) {
  return Boolean(descendantByLocalName(shapeNode, 'txbxContent'));
}

function parseVmlShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  scale?: { x?: number; y?: number },
): DocxShapeItem | undefined {
  const size = parseVmlShapeSize(shapeNode, scale);
  if (!size.width || !size.height) return undefined;

  const isEllipse =
    matchesLocalName(shapeNode, 'shape') &&
    ((attr(shapeNode, 'o:spt') ?? attr(shapeNode, 'spt')) === '3' || (attr(shapeNode, 'type') ?? '').includes('_x0000_t3'));
  const fillColor = parseVmlFillColor(shapeNode);
  const stroke = parseVmlStroke(shapeNode);
  const path = convertVmlPathToSvgPath(attr(shapeNode, 'path'), size.width, size.height, shapeNode);
  const id = `vml-item-${context.shapeIndex + 1}-${index + 1}`;
  const paragraphs = parseVmlTextBoxParagraphs(shapeNode, context, id).filter(
    (paragraph) => paragraph.text || paragraph.inlines.length,
  );

  return {
    id,
    kind: isEllipse ? 'ellipse' : 'rect',
    ...size,
    path,
    viewBox: path ? `0 0 ${size.width} ${size.height}` : undefined,
    fillColor,
    ...stroke,
    borderRadius: isEllipse ? '50%' : matchesLocalName(shapeNode, 'roundrect') ? 8 : undefined,
    textVerticalAlign: readVmlTextAnchor(shapeNode),
    paragraphs: paragraphs.length ? paragraphs : undefined,
  };
}

function parseVmlShape(node: Element, context: ParseContext): DocxShape | undefined {
  const group = matchesLocalName(node, 'group') ? node : descendantByLocalName(node, 'group');
  const shapeRoot = group ?? node;
  const rootSize = parseVmlShapeSize(shapeRoot);
  const scale = parseVmlCoordSize(shapeRoot, rootSize.width, rootSize.height);
  const rawItems = Array.from(shapeRoot.children).filter(
    (child) =>
      (matchesLocalName(child, 'shape') || matchesLocalName(child, 'rect') || matchesLocalName(child, 'roundrect')) &&
      (child.hasAttribute('fillcolor') ||
        child.hasAttribute('strokecolor') ||
        attr(child, 'filled') !== 'f' ||
        attr(child, 'stroked') !== 'f' ||
        hasVmlTextBox(child)),
  );
  const items = rawItems
    .map((child, index) => parseVmlShapeItem(child, index, context, scale))
    .filter((item): item is DocxShapeItem => Boolean(item));

  if (!items.length) return undefined;

  const maxRight = Math.max(...items.map((item) => item.left + item.width));
  const maxBottom = Math.max(...items.map((item) => item.top + item.height));
  context.shapeIndex += 1;

  return {
    id: `docx-shape-${context.shapeIndex}`,
    width: rootSize.width || maxRight,
    height: rootSize.height || maxBottom,
    items,
  };
}

function parseRun(runNode: Element, paragraphStyle: DocxTextStyle | undefined, context: ParseContext): DocxInline[] {
  const runStyle = mergeTextStyle(
    inlineInheritedStyle(paragraphStyle),
    resolveRunStyle(childByLocalName(runNode, 'rPr'), context.styles, context.theme),
  );
  const inlines: DocxInline[] = [];

  Array.from(runNode.children).forEach((child) => {
    if (matchesLocalName(child, 't')) {
      inlines.push({ type: 'text', text: textContent(child), style: runStyle });
      return;
    }
    if (matchesLocalName(child, 'tab')) {
      inlines.push({ type: 'text', text: '\t', style: runStyle });
      return;
    }
    if (matchesLocalName(child, 'br') || matchesLocalName(child, 'cr')) {
      inlines.push({ type: 'break' });
      return;
    }
    if (matchesLocalName(child, 'drawing')) {
      const shape = parseWpgShape(child, context);
      if (shape) {
        inlines.push({ type: 'shape', shape });
        return;
      }
      const webExtensionChart = parseWpsWebExtensionChart(child, context);
      if (webExtensionChart) {
        inlines.push({ type: 'chart', chart: webExtensionChart });
        return;
      }
      const chart = parseChartElement(child, context);
      if (chart) {
        inlines.push({ type: 'chart', chart });
        return;
      }
      const image = parseDrawingImage(child, context);
      if (image) {
        inlines.push({ type: 'image', image });
        return;
      }
    }
    if (matchesLocalName(child, 'pict') || matchesLocalName(child, 'alternateContent')) {
      const shape = matchesLocalName(child, 'pict') ? parseVmlShape(child, context) : parseAlternateContentShape(child, context);
      if (shape) {
        inlines.push({ type: 'shape', shape });
      }
    }
  });

  return inlines;
}

function readParagraphBlocks(pNode: Element, id: string, context: ParseContext): DocxParagraphBlock[] {
  const paragraph = parseParagraph(pNode, id, context);
  return [paragraph];
}

function readParagraphRuns(pNode: Element, paragraphStyle: DocxTextStyle | undefined, context: ParseContext) {
  const inlines: DocxInline[] = [];

  Array.from(pNode.children).forEach((child) => {
    if (matchesLocalName(child, 'r')) {
      inlines.push(...parseRun(child, paragraphStyle, context));
    }
    if (matchesLocalName(child, 'hyperlink')) {
      childrenByLocalName(child, 'r').forEach((runNode) => {
        inlines.push(...parseRun(runNode, paragraphStyle, context));
      });
    }
  });

  return inlines;
}

function textFromInlines(inlines: DocxInline[]) {
  return inlines.map((inline) => (inline.type === 'text' ? inline.text : '')).join('');
}

function parseParagraph(pNode: Element, id: string, context: ParseContext): DocxParagraphBlock {
  const pPr = childByLocalName(pNode, 'pPr');
  const style = resolveParagraphStyle(pPr, context.styles, context.theme);
  const inlines = readParagraphRuns(pNode, style.style, context);
  const text = textFromInlines(inlines).trim();

  return {
    id,
    type: 'paragraph',
    inlines,
    text,
    align: style.align,
    lineHeight: style.lineHeight,
    style: style.style,
    spacingBefore: style.spacingBefore,
    spacingAfter: style.spacingAfter,
    indentLeft: style.indentLeft,
    indentRight: style.indentRight,
    firstLineIndent: style.firstLineIndent,
    backgroundColor: style.backgroundColor,
    borderTop: style.borderTop,
    borderRight: style.borderRight,
    borderBottom: style.borderBottom,
    borderLeft: style.borderLeft,
    paddingTop: style.paddingTop,
    paddingRight: style.paddingRight,
    paddingBottom: style.paddingBottom,
    paddingLeft: style.paddingLeft,
  };
}

function readCellMargins(tcPr: Element | null | undefined) {
  const tcMar = childByLocalName(tcPr, 'tcMar') ?? childByLocalName(tcPr, 'tblCellMar');
  const readMargin = (name: string) => {
    const node = childByLocalName(tcMar, name);
    return positiveTwipToPx(attr(node, 'w:w') ?? attr(node, 'w'));
  };
  return {
    paddingTop: readMargin('top'),
    paddingRight: readMargin('right'),
    paddingBottom: readMargin('bottom'),
    paddingLeft: readMargin('left'),
  };
}

function mergeCellMargins(
  base: Pick<DocxTableCell, 'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'>,
  next: Pick<DocxTableCell, 'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'>,
) {
  return {
    paddingTop: next.paddingTop ?? base.paddingTop,
    paddingRight: next.paddingRight ?? base.paddingRight,
    paddingBottom: next.paddingBottom ?? base.paddingBottom,
    paddingLeft: next.paddingLeft ?? base.paddingLeft,
  };
}

function readCellBorders(tcPr: Element | null | undefined) {
  const tcBorders = childByLocalName(tcPr, 'tcBorders');
  const top = childByLocalName(tcBorders, 'top');
  const right = childByLocalName(tcBorders, 'right');
  const bottom = childByLocalName(tcBorders, 'bottom');
  const left = childByLocalName(tcBorders, 'left');
  return {
    borderTop: readBorder(top),
    borderRight: readBorder(right),
    borderBottom: readBorder(bottom),
    borderLeft: readBorder(left),
    hasBorderTop: Boolean(top),
    hasBorderRight: Boolean(right),
    hasBorderBottom: Boolean(bottom),
    hasBorderLeft: Boolean(left),
  };
}

function readCellStyle(
  tcNode: Element,
  defaultMargins: Pick<DocxTableCell, 'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'>,
  theme: OfficeTheme,
): Omit<DocxTableCell, 'id' | 'blocks'> {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const gridSpan = childByLocalName(tcPr, 'gridSpan');
  const width = childByLocalName(tcPr, 'tcW');
  const vAlign = attr(childByLocalName(tcPr, 'vAlign'), 'w:val') ?? attr(childByLocalName(tcPr, 'vAlign'), 'val');
  const shading = childByLocalName(tcPr, 'shd');
  const margins = mergeCellMargins(defaultMargins, readCellMargins(tcPr));
  return {
    colSpan: Number(attr(gridSpan, 'w:val') ?? attr(gridSpan, 'val') ?? 1),
    width: twipToPx(attr(width, 'w:w') ?? attr(width, 'w')),
    verticalAlign: vAlign === 'center' ? 'middle' : vAlign === 'bottom' ? 'bottom' : 'top',
    backgroundColor: readShading(shading, theme),
    noWrap: readOnOff(childByLocalName(tcPr, 'noWrap')),
    ...readCellBorders(tcPr),
    ...margins,
  };
}

function parseTable(tblNode: Element, id: string, context: ParseContext): DocxTableBlock {
  const tblPr = childByLocalName(tblNode, 'tblPr');
  const tblW = childByLocalName(tblPr, 'tblW');
  const align = mapAlignment(readVal(childByLocalName(tblPr, 'jc')));
  const columns = childrenByLocalName(childByLocalName(tblNode, 'tblGrid'), 'gridCol')
    .map((col) => positiveTwipToPx(attr(col, 'w:w') ?? attr(col, 'w')))
    .filter((width): width is number => width !== undefined);
  const tableMargins = readCellMargins(tblPr);
  return {
    id,
    type: 'table',
    width: positiveTwipToPx(attr(tblW, 'w:w') ?? attr(tblW, 'w')),
    align: align === 'center' || align === 'right' ? align : 'left',
    columns,
    rows: childrenByLocalName(tblNode, 'tr').map((rowNode, rowIndex) => ({
      id: `${id}-row-${rowIndex + 1}`,
      cells: childrenByLocalName(rowNode, 'tc').map((cellNode, cellIndex): DocxTableCell => ({
        id: `${id}-cell-${rowIndex + 1}-${cellIndex + 1}`,
        ...readCellStyle(cellNode, tableMargins, context.theme),
        blocks: childrenByLocalName(cellNode, 'p').flatMap((pNode, paragraphIndex) =>
          readParagraphBlocks(pNode, `${id}-cell-${rowIndex + 1}-${cellIndex + 1}-p-${paragraphIndex + 1}`, context),
        ),
      })),
    })),
  };
}

function readPage(bodyNode: Element | null | undefined): DocxPage {
  const sectPr = childByLocalName(bodyNode, 'sectPr');
  const pgSz = childByLocalName(sectPr, 'pgSz');
  const pgMar = childByLocalName(sectPr, 'pgMar');
  const pgBorders = childByLocalName(sectPr, 'pgBorders');

  return {
    width: Math.round(twipToPx(attr(pgSz, 'w:w') ?? attr(pgSz, 'w')) ?? DEFAULT_PAGE.width),
    minHeight: Math.round(twipToPx(attr(pgSz, 'w:h') ?? attr(pgSz, 'h')) ?? DEFAULT_PAGE.minHeight),
    marginTop: Math.round(twipToPx(attr(pgMar, 'w:top') ?? attr(pgMar, 'top')) ?? DEFAULT_PAGE.marginTop),
    marginRight: Math.round(twipToPx(attr(pgMar, 'w:right') ?? attr(pgMar, 'right')) ?? DEFAULT_PAGE.marginRight),
    marginBottom: Math.round(twipToPx(attr(pgMar, 'w:bottom') ?? attr(pgMar, 'bottom')) ?? DEFAULT_PAGE.marginBottom),
    marginLeft: Math.round(twipToPx(attr(pgMar, 'w:left') ?? attr(pgMar, 'left')) ?? DEFAULT_PAGE.marginLeft),
    borderTop: readBorder(childByLocalName(pgBorders, 'top')),
    borderRight: readBorder(childByLocalName(pgBorders, 'right')),
    borderBottom: readBorder(childByLocalName(pgBorders, 'bottom')),
    borderLeft: readBorder(childByLocalName(pgBorders, 'left')),
  };
}

function markTitle(blocks: DocxBlock[]) {
  const firstParagraph = blocks.find(
    (block): block is DocxParagraphBlock => block.type === 'paragraph' && Boolean(block.text),
  );
  return firstParagraph?.text ?? 'DOCX 文档';
}

export async function parseDocx(file: File): Promise<DocxDocument> {
  const entries = await loadDocxEntries(file);
  const packageState = buildPackageState(entries);
  const theme = readOfficeTheme(readXml(entries, 'word/theme/theme1.xml'));
  const documentXml = readXml(entries, 'word/document.xml');
  const documentDoc = parseXml(documentXml);
  const bodyNode = childByLocalName(documentDoc.documentElement, 'body');
  const context: ParseContext = {
    packageState,
    documentRels: packageState.relationships['word/_rels/document.xml.rels'] ?? {},
    theme,
    styles: readDocxStyles(entries, theme),
    images: [],
    imageIndex: 0,
    chartIndex: 0,
    shapeIndex: 0,
  };

  const blocks: DocxBlock[] = [];
  Array.from(bodyNode?.children ?? []).forEach((child, index) => {
    if (matchesLocalName(child, 'p')) {
      blocks.push(...readParagraphBlocks(child, `p-${index + 1}`, context));
    }
    if (matchesLocalName(child, 'tbl')) {
      blocks.push(parseTable(child, `table-${index + 1}`, context));
    }
  });

  return {
    title: markTitle(blocks),
    page: readPage(bodyNode),
    blocks,
    images: context.images,
  };
}
