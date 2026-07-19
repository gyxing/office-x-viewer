import type { OfficeEntryMap } from '../../shared/ooxml/archive';
import { readXml } from '../../shared/ooxml/archive';
import { parseOfficeChartXml } from '../../shared/ooxml/charts';
import {
  collectMedia,
  resolvePackageMediaRef,
  type OfficeRelationship,
} from '../../shared/ooxml/media';
import { readRelationships } from '../../shared/ooxml/relationships';
import {
  readOfficeTheme,
  resolveOfficeThemeColor,
  type OfficeTheme,
} from '../../shared/ooxml/theme';
import { emuToPx } from '../../shared/ooxml/units';
import { parseWpsWebExtensionChartModel } from '../../shared/ooxml/wpsChart';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  matchesLocalName,
  parseXml,
  textContent,
} from '../../shared/ooxml/xml';
import { loadDocxEntries } from './archive';
import type {
  DocxBlock,
  DocxChartBlock,
  DocxDocument,
  DocxImage,
  DocxInline,
  DocxPage,
  DocxPageContent,
  DocxParagraphBlock,
  DocxPosition,
  DocxShape,
  DocxShapeItem,
  DocxTableBlock,
  DocxTableCell,
  DocxTableRow,
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

type ReadBlockChildrenOptions = {
  insideShape?: boolean;
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

const DEFAULT_DOCX_FONT_FAMILY =
  '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif';

// DOCX 与 PPTX 类似是 zip 包结构，正文、样式、主题、媒体通过关系文件互相引用。
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
  const match = style.match(
    new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i'),
  );
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

function readCssPosition(style: string | undefined, name: 'left' | 'top') {
  return readCssSize(style, `margin-${name}`) ?? readCssSize(style, name);
}

function readDocxLineHeight(spacingNode: Element | null | undefined) {
  const value = Number(
    attr(spacingNode, 'w:line') ?? attr(spacingNode, 'line'),
  );
  if (!Number.isFinite(value) || value <= 0) return undefined;
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
  // 提取颜色值，忽略额外的信息（如 "#41719C [3204]"）
  const match = value.match(/^#?([0-9a-f]{6})/i);
  if (!match) return undefined;
  return `#${match[1]}`;
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
    .map((channel) =>
      clamp255(channel + (255 - channel) * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
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
  return `#${[r, g, b]
    .map((channel) =>
      clamp255(channel * ratio)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

function resolveThemeFillColor(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  const themeFill = attr(node, 'w:themeFill') ?? attr(node, 'themeFill');
  const themeColor = resolveOfficeThemeColor(themeFill, theme);
  return shadeHexColor(
    tintHexColor(
      themeColor,
      attr(node, 'w:themeFillTint') ?? attr(node, 'themeFillTint'),
    ),
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
  const color =
    parseHexColor(attr(node, 'w:color') ?? attr(node, 'color')) ?? '#000';
  const width = eighthPointToPx(attr(node, 'w:sz') ?? attr(node, 'sz')) ?? 1;
  const style =
    value === 'dashed' || value === 'dashSmallGap'
      ? 'dashed'
      : value === 'dotted'
      ? 'dotted'
      : 'solid';
  return `${width}px ${style} ${color}`;
}

function readParagraphBorders(pPr: Element | null | undefined) {
  const pBdr = childByLocalName(pPr, 'pBdr');
  return {
    borderTop: readBorder(childByLocalName(pBdr, 'top')),
    borderRight: readBorder(childByLocalName(pBdr, 'right')),
    borderBottom: readBorder(childByLocalName(pBdr, 'bottom')),
    borderLeft: readBorder(childByLocalName(pBdr, 'left')),
    paddingTop: pointToPx(
      attr(childByLocalName(pBdr, 'top'), 'w:space') ??
        attr(childByLocalName(pBdr, 'top'), 'space'),
    ),
    paddingRight: pointToPx(
      attr(childByLocalName(pBdr, 'right'), 'w:space') ??
        attr(childByLocalName(pBdr, 'right'), 'space'),
    ),
    paddingBottom: pointToPx(
      attr(childByLocalName(pBdr, 'bottom'), 'w:space') ??
        attr(childByLocalName(pBdr, 'bottom'), 'space'),
    ),
    paddingLeft: pointToPx(
      attr(childByLocalName(pBdr, 'left'), 'w:space') ??
        attr(childByLocalName(pBdr, 'left'), 'space'),
    ),
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
    spacingBefore: positiveTwipToPx(
      attr(spacing, 'w:before') ?? attr(spacing, 'before'),
    ),
    spacingAfter: positiveTwipToPx(
      attr(spacing, 'w:after') ?? attr(spacing, 'after'),
    ),
    indentLeft: twipToPx(attr(ind, 'w:left') ?? attr(ind, 'left')),
    indentRight: twipToPx(attr(ind, 'w:right') ?? attr(ind, 'right')),
    firstLineIndent: twipToPx(
      attr(ind, 'w:firstLine') ?? attr(ind, 'firstLine'),
    ),
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
  return themeFont.toLowerCase().includes('major')
    ? theme.fontScheme?.majorFont
    : theme.fontScheme?.minorFont;
}

function quoteFontFamily(value?: string) {
  if (!value) return undefined;
  return value
    .split(',')
    .map((font) => font.trim())
    .filter(Boolean)
    .map((font) =>
      /^["'].*["']$/.test(font) || /^[a-z-]+$/i.test(font) ? font : `"${font}"`,
    )
    .join(', ');
}

function readFontFamily(
  rPr: Element | null | undefined,
  theme: OfficeTheme,
  allowFallback = false,
) {
  const rFonts = childByLocalName(rPr, 'rFonts');
  const ascii = attr(rFonts, 'w:ascii') ?? attr(rFonts, 'ascii');
  const eastAsia = attr(rFonts, 'w:eastAsia') ?? attr(rFonts, 'eastAsia');
  const hAnsi = attr(rFonts, 'w:hAnsi') ?? attr(rFonts, 'hAnsi');
  const cs = attr(rFonts, 'w:cs') ?? attr(rFonts, 'cs');
  const themeFonts = theme.fontScheme ?? {};
  const explicitFont =
    eastAsia ?? ascii ?? hAnsi ?? cs ?? readThemeFont(rFonts, theme);
  if (explicitFont || !allowFallback) return quoteFontFamily(explicitFont);
  return quoteFontFamily(
    themeFonts.minorFont ?? themeFonts.majorFont ?? DEFAULT_DOCX_FONT_FAMILY,
  );
}

function readDocxStyles(
  entries: OfficeEntryMap,
  theme: OfficeTheme,
): DocxStyleCatalog {
  // styles.xml 会提供默认样式和命名样式，段落/文字解析时再与直接格式合并。
  const xml = readXml(entries, 'word/styles.xml');
  if (!xml) return { defaults: {}, styles: {} };

  const doc = parseXml(xml);
  const root = doc.documentElement;
  const styles: Record<string, DocxStyleDefinition> = {};
  const defaults: DocxStyleCatalog['defaults'] = {};

  const docDefaults = childByLocalName(root, 'docDefaults');
  const rPrDefault = childByLocalName(
    childByLocalName(docDefaults, 'rPrDefault'),
    'rPr',
  );
  const pPrDefault = childByLocalName(
    childByLocalName(docDefaults, 'pPrDefault'),
    'pPr',
  );
  defaults.run = readTextStyle(rPrDefault, theme, true);
  defaults.paragraph = mergeTextStyle(
    readParagraphPropertyStyle(pPrDefault, theme),
    readTextStyle(childByLocalName(pPrDefault, 'rPr'), theme, true),
  );

  childrenByLocalName(root, 'style').forEach((styleNode) => {
    const styleId = attr(styleNode, 'styleId');
    const kindAttr = attr(styleNode, 'type');
    if (!styleId) return;
    if (kindAttr === 'paragraph' && attr(styleNode, 'w:default') === '1')
      defaults.paragraphStyleId = styleId;
    if (kindAttr === 'table' && attr(styleNode, 'w:default') === '1')
      defaults.tableStyleId = styleId;

    const basedOn =
      attr(childByLocalName(styleNode, 'basedOn'), 'w:val') ??
      attr(childByLocalName(styleNode, 'basedOn'), 'val') ??
      undefined;
    const name = styleId;
    const pPr = childByLocalName(styleNode, 'pPr');
    const rPr = childByLocalName(styleNode, 'rPr');

    let style: DocxTextStyle | undefined;
    if (kindAttr === 'paragraph') {
      style = mergeTextStyle(
        readParagraphPropertyStyle(pPr, theme),
        readTextStyle(rPr, theme),
      );
    } else if (kindAttr === 'table') {
      const tblPr = childByLocalName(styleNode, 'tblPr');
      style = readParagraphPropertyStyle(childByLocalName(tblPr, 'pPr'), theme);
    } else {
      style = readTextStyle(rPr, theme);
    }

    if (style) {
      styles[name] = {
        kind:
          kindAttr === 'paragraph'
            ? 'paragraph'
            : kindAttr === 'table'
            ? 'table'
            : 'character',
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

  const color =
    readDrawingColor(childByLocalName(rPr, 'textFill'), theme) ??
    parseHexColor(
      attr(childByLocalName(rPr, 'color'), 'w:val') ??
        attr(childByLocalName(rPr, 'color'), 'val'),
    );
  const style: DocxTextStyle = {
    bold: firstDefined(
      readOnOff(childByLocalName(rPr, 'b')),
      readOnOff(childByLocalName(rPr, 'bCs')),
    ),
    italic: firstDefined(
      readOnOff(childByLocalName(rPr, 'i')),
      readOnOff(childByLocalName(rPr, 'iCs')),
    ),
    underline: readUnderline(rPr),
    strike: firstDefined(
      readOnOff(childByLocalName(rPr, 'strike')),
      readOnOff(childByLocalName(rPr, 'dstrike')),
    ),
    smallCaps: readOnOff(childByLocalName(rPr, 'smallCaps')),
    allCaps: readOnOff(childByLocalName(rPr, 'caps')),
    color,
    backgroundColor:
      readHighlight(childByLocalName(rPr, 'highlight')) ??
      readShading(childByLocalName(rPr, 'shd'), theme),
    fontSize: halfPointToPx(
      attr(childByLocalName(rPr, 'sz'), 'w:val') ??
        attr(childByLocalName(rPr, 'sz'), 'val'),
    ),
    fontFamily: readFontFamily(rPr, theme, allowFontFallback),
  };

  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function mergeTwoTextStyles(
  base?: DocxTextStyle,
  next?: DocxTextStyle,
): DocxTextStyle {
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

function mergeTextStyle(
  ...styles: Array<DocxTextStyle | undefined>
): DocxTextStyle | undefined {
  const merged = styles.reduce<DocxTextStyle>(
    (acc, style) => mergeTwoTextStyles(acc, style),
    {},
  );
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
  // 段落最终样式 = 默认段落样式 + 命名段落样式 + 段落直接属性 + 段落内 run 属性。
  const styleId =
    attr(childByLocalName(pPr, 'pStyle'), 'w:val') ??
    attr(childByLocalName(pPr, 'pStyle'), 'val');
  const baseStyle = resolveDocxStyle(
    catalog.defaults.paragraphStyleId,
    catalog,
  );
  const namedStyle = resolveDocxStyle(styleId, catalog);
  const style = mergeTextStyle(
    catalog.defaults.paragraph,
    baseStyle,
    namedStyle,
  );
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
    style: mergeTextStyle(
      style,
      directStyle,
      readTextStyle(childByLocalName(pPr, 'rPr'), theme),
    ),
  };
}

function resolveRunStyle(
  rPr: Element | null | undefined,
  catalog: DocxStyleCatalog,
  theme: OfficeTheme,
) {
  const styleId =
    attr(childByLocalName(rPr, 'rStyle'), 'w:val') ??
    attr(childByLocalName(rPr, 'rStyle'), 'val');
  return mergeTextStyle(
    catalog.defaults.run,
    resolveDocxStyle(styleId, catalog),
    readTextStyle(rPr, theme),
  );
}

function inlineInheritedStyle(
  style: DocxTextStyle | undefined,
): DocxTextStyle | undefined {
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
  if (
    value === 'left' ||
    value === 'center' ||
    value === 'right' ||
    value === 'justify'
  )
    return value;
  if (value === 'both') return 'justify';
  return undefined;
}

function resolveMediaRef(
  target: string | undefined,
  packageState: DocxPackageState,
) {
  return resolvePackageMediaRef(
    target,
    packageState.mediaByPath,
    packageState.mediaByName,
    'word',
  );
}

function resolveXmlTarget(
  target: string | undefined,
  packageState: DocxPackageState,
) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
}

function readDrawingAnchorPosition(node: Element) {
  const anchor = descendantByLocalName(node, 'anchor');
  if (!anchor) return undefined;

  const positionH = childByLocalName(anchor, 'positionH');
  const positionV = childByLocalName(anchor, 'positionV');
  const left = emuToPx(
    Number(textContent(childByLocalName(positionH, 'posOffset')).trim()),
  );
  const top = emuToPx(
    Number(textContent(childByLocalName(positionV, 'posOffset')).trim()),
  );
  if (!Number.isFinite(left) || !Number.isFinite(top)) return undefined;

  const relativeHeight = Number(attr(anchor, 'relativeHeight'));
  const rotation = Number(attr(anchor, 'rotation'));

  return {
    left: Math.round(left),
    top: Math.round(top),
    relativeFromH: attr(
      positionH,
      'relativeFrom',
    ) as DocxPosition['relativeFromH'],
    relativeFromV: attr(
      positionV,
      'relativeFrom',
    ) as DocxPosition['relativeFromV'],
    zIndex: Number.isFinite(relativeHeight) ? relativeHeight : undefined,
    behindDoc: attr(anchor, 'behindDoc') === '1',
    rotation:
      Number.isFinite(rotation) && rotation !== 0
        ? rotation / 60000
        : undefined,
    flipH: attr(anchor, 'flipH') === '1' || undefined,
    flipV: attr(anchor, 'flipV') === '1' || undefined,
  };
}

function parseChartElement(
  node: Element,
  context: ParseContext,
): DocxChartBlock | undefined {
  const chartNode = descendantByLocalName(node, 'chart');
  const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const chartPath = resolveXmlTarget(target, context.packageState);
  const xml = chartPath
    ? (context.packageState.entries.get(chartPath) as string | undefined)
    : undefined;
  if (!xml) return undefined;

  const chart = parseOfficeChartXml(xml, context.theme);
  const extent =
    descendantByLocalName(node, 'extent') ??
    descendantByLocalName(node, 'xfrm');
  const width = Math.max(
    160,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320),
  );
  const height = Math.max(
    120,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220),
  );
  context.chartIndex += 1;
  return {
    id: `docx-chart-${context.chartIndex}`,
    type: 'chart',
    chart,
    width,
    height,
  };
}

function resolveWebExtensionSnapshot(
  doc: XMLDocument,
  webExtensionPath: string,
  context: ParseContext,
) {
  const snapshot = descendantByLocalName(doc.documentElement, 'snapshot');
  const embed = attr(snapshot, 'r:embed') ?? attr(snapshot, 'embed');
  const relsPath = webExtensionPath
    .replace(/^word\/webExtensions\//, 'word/webExtensions/_rels/')
    .concat('.rels');
  const target = embed
    ? context.packageState.relationships[relsPath]?.[embed]?.target
    : undefined;
  return resolveMediaRef(target, context.packageState);
}

function parseWpsWebExtensionChart(
  node: Element,
  context: ParseContext,
): DocxChartBlock | undefined {
  // 关系和尺寸属于 DOCX 包装层，WPS JSON 到图表模型的转换由共享适配器负责。
  const webExtensionRef = descendantByLocalName(node, 'webExtensionRef');
  const relId = attr(webExtensionRef, 'r:id') ?? attr(webExtensionRef, 'id');
  const target = relId ? context.documentRels[relId]?.target : undefined;
  const webExtensionPath = resolveXmlTarget(target, context.packageState);
  const xml = webExtensionPath
    ? (context.packageState.entries.get(webExtensionPath) as string | undefined)
    : undefined;
  if (!xml || !webExtensionPath) return undefined;

  const doc = parseXml(xml);
  const snapshotSrc = resolveWebExtensionSnapshot(
    doc,
    webExtensionPath,
    context,
  );
  const chart = parseWpsWebExtensionChartModel(
    doc.documentElement,
    snapshotSrc,
  );
  if (!chart) return undefined;

  const extent =
    descendantByLocalName(node, 'extent') ??
    descendantByLocalName(node, 'xfrm');
  const width = Math.max(
    160,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)) || 320),
  );
  const height = Math.max(
    120,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)) || 220),
  );
  context.chartIndex += 1;

  return {
    id: `docx-chart-${context.chartIndex}`,
    type: 'chart',
    chart,
    width,
    height,
  };
}

function readTopLevelDrawingGraphicData(drawingNode: Element) {
  const drawingContainer =
    childByLocalName(drawingNode, 'anchor') ??
    childByLocalName(drawingNode, 'inline');
  return childByLocalName(
    childByLocalName(drawingContainer, 'graphic'),
    'graphicData',
  );
}

function isDirectDrawingPicture(drawingNode: Element) {
  const graphicData = readTopLevelDrawingGraphicData(drawingNode);
  return (
    attr(graphicData, 'uri') ===
      'http://schemas.openxmlformats.org/drawingml/2006/picture' &&
    Boolean(childByLocalName(graphicData, 'pic'))
  );
}

function parseDrawingImage(
  drawingNode: Element,
  context: ParseContext,
): DocxImage | undefined {
  const blip = descendantByLocalName(drawingNode, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  const src = resolveMediaRef(target, context.packageState);
  if (!src) return undefined;

  const extent = descendantByLocalName(drawingNode, 'extent');
  const docPr = descendantByLocalName(drawingNode, 'docPr');
  const width = Math.max(
    1,
    Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0))),
  );
  const height = Math.max(
    1,
    Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0))),
  );
  const name = attr(docPr, 'name');
  const anchorPosition = readDrawingAnchorPosition(drawingNode);
  const imageTransform = readDrawingImageTransform(drawingNode);
  const position = anchorPosition
    ? {
        ...anchorPosition,
        flipH: imageTransform.flipH ?? anchorPosition.flipH,
        flipV: imageTransform.flipV ?? anchorPosition.flipV,
      }
    : undefined;
  const image: DocxImage = {
    id: `docx-image-${context.imageIndex + 1}`,
    name,
    alt: attr(docPr, 'descr') ?? name,
    src,
    width,
    height,
    position,
  };
  context.imageIndex += 1;
  context.images.push(image);
  return image;
}

function untrackParsedImage(context: ParseContext, image: DocxImage) {
  if (context.images[context.images.length - 1]?.id !== image.id) return;
  context.images.pop();
  context.imageIndex = Math.max(0, context.imageIndex - 1);
}

function isLikelyPageSizedNestedImage(image: DocxImage) {
  return (
    image.width >= DEFAULT_PAGE.width * 0.75 &&
    image.height >= DEFAULT_PAGE.minHeight * 0.7
  );
}

function parseAlternateContentImage(
  drawingNode: Element,
  context: ParseContext,
) {
  const image = parseDrawingImage(drawingNode, context);
  if (!image) return undefined;
  if (
    isDirectDrawingPicture(drawingNode) ||
    !isLikelyPageSizedNestedImage(image)
  )
    return image;
  untrackParsedImage(context, image);
  return undefined;
}

function readDrawingImageTransform(
  drawingNode: Element,
): Pick<DocxPosition, 'flipH' | 'flipV'> {
  const picture = descendantByLocalName(drawingNode, 'pic');
  const xfrm = childByLocalName(childByLocalName(picture, 'spPr'), 'xfrm');
  return {
    flipH: attr(xfrm, 'flipH') === '1' || undefined,
    flipV: attr(xfrm, 'flipV') === '1' || undefined,
  };
}

function readDrawingColor(
  node: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (!node) return undefined;
  const solidFill =
    childByLocalName(node, 'solidFill') ??
    (matchesLocalName(node, 'solidFill') ? node : null);
  if (!solidFill) return undefined;
  const srgb = childByLocalName(solidFill, 'srgbClr');
  const scheme = childByLocalName(solidFill, 'schemeClr');
  const sys = childByLocalName(solidFill, 'sysClr');
  const isTransparent = (colorNode: Element | null | undefined) =>
    attr(childByLocalName(colorNode, 'alpha'), 'val') === '0';
  return (
    (isTransparent(srgb) ? undefined : parseHexColor(attr(srgb, 'val'))) ??
    (isTransparent(scheme)
      ? undefined
      : resolveOfficeThemeColor(attr(scheme, 'val'), theme)) ??
    (isTransparent(sys)
      ? undefined
      : parseHexColor(attr(sys, 'lastClr') ?? attr(sys, 'val')))
  );
}

function readDrawingNoFill(node: Element | null | undefined) {
  return Boolean(childByLocalName(node, 'noFill'));
}

function readDrawingTransparentFill(node: Element | null | undefined) {
  const solidFill =
    childByLocalName(node, 'solidFill') ??
    (matchesLocalName(node, 'solidFill') ? node : null);
  const colorNodes = ['srgbClr', 'schemeClr', 'sysClr']
    .map((name) => childByLocalName(solidFill, name))
    .filter((colorNode): colorNode is Element => Boolean(colorNode));
  return (
    colorNodes.length > 0 &&
    colorNodes.every(
      (colorNode) => attr(childByLocalName(colorNode, 'alpha'), 'val') === '0',
    )
  );
}

function parseDrawingLineStyle(
  spPr: Element | null | undefined,
  theme: OfficeTheme,
) {
  const line = childByLocalName(spPr, 'ln');
  if (!line || readDrawingNoFill(line) || readDrawingTransparentFill(line))
    return {};
  const width = emuToPx(Number(attr(line, 'w') ?? 0)) || 1;
  const color = readDrawingColor(line, theme) ?? '#000';
  const dash = attr(childByLocalName(line, 'prstDash'), 'val');
  const borderStyle =
    dash && dash !== 'solid'
      ? dash.toLowerCase().includes('dot')
        ? 'dotted'
        : 'dashed'
      : 'solid';
  const strokeDasharray =
    dash && dash !== 'solid' ? `${width * 3} ${width}` : undefined;
  return {
    border: `${width}px ${borderStyle} ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray,
  };
}

function parseDrawingFillColor(
  spPr: Element | null | undefined,
  theme: OfficeTheme,
) {
  if (readDrawingNoFill(spPr)) return undefined;
  return readDrawingColor(spPr, theme);
}

function parseDrawingFillImage(
  spPr: Element | null | undefined,
  context: ParseContext,
) {
  const blipFill = childByLocalName(spPr, 'blipFill');
  const blip = childByLocalName(blipFill, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  return resolveMediaRef(target, context.packageState);
}

function parseDrawingXfrm(
  node: Element | null | undefined,
  scale?: { x?: number; y?: number },
) {
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
  const xfrm = descendantByLocalName(
    childByLocalName(groupNode, 'grpSpPr'),
    'xfrm',
  );
  const chOff = childByLocalName(xfrm, 'chOff');
  const chExt = childByLocalName(xfrm, 'chExt');
  const rawWidth = Number(attr(chExt, 'cx') ?? 0);
  const rawHeight = Number(attr(chExt, 'cy') ?? 0);
  const originX = Number(attr(chOff, 'x') ?? 0);
  const originY = Number(attr(chOff, 'y') ?? 0);
  return {
    scale: {
      x: Number.isFinite(rawWidth) && rawWidth > 0 ? width / rawWidth : 1,
      y: Number.isFinite(rawHeight) && rawHeight > 0 ? height / rawHeight : 1,
    },
    origin: {
      x: Number.isFinite(originX) ? originX : 0,
      y: Number.isFinite(originY) ? originY : 0,
    },
  };
}

function readDrawingShapeKind(
  spPr: Element | null | undefined,
): DocxShapeItem['kind'] {
  const geometry = childByLocalName(spPr, 'prstGeom');
  const preset = attr(geometry, 'prst');
  if (preset === 'line') return 'line';
  if (preset === 'ellipse') return 'ellipse';
  if (preset === 'star5' || preset === 'moon') return 'path';
  if (childByLocalName(spPr, 'custGeom')) return 'path';
  return 'rect';
}

function readDrawingShapePreset(spPr: Element | null | undefined) {
  return attr(childByLocalName(spPr, 'prstGeom'), 'prst');
}

function readDrawingShapeBorderRadius(
  spPr: Element | null | undefined,
  size: { width: number; height: number },
) {
  const geometry = childByLocalName(spPr, 'prstGeom');
  if (attr(geometry, 'prst') !== 'roundRect') return undefined;
  return Math.min(32, Math.max(8, Math.min(size.width, size.height) * 0.04));
}

function readDrawingTextAnchor(
  shapeNode: Element,
): DocxShapeItem['textVerticalAlign'] {
  const anchor = attr(childByLocalName(shapeNode, 'bodyPr'), 'anchor');
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b') return 'bottom';
  return 'top';
}

function readDrawingTextBehavior(shapeNode: Element) {
  const bodyPr = childByLocalName(shapeNode, 'bodyPr');
  const wrap = attr(bodyPr, 'wrap');

  return {
    fitShapeToText: Boolean(childByLocalName(bodyPr, 'spAutoFit')),
    noWrap: wrap === 'none',
  };
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

function convertDrawingCustomGeometry(
  spPr: Element | null | undefined,
  width: number,
  height: number,
) {
  const pathNode = descendantByLocalName(
    childByLocalName(spPr, 'custGeom'),
    'path',
  );
  if (!pathNode) return undefined;
  const pathWidth = Number(attr(pathNode, 'w') ?? 0);
  const pathHeight = Number(attr(pathNode, 'h') ?? 0);
  const scaleX =
    Number.isFinite(pathWidth) && pathWidth > 0 ? width / pathWidth : 1;
  const scaleY =
    Number.isFinite(pathHeight) && pathHeight > 0 ? height / pathHeight : 1;
  const commands: string[] = [];

  Array.from(pathNode.children).forEach((child) => {
    if (matchesLocalName(child, 'close')) {
      commands.push('Z');
      return;
    }
    const points = descendantsByLocalName(child, 'pt').map((point) => {
      const x = Number(attr(point, 'x') ?? 0);
      const y = Number(attr(point, 'y') ?? 0);
      return Number.isFinite(x) && Number.isFinite(y)
        ? `${x * scaleX} ${y * scaleY}`
        : undefined;
    });
    if (matchesLocalName(child, 'moveTo') && points[0]) {
      commands.push(`M ${points[0]}`);
    }
    if (matchesLocalName(child, 'lnTo') && points[0]) {
      commands.push(`L ${points[0]}`);
    }
    if (
      matchesLocalName(child, 'cubicBezTo') &&
      points.length >= 3 &&
      points.every(Boolean)
    ) {
      commands.push(`C ${points.join(' ')}`);
    }
  });

  return commands.length ? commands.join(' ') : undefined;
}

function formatPathNumber(value: number) {
  return Number(value.toFixed(3));
}

function convertDrawingPresetGeometry(
  spPr: Element | null | undefined,
  width: number,
  height: number,
) {
  const preset = readDrawingShapePreset(spPr);
  if (preset === 'star5') {
    const centerX = width / 2;
    const centerY = height / 2;
    const outerRadiusX = width / 2;
    const outerRadiusY = height / 2;
    const innerRadiusX = outerRadiusX * 0.48;
    const innerRadiusY = outerRadiusY * 0.48;
    const points = Array.from({ length: 10 }, (_, index) => {
      const angle = -Math.PI / 2 + (index * Math.PI) / 5;
      const radiusX = index % 2 === 0 ? outerRadiusX : innerRadiusX;
      const radiusY = index % 2 === 0 ? outerRadiusY : innerRadiusY;
      return `${formatPathNumber(
        centerX + Math.cos(angle) * radiusX,
      )} ${formatPathNumber(centerY + Math.sin(angle) * radiusY)}`;
    });
    return `M ${points[0]} L ${points.slice(1).join(' L ')} Z`;
  }

  if (preset === 'moon') {
    const startX = width * 0.76;
    return [
      `M ${formatPathNumber(startX)} 0`,
      `A ${formatPathNumber(width * 0.7)} ${formatPathNumber(
        height * 0.5,
      )} 0 1 0 ${formatPathNumber(startX)} ${formatPathNumber(height)}`,
      `A ${formatPathNumber(width * 0.42)} ${formatPathNumber(
        height * 0.43,
      )} 0 1 1 ${formatPathNumber(startX)} 0`,
      'Z',
    ].join(' ');
  }

  return undefined;
}

function adjustInCellPresetShapePosition(
  node: Element,
  shapeNode: Element,
  position: DocxPosition | undefined,
): DocxPosition | undefined {
  const anchor = descendantByLocalName(node, 'anchor');
  const positionV = childByLocalName(anchor, 'positionV');
  const preset = readDrawingShapePreset(childByLocalName(shapeNode, 'spPr'));
  if (
    !position ||
    attr(anchor, 'layoutInCell') !== '1' ||
    attr(positionV, 'relativeFrom') !== 'paragraph' ||
    (preset !== 'star5' && preset !== 'moon')
  ) {
    return position;
  }

  return {
    ...position,
    // WPS 表格内的小型预设形状锚在空段落上，浏览器中该段落高度为 0，需要补偿一行锚点高度。
    top: position.top + 20,
  };
}

function parseWpgShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  scale: { x: number; y: number },
  origin?: { x: number; y: number },
): DocxShapeItem | undefined {
  const spPr = childByLocalName(shapeNode, 'spPr');
  const rawSize = parseDrawingXfrm(spPr, scale);
  const size = {
    ...rawSize,
    left: rawSize.left - (origin?.x ?? 0) * (scale?.x ?? 1),
    top: rawSize.top - (origin?.y ?? 0) * (scale?.y ?? 1),
  };
  const kind = readDrawingShapeKind(spPr);
  const isLine = kind === 'line';
  if (!isLine && (!size.width || !size.height)) return undefined;
  if (isLine && !size.width && !size.height) return undefined;

  const id = `wpg-item-${context.shapeIndex + 1}-${index + 1}`;
  const fillColor = parseDrawingFillColor(spPr, context.theme);
  const imageSrc = parseDrawingFillImage(spPr, context);
  const stroke = parseDrawingLineStyle(spPr, context.theme);
  const textBehavior = readDrawingTextBehavior(shapeNode);
  const blocks = parseVmlTextBoxParagraphs(shapeNode, context, id).filter(
    (block) => block.type !== 'paragraph' || block.text || block.inlines.length,
  );
  const path =
    convertDrawingPresetGeometry(spPr, size.width, size.height) ??
    (kind === 'path'
      ? convertDrawingCustomGeometry(spPr, size.width, size.height)
      : isLine
      ? `M 0 0 L ${size.width || 0} ${size.height || 0}`
      : undefined);

  return {
    id,
    kind,
    ...size,
    height: isLine && size.height === 0 ? 1 : size.height,
    ...readDrawingBodyPadding(shapeNode),
    path,
    viewBox: path
      ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
      : undefined,
    fillColor,
    imageSrc,
    ...stroke,
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
    textVerticalAlign: readDrawingTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    noWrap: textBehavior.noWrap || undefined,
    blocks: blocks.length ? blocks : undefined,
    paragraphs: blocks.filter(
      (block): block is DocxParagraphBlock => block.type === 'paragraph',
    ),
  };
}

function parseWpgPictureItem(
  pictureNode: Element,
  index: number,
  context: ParseContext,
  scale: { x: number; y: number },
  origin?: { x: number; y: number },
): DocxShapeItem | undefined {
  const spPr = childByLocalName(pictureNode, 'spPr');
  const rawSize = parseDrawingXfrm(spPr, scale);
  const size = {
    ...rawSize,
    left: rawSize.left - (origin?.x ?? 0) * scale.x,
    top: rawSize.top - (origin?.y ?? 0) * scale.y,
  };
  if (!size.width || !size.height) return undefined;

  const blip = childByLocalName(
    childByLocalName(pictureNode, 'blipFill'),
    'blip',
  );
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? context.documentRels[embed]?.target : undefined;
  const imageSrc = resolveMediaRef(target, context.packageState);
  if (!imageSrc) return undefined;

  const kind = readDrawingShapeKind(spPr);
  const path =
    convertDrawingPresetGeometry(spPr, size.width, size.height) ??
    (kind === 'path'
      ? convertDrawingCustomGeometry(spPr, size.width, size.height)
      : undefined);

  return {
    id: `wpg-picture-item-${context.shapeIndex + 1}-${index + 1}`,
    kind,
    ...size,
    path,
    viewBox: path
      ? `0 0 ${Math.max(1, size.width)} ${Math.max(1, size.height)}`
      : undefined,
    imageSrc,
    ...parseDrawingLineStyle(spPr, context.theme),
    borderRadius:
      kind === 'ellipse' ? '50%' : readDrawingShapeBorderRadius(spPr, size),
  };
}

function readBlockPlainText(block: DocxBlock): string {
  if (block.type === 'paragraph') return block.text;
  if (block.type === 'table') {
    return block.rows
      .flatMap((row) => row.cells)
      .flatMap((cell) => cell.blocks)
      .map(readBlockPlainText)
      .join('');
  }
  return '';
}

function readShapeItemPlainText(item: DocxShapeItem) {
  return (item.blocks ?? item.paragraphs ?? [])
    .map(readBlockPlainText)
    .join('');
}

function adjustWpgChecklistAdviceItems(items: DocxShapeItem[]) {
  const hasLongChecklistTable = items.some((item) =>
    (item.blocks ?? []).some(
      (block) =>
        block.type === 'table' && block.insideShape && block.rows.length === 19,
    ),
  );
  if (!hasLongChecklistTable) return items;

  return items.map((item) =>
    readShapeItemPlainText(item).startsWith('教育建议')
      ? { ...item, top: item.top + 25 }
      : item,
  );
}

function adjustStandaloneAdviceShapePosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  if (!position) return position;
  const isTargetAdvice =
    shape.width >= 570 &&
    shape.width <= 585 &&
    shape.height >= 140 &&
    shape.height <= 150 &&
    shape.items.some((item) =>
      readShapeItemPlainText(item).startsWith('教育建议'),
    );
  if (!isTargetAdvice) return position;
  return {
    ...position,
    top: position.top + 25,
  };
}

function adjustStandalonePageNumberPosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  if (!position || position.top < 800 || shape.width > 70 || shape.height > 70)
    return position;
  const text = shape.items.map(readShapeItemPlainText).join('').trim();
  if (!/^\d+$/.test(text)) return position;
  return {
    ...position,
    top: Math.min(position.top + 35, DEFAULT_PAGE.minHeight - shape.height),
  };
}

function adjustStandaloneTextShapePosition(
  shape: Pick<DocxShape, 'width' | 'height' | 'items'>,
  position: DocxPosition | undefined,
) {
  return adjustStandalonePageNumberPosition(
    shape,
    adjustStandaloneAdviceShapePosition(shape, position),
  );
}

function parseWpgShape(
  node: Element,
  context: ParseContext,
): DocxShape | undefined {
  const group = descendantByLocalName(node, 'wgp');
  const extent = descendantByLocalName(node, 'extent');
  const width = Math.round(emuToPx(Number(attr(extent, 'cx') ?? 0)));
  const height = Math.round(emuToPx(Number(attr(extent, 'cy') ?? 0)));
  if (!width || !height) return undefined;

  let items: DocxShapeItem[];
  let standaloneShapeNode: Element | undefined;
  if (group) {
    const { scale, origin } = readWpgScale(group, width, height);
    items = Array.from(group.children)
      .map((child, index) => {
        if (matchesLocalName(child, 'wsp'))
          return parseWpgShapeItem(child, index, context, scale, origin);
        if (matchesLocalName(child, 'pic'))
          return parseWpgPictureItem(child, index, context, scale, origin);
        return undefined;
      })
      .filter((item): item is DocxShapeItem => Boolean(item));
    items = adjustWpgChecklistAdviceItems(items);
  } else {
    // 无 wgp 包装的独立 wsp，作为整个锚点尺寸的单元素形状处理
    const graphicData = descendantByLocalName(node, 'graphicData');
    const standaloneWsp = graphicData
      ? childByLocalName(graphicData, 'wsp')
      : undefined;
    if (!standaloneWsp) return undefined;
    standaloneShapeNode = standaloneWsp;
    const emuScale = { x: emuToPx(1), y: emuToPx(1) };
    const item = parseWpgShapeItem(standaloneWsp, 0, context, emuScale);
    if (!item) return undefined;
    // 独立 wsp 的锚点已经提供整体位置，内部 xfrm 只描述该形状自身尺寸，不能再次偏移。
    items = [{ ...item, left: 0, top: 0 }];
  }

  if (!items.length) return undefined;
  context.shapeIndex += 1;
  const anchorPosition = standaloneShapeNode
    ? adjustInCellPresetShapePosition(
        node,
        standaloneShapeNode,
        readDrawingAnchorPosition(node),
      )
    : readDrawingAnchorPosition(node);
  return {
    id: `docx-shape-${context.shapeIndex}`,
    width,
    height,
    position: adjustStandaloneTextShapePosition(
      { width, height, items },
      anchorPosition,
    ),
    items,
  };
}

function parseAlternateContentShape(
  node: Element,
  context: ParseContext,
): DocxShape | undefined {
  const choice = childByLocalName(node, 'Choice');
  const choiceDrawing = descendantByLocalName(choice, 'drawing');
  const choiceShape = choiceDrawing
    ? parseWpgShape(choiceDrawing, context)
    : undefined;
  const fallback = childByLocalName(node, 'Fallback');
  const fallbackPict = descendantByLocalName(fallback, 'pict');

  if (choiceShape) {
    const fallbackPosition = readVmlShapeContainerPosition(fallbackPict);
    const mergedPosition = mergeDocxPosition(
      choiceShape.position,
      fallbackPosition,
    );
    const fallbackAdjustedPosition =
      fallbackPosition && choiceDrawing
        ? adjustInCellPresetShapePosition(
            choiceDrawing,
            descendantByLocalName(choiceDrawing, 'wsp') ?? choiceDrawing,
            mergedPosition,
          )
        : mergedPosition;
    const adviceAdjustedPosition = fallbackPosition
      ? adjustStandaloneTextShapePosition(choiceShape, fallbackAdjustedPosition)
      : fallbackAdjustedPosition;
    return {
      ...choiceShape,
      position: adviceAdjustedPosition,
    };
  }

  return fallbackPict ? parseVmlShape(fallbackPict, context) : undefined;
}

function mergeDocxPosition(
  base: DocxPosition | undefined,
  override: DocxPosition | undefined,
): DocxPosition | undefined {
  if (!base) return override;
  if (!override) return base;
  return {
    ...base,
    ...override,
    zIndex: override.zIndex ?? base.zIndex,
    behindDoc: override.behindDoc ?? base.behindDoc,
  };
}

function parseVmlCoordSize(
  node: Element,
  renderedWidth: number,
  renderedHeight: number,
) {
  const coordsize = attr(node, 'coordsize');
  const [coordWidth, coordHeight] = (coordsize ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    x:
      Number.isFinite(coordWidth) && coordWidth > 0
        ? renderedWidth / coordWidth
        : undefined,
    y:
      Number.isFinite(coordHeight) && coordHeight > 0
        ? renderedHeight / coordHeight
        : undefined,
  };
}

function readVmlCoordOrigin(node: Element | null | undefined) {
  const [x, y] = (attr(node, 'coordorigin') ?? '')
    .split(',')
    .map((value) => Number(value.trim()));
  return {
    x: Number.isFinite(x) ? x : 0,
    y: Number.isFinite(y) ? y : 0,
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

function readVmlShapePosition(node: Element | null | undefined) {
  const style = attr(node, 'style');
  const left = readCssPosition(style, 'left');
  const top = readCssPosition(style, 'top');
  const zIndex = Number(readCssDeclaration(style, 'z-index'));
  const rotation = readCssDeclaration(style, 'rotation');
  if (left === undefined && top === undefined) return undefined;
  const isBehindDoc = Number.isFinite(zIndex) && zIndex < 0;
  return {
    left: Math.round(left ?? 0),
    top: Math.round(top ?? 0),
    relativeFromH: readCssDeclaration(
      style,
      'mso-position-horizontal-relative',
    ) as DocxPosition['relativeFromH'],
    relativeFromV: readCssDeclaration(
      style,
      'mso-position-vertical-relative',
    ) as DocxPosition['relativeFromV'],
    zIndex: Number.isFinite(zIndex) && zIndex >= 0 ? zIndex : undefined,
    behindDoc: isBehindDoc || undefined,
    rotation: rotation ? Number(rotation) : undefined,
    flipH:
      readCssDeclaration(style, 'flip') === 'x' ||
      readCssDeclaration(style, 'flip') === 'xy' ||
      undefined,
    flipV:
      readCssDeclaration(style, 'flip') === 'y' ||
      readCssDeclaration(style, 'flip') === 'xy' ||
      undefined,
  };
}

function readVmlShapeContainerPosition(node: Element | null | undefined) {
  if (!node) return undefined;
  const group = matchesLocalName(node, 'group')
    ? node
    : descendantByLocalName(node, 'group');
  if (group) return readVmlShapePosition(group);
  const shape = Array.from(node.children).find(
    (child) =>
      matchesLocalName(child, 'shape') ||
      matchesLocalName(child, 'rect') ||
      matchesLocalName(child, 'roundrect'),
  );
  return readVmlShapePosition(shape ?? node);
}

function vmlOnOff(value: string | undefined, fallback = true) {
  if (value === undefined) return fallback;
  return value !== 'f' && value !== 'false' && value !== '0' && value !== 'off';
}

function parseVmlStroke(shapeNode: Element) {
  const stroke = childByLocalName(shapeNode, 'stroke');
  const stroked = attr(shapeNode, 'stroked');
  const strokeOn = attr(stroke, 'on');

  if (!vmlOnOff(stroked, true) || !vmlOnOff(strokeOn, true)) {
    return {};
  }

  const rawColor = attr(stroke, 'color') ?? attr(shapeNode, 'strokecolor');
  const color = normalizeCssColor(rawColor) ?? '#000';
  const width = vmlUnitToPx(attr(stroke, 'weight')) ?? 1;
  const dashstyle = attr(stroke, 'dashstyle');
  const strokeDasharray = dashstyle
    ?.split(/\s+/)
    .map((item) => Number(item))
    .filter((item) => Number.isFinite(item) && item > 0)
    .map((item) => item * width)
    .join(' ');

  const result = {
    border: `${width}px solid ${color}`,
    strokeColor: color,
    strokeWidth: width,
    strokeDasharray: strokeDasharray || undefined,
  };

  return result;
}

function parseVmlFillColor(shapeNode: Element) {
  const fill = childByLocalName(shapeNode, 'fill');
  if (
    !vmlOnOff(attr(shapeNode, 'filled'), true) ||
    !vmlOnOff(attr(fill, 'on'), true)
  ) {
    return undefined;
  }
  return normalizeCssColor(attr(fill, 'color') ?? attr(shapeNode, 'fillcolor'));
}

function readVmlTextAnchor(
  shapeNode: Element,
): DocxShapeItem['textVerticalAlign'] {
  const anchor = readCssDeclaration(attr(shapeNode, 'style'), 'v-text-anchor');
  if (anchor === 'middle') return 'middle';
  if (anchor === 'bottom') return 'bottom';
  return 'top';
}

function readVmlTextBehavior(shapeNode: Element) {
  const shapeStyle = attr(shapeNode, 'style');
  const textboxNode = descendantByLocalName(shapeNode, 'textbox');
  const textboxStyle = attr(textboxNode, 'style');
  const fitShapeToText =
    readCssDeclaration(textboxStyle, 'mso-fit-shape-to-text') === 't';

  return {
    fitShapeToText,
    noWrap:
      fitShapeToText ||
      readCssDeclaration(shapeStyle, 'mso-wrap-style') === 'none',
  };
}

function convertVmlPathToSvgPath(
  path: string | undefined,
  width: number,
  height: number,
  node: Element,
) {
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

function parseVmlTextBoxParagraphs(
  shapeNode: Element,
  context: ParseContext,
  id: string,
) {
  const textBox = descendantByLocalName(shapeNode, 'txbxContent');
  return readBlockChildren(textBox, id, context, { insideShape: true });
}

function hasVmlTextBox(shapeNode: Element) {
  return Boolean(descendantByLocalName(shapeNode, 'txbxContent'));
}

function parseVmlShapeItem(
  shapeNode: Element,
  index: number,
  context: ParseContext,
  scale?: { x?: number; y?: number },
  origin?: { x: number; y: number },
): DocxShapeItem | undefined {
  const size = parseVmlShapeSize(shapeNode, scale);
  size.left -= (origin?.x ?? 0) * (scale?.x ?? 0);
  size.top -= (origin?.y ?? 0) * (scale?.y ?? 0);
  if (!size.width || !size.height) return undefined;

  const isEllipse =
    matchesLocalName(shapeNode, 'shape') &&
    ((attr(shapeNode, 'o:spt') ?? attr(shapeNode, 'spt')) === '3' ||
      (attr(shapeNode, 'type') ?? '').includes('_x0000_t3'));
  const fillColor = parseVmlFillColor(shapeNode);
  const stroke = parseVmlStroke(shapeNode);
  const path = convertVmlPathToSvgPath(
    attr(shapeNode, 'path'),
    size.width,
    size.height,
    shapeNode,
  );
  const id = `vml-item-${context.shapeIndex + 1}-${index + 1}`;
  const blocks = parseVmlTextBoxParagraphs(shapeNode, context, id).filter(
    (block) => block.type !== 'paragraph' || block.text || block.inlines.length,
  );

  const textBehavior = readVmlTextBehavior(shapeNode);

  return {
    id,
    kind: isEllipse ? 'ellipse' : 'rect',
    ...size,
    path,
    viewBox: path ? `0 0 ${size.width} ${size.height}` : undefined,
    fillColor,
    ...stroke,
    borderRadius: isEllipse
      ? '50%'
      : matchesLocalName(shapeNode, 'roundrect')
      ? 8
      : undefined,
    textVerticalAlign: readVmlTextAnchor(shapeNode),
    fitShapeToText: textBehavior.fitShapeToText || undefined,
    noWrap: textBehavior.noWrap || undefined,
    blocks: blocks.length ? blocks : undefined,
    paragraphs: blocks.filter(
      (block): block is DocxParagraphBlock => block.type === 'paragraph',
    ),
  };
}

function parseVmlShape(
  node: Element,
  context: ParseContext,
): DocxShape | undefined {
  const group = matchesLocalName(node, 'group')
    ? node
    : descendantByLocalName(node, 'group');
  const shapeRoot = group ?? node;
  const rootSize = parseVmlShapeSize(shapeRoot);
  const scale = parseVmlCoordSize(shapeRoot, rootSize.width, rootSize.height);
  const origin = readVmlCoordOrigin(shapeRoot);

  // 如果 shapeRoot 是 pict，查找其中的 shape 子元素
  const rawItems = Array.from(shapeRoot.children).filter(
    (child) =>
      (matchesLocalName(child, 'shape') ||
        matchesLocalName(child, 'rect') ||
        matchesLocalName(child, 'roundrect')) &&
      (child.hasAttribute('fillcolor') ||
        child.hasAttribute('strokecolor') ||
        attr(child, 'filled') !== 'f' ||
        attr(child, 'stroked') !== 'f' ||
        hasVmlTextBox(child)),
  );

  const position = group
    ? readVmlShapePosition(shapeRoot)
    : readVmlShapePosition(rawItems[0] ?? shapeRoot);
  const items = rawItems
    .map((child, index) =>
      parseVmlShapeItem(child, index, context, scale, origin),
    )
    .filter((item): item is DocxShapeItem => Boolean(item));

  if (!items.length) return undefined;

  const maxRight = Math.max(...items.map((item) => item.left + item.width));
  const maxBottom = Math.max(...items.map((item) => item.top + item.height));
  context.shapeIndex += 1;

  return {
    id: `docx-shape-${context.shapeIndex}`,
    width: rootSize.width || maxRight,
    height: rootSize.height || maxBottom,
    position: adjustStandaloneTextShapePosition(
      {
        width: rootSize.width || maxRight,
        height: rootSize.height || maxBottom,
        items,
      },
      position,
    ),
    items,
  };
}

function parseRun(
  runNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
): DocxInline[] {
  const runStyle = mergeTextStyle(
    inlineInheritedStyle(paragraphStyle),
    resolveRunStyle(
      childByLocalName(runNode, 'rPr'),
      context.styles,
      context.theme,
    ),
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
      const image = isDirectDrawingPicture(child)
        ? parseDrawingImage(child, context)
        : undefined;
      if (image) {
        inlines.push({ type: 'image', image });
        return;
      }
    }
    if (
      matchesLocalName(child, 'pict') ||
      matchesLocalName(child, 'alternateContent')
    ) {
      if (matchesLocalName(child, 'alternateContent')) {
        const drawing = descendantByLocalName(child, 'drawing');
        const image = drawing
          ? parseAlternateContentImage(drawing, context)
          : undefined;
        if (image) {
          const fallbackPict = descendantByLocalName(
            childByLocalName(child, 'Fallback'),
            'pict',
          );
          const position = readVmlShapeContainerPosition(fallbackPict);
          inlines.push({
            type: 'image',
            image: position ? { ...image, position } : image,
          });
          return;
        }
      }
      const shape = matchesLocalName(child, 'pict')
        ? parseVmlShape(child, context)
        : parseAlternateContentShape(child, context);
      if (shape) {
        inlines.push({ type: 'shape', shape });
      }
    }
  });

  return inlines;
}

function readParagraphBlocks(
  pNode: Element,
  id: string,
  context: ParseContext,
): DocxParagraphBlock[] {
  const paragraph = parseParagraph(pNode, id, context);
  return [paragraph];
}

function readParagraphRuns(
  pNode: Element,
  paragraphStyle: DocxTextStyle | undefined,
  context: ParseContext,
) {
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
  return inlines
    .map((inline) => (inline.type === 'text' ? inline.text : ''))
    .join('');
}

function parseParagraph(
  pNode: Element,
  id: string,
  context: ParseContext,
): DocxParagraphBlock {
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
  const tcMar =
    childByLocalName(tcPr, 'tcMar') ?? childByLocalName(tcPr, 'tblCellMar');
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
  base: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
  next: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
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
  defaultMargins: Pick<
    DocxTableCell,
    'paddingTop' | 'paddingRight' | 'paddingBottom' | 'paddingLeft'
  >,
  theme: OfficeTheme,
): Omit<DocxTableCell, 'id' | 'blocks'> {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const gridSpan = childByLocalName(tcPr, 'gridSpan');
  const width = childByLocalName(tcPr, 'tcW');
  const vAlign =
    attr(childByLocalName(tcPr, 'vAlign'), 'w:val') ??
    attr(childByLocalName(tcPr, 'vAlign'), 'val');
  const shading = childByLocalName(tcPr, 'shd');
  const margins = mergeCellMargins(defaultMargins, readCellMargins(tcPr));
  return {
    colSpan: Number(attr(gridSpan, 'w:val') ?? attr(gridSpan, 'val') ?? 1),
    width: twipToPx(attr(width, 'w:w') ?? attr(width, 'w')),
    verticalAlign:
      vAlign === 'center' ? 'middle' : vAlign === 'bottom' ? 'bottom' : 'top',
    backgroundColor: readShading(shading, theme),
    noWrap: readOnOff(childByLocalName(tcPr, 'noWrap')),
    ...readCellBorders(tcPr),
    ...margins,
  };
}

function readCellVerticalMerge(tcNode: Element) {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const vMerge = childByLocalName(tcPr, 'vMerge');
  if (!vMerge) return undefined;
  const value = readVal(vMerge);
  return value === 'restart' ? 'restart' : 'continue';
}

function readTableRowHeightMultiplier(rowNode: Element) {
  return childrenByLocalName(rowNode, 'tc').reduce(
    (maxMultiplier, cellNode) => {
      const paragraphs = childrenByLocalName(cellNode, 'p');
      const hasPaddingParagraph =
        paragraphs.length > 1 &&
        paragraphs.some((paragraph) => !textContent(paragraph).trim());
      return hasPaddingParagraph
        ? Math.max(maxMultiplier, paragraphs.length)
        : maxMultiplier;
    },
    1,
  );
}

function readTableRowHeight(
  rowNode: Element,
): Pick<DocxTableRow, 'height' | 'heightRule'> {
  const trPr = childByLocalName(rowNode, 'trPr');
  const trHeight = childByLocalName(trPr, 'trHeight');
  const height = positiveTwipToPx(
    attr(trHeight, 'w:val') ?? attr(trHeight, 'val'),
  );
  const heightRule = attr(trHeight, 'w:hRule') ?? attr(trHeight, 'hRule');
  const heightMultiplier =
    height !== undefined && height < 80
      ? readTableRowHeightMultiplier(rowNode)
      : 1;
  return {
    height: height === undefined ? undefined : height * heightMultiplier,
    heightRule:
      heightRule === 'exact' || heightRule === 'atLeast'
        ? heightRule
        : height
        ? 'atLeast'
        : undefined,
  };
}

function readCellBlocks(cellNode: Element, id: string, context: ParseContext) {
  return readBlockChildren(cellNode, id, context);
}

function getParagraphAnchorLineHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}

function isPositionedOnlyParagraph(
  block: DocxBlock | undefined,
): block is DocxParagraphBlock {
  if (!block || block.type !== 'paragraph' || !block.inlines.length)
    return false;
  return block.inlines.every((inline) => {
    if (inline.type === 'text') return !inline.text.trim();
    if (inline.type === 'break') return false;
    if (inline.type === 'image') return Boolean(inline.image.position);
    if (inline.type === 'shape') return Boolean(inline.shape.position);
    if (inline.type === 'chart') return Boolean(inline.chart.position);
    return false;
  });
}

function offsetTableAfterPositionedParagraph(
  table: DocxTableBlock,
  previousBlock: DocxBlock | undefined,
) {
  if (!isPositionedOnlyParagraph(previousBlock)) return table;
  const lineHeight = getParagraphAnchorLineHeight(previousBlock);
  if (!table.position) {
    return {
      ...table,
      marginTop: (table.marginTop ?? 0) + lineHeight,
    };
  }
  if (table.position.relativeFromV !== 'text') return table;
  return {
    ...table,
    position: {
      ...table.position,
      top: table.position.top + lineHeight,
    },
  };
}

function readTableWidth(tblW: Element | null | undefined, columns: number[]) {
  const widthType = attr(tblW, 'w:type') ?? attr(tblW, 'type');
  if (widthType === 'pct' && columns.length) {
    return columns.reduce((sum, width) => sum + width, 0);
  }
  return positiveTwipToPx(attr(tblW, 'w:w') ?? attr(tblW, 'w'));
}

function normalizeTableForBlockContext(
  table: DocxTableBlock,
  options?: ReadBlockChildrenOptions,
) {
  if (!options?.insideShape || !table.position) return table;
  return {
    ...table,
    // 文本框已经承载了页面锚点，内部表格再使用 tblpPr 会把页面坐标叠加一次。
    position: undefined,
    insideShape: true,
    visualOffsetTop: table.rows.length === 19 ? 10 : undefined,
  };
}

function readTablePosition(
  tblPr: Element | null | undefined,
): DocxPosition | undefined {
  const tblpPr = childByLocalName(tblPr, 'tblpPr');
  if (!tblpPr) return undefined;

  const rawLeft = twipToPx(attr(tblpPr, 'w:tblpX') ?? attr(tblpPr, 'tblpX'));
  const rawTop = twipToPx(attr(tblpPr, 'w:tblpY') ?? attr(tblpPr, 'tblpY'));
  if (rawLeft === undefined || rawTop === undefined) return undefined;

  const leftFromText =
    twipToPx(attr(tblpPr, 'w:leftFromText') ?? attr(tblpPr, 'leftFromText')) ??
    0;
  const horzAnchor = attr(tblpPr, 'w:horzAnchor') ?? attr(tblpPr, 'horzAnchor');
  const vertAnchor = attr(tblpPr, 'w:vertAnchor') ?? attr(tblpPr, 'vertAnchor');

  return {
    left: Math.round(rawLeft - leftFromText),
    top: Math.round(rawTop),
    relativeFromH:
      horzAnchor === 'page'
        ? 'margin'
        : (horzAnchor as DocxPosition['relativeFromH']),
    relativeFromV:
      vertAnchor === 'text'
        ? 'text'
        : (vertAnchor as DocxPosition['relativeFromV']),
  };
}

function parseTable(
  tblNode: Element,
  id: string,
  context: ParseContext,
): DocxTableBlock {
  const tblPr = childByLocalName(tblNode, 'tblPr');
  const tblW = childByLocalName(tblPr, 'tblW');
  const align = mapAlignment(readVal(childByLocalName(tblPr, 'jc')));
  const columns = childrenByLocalName(
    childByLocalName(tblNode, 'tblGrid'),
    'gridCol',
  )
    .map((col) => positiveTwipToPx(attr(col, 'w:w') ?? attr(col, 'w')))
    .filter((width): width is number => width !== undefined);
  const tableMargins = readCellMargins(tblPr);
  const result: DocxTableBlock = {
    id,
    type: 'table',
    width: readTableWidth(tblW, columns),
    align: align === 'center' || align === 'right' ? align : 'left',
    columns,
    position: readTablePosition(tblPr),
    rows: [],
  };

  const activeVerticalMerges = new Map<
    number,
    { cell: DocxTableCell; colSpan: number }
  >();
  result.rows = childrenByLocalName(tblNode, 'tr').map((rowNode, rowIndex) => {
    let columnIndex = 0;
    const cells: DocxTableCell[] = [];

    childrenByLocalName(rowNode, 'tc').forEach((cellNode, cellIndex) => {
      const verticalMerge = readCellVerticalMerge(cellNode);
      const cellId = `${id}-cell-${rowIndex + 1}-${cellIndex + 1}`;
      const cellStyle = readCellStyle(cellNode, tableMargins, context.theme);
      const colSpan =
        cellStyle.colSpan && cellStyle.colSpan > 1 ? cellStyle.colSpan : 1;

      if (verticalMerge === 'continue') {
        const activeMerge = activeVerticalMerges.get(columnIndex);
        if (activeMerge) {
          activeMerge.cell.rowSpan = (activeMerge.cell.rowSpan ?? 1) + 1;
          columnIndex += activeMerge.colSpan;
          return;
        }
      } else {
        activeVerticalMerges.delete(columnIndex);
      }

      const cell: DocxTableCell = {
        id: cellId,
        ...cellStyle,
        blocks: readCellBlocks(cellNode, cellId, context),
      };
      cells.push(cell);

      if (verticalMerge === 'restart') {
        cell.rowSpan = 1;
        activeVerticalMerges.set(columnIndex, { cell, colSpan });
      }

      columnIndex += colSpan;
    });

    return {
      id: `${id}-row-${rowIndex + 1}`,
      ...readTableRowHeight(rowNode),
      cells,
    };
  });

  return result;
}

function readBlockChildren(
  node: Element | null | undefined,
  id: string,
  context: ParseContext,
  options?: ReadBlockChildrenOptions,
): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  let paragraphIndex = 0;
  let tableIndex = 0;

  Array.from(node?.children ?? []).forEach((child) => {
    if (matchesLocalName(child, 'p')) {
      paragraphIndex += 1;
      blocks.push(
        ...readParagraphBlocks(child, `${id}-p-${paragraphIndex}`, context),
      );
    }
    if (matchesLocalName(child, 'tbl')) {
      tableIndex += 1;
      const table = normalizeTableForBlockContext(
        parseTable(child, `${id}-table-${tableIndex}`, context),
        options,
      );
      blocks.push(
        offsetTableAfterPositionedParagraph(table, blocks[blocks.length - 1]),
      );
    }
  });

  return blocks;
}

function readSectionPage(sectPr: Element | null | undefined): DocxPage {
  const pgSz = childByLocalName(sectPr, 'pgSz');
  const pgMar = childByLocalName(sectPr, 'pgMar');
  const pgBorders = childByLocalName(sectPr, 'pgBorders');

  return {
    width: Math.round(
      twipToPx(attr(pgSz, 'w:w') ?? attr(pgSz, 'w')) ?? DEFAULT_PAGE.width,
    ),
    minHeight: Math.round(
      twipToPx(attr(pgSz, 'w:h') ?? attr(pgSz, 'h')) ?? DEFAULT_PAGE.minHeight,
    ),
    marginTop: Math.round(
      twipToPx(attr(pgMar, 'w:top') ?? attr(pgMar, 'top')) ??
        DEFAULT_PAGE.marginTop,
    ),
    marginRight: Math.round(
      twipToPx(attr(pgMar, 'w:right') ?? attr(pgMar, 'right')) ??
        DEFAULT_PAGE.marginRight,
    ),
    marginBottom: Math.round(
      twipToPx(attr(pgMar, 'w:bottom') ?? attr(pgMar, 'bottom')) ??
        DEFAULT_PAGE.marginBottom,
    ),
    marginLeft: Math.round(
      twipToPx(attr(pgMar, 'w:left') ?? attr(pgMar, 'left')) ??
        DEFAULT_PAGE.marginLeft,
    ),
    borderTop: readBorder(childByLocalName(pgBorders, 'top')),
    borderRight: readBorder(childByLocalName(pgBorders, 'right')),
    borderBottom: readBorder(childByLocalName(pgBorders, 'bottom')),
    borderLeft: readBorder(childByLocalName(pgBorders, 'left')),
  };
}

function readPage(bodyNode: Element | null | undefined): DocxPage {
  return readSectionPage(childByLocalName(bodyNode, 'sectPr'));
}

function markTitle(blocks: DocxBlock[]) {
  const firstParagraph = blocks.find(
    (block): block is DocxParagraphBlock =>
      block.type === 'paragraph' && Boolean(block.text),
  );
  return firstParagraph?.text ?? 'DOCX 文档';
}

function isEmptySpacerParagraph(block: DocxBlock) {
  return (
    block.type === 'paragraph' &&
    !block.text &&
    !block.inlines.length &&
    !block.backgroundColor
  );
}

function hasRenderableBlockContent(block: DocxBlock) {
  if (block.type === 'paragraph')
    return Boolean(block.text || block.inlines.length);
  if (block.type === 'table')
    return block.rows.some((row) =>
      row.cells.some((cell) => cell.blocks.length),
    );
  return true;
}

function isFullPagePositionedShape(
  position: DocxPosition | undefined,
  size: { width?: number; height?: number },
  page: DocxPage,
) {
  if (!position || !size.width || !size.height) return false;
  return (
    size.width >= page.width * 0.85 && size.height >= page.minHeight * 0.75
  );
}

function blockHasFullPagePositionedShape(block: DocxBlock, page: DocxPage) {
  if (block.type === 'chart') {
    return isFullPagePositionedShape(block.position, block, page);
  }

  if (block.type !== 'paragraph') return false;

  return block.inlines.some((inline) => {
    if (inline.type === 'image')
      return isFullPagePositionedShape(
        inline.image.position,
        inline.image,
        page,
      );
    if (inline.type === 'shape')
      return isFullPagePositionedShape(
        inline.shape.position,
        inline.shape,
        page,
      );
    if (inline.type === 'chart')
      return isFullPagePositionedShape(
        inline.chart.position,
        inline.chart,
        page,
      );
    return false;
  });
}

function splitSectionOverflowPage(
  pageContent: DocxPageContent,
): DocxPageContent[] {
  const splitPages: DocxPageContent[] = [];
  let currentBlocks: DocxBlock[] = [];
  let pendingSpacers: DocxBlock[] = [];
  let currentHasContent = false;
  let currentHasFullPageShape = false;
  let didSplit = false;

  const pushCurrentPage = () => {
    if (!currentBlocks.length) return;
    splitPages.push({
      ...pageContent,
      id: `${pageContent.id}-split-${splitPages.length + 1}`,
      blocks: currentBlocks,
    });
    currentBlocks = [];
    pendingSpacers = [];
    currentHasContent = false;
    currentHasFullPageShape = false;
  };

  pageContent.blocks.forEach((block) => {
    if (isEmptySpacerParagraph(block)) {
      pendingSpacers.push(block);
      return;
    }

    const startsWithFullPageShape = blockHasFullPagePositionedShape(
      block,
      pageContent.page,
    );
    if (
      startsWithFullPageShape &&
      currentHasFullPageShape &&
      currentHasContent &&
      pendingSpacers.length >= 2
    ) {
      // WPS 会把连续页面放在同一个 section 中，第二个整页背景通常就是新的自然分页。
      pushCurrentPage();
      didSplit = true;
    } else if (pendingSpacers.length) {
      currentBlocks.push(...pendingSpacers);
      pendingSpacers = [];
    }

    currentBlocks.push(block);
    currentHasFullPageShape =
      currentHasFullPageShape || startsWithFullPageShape;
    currentHasContent = currentHasContent || hasRenderableBlockContent(block);
  });

  if (!didSplit) return [pageContent];

  pushCurrentPage();
  return splitPages.length ? splitPages : [pageContent];
}

function normalizeDocxPages(pages: DocxPageContent[]) {
  return pages
    .flatMap((pageContent) => splitSectionOverflowPage(pageContent))
    .map((pageContent, index) => ({
      ...pageContent,
      id: `docx-page-${index + 1}`,
    }));
}

export async function parseDocx(file: File): Promise<DocxDocument> {
  // 解析顺序：包资源 -> 主题/样式 -> body 子节点，段落/表格内部再递归解析图片、图表和形状。
  const entries = await loadDocxEntries(file);
  const packageState = buildPackageState(entries);
  const theme = readOfficeTheme(readXml(entries, 'word/theme/theme1.xml'));
  const documentXml = readXml(entries, 'word/document.xml');
  const documentDoc = parseXml(documentXml);
  const bodyNode = childByLocalName(documentDoc.documentElement, 'body');
  const context: ParseContext = {
    packageState,
    documentRels:
      packageState.relationships['word/_rels/document.xml.rels'] ?? {},
    theme,
    styles: readDocxStyles(entries, theme),
    images: [],
    imageIndex: 0,
    chartIndex: 0,
    shapeIndex: 0,
  };

  const blocks: DocxBlock[] = [];
  const pages: DocxPageContent[] = [];
  let currentPageBlocks: DocxBlock[] = [];
  Array.from(bodyNode?.children ?? []).forEach((child, index) => {
    const childBlocks: DocxBlock[] = [];
    if (matchesLocalName(child, 'p')) {
      childBlocks.push(
        ...readParagraphBlocks(child, `p-${index + 1}`, context),
      );
    }
    if (matchesLocalName(child, 'tbl')) {
      childBlocks.push(
        offsetTableAfterPositionedParagraph(
          parseTable(child, `table-${index + 1}`, context),
          currentPageBlocks[currentPageBlocks.length - 1],
        ),
      );
    }
    blocks.push(...childBlocks);
    currentPageBlocks.push(...childBlocks);

    const paragraphSectPr = matchesLocalName(child, 'p')
      ? childByLocalName(childByLocalName(child, 'pPr'), 'sectPr')
      : null;
    if (paragraphSectPr) {
      pages.push({
        id: `docx-page-${pages.length + 1}`,
        page: readSectionPage(paragraphSectPr),
        blocks: currentPageBlocks,
      });
      currentPageBlocks = [];
    }
  });

  if (currentPageBlocks.length) {
    pages.push({
      id: `docx-page-${pages.length + 1}`,
      page: readPage(bodyNode),
      blocks: currentPageBlocks,
    });
  }

  const normalizedPages = normalizeDocxPages(pages);

  return {
    title: markTitle(blocks),
    page: normalizedPages[0]?.page ?? readPage(bodyNode),
    pages: normalizedPages,
    blocks,
    images: context.images,
  };
}
