import { loadPptxEntries } from './archive';
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
import { emuToPx } from '../office/units';
import { alphaToOpacity, alphaToRatio, resolveThemeColor, toHexColor, transformColor } from './colors';
import { collectMedia, resolvePackageMediaRef, type OfficeRelationship } from '../office/media';
import { readRelationships } from '../office/relationships';
import { decodeMojibake, parseOfficeChartXml } from '../office/charts';
import type {
  ChartElement,
  GradientFill,
  ImageCrop,
  ImageElement,
  PptxDocument,
  ShapeElement,
  SlideElement,
  SlideModel,
  SlideBackground,
  TableElement,
  ThemeModel,
  TextElement,
  TextParagraph,
  TextRun,
  TextStyle,
  ShadowStyle,
  UnsupportedElement,
} from './types';

type RelationshipMap = Record<string, Record<string, string>>;

type PackageState = {
  entries: OfficeEntryMap;
  relationships: RelationshipMap;
  mediaByName: Record<string, string>;
  mediaByPath: Record<string, string>;
};

type TableStyleVariantName =
  | 'wholeTbl'
  | 'band1H'
  | 'band2H'
  | 'band1V'
  | 'band2V'
  | 'firstRow'
  | 'lastRow'
  | 'firstCol'
  | 'lastCol';

type TableCellStyle = {
  text?: TextStyle;
  backgroundColor?: string | null;
  backgroundOpacity?: number;
  borderColor?: string | null;
  borderOpacity?: number;
  borderWidth?: number;
};

type TableStyleDefinition = {
  styleId: string;
  styleName?: string;
  variants: Partial<Record<TableStyleVariantName, TableCellStyle>>;
};

type TableStyleMap = Record<string, TableStyleDefinition>;

type LayoutDefinition = {
  path: string;
  masterPath: string;
  placeholders: Record<string, PlaceholderStyle>;
  textPresets: Record<string, PlaceholderStyle>;
  background?: SlideBackground;
  elements: SlideElement[];
};

type MasterDefinition = {
  path: string;
  placeholders: Record<string, PlaceholderStyle>;
  textPresets: Record<string, PlaceholderStyle>;
  background?: SlideBackground;
  elements: SlideElement[];
};

type PlaceholderStyle = {
  type?: string;
  idx?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  fill?: string | null;
  fillOpacity?: number;
  stroke?: string | null;
  strokeOpacity?: number;
  strokeWidth?: number;
  strokeDash?: string;
  shadow?: ShadowStyle;
  text?: TextStyle;
  body?: TextStyle;
  levels?: Record<number, TextStyle>;
};

function relationshipTargets(rels: Record<string, OfficeRelationship>) {
  const map: Record<string, string> = {};
  Object.entries(rels).forEach(([id, rel]) => {
    map[id] = rel.target;
  });
  return map;
}

function buildPackageState(entries: OfficeEntryMap): PackageState {
  const relationships: RelationshipMap = {};
  for (const [path, value] of entries) {
    if (typeof value === 'string' && path.endsWith('.rels')) {
      relationships[path] = relationshipTargets(readRelationships(value, path));
    }
  }

  const media = collectMedia(entries, 'ppt/media/');

  return { entries, relationships, mediaByName: media.byName, mediaByPath: media.byPath };
}

export function debugPptxPackage(entries: OfficeEntryMap) {
  const packageState = buildPackageState(entries);
  return {
    relsCount: Object.keys(packageState.relationships).length,
    mediaCount: Object.keys(packageState.mediaByPath).length,
    mediaSample: Object.entries(packageState.mediaByPath)
      .slice(0, 5)
      .map(([path, src]) => ({ path, hasSrc: Boolean(src), prefix: src.slice(0, 30) })),
  };
}

function readPresentationSize(xml: string) {
  const doc = parseXml(xml);
  const sldSz = doc.querySelector('sldSz, p\\:sldSz');
  const cx = Number(attr(sldSz, 'cx') ?? 12800000);
  const cy = Number(attr(sldSz, 'cy') ?? 7200000);
  return { width: emuToPx(cx), height: emuToPx(cy) };
}

function readTheme(xml: string): ThemeModel {
  const doc = parseXml(xml);
  const colorScheme: Record<string, string> = {};
  const fontScheme: Record<string, string> = {};
  const colorMap = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'dk2',
    tx2: 'lt2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
  };

  const colorNode = doc.querySelector('clrScheme, a\\:clrScheme');
  descendantsByLocalName(colorNode, 'dk1')
    .concat(descendantsByLocalName(colorNode, 'lt1'))
    .concat(descendantsByLocalName(colorNode, 'dk2'))
    .concat(descendantsByLocalName(colorNode, 'lt2'))
    .concat(descendantsByLocalName(colorNode, 'accent1'))
    .concat(descendantsByLocalName(colorNode, 'accent2'))
    .concat(descendantsByLocalName(colorNode, 'accent3'))
    .concat(descendantsByLocalName(colorNode, 'accent4'))
    .concat(descendantsByLocalName(colorNode, 'accent5'))
    .concat(descendantsByLocalName(colorNode, 'accent6'))
    .concat(descendantsByLocalName(colorNode, 'hlink'))
    .concat(descendantsByLocalName(colorNode, 'folHlink'))
    .forEach((node) => {
      const child = node.firstElementChild;
      if (!child) return;
      const childName = child.localName.split(':').pop()?.toLowerCase();
      const value = childName === 'sysclr'
        ? attr(child, 'lastClr') ?? attr(child, 'val')
        : attr(child, 'val') ?? attr(child, 'lastClr');
      if (value) colorScheme[node.localName] = value;
    });

  const fontNode = doc.querySelector('fontScheme, a\\:fontScheme');
  ['majorFont', 'minorFont'].forEach((bucket) => {
    const node = childByLocalName(fontNode, bucket);
    if (!node) return;
    const latin = childByLocalName(node, 'latin');
    const ea = childByLocalName(node, 'ea');
    const cs = childByLocalName(node, 'cs');
    fontScheme[bucket] = [attr(latin, 'typeface'), attr(ea, 'typeface'), attr(cs, 'typeface')]
      .filter(Boolean)
      .join(', ');
  });

  return { colorScheme, fontScheme, colorMap };
}

function emuValue(node: Element | null, name: string) {
  const value = attr(node, name);
  return value ? emuToPx(Number(value)) : undefined;
}

function pointToPx(point?: string) {
  if (!point) return undefined;
  const value = Number(point);
  if (!Number.isFinite(value)) return undefined;
  return (value / 100) * (96 / 72);
}

function pctToRatio(value?: string) {
  if (!value) return undefined;
  const next = Number(value);
  if (!Number.isFinite(next)) return undefined;
  return next / 100000;
}

function clamp01(value: number) {
  return Math.max(0, Math.min(1, value));
}

function boolAttr(node: Element | null, name: string) {
  const value = attr(node, name);
  if (value === undefined) return undefined;
  return value === '1' || value === 'true';
}

function mergeTextStyles(...styles: Array<TextStyle | undefined>) {
  return styles.reduce<TextStyle>((acc, style) => {
    if (!style) return acc;
    return {
      ...acc,
      ...Object.fromEntries(
        Object.entries(style).filter(([, value]) => value !== undefined),
      ),
    };
  }, {});
}

function mergePlaceholderStyle(base?: PlaceholderStyle, override?: PlaceholderStyle): PlaceholderStyle {
  if (!base && !override) return {};
  if (!base) return { ...override };
  if (!override) return { ...base };
  return {
    ...base,
    ...override,
    fill: override.fill !== undefined ? override.fill : base.fill,
    fillOpacity: override.fillOpacity !== undefined ? override.fillOpacity : base.fillOpacity,
    stroke: override.stroke !== undefined ? override.stroke : base.stroke,
    strokeOpacity: override.strokeOpacity !== undefined ? override.strokeOpacity : base.strokeOpacity,
    strokeWidth: override.strokeWidth !== undefined ? override.strokeWidth : base.strokeWidth,
    strokeDash: override.strokeDash !== undefined ? override.strokeDash : base.strokeDash,
    shadow: override.shadow ?? base.shadow,
    text: mergeTextStyles(base.text, override.text),
    body: mergeTextStyles(base.body, override.body),
    levels: {
      ...(base.levels ?? {}),
      ...(override.levels ?? {}),
    },
  };
}

function mergePlaceholderMap(
  base: Record<string, PlaceholderStyle>,
  override: Record<string, PlaceholderStyle>,
) {
  const result = { ...base };
  Object.entries(override).forEach(([key, value]) => {
    result[key] = mergePlaceholderStyle(result[key], value);
  });
  return result;
}

function readBodyPrStyle(bodyPr: Element | null): TextStyle {
  return {
    verticalAlign:
      attr(bodyPr, 'anchor') === 'ctr'
        ? 'middle'
        : attr(bodyPr, 'anchor') === 'b'
          ? 'bottom'
          : 'top',
    writingMode:
      attr(bodyPr, 'vert') === 'vert'
        ? 'vertical-rl'
        : attr(bodyPr, 'vert') === 'vert270'
          ? 'vertical-lr'
          : 'horizontal-tb',
    fit: childByLocalName(bodyPr, 'spAutoFit')
      ? 'resizeShape'
      : childByLocalName(bodyPr, 'normAutofit')
        ? 'shrinkText'
        : childByLocalName(bodyPr, 'noAutofit')
          ? 'none'
          : undefined,
    marginLeft: attr(bodyPr, 'lIns') ? emuToPx(Number(attr(bodyPr, 'lIns'))) : undefined,
    marginRight: attr(bodyPr, 'rIns') ? emuToPx(Number(attr(bodyPr, 'rIns'))) : undefined,
    marginTop: attr(bodyPr, 'tIns') ? emuToPx(Number(attr(bodyPr, 'tIns'))) : undefined,
    marginBottom: attr(bodyPr, 'bIns') ? emuToPx(Number(attr(bodyPr, 'bIns'))) : undefined,
  };
}

function readDefaultRunStyle(node: Element | null, theme: ThemeModel): TextStyle {
  if (!node) return {};
  const solidFill = childByLocalName(node, 'solidFill');
  const fontNode = childByLocalName(node, 'latin') ?? childByLocalName(node, 'ea') ?? childByLocalName(node, 'cs');
  return {
    fontFamily: attr(fontNode, 'typeface'),
    fontSize: attr(node, 'sz') ? Number(attr(node, 'sz')) / 100 : undefined,
    bold: boolAttr(node, 'b'),
    italic: boolAttr(node, 'i'),
    underline: attr(node, 'u') === 'sng' || attr(node, 'u') === '1',
    strike:
      attr(node, 'strike') === 'dblStrike'
        ? 'dblStrike'
        : attr(node, 'strike') === 'sngStrike'
          ? 'sngStrike'
          : attr(node, 'strike') === 'none'
            ? 'none'
            : undefined,
    smallCaps: boolAttr(node, 'smCap'),
    allCaps: boolAttr(node, 'cap'),
    color: parseColorNode(solidFill, theme),
    opacity: parseAlphaNode(solidFill),
    charSpace: attr(node, 'spc') ? Number(attr(node, 'spc')) / 100 : undefined,
    baseline: attr(node, 'baseline') ? Number(attr(node, 'baseline')) / 1000 : undefined,
  };
}

function readParagraphLevelStyle(node: Element | null, theme: ThemeModel): TextStyle {
  if (!node) return {};
  const solidFill = childByLocalName(node, 'solidFill');
  const bulletChar = attr(childByLocalName(node, 'buChar'), 'char');
  const bulletSize = pointToPx(attr(childByLocalName(node, 'buSzPts'), 'val'));
  const bulletColorNode = childByLocalName(node, 'buClr');
  const bulletColor = parseColorNode(childByLocalName(bulletColorNode, 'solidFill') ?? bulletColorNode, theme);
  const bulletNone = Boolean(childByLocalName(node, 'buNone'));
  const lineSpace = childByLocalName(node, 'lnSpc');
  const spcPct = childByLocalName(lineSpace, 'spcPct');
  const spcPts = childByLocalName(lineSpace, 'spcPts');
  const spcBef = childByLocalName(node, 'spcBef');
  const spcAft = childByLocalName(node, 'spcAft');
  const spaceBefore = pointToPx(attr(childByLocalName(spcBef, 'spcPts'), 'val'));
  const spaceAfter = pointToPx(attr(childByLocalName(spcAft, 'spcPts'), 'val'));

  return {
    align:
      attr(node, 'algn') === 'ctr'
        ? 'center'
        : attr(node, 'algn') === 'r'
          ? 'right'
          : attr(node, 'algn') === 'just'
            ? 'justify'
            : undefined,
    lineHeight: pctToRatio(attr(spcPct, 'val')) ?? pointToPx(attr(spcPts, 'val')),
    marginLeft: attr(node, 'marL') ? emuToPx(Number(attr(node, 'marL'))) : undefined,
    textIndent: attr(node, 'indent') ? emuToPx(Number(attr(node, 'indent'))) : undefined,
    spaceBefore,
    spaceAfter,
    color: parseColorNode(solidFill, theme),
    bullet: bulletChar || bulletNone
      ? {
          char: bulletChar,
          color: bulletColor,
          size: bulletSize,
          none: bulletNone,
        }
      : undefined,
  };
}

function readLevelStyles(txBody: Element | null, theme: ThemeModel) {
  const listStyle = childByLocalName(txBody, 'lstStyle');
  const levels: Record<number, TextStyle> = {};
  for (let level = 1; level <= 9; level += 1) {
    const node = childByLocalName(listStyle, `lvl${level}pPr`);
    if (!node) continue;
    levels[level - 1] = mergeTextStyles(
      readParagraphLevelStyle(node, theme),
      readDefaultRunStyle(childByLocalName(node, 'defRPr'), theme),
    );
  }
  return levels;
}

function readTextStyleFamily(styleNode: Element | null, theme: ThemeModel): PlaceholderStyle {
  if (!styleNode) return {};
  const defaultParagraph = childByLocalName(styleNode, 'defPPr');
  const defaultRun = childByLocalName(defaultParagraph, 'defRPr');
  const body = mergeTextStyles(readParagraphLevelStyle(defaultParagraph, theme), readBodyPrStyle(null));
  const text = mergeTextStyles(body, readDefaultRunStyle(defaultRun, theme));
  const levels: Record<number, TextStyle> = {};

  for (let level = 1; level <= 9; level += 1) {
    const node = childByLocalName(styleNode, `lvl${level}pPr`);
    if (!node) continue;
    levels[level - 1] = mergeTextStyles(
      readParagraphLevelStyle(node, theme),
      readDefaultRunStyle(childByLocalName(node, 'defRPr'), theme),
    );
  }

  return { text, body, levels };
}

function readTextPresetMap(txStyles: Element | null, theme: ThemeModel) {
  const presets: Record<string, PlaceholderStyle> = {};
  const titleStyle = childByLocalName(txStyles, 'titleStyle');
  const bodyStyle = childByLocalName(txStyles, 'bodyStyle');
  const otherStyle = childByLocalName(txStyles, 'otherStyle');

  const title = readTextStyleFamily(titleStyle, theme);
  const body = readTextStyleFamily(bodyStyle, theme);
  const other = readTextStyleFamily(otherStyle, theme);

  ['title:0', 'ctrTitle:0', 'title:1', 'ctrTitle:1'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], title);
  });
  ['subTitle:0', 'subTitle:1', 'body:0'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], body);
  });
  ['dt:10', 'ftr:11', 'sldNum:12', 'other:0'].forEach((key) => {
    presets[key] = mergePlaceholderStyle(presets[key], other);
  });

  return presets;
}

function readTableCellStyle(node: Element | null, theme: ThemeModel): TableCellStyle {
  if (!node) return {};
  const tcStyle = childByLocalName(node, 'tcStyle') ?? node;
  const tcTxStyle = childByLocalName(node, 'tcTxStyle') ?? node;
  const fillNode = childByLocalName(tcStyle, 'fill');
  const solidFill = childByLocalName(fillNode, 'solidFill');
  const noFill = Boolean(childByLocalName(fillNode, 'noFill'));
  const borderNode = childByLocalName(tcStyle, 'tcBdr');
  const borderLine =
    childByLocalName(borderNode, 'ln') ??
    childByLocalName(borderNode, 'left')?.firstElementChild ??
    childByLocalName(borderNode, 'right')?.firstElementChild ??
    childByLocalName(borderNode, 'top')?.firstElementChild ??
    childByLocalName(borderNode, 'bottom')?.firstElementChild ??
    childByLocalName(borderNode, 'insideH')?.firstElementChild ??
    childByLocalName(borderNode, 'insideV')?.firstElementChild;
  const fill = noFill || !solidFill ? undefined : parseColorNode(solidFill, theme);
  const borderFill = childByLocalName(borderLine, 'solidFill');
  const textColorNode =
    childByLocalName(tcTxStyle, 'solidFill') ??
    childByLocalName(tcTxStyle, 'srgbClr') ??
    childByLocalName(tcTxStyle, 'schemeClr') ??
    childByLocalName(tcTxStyle, 'prstClr');
  return {
    text: mergeTextStyles(readDefaultRunStyle(tcTxStyle, theme), {
      color: parseColorNode(textColorNode, theme),
      opacity: parseAlphaNode(textColorNode),
    }),
    backgroundColor: fill,
    backgroundOpacity: parseAlphaNode(solidFill),
    borderColor: childByLocalName(borderLine, 'noFill') ? null : parseColorNode(borderFill ?? borderLine, theme),
    borderOpacity: parseAlphaNode(borderFill ?? borderLine),
    borderWidth: attr(borderLine, 'w') ? Number(attr(borderLine, 'w')) / 12700 : undefined,
  };
}

function readTableStyles(xml: string, theme: ThemeModel): TableStyleMap {
  if (!xml) return {};
  const doc = parseXml(xml);
  const result: TableStyleMap = {};
  descendantsByLocalName(doc.documentElement, 'tblStyle').forEach((styleNode) => {
    const styleId = attr(styleNode, 'styleId');
    if (!styleId) return;
    const variants: Partial<Record<TableStyleVariantName, TableCellStyle>> = {};
    ([
      'wholeTbl',
      'band1H',
      'band2H',
      'band1V',
      'band2V',
      'firstRow',
      'lastRow',
      'firstCol',
      'lastCol',
    ] as TableStyleVariantName[]).forEach((variantName) => {
      const variantNode = childByLocalName(styleNode, variantName);
      if (!variantNode) return;
      variants[variantName] = readTableCellStyle(variantNode, theme);
    });
    result[styleId] = {
      styleId,
      styleName: attr(styleNode, 'styleName') ?? undefined,
      variants,
    };
  });
  return result;
}

function mergeTableCellStyle(...styles: Array<TableCellStyle | undefined>) {
  return styles.reduce<TableCellStyle>((acc, style) => {
    if (!style) return acc;
    return {
      text: mergeTextStyles(acc.text, style.text),
      backgroundColor: style.backgroundColor !== undefined ? style.backgroundColor : acc.backgroundColor,
      backgroundOpacity: style.backgroundOpacity !== undefined ? style.backgroundOpacity : acc.backgroundOpacity,
      borderColor: style.borderColor !== undefined ? style.borderColor : acc.borderColor,
      borderOpacity: style.borderOpacity !== undefined ? style.borderOpacity : acc.borderOpacity,
      borderWidth: style.borderWidth !== undefined ? style.borderWidth : acc.borderWidth,
    };
  }, {});
}

function readCustomGeometry(spPr: Element | null) {
  const custGeom = childByLocalName(spPr, 'custGeom');
  if (!custGeom) return {};

  const paths = descendantsByLocalName(custGeom, 'path');
  const pathData: string[] = [];
  let viewBox: string | undefined;

  paths.forEach((pathNode) => {
    const width = attr(pathNode, 'w');
    const height = attr(pathNode, 'h');
    if (!viewBox && width && height) {
      viewBox = `0 0 ${width} ${height}`;
    }

    const commands: string[] = [];
    Array.from(pathNode.children).forEach((child) => {
      if (matchesLocalName(child, 'close')) {
        commands.push('Z');
        return;
      }

      const points = descendantsByLocalName(child, 'pt').map((point) => {
        const x = Number(attr(point, 'x') ?? Number.NaN);
        const y = Number(attr(point, 'y') ?? Number.NaN);
        return Number.isFinite(x) && Number.isFinite(y) ? `${x} ${y}` : undefined;
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

    if (commands.length) {
      pathData.push(commands.join(' '));
    }
  });

  return {
    path: pathData.length ? pathData.join(' ') : undefined,
    viewBox,
  };
}

function readShapeVisualStyle(spPr: Element | null, theme: ThemeModel) {
  const xfrm = childByLocalName(spPr, 'xfrm');
  const noFill = Boolean(childByLocalName(spPr, 'noFill'));
  const line = childByLocalName(spPr, 'ln');
  const customGeometry = readCustomGeometry(spPr);
  const shape = customGeometry.path ? 'path' : attr(childByLocalName(spPr, 'prstGeom'), 'prst') ?? 'rect';
  const fillNode = childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill') ?? childByLocalName(spPr, 'pattFill');
  const fill = noFill || !fillNode ? null : parsePaintNode(fillNode, theme);
  const strokeNone = !line || Boolean(childByLocalName(line, 'noFill'));
  const strokeNode = childByLocalName(line, 'solidFill') ?? childByLocalName(line, 'gradFill') ?? childByLocalName(line, 'pattFill');
  const stroke = strokeNone || !strokeNode ? null : parseColorNode(strokeNode ?? line, theme);
  const shadow = parseShadowNode(childByLocalName(spPr, 'effectLst') ?? childByLocalName(spPr, 'effectDag'), theme);

  return {
    shape,
    fill,
    fillOpacity: typeof fill === 'string' ? parseAlphaNode(fillNode) : undefined,
    stroke,
    strokeOpacity: parseAlphaNode(strokeNode ?? line),
    strokeWidth: attr(line, 'w') ? Number(attr(line, 'w')) / 12700 : undefined,
    strokeDash: attr(line, 'prstDash') ?? undefined,
    shadow,
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    borderRadius: readBorderRadius(spPr),
    path: customGeometry.path,
    viewBox: customGeometry.viewBox,
  };
}

function readBorderRadius(spPr: Element | null) {
  const geom = childByLocalName(spPr, 'prstGeom');
  if (attr(geom, 'prst') !== 'roundRect') return undefined;
  const adj = descendantByLocalName(geom, 'gd');
  const value = attr(adj, 'fmla');
  if (!value) return undefined;
  const match = value.match(/val\s+(\d+)/i);
  const ratio = match ? Number(match[1]) / 100000 : undefined;
  return ratio;
}

function colorWithOpacity(color: string, opacity?: number) {
  if (opacity === undefined || opacity >= 1) return color;
  const normalized = color.replace('#', '');
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return color;
  const value = Number.parseInt(normalized, 16);
  const r = (value >> 16) & 255;
  const g = (value >> 8) & 255;
  const b = value & 255;
  return `rgba(${r}, ${g}, ${b}, ${opacity})`;
}

function readGradientFill(node: Element | null, theme: ThemeModel): GradientFill | undefined {
  if (!node || !matchesLocalName(node, 'gradFill')) return undefined;

  const stops = childrenByLocalName(childByLocalName(node, 'gsLst'), 'gs')
    .map((stop) => {
      const colorNode =
        childByLocalName(stop, 'srgbClr') ??
        childByLocalName(stop, 'schemeClr') ??
        childByLocalName(stop, 'sysClr') ??
        childByLocalName(stop, 'prstClr');
      const color = parseColorNode(colorNode, theme);
      if (!color) return undefined;
      return {
        offset: clamp01(Number(attr(stop, 'pos') ?? 0) / 100000),
        color: colorWithOpacity(color, parseAlphaNode(colorNode)),
      };
    })
    .filter((stop): stop is { offset: number; color: string } => Boolean(stop))
    .sort((a, b) => a.offset - b.offset);

  if (!stops.length) return undefined;

  return {
    type: 'linear',
    angle: Number(attr(childByLocalName(node, 'lin'), 'ang') ?? 0) / 60000,
    stops,
  };
}

function pickGradientColorNode(node: Element | null) {
  if (!node || !matchesLocalName(node, 'gradFill')) return node;
  const stops = descendantsByLocalName(node, 'gs');
  for (let index = stops.length - 1; index >= 0; index -= 1) {
    const colorNode =
      childByLocalName(stops[index], 'srgbClr') ??
      childByLocalName(stops[index], 'schemeClr') ??
      childByLocalName(stops[index], 'sysClr') ??
      childByLocalName(stops[index], 'prstClr');
    if (colorNode) return colorNode;
  }
  return node;
}

function parsePaintNode(node: Element | null, theme: ThemeModel) {
  return readGradientFill(node, theme) ?? parseColorNode(node, theme);
}

function parseColorNode(node: Element | null, theme: ThemeModel) {
  const sourceNode = pickGradientColorNode(node);
  if (!sourceNode) return undefined;
  const srgb = matchesLocalName(sourceNode, 'srgbClr') ? sourceNode : childByLocalName(sourceNode, 'srgbClr');
  const scheme = matchesLocalName(sourceNode, 'schemeClr') ? sourceNode : childByLocalName(sourceNode, 'schemeClr');
  const sys = matchesLocalName(sourceNode, 'sysClr') ? sourceNode : childByLocalName(sourceNode, 'sysClr');
  const prst = matchesLocalName(sourceNode, 'prstClr') ? sourceNode : childByLocalName(sourceNode, 'prstClr');
  const colorNode = srgb ?? scheme ?? sys ?? prst;
  const base =
    attr(srgb, 'val') ??
    resolveThemeColor(attr(scheme, 'val'), theme) ??
    attr(sys, 'lastClr') ??
    attr(prst, 'val');
  const transforms = descendantsByLocalName(colorNode, 'tint')
    .concat(descendantsByLocalName(colorNode, 'shade'))
    .concat(descendantsByLocalName(colorNode, 'lumMod'))
    .concat(descendantsByLocalName(colorNode, 'lumOff'))
    .map((item) => ({ type: item.localName, val: Number(attr(item, 'val') ?? 0) }));
  const raw = transformColor(toHexColor(base), transforms);
  return raw;
}

function parseAlphaNode(node: Element | null) {
  return alphaToOpacity(attr(descendantByLocalName(pickGradientColorNode(node), 'alpha'), 'val'));
}

function parseRatioNode(node: Element | null) {
  return alphaToRatio(attr(descendantByLocalName(node, 'alpha'), 'val'));
}

function parseShadowNode(node: Element | null, theme: ThemeModel): ShadowStyle | undefined {
  if (!node) return undefined;
  const colorNode = childByLocalName(node, 'srgbClr') ?? childByLocalName(node, 'schemeClr');
  const color = parseColorNode(colorNode, theme);
  const opacity = parseRatioNode(colorNode);
  const blur = attr(node, 'blurRad') ? Number(attr(node, 'blurRad')) / 12700 : undefined;
  const dist = attr(node, 'dist') ? Number(attr(node, 'dist')) / 12700 : undefined;
  const dir = attr(node, 'dir') ? (Number(attr(node, 'dir')) / 60000) * (Math.PI / 180) : undefined;
  return {
    color,
    opacity,
    blur,
    offsetX: dist && dir !== undefined ? Math.cos(dir) * dist : undefined,
    offsetY: dist && dir !== undefined ? Math.sin(dir) * dist : undefined,
  };
}

function resolveMediaRef(
  target: string | undefined,
  packageState: PackageState,
) {
  return resolvePackageMediaRef(target, packageState.mediaByPath, packageState.mediaByName, 'ppt');
}

function resolveXmlTarget(target: string | undefined, packageState: PackageState) {
  if (!target) return undefined;
  const normalized = target.replace(/^\.\.\//, '');
  return packageState.entries.get(normalized) ? normalized : target;
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

function normalizeLegendPosition(value: unknown): ChartElement['chart']['legendPosition'] | undefined {
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

function readWpsLegendStyle(legend: unknown): ChartElement['chart']['legendStyle'] | undefined {
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

function resolveWebExtensionSnapshot(doc: XMLDocument, webExtensionPath: string, packageState: PackageState) {
  const snapshot = descendantByLocalName(doc.documentElement, 'snapshot');
  const embed = attr(snapshot, 'r:embed') ?? attr(snapshot, 'embed');
  const relsPath = webExtensionPath.replace(/^ppt\/webExtensions\//, 'ppt/webExtensions/_rels/').concat('.rels');
  const target = embed ? packageState.relationships[relsPath]?.[embed] : undefined;
  return resolveMediaRef(target, packageState);
}

function readSlideBackground(
  bgNode: Element | null,
  theme: ThemeModel,
  packageState: PackageState,
  slideRels?: Record<string, string>,
): SlideBackground | undefined {
  if (!bgNode) return undefined;
  const bgPr = childByLocalName(bgNode, 'bgPr');
  if (!bgPr) return undefined;
  const solidFill = childByLocalName(bgPr, 'solidFill');
  const fill = parseColorNode(solidFill, theme);
  const blip = descendantByLocalName(bgPr, 'blip');
  const embed = attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? slideRels?.[embed] : undefined;
  return {
    fill,
    fillOpacity: parseAlphaNode(solidFill),
    imageRef: resolveMediaRef(target, packageState),
  };
}

function readPlaceholder(node: Element | null, theme: ThemeModel): PlaceholderStyle {
  if (!node) return {};
  const ph = descendantByLocalName(node, 'ph');
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const visual = readShapeVisualStyle(spPr, theme);
  const textBody = childByLocalName(node, 'txBody');
  const bodyPr = childByLocalName(textBody, 'bodyPr');
  const defRPr = descendantByLocalName(textBody, 'defRPr') ?? descendantByLocalName(textBody, 'endParaRPr');

  return {
    type: attr(ph, 'type') ?? undefined,
    idx: attr(ph, 'idx') ?? undefined,
    x: emuValue(off, 'x'),
    y: emuValue(off, 'y'),
    width: emuValue(ext, 'cx'),
    height: emuValue(ext, 'cy'),
    ...visual,
    text: mergeTextStyles(readDefaultRunStyle(defRPr, theme), readBodyPrStyle(bodyPr)),
    body: mergeTextStyles(readBodyPrStyle(bodyPr)),
    levels: readLevelStyles(textBody, theme),
  };
}

function resolvePlaceholderStyle(
  ph: Element,
  placeholderStyles?: Record<string, PlaceholderStyle>,
) {
  if (!placeholderStyles) return undefined;
  const type = attr(ph, 'type') ?? 'body';
  const idx = attr(ph, 'idx') ?? '0';
  const aliases = [
    `${type}:${idx}`,
    `${type}:0`,
    type === 'ctrTitle' ? 'title:0' : undefined,
    type === 'title' ? 'ctrTitle:0' : undefined,
    type === 'subTitle' ? 'body:0' : undefined,
    type === 'subTitle' ? 'title:0' : undefined,
    type === 'dt' ? 'other:0' : undefined,
    type === 'ftr' ? 'other:0' : undefined,
    type === 'sldNum' ? 'other:0' : undefined,
    `body:${idx}`,
    'body:0',
  ].filter((key): key is string => Boolean(key));

  for (const key of aliases) {
    if (placeholderStyles[key]) return placeholderStyles[key];
  }

  return undefined;
}

function translateElement(node: { x: number; y: number; width: number; height: number }, dx = 0, dy = 0) {
  return {
    x: node.x + dx,
    y: node.y + dy,
  };
}

function transformGroupedElement(
  element: { x: number; y: number; width: number; height: number },
  group: { x: number; y: number; width: number; height: number; childX: number; childY: number; childWidth: number; childHeight: number },
) {
  const scaleX = group.childWidth ? group.width / group.childWidth : 1;
  const scaleY = group.childHeight ? group.height / group.childHeight : 1;
  return {
    x: group.x + (element.x - group.childX) * scaleX,
    y: group.y + (element.y - group.childY) * scaleY,
    width: element.width * scaleX,
    height: element.height * scaleY,
  };
}

function parseGroupElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, string>,
  sourcePrefix: string,
  placeholderStyles?: Record<string, PlaceholderStyle>,
  tableStyles?: TableStyleMap,
  includePlaceholders = true,
) {
  const spPr = childByLocalName(node, 'grpSpPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const offsetX = emuValue(childByLocalName(xfrm, 'off'), 'x') ?? 0;
  const offsetY = emuValue(childByLocalName(xfrm, 'off'), 'y') ?? 0;
  const childX = emuValue(childByLocalName(xfrm, 'chOff'), 'x') ?? 0;
  const childY = emuValue(childByLocalName(xfrm, 'chOff'), 'y') ?? 0;
  const childWidth = emuValue(childByLocalName(xfrm, 'chExt'), 'cx') ?? emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0;
  const childHeight = emuValue(childByLocalName(xfrm, 'chExt'), 'cy') ?? emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0;
  const width = emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0;
  const height = emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0;
  const inner = childByLocalName(node, 'spTree') ?? node;
  const childElements = parseVisualTree(
    inner,
    theme,
    packageState,
    rels,
    `${sourcePrefix}-group-${index}`,
    placeholderStyles,
    tableStyles,
    includePlaceholders,
  );
  return childElements.map((element) => {
    const translated = transformGroupedElement(element, {
      x: offsetX,
      y: offsetY,
      width,
      height,
      childX,
      childY,
      childWidth,
      childHeight,
    });
    return {
      ...element,
      id: `${sourcePrefix}-group-${index}-${element.id}`,
      ...translated,
    };
  });
}

function parseVisualTree(
  spTree: Element | null,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, string>,
  sourcePrefix: string,
  placeholderStyles?: Record<string, PlaceholderStyle>,
  tableStyles?: TableStyleMap,
  includePlaceholders = true,
) {
  const elements: SlideElement[] = [];
  const nodes = childrenByLocalName(spTree, 'sp')
    .concat(childrenByLocalName(spTree, 'pic'))
    .concat(childrenByLocalName(spTree, 'graphicFrame'))
    .concat(childrenByLocalName(spTree, 'cxnSp'))
    .concat(childrenByLocalName(spTree, 'grpSp'))
    .sort((a, b) => Array.from(spTree?.children ?? []).indexOf(a) - Array.from(spTree?.children ?? []).indexOf(b));

  nodes.forEach((node, elementIndex) => {
    if (node.localName === 'pic') {
      const chart = parseWpsWebExtensionChart(node, elementIndex, packageState, rels);
      if (chart) {
        chart.id = `${sourcePrefix}-${chart.id}`;
        elements.push(chart);
        return;
      }
      const image = parseImageElement(node, elementIndex, packageState, rels);
      image.id = `${sourcePrefix}-${image.id}`;
      elements.push(image);
      return;
    }

    if (node.localName === 'graphicFrame') {
      const chart = parseChartElement(node, elementIndex, theme, packageState, rels);
      const tbl = descendantByLocalName(node, 'tbl');
      const element = chart ?? (tbl ? parseTableElement(node, elementIndex, theme, tableStyles) : parseUnsupportedElement(elementIndex, 'Unsupported graphic frame'));
      element.id = `${sourcePrefix}-${element.id}`;
      elements.push(element);
      return;
    }

    if (node.localName === 'grpSp') {
      const groupElements = parseGroupElement(
        node,
        elementIndex,
        theme,
        packageState,
        rels,
        sourcePrefix,
        placeholderStyles,
        tableStyles,
        includePlaceholders,
      );
      elements.push(...groupElements);
      return;
    }

    const ph = descendantByLocalName(node, 'ph');
    if (ph && !includePlaceholders) {
      return;
    }
    const inherited = ph ? resolvePlaceholderStyle(ph, placeholderStyles) : undefined;
    const hasText = Boolean(node.querySelector('txBody'));
    const visualNode = childByLocalName(node, 'spPr');
    const visual = visualNode ? readShapeVisualStyle(visualNode, theme) : undefined;
    const hasVisibleVisual = Boolean(
      visual &&
        ((visual.fill !== undefined && visual.fill !== null) ||
          (visual.stroke !== undefined && visual.stroke !== null) ||
          visual.shadow),
    );
    if (ph && !placeholderStyles && !hasText && !hasVisibleVisual) {
      return;
    }
    if (ph && !hasText && !inherited?.fill && !inherited?.stroke && !inherited?.shadow) {
      return;
    }
    elements.push(hasText
      ? parseTextElement(node, elementIndex, theme, inherited)
      : parseShapeElement(node, elementIndex, theme, inherited));
    elements[elements.length - 1].id = `${sourcePrefix}-${elements[elements.length - 1].id}`;
  });

  return elements;
}

function readMaster(
  xml: string,
  theme: ThemeModel,
  relPath: string,
  packageState: PackageState,
  rels: Record<string, string>,
  tableStyles?: TableStyleMap,
): MasterDefinition {
  const doc = parseXml(xml);
  const cSld = childByLocalName(doc.documentElement, 'cSld');
  const bg = childByLocalName(cSld, 'bg');
  const background = readSlideBackground(bg, theme, packageState, rels);
  const placeholders: Record<string, PlaceholderStyle> = {};
  descendantsByLocalName(cSld, 'sp').forEach((node) => {
    const ph = descendantByLocalName(node, 'ph');
    if (!ph) return;
    const style = readPlaceholder(node, theme);
    const key = `${style.type ?? 'body'}:${style.idx ?? '0'}`;
    placeholders[key] = style;
  });
  const textPresets = readTextPresetMap(childByLocalName(doc.documentElement, 'txStyles'), theme);
  const elements = parseVisualTree(
    childByLocalName(cSld, 'spTree'),
    theme,
    packageState,
    rels,
    `master-${relPath}`,
    undefined,
    tableStyles,
    false,
  );
  return { path: relPath, placeholders, textPresets, background, elements };
}

function readLayout(
  xml: string,
  theme: ThemeModel,
  relPath: string,
  masterPath: string,
  packageState: PackageState,
  rels: Record<string, string>,
  tableStyles?: TableStyleMap,
): LayoutDefinition {
  const doc = parseXml(xml);
  const cSld = childByLocalName(doc.documentElement, 'cSld');
  const bg = childByLocalName(cSld, 'bg');
  const background = readSlideBackground(bg, theme, packageState, rels);
  const placeholders: Record<string, PlaceholderStyle> = {};
  descendantsByLocalName(cSld, 'sp').forEach((node) => {
    const ph = descendantByLocalName(node, 'ph');
    if (!ph) return;
    const style = readPlaceholder(node, theme);
    const key = `${style.type ?? 'body'}:${style.idx ?? '0'}`;
    placeholders[key] = style;
  });
  const textPresets = readTextPresetMap(childByLocalName(doc.documentElement, 'txStyles'), theme);
  const elements = parseVisualTree(
    childByLocalName(cSld, 'spTree'),
    theme,
    packageState,
    rels,
    `layout-${relPath}`,
    undefined,
    tableStyles,
    false,
  );
  return { path: relPath, masterPath, placeholders, textPresets, background, elements };
}

function readPresentationLayouts(entries: OfficeEntryMap, packageState: PackageState, theme: ThemeModel, tableStyles?: TableStyleMap) {
  const presentationRels = packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const masterDefinitions: MasterDefinition[] = [];
  const masterLayoutDefinitions: Record<string, LayoutDefinition[]> = {};

  Object.entries(presentationRels).forEach(([relId, target]) => {
    if (!target.includes('slideMasters/')) return;
    const xmlPath = target.startsWith('ppt/') ? target : `ppt/${target}`;
    const relPath = xmlPath.replace(/^ppt\/slideMasters\//, 'ppt/slideMasters/_rels/').replace(/\.xml$/, '.xml.rels');
    const masterRels = packageState.relationships[relPath] ?? {};
    const master = readMaster(readXml(entries, xmlPath), theme, xmlPath, packageState, masterRels, tableStyles);
    masterDefinitions.push(master);
    masterLayoutDefinitions[xmlPath] = Object.values(masterRels)
      .filter((item) => item.includes('slideLayouts/'))
      .map((layoutTarget) => {
        const layoutPath = layoutTarget.startsWith('ppt/') ? layoutTarget : `ppt/${layoutTarget}`;
        const layoutRelPath = layoutPath.replace(/^ppt\/slideLayouts\//, 'ppt/slideLayouts/_rels/').concat('.rels');
        return readLayout(
          readXml(entries, layoutPath),
          theme,
          layoutPath,
          xmlPath,
          packageState,
          packageState.relationships[layoutRelPath] ?? {},
          tableStyles,
        );
      });
  });

  return { masterDefinitions, masterLayoutDefinitions };
}

function mergePlaceholder(
  key: string,
  layoutPlaceholders: Record<string, PlaceholderStyle>,
  masterPlaceholders: Record<string, PlaceholderStyle>,
) {
  return mergePlaceholderStyle(masterPlaceholders[key], layoutPlaceholders[key]);
}

function parseTextElement(node: Element, index: number, theme: ThemeModel, inherited?: PlaceholderStyle): TextElement {
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const txBody = childByLocalName(node, 'txBody');
  const bodyPr = childByLocalName(txBody, 'bodyPr');
  const visual = readShapeVisualStyle(spPr, theme);
  const localLevels = readLevelStyles(txBody, theme);
  const bodyStyle = mergeTextStyles(inherited?.body, readBodyPrStyle(bodyPr));

  const paragraphs: TextParagraph[] = childrenByLocalName(txBody, 'p').map((paragraphNode) => {
    const paragraphProps = childByLocalName(paragraphNode, 'pPr');
    const level = Number(attr(paragraphProps, 'lvl') ?? 0);
    const levelStyle = mergeTextStyles(
      inherited?.levels?.[level],
      localLevels[level],
      readParagraphLevelStyle(paragraphProps, theme),
    );
    const defaultRunStyle = mergeTextStyles(
      inherited?.text,
      inherited?.body,
      inherited?.levels?.[level],
      localLevels[level],
      readDefaultRunStyle(childByLocalName(paragraphProps, 'defRPr'), theme),
      readDefaultRunStyle(childByLocalName(paragraphNode, 'endParaRPr'), theme),
    );

    const runs: TextRun[] = [];
    Array.from(paragraphNode.children).forEach((child) => {
      if (child.localName === 'r') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: textContent(childByLocalName(child, 't')),
          style: mergeTextStyles(defaultRunStyle, readDefaultRunStyle(runProps, theme)),
        });
        return;
      }

      if (child.localName === 'fld') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: textContent(childByLocalName(child, 't')) || child.textContent || '',
          style: mergeTextStyles(defaultRunStyle, readDefaultRunStyle(runProps, theme)),
        });
        return;
      }

      if (child.localName === 'br') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: '\n',
          style: mergeTextStyles(defaultRunStyle, readDefaultRunStyle(runProps, theme)),
        });
      }
    });

    if (!runs.length) {
      runs.push({ text: paragraphNode.textContent ?? '', style: defaultRunStyle });
    }

    return {
      level,
      runs,
      style: mergeTextStyles(levelStyle, {
        align: levelStyle.align ?? bodyStyle.align,
      }),
      bullet: levelStyle.bullet,
    };
  });

  const firstRunStyle = paragraphs.flatMap((paragraph) => paragraph.runs).find(Boolean)?.style ?? {};
  const fallbackStyle = mergeTextStyles(inherited?.text, bodyStyle);

  return {
    id: `text-${index}`,
    type: 'text',
    x: emuValue(off, 'x') ?? inherited?.x ?? 0,
    y: emuValue(off, 'y') ?? inherited?.y ?? 0,
    width: emuValue(ext, 'cx') ?? inherited?.width ?? 0,
    height: emuValue(ext, 'cy') ?? inherited?.height ?? 0,
    placeholderType: inherited?.type,
    placeholderIdx: inherited?.idx,
    paragraphs,
    shape: visual.shape,
    path: visual.path,
    viewBox: visual.viewBox,
    fill: visual.fill !== undefined ? visual.fill : inherited?.fill,
    fillOpacity: visual.fillOpacity ?? inherited?.fillOpacity,
    stroke: visual.stroke !== undefined ? visual.stroke : inherited?.stroke,
    strokeOpacity: visual.strokeOpacity ?? inherited?.strokeOpacity,
    strokeWidth: visual.strokeWidth ?? inherited?.strokeWidth,
    strokeDash: visual.strokeDash ?? inherited?.strokeDash,
    shadow: visual.shadow ?? inherited?.shadow,
    borderRadius: visual.borderRadius,
    boxStyle: {
      fontFamily: firstRunStyle.fontFamily ?? fallbackStyle.fontFamily,
      fontSize: firstRunStyle.fontSize ?? fallbackStyle.fontSize,
      bold: firstRunStyle.bold ?? fallbackStyle.bold,
      italic: firstRunStyle.italic ?? fallbackStyle.italic,
      underline: firstRunStyle.underline ?? fallbackStyle.underline,
      color: firstRunStyle.color ?? fallbackStyle.color,
      opacity: firstRunStyle.opacity ?? fallbackStyle.opacity,
      align: paragraphs[0]?.style?.align ?? bodyStyle.align ?? fallbackStyle.align,
      lineHeight: paragraphs[0]?.style?.lineHeight ?? fallbackStyle.lineHeight,
      marginLeft: bodyStyle.marginLeft,
      marginRight: bodyStyle.marginRight,
      marginTop: bodyStyle.marginTop,
      marginBottom: bodyStyle.marginBottom,
      verticalAlign: bodyStyle.verticalAlign ?? fallbackStyle.verticalAlign ?? 'top',
      writingMode: bodyStyle.writingMode ?? fallbackStyle.writingMode,
      fit: bodyStyle.fit ?? fallbackStyle.fit,
    },
  };
}

function parseShapeElement(node: Element, index: number, theme: ThemeModel, inherited?: PlaceholderStyle): ShapeElement {
  const spPr = childByLocalName(node, 'spPr');
  const xfrm = childByLocalName(spPr, 'xfrm');
  const ph = descendantByLocalName(node, 'ph');
  const visual = readShapeVisualStyle(spPr, theme);

  return {
    id: `shape-${index}`,
    type: 'shape',
    shape: visual.shape,
    path: visual.path,
    viewBox: visual.viewBox,
    x: emuValue(childByLocalName(xfrm, 'off'), 'x') ?? inherited?.x ?? 0,
    y: emuValue(childByLocalName(xfrm, 'off'), 'y') ?? inherited?.y ?? 0,
    width: emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? inherited?.width ?? 0,
    height: emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? inherited?.height ?? 0,
    rotate: visual.rotate,
    flipH: visual.flipH,
    flipV: visual.flipV,
    fill: visual.fill !== undefined ? visual.fill : inherited?.fill,
    fillOpacity: visual.fillOpacity ?? inherited?.fillOpacity,
    stroke: visual.stroke !== undefined ? visual.stroke : inherited?.stroke,
    strokeOpacity: visual.strokeOpacity ?? inherited?.strokeOpacity,
    strokeWidth: visual.strokeWidth ?? inherited?.strokeWidth,
    opacity: visual.fillOpacity,
    strokeDash: visual.strokeDash ?? inherited?.strokeDash,
    shadow: visual.shadow ?? inherited?.shadow,
    placeholderType: attr(ph, 'type') ?? inherited?.type,
    placeholderIdx: attr(ph, 'idx') ?? inherited?.idx,
    borderRadius: visual.borderRadius,
  };
}

function parseImageElement(
  node: Element,
  index: number,
  packageState: PackageState,
  slideRels: Record<string, string>,
): ImageElement {
  const xfrm = childByLocalName(childByLocalName(node, 'spPr') ?? node, 'xfrm');
  const blip = descendantByLocalName(node, 'blip');
  const svgBlip = descendantByLocalName(node, 'svgBlip');
  const blipFill = childByLocalName(node, 'blipFill');
  const srcRect = childByLocalName(blipFill, 'srcRect');
  const embed = attr(svgBlip, 'r:embed') ?? attr(svgBlip, 'embed') ?? attr(blip, 'r:embed') ?? attr(blip, 'embed');
  const target = embed ? slideRels[embed] : undefined;
  const resolved = resolveMediaRef(target, packageState);
  const crop: ImageCrop | undefined = srcRect
    ? {
        left: attr(srcRect, 'l') ? Number(attr(srcRect, 'l')) / 100000 : undefined,
        top: attr(srcRect, 't') ? Number(attr(srcRect, 't')) / 100000 : undefined,
        right: attr(srcRect, 'r') ? Number(attr(srcRect, 'r')) / 100000 : undefined,
        bottom: attr(srcRect, 'b') ? Number(attr(srcRect, 'b')) / 100000 : undefined,
      }
    : undefined;

  return {
    id: `image-${index}`,
    type: 'image',
    x: emuValue(childByLocalName(xfrm, 'off'), 'x') ?? 0,
    y: emuValue(childByLocalName(xfrm, 'off'), 'y') ?? 0,
    width: emuValue(childByLocalName(xfrm, 'ext'), 'cx') ?? 0,
    height: emuValue(childByLocalName(xfrm, 'ext'), 'cy') ?? 0,
    src: resolved || '',
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    crop,
    alt: target?.split('/').pop(),
  };
}

function parseWpsWebExtensionChart(
  node: Element,
  index: number,
  packageState: PackageState,
  rels: Record<string, string>,
): ChartElement | undefined {
  const webExtensionRef = descendantByLocalName(node, 'webExtensionRef');
  const relId = attr(webExtensionRef, 'r:id') ?? attr(webExtensionRef, 'id');
  const target = relId ? rels[relId] : undefined;
  const webExtensionPath = resolveXmlTarget(target, packageState);
  const xml = webExtensionPath ? (packageState.entries.get(webExtensionPath) as string | undefined) : undefined;
  if (!xml || !webExtensionPath) return undefined;

  const doc = parseXml(xml);
  const snapshotSrc = resolveWebExtensionSnapshot(doc, webExtensionPath, packageState);
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
  const showDataLabels = Boolean(
    (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { show?: unknown }).show) ||
      (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { numberLabel?: { show?: unknown }; textLabel?: { show?: unknown } }).numberLabel?.show) ||
      (chartStyle?.label && typeof chartStyle.label === 'object' && (chartStyle.label as { numberLabel?: { show?: unknown }; textLabel?: { show?: unknown } }).textLabel?.show),
  );
  const titleText =
    typeof title?.show === 'boolean' && title.show !== false && typeof title?.text === 'string'
      ? decodeMojibake(title.text)
      : undefined;
  const xfrm = childByLocalName(childByLocalName(node, 'spPr') ?? node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const x = emuValue(off, 'x') ?? 0;
  const y = emuValue(off, 'y') ?? 0;
  const width = emuValue(ext, 'cx') ?? 0;
  const height = emuValue(ext, 'cy') ?? 0;

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
    const seriesNames = headers.slice(1).map((header, seriesIndex) =>
      decodeMojibake(String(header ?? `Series ${seriesIndex + 1}`).trim()),
    );
    const palette = collectChartColors(chartStyle, 'seriesThemeColor');
    const chartType = isPie && (pieType.includes('doughnut') || radius || roseType) ? 'doughnut' : isPie ? 'pie' : 'line';
    const isPieChart = chartType === 'pie' || chartType === 'doughnut';
    const piePointStyles = isPieChart ? readWpsPiePointStyles(seriesStyle, categories.length) : undefined;
    const sourceSeries = seriesNames.length
      ? seriesNames.map((name, seriesIndex) => ({
          name,
          values: rows.map((row) => Number(row[seriesIndex + 1] ?? 0) || 0),
          type: isPieChart ? ('pie' as const) : ((style?.areaStyle && typeof style.areaStyle === 'object' && (style.areaStyle as { show?: unknown }).show) ? ('area' as const) : ('line' as const)),
          color: palette[seriesIndex],
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

    return {
      id: `chart-${index}`,
      type: 'chart',
      chart: {
        type: chartType,
        title: titleText,
        categories,
        series: sourceSeries.length
          ? sourceSeries.map((series) =>
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
        holeSize:
          chartType === 'doughnut'
            ? (() => {
                const parsed = readPercent(radius?.[0]);
                return Number.isFinite(parsed) ? parsed : undefined;
              })()
            : undefined,
        radius: roseType ? radius : undefined,
        roseType,
        startAngle: Number.isFinite(Number(style?.startAngle)) ? Number(style?.startAngle) : undefined,
        snapshotSrc,
      },
      chartId: relId,
      chartPath: webExtensionPath,
      x,
      y,
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

    return {
      id: `chart-${index}`,
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
      chartId: relId,
      chartPath: webExtensionPath,
      x,
      y,
      width,
      height,
    };
  }

  return undefined;
}

function parseChartElement(
  node: Element,
  index: number,
  theme: ThemeModel,
  packageState: PackageState,
  rels: Record<string, string>,
): ChartElement | undefined {
  const chartNode = descendantByLocalName(node, 'chart');
  const relId = attr(chartNode, 'r:id') ?? attr(chartNode, 'id');
  const target = relId ? rels[relId] : undefined;
  const chartPath = resolveXmlTarget(target, packageState);
  const xml = chartPath ? (packageState.entries.get(chartPath) as string | undefined) : undefined;
  if (!xml) return undefined;

  const chart = parseOfficeChartXml(xml, theme);
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');

  return {
    id: `chart-${index}`,
    type: 'chart',
    chart,
    chartId: relId,
    chartPath,
    x: emuValue(off, 'x') ?? 0,
    y: emuValue(off, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
  };
}

function parseTableCellText(cellNode: Element, theme: ThemeModel) {
  const txBody = childByLocalName(cellNode, 'txBody');
  const paragraphs = childrenByLocalName(txBody, 'p').map((paragraphNode) => {
    const paragraphProps = childByLocalName(paragraphNode, 'pPr');
    const paragraphStyle = readParagraphLevelStyle(paragraphProps, theme);
    const defaultRunStyle = mergeTextStyles(
      readDefaultRunStyle(childByLocalName(paragraphProps, 'defRPr'), theme),
      readDefaultRunStyle(childByLocalName(paragraphNode, 'endParaRPr'), theme),
    );
    const runs: TextRun[] = [];

    Array.from(paragraphNode.children).forEach((child) => {
      if (child.localName === 'r' || child.localName === 'fld') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: textContent(childByLocalName(child, 't')) || child.textContent || '',
          style: mergeTextStyles(defaultRunStyle, readDefaultRunStyle(runProps, theme)),
        });
        return;
      }

      if (child.localName === 'br') {
        const runProps = childByLocalName(child, 'rPr');
        runs.push({
          text: '\n',
          style: mergeTextStyles(defaultRunStyle, readDefaultRunStyle(runProps, theme)),
        });
      }
    });

    if (!runs.length && paragraphNode.textContent) {
      runs.push({ text: paragraphNode.textContent, style: defaultRunStyle });
    }

    return {
      runs,
      style: paragraphStyle,
      bullet: paragraphStyle.bullet,
    };
  });

  const text = paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
  const firstRunStyle = paragraphs.flatMap((paragraph) => paragraph.runs).find(Boolean)?.style;

  return { text, paragraphs, firstRunStyle };
}

function parseTableBorder(tcPr: Element | null, theme: ThemeModel) {
  const line =
    childByLocalName(tcPr, 'ln') ??
    childByLocalName(tcPr, 'lnL') ??
    childByLocalName(tcPr, 'lnR') ??
    childByLocalName(tcPr, 'lnT') ??
    childByLocalName(tcPr, 'lnB');
  if (!line) return {};
  if (childByLocalName(line, 'noFill')) {
    return {
      borderColor: null,
    };
  }
  const fill = childByLocalName(line, 'solidFill') ?? line;
  return {
    borderColor: parseColorNode(fill, theme),
    borderOpacity: parseAlphaNode(fill),
    borderWidth: attr(line, 'w') ? Number(attr(line, 'w')) / 12700 : undefined,
  };
}

function tableFlag(node: Element | null, name: string) {
  return attr(node, name) === '1' || attr(node, name) === 'true';
}

function resolveTableCellStyle(
  style: TableStyleDefinition | undefined,
  tblPr: Element | null,
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number,
) {
  if (!style) return {};
  const isFirstRow = tableFlag(tblPr, 'firstRow') && rowIndex === 0;
  const isLastRow = tableFlag(tblPr, 'lastRow') && rowIndex === rowCount - 1;
  const isFirstCol = tableFlag(tblPr, 'firstCol') && columnIndex === 0;
  const isLastCol = tableFlag(tblPr, 'lastCol') && columnIndex === columnCount - 1;
  const bandRowOffset = tableFlag(tblPr, 'firstRow') ? 1 : 0;
  const bandColOffset = tableFlag(tblPr, 'firstCol') ? 1 : 0;
  const rowBand =
    tableFlag(tblPr, 'bandRow') && !isFirstRow && !isLastRow
      ? (rowIndex - bandRowOffset) % 2 === 0
        ? style.variants.band1H
        : style.variants.band2H
      : undefined;
  const colBand =
    tableFlag(tblPr, 'bandCol') && !isFirstCol && !isLastCol
      ? (columnIndex - bandColOffset) % 2 === 0
        ? style.variants.band1V
        : style.variants.band2V
      : undefined;

  return mergeTableCellStyle(
    style.variants.wholeTbl,
    rowBand,
    colBand,
    isFirstRow ? style.variants.firstRow : undefined,
    isLastRow ? style.variants.lastRow : undefined,
    isFirstCol ? style.variants.firstCol : undefined,
    isLastCol ? style.variants.lastCol : undefined,
  );
}

function parseTableElement(node: Element, index: number, theme: ThemeModel, tableStyles?: TableStyleMap): TableElement {
  const xfrm = childByLocalName(node, 'xfrm');
  const off = childByLocalName(xfrm, 'off');
  const ext = childByLocalName(xfrm, 'ext');
  const tbl = descendantByLocalName(node, 'tbl');
  const tblPr = childByLocalName(tbl, 'tblPr');
  const styleId = textContent(childByLocalName(tblPr, 'tableStyleId')).trim();
  const tableStyle = styleId ? tableStyles?.[styleId] : undefined;
  const columnWidths = childrenByLocalName(childByLocalName(tbl, 'tblGrid'), 'gridCol')
    .map((col) => emuValue(col, 'w') ?? 0);
  const rowNodes = childrenByLocalName(tbl, 'tr');
  const rowHeights = rowNodes.map((rowNode) => emuValue(rowNode, 'h') ?? 0);
  const rows = rowNodes.map((rowNode, rowIndex) =>
    childrenByLocalName(rowNode, 'tc').map((cellNode, columnIndex) => {
      const tcPr = childByLocalName(cellNode, 'tcPr');
      const fillNode = childByLocalName(tcPr, 'solidFill') ?? childByLocalName(tcPr, 'gradFill');
      const { text, paragraphs, firstRunStyle } = parseTableCellText(cellNode, theme);
      const explicitBorder = parseTableBorder(tcPr, theme);
      const styled = resolveTableCellStyle(
        tableStyle,
        tblPr,
        rowIndex,
        columnIndex,
        rowNodes.length,
        columnWidths.length,
      );
      const explicitBackgroundColor = childByLocalName(tcPr, 'noFill') ? null : parseColorNode(fillNode, theme);
      const explicitBackgroundOpacity = childByLocalName(tcPr, 'noFill') ? undefined : parseAlphaNode(fillNode);
      return {
        text,
        paragraphs,
        style: mergeTextStyles(styled.text, firstRunStyle),
        backgroundColor: explicitBackgroundColor !== undefined ? explicitBackgroundColor : styled.backgroundColor ?? undefined,
        backgroundOpacity: explicitBackgroundOpacity !== undefined ? explicitBackgroundOpacity : styled.backgroundOpacity,
        borderColor: explicitBorder.borderColor !== undefined ? explicitBorder.borderColor : styled.borderColor ?? undefined,
        borderOpacity: explicitBorder.borderOpacity !== undefined ? explicitBorder.borderOpacity : styled.borderOpacity,
        borderWidth: explicitBorder.borderWidth !== undefined ? explicitBorder.borderWidth : styled.borderWidth,
        margins: {
          left: emuValue(tcPr, 'marL'),
          right: emuValue(tcPr, 'marR'),
          top: emuValue(tcPr, 'marT'),
          bottom: emuValue(tcPr, 'marB'),
        },
        verticalAlign:
          attr(tcPr, 'anchor') === 'b'
            ? 'bottom'
            : attr(tcPr, 'anchor') === 'ctr'
              ? 'middle'
              : 'top',
      };
    }),
  );

  return {
    id: `table-${index}`,
    type: 'table',
    x: emuValue(off, 'x') ?? 0,
    y: emuValue(off, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
    columnWidths,
    rowHeights,
    rows,
  };
}

function parseUnsupportedElement(index: number, reason: string): UnsupportedElement {
  return {
    id: `unsupported-${index}`,
    type: 'unsupported',
    x: 0,
    y: 0,
    width: 120,
    height: 32,
    reason,
  };
}

function findLayoutForSlide(
  slideRels: Record<string, string>,
  layoutDefinitions: LayoutDefinition[],
) {
  const layoutTarget = Object.values(slideRels).find((target) => target.includes('slideLayout'));
  if (!layoutTarget) return undefined;
  const layoutPath = layoutTarget.startsWith('ppt/') ? layoutTarget : `ppt/${layoutTarget.replace(/^\.\.\//, '')}`;
  return layoutDefinitions.find((layout) => layout.path === layoutTarget || layout.path === layoutPath);
}

function parseSlideXml(
  xml: string,
  index: number,
  width: number,
  height: number,
  packageState: PackageState,
  theme: ThemeModel,
  relPath: string,
  layoutDefinitions: LayoutDefinition[],
  masterDefinitions: MasterDefinition[],
  tableStyles?: TableStyleMap,
): SlideModel {
  const doc = parseXml(xml);
  const slide = doc.documentElement;
  const cSld = childByLocalName(slide, 'cSld');
  const spTree = childByLocalName(cSld, 'spTree');
  const bg = childByLocalName(cSld, 'bg');
  const slideRels = packageState.relationships[relPath] ?? {};
  const layout = findLayoutForSlide(slideRels, layoutDefinitions);
  const master = layout
    ? masterDefinitions.find((item) => item.path === layout.masterPath)
    : masterDefinitions[0];
  const slideBackground = readSlideBackground(bg, theme, packageState, slideRels);
  const background = slideBackground?.fill || slideBackground?.imageRef
    ? slideBackground
    : layout?.background?.fill || layout?.background?.imageRef
      ? layout.background
      : master?.background;

  const elements: SlideElement[] = [...(master?.elements ?? []), ...(layout?.elements ?? [])];
  const placeholderStyles = { ...(master?.placeholders ?? {}) };
  Object.keys(layout?.placeholders ?? {}).forEach((key) => {
    placeholderStyles[key] = mergePlaceholder(key, layout?.placeholders ?? {}, master?.placeholders ?? {});
  });
  Object.entries(master?.textPresets ?? {}).forEach(([key, preset]) => {
    placeholderStyles[key] = mergePlaceholderStyle(placeholderStyles[key], preset);
  });
  Object.entries(layout?.textPresets ?? {}).forEach(([key, preset]) => {
    placeholderStyles[key] = mergePlaceholderStyle(placeholderStyles[key], preset);
  });

  elements.push(...parseVisualTree(spTree, theme, packageState, slideRels, `slide-${index}`, placeholderStyles, tableStyles));

  return {
    id: `slide-${index}`,
    index,
    width,
    height,
    background,
    elements,
  };
}

export async function parsePptx(file: File): Promise<PptxDocument> {
  const entries = await loadPptxEntries(file);
  const packageState = buildPackageState(entries);
  const presentationXml = readXml(entries, 'ppt/presentation.xml');
  const presentationDoc = parseXml(presentationXml);
  const themeXml = readXml(entries, 'ppt/theme/theme1.xml');
  const theme = themeXml ? readTheme(themeXml) : { colorScheme: {}, fontScheme: {}, colorMap: {} };
  const tableStyles = readTableStyles(readXml(entries, 'ppt/tableStyles.xml'), theme);
  const size = presentationXml ? readPresentationSize(presentationXml) : { width: 960, height: 540 };
  const presentationRels = packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const slideIds = childrenByLocalName(presentationDoc.querySelector('sldIdLst, p\\:sldIdLst'), 'sldId');
  const { masterDefinitions, masterLayoutDefinitions } = readPresentationLayouts(entries, packageState, theme, tableStyles);
  const layoutDefinitions = Object.values(masterLayoutDefinitions).flat();
  const slides: SlideModel[] = [];

  slideIds.forEach((node, index) => {
    const relId = attr(node, 'r:id');
    const relTarget = relId ? presentationRels[relId] : undefined;
    const relPath = relTarget ? relTarget.replace(/^ppt\//, '') : `slides/slide${index + 1}.xml`;
    const slidePath = relTarget ?? `ppt/${relPath}`;
    const relsPath = `ppt/${relPath.replace(/^slides\//, 'slides/_rels/')}.rels`;
    slides.push(
      parseSlideXml(
        readXml(entries, slidePath),
        index + 1,
        size.width,
        size.height,
        packageState,
        theme,
        relsPath,
        layoutDefinitions,
        masterDefinitions,
        tableStyles,
      ),
    );
  });

  return { width: size.width, height: size.height, theme, slides };
}
