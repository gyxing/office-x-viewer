import { loadPptxEntries } from './archive';
import type { OfficeEntryMap } from '../office/archive';
import { readXml } from '../office/archive';
import {
  attr,
  childByLocalName,
  childrenByLocalName,
  descendantByLocalName,
  descendantsByLocalName,
  parseXml,
  textContent,
} from '../office/xml';
import { emuToPx } from '../office/units';
import { alphaToOpacity, alphaToRatio, resolveThemeColor, toHexColor, transformColor } from './colors';
import { collectMedia, resolvePackageMediaRef, type OfficeRelationship } from '../office/media';
import { readRelationships } from '../office/relationships';
import { parseOfficeChartXml } from '../office/charts';
import type {
  ChartElement,
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

export function debugPptxPackage(entries: EntryMap) {
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
      const value = attr(child, 'val') ?? attr(child, 'lastClr');
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

function readShapeVisualStyle(spPr: Element | null, theme: ThemeModel) {
  const xfrm = childByLocalName(spPr, 'xfrm');
  const solidFill = childByLocalName(spPr, 'solidFill');
  const noFill = Boolean(childByLocalName(spPr, 'noFill'));
  const line = childByLocalName(spPr, 'ln');
  const shape = attr(childByLocalName(spPr, 'prstGeom'), 'prst') ?? 'rect';
  const fillNode = childByLocalName(spPr, 'solidFill') ?? childByLocalName(spPr, 'gradFill') ?? childByLocalName(spPr, 'pattFill');
  const fill = noFill || !fillNode ? null : parseColorNode(fillNode, theme);
  const strokeNone = !line || Boolean(childByLocalName(line, 'noFill'));
  const strokeNode = childByLocalName(line, 'solidFill');
  const stroke = strokeNone || !strokeNode ? null : parseColorNode(strokeNode ?? line, theme);
  const shadow = parseShadowNode(childByLocalName(spPr, 'effectLst') ?? childByLocalName(spPr, 'effectDag'), theme);

  return {
    shape,
    fill,
    fillOpacity: parseAlphaNode(solidFill),
    stroke,
    strokeOpacity: parseAlphaNode(strokeNode ?? line),
    strokeWidth: attr(line, 'w') ? Number(attr(line, 'w')) / 12700 : undefined,
    strokeDash: attr(line, 'prstDash') ?? undefined,
    shadow,
    rotate: attr(xfrm, 'rot') ? Number(attr(xfrm, 'rot')) / 60000 : undefined,
    flipH: attr(xfrm, 'flipH') === '1',
    flipV: attr(xfrm, 'flipV') === '1',
    borderRadius: readBorderRadius(spPr),
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

function parseColorNode(node: Element | null, theme: ThemeModel) {
  if (!node) return undefined;
  const srgb = node.localName === 'srgbClr' ? node : childByLocalName(node, 'srgbClr');
  const scheme = node.localName === 'schemeClr' ? node : childByLocalName(node, 'schemeClr');
  const sys = node.localName === 'sysClr' ? node : childByLocalName(node, 'sysClr');
  const prst = node.localName === 'prstClr' ? node : childByLocalName(node, 'prstClr');
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
  return alphaToOpacity(attr(descendantByLocalName(node, 'alpha'), 'val'));
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
  const childElements = parseVisualTree(inner, theme, packageState, rels, `${sourcePrefix}-group-${index}`, placeholderStyles);
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
) {
  const elements: SlideElement[] = [];
  const nodes = childrenByLocalName(spTree, 'sp')
    .concat(childrenByLocalName(spTree, 'pic'))
    .concat(childrenByLocalName(spTree, 'graphicFrame'))
    .concat(childrenByLocalName(spTree, 'grpSp'))
    .sort((a, b) => Array.from(spTree?.children ?? []).indexOf(a) - Array.from(spTree?.children ?? []).indexOf(b));

  nodes.forEach((node, elementIndex) => {
    if (node.localName === 'pic') {
      const image = parseImageElement(node, elementIndex, packageState, rels);
      image.id = `${sourcePrefix}-${image.id}`;
      elements.push(image);
      return;
    }

    if (node.localName === 'graphicFrame') {
      const chart = parseChartElement(node, elementIndex, theme, packageState, rels);
      const tbl = childByLocalName(node, 'tbl');
      const element = chart ?? (tbl ? parseTableElement(node, elementIndex) : parseUnsupportedElement(elementIndex, 'Unsupported graphic frame'));
      element.id = `${sourcePrefix}-${element.id}`;
      elements.push(element);
      return;
    }

    if (node.localName === 'grpSp') {
      const groupElements = parseGroupElement(node, elementIndex, theme, packageState, rels, sourcePrefix, placeholderStyles);
      elements.push(...groupElements);
      return;
    }

    const ph = descendantByLocalName(node, 'ph');
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

function readMaster(xml: string, theme: ThemeModel, relPath: string, packageState: PackageState, rels: Record<string, string>): MasterDefinition {
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
  const elements = parseVisualTree(childByLocalName(cSld, 'spTree'), theme, packageState, rels, `master-${relPath}`);
  return { path: relPath, placeholders, textPresets, background, elements };
}

function readLayout(xml: string, theme: ThemeModel, relPath: string, masterPath: string, packageState: PackageState, rels: Record<string, string>): LayoutDefinition {
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
  const elements = parseVisualTree(childByLocalName(cSld, 'spTree'), theme, packageState, rels, `layout-${relPath}`);
  return { path: relPath, masterPath, placeholders, textPresets, background, elements };
}

function readPresentationLayouts(entries: OfficeEntryMap, packageState: PackageState, theme: ThemeModel) {
  const presentationRels = packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const masterDefinitions: MasterDefinition[] = [];
  const masterLayoutDefinitions: Record<string, LayoutDefinition[]> = {};

  Object.entries(presentationRels).forEach(([relId, target]) => {
    if (!target.includes('slideMasters/')) return;
    const xmlPath = target.startsWith('ppt/') ? target : `ppt/${target}`;
    const relPath = xmlPath.replace(/^ppt\/slideMasters\//, 'ppt/slideMasters/_rels/').replace(/\.xml$/, '.xml.rels');
    const masterRels = packageState.relationships[relPath] ?? {};
    const master = readMaster(readXml(entries, xmlPath), theme, xmlPath, packageState, masterRels);
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

function parseTableElement(node: Element, index: number): TableElement {
  const xfrm = childByLocalName(childByLocalName(node, 'xfrm'), 'off');
  const ext = childByLocalName(childByLocalName(node, 'xfrm'), 'ext');
  const tbl = childByLocalName(node, 'tbl');
  const rows = childrenByLocalName(tbl, 'tr').map((rowNode) =>
    childrenByLocalName(rowNode, 'tc').map((cellNode) => ({
      text: cellNode.textContent ?? '',
      style: undefined,
    })),
  );

  return {
    id: `table-${index}`,
    type: 'table',
    x: emuValue(xfrm, 'x') ?? 0,
    y: emuValue(xfrm, 'y') ?? 0,
    width: emuValue(ext, 'cx') ?? 0,
    height: emuValue(ext, 'cy') ?? 0,
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

  elements.push(...parseVisualTree(spTree, theme, packageState, slideRels, `slide-${index}`, placeholderStyles));

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
  const size = presentationXml ? readPresentationSize(presentationXml) : { width: 960, height: 540 };
  const presentationRels = packageState.relationships['ppt/_rels/presentation.xml.rels'] ?? {};
  const slideIds = childrenByLocalName(presentationDoc.querySelector('sldIdLst, p\\:sldIdLst'), 'sldId');
  const { masterDefinitions, masterLayoutDefinitions } = readPresentationLayouts(entries, packageState, theme);
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
      ),
    );
  });

  return { width: size.width, height: size.height, theme, slides };
}
