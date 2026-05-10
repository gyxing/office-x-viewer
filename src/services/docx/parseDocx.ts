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
import { parseOfficeChartXml } from '../office/charts';
import { readOfficeTheme, type OfficeTheme } from '../office/theme';
import type {
  DocxBlock,
  DocxChartBlock,
  DocxDocument,
  DocxImage,
  DocxInline,
  DocxPage,
  DocxParagraphBlock,
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
  images: DocxImage[];
  imageIndex: number;
  chartIndex: number;
};

const DEFAULT_PAGE: DocxPage = {
  width: 794,
  minHeight: 1123,
  marginTop: 96,
  marginRight: 120,
  marginBottom: 96,
  marginLeft: 120,
};

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

function halfPointToPx(value?: string) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue)) return undefined;
  return (numberValue / 2) * (96 / 72);
}

function parseHexColor(value?: string) {
  if (!value || value === 'auto') return undefined;
  return value.startsWith('#') ? value : `#${value}`;
}

function readTextStyle(rPr: Element | null | undefined): DocxTextStyle | undefined {
  if (!rPr) return undefined;

  const color = parseHexColor(attr(childByLocalName(rPr, 'color'), 'w:val') ?? attr(childByLocalName(rPr, 'color'), 'val'));
  const style: DocxTextStyle = {
    bold: Boolean(childByLocalName(rPr, 'b') || childByLocalName(rPr, 'bCs')),
    italic: Boolean(childByLocalName(rPr, 'i') || childByLocalName(rPr, 'iCs')),
    underline: Boolean(childByLocalName(rPr, 'u')),
    color,
    fontSize: halfPointToPx(attr(childByLocalName(rPr, 'sz'), 'w:val') ?? attr(childByLocalName(rPr, 'sz'), 'val')),
  };

  const cleaned = Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined && value !== false),
  ) as DocxTextStyle;
  return Object.keys(cleaned).length ? cleaned : undefined;
}

function mergeTextStyle(base?: DocxTextStyle, next?: DocxTextStyle): DocxTextStyle | undefined {
  const merged = { ...base, ...next };
  return Object.keys(merged).length ? merged : undefined;
}

function mapAlignment(value?: string) {
  if (value === 'center' || value === 'right' || value === 'justify') return value;
  if (value === 'both') return 'justify';
  return 'left';
}

function readParagraphStyle(pPr: Element | null | undefined) {
  const jc = childByLocalName(pPr, 'jc');
  const spacing = childByLocalName(pPr, 'spacing');
  const ind = childByLocalName(pPr, 'ind');
  return {
    align: mapAlignment(attr(jc, 'w:val') ?? attr(jc, 'val')),
    spacingBefore: twipToPx(attr(spacing, 'w:before') ?? attr(spacing, 'before')),
    spacingAfter: twipToPx(attr(spacing, 'w:after') ?? attr(spacing, 'after')),
    indentLeft: twipToPx(attr(ind, 'w:left') ?? attr(ind, 'left')),
    style: readTextStyle(childByLocalName(pPr, 'rPr')),
  };
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

function parseRun(runNode: Element, paragraphStyle: DocxTextStyle | undefined, context: ParseContext): DocxInline[] {
  const runStyle = mergeTextStyle(paragraphStyle, readTextStyle(childByLocalName(runNode, 'rPr')));
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
      const image = parseDrawingImage(child, context);
      if (image) {
        inlines.push({ type: 'image', image });
      }
    }
  });

  return inlines;
}

function readParagraphBlocks(pNode: Element, id: string, context: ParseContext): DocxBlock[] {
  const paragraph = parseParagraph(pNode, id, context);
  const chartBlocks = childrenByLocalName(pNode, 'r')
    .flatMap((runNode) => childrenByLocalName(runNode, 'drawing'))
    .map((drawingNode) => parseChartElement(drawingNode, context))
    .filter(Boolean) as DocxChartBlock[];

  if (!paragraph.text && !paragraph.inlines.length) {
    return chartBlocks;
  }

  return chartBlocks.length ? [paragraph, ...chartBlocks] : [paragraph];
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
  const style = readParagraphStyle(childByLocalName(pNode, 'pPr'));
  const inlines = readParagraphRuns(pNode, style.style, context);
  const text = textFromInlines(inlines).trim();

  return {
    id,
    type: 'paragraph',
    inlines,
    text,
    align: style.align,
    style: style.style,
    spacingBefore: style.spacingBefore,
    spacingAfter: style.spacingAfter,
    indentLeft: style.indentLeft,
  };
}

function readCellStyle(tcNode: Element): Pick<DocxTableCell, 'colSpan' | 'width' | 'verticalAlign' | 'backgroundColor'> {
  const tcPr = childByLocalName(tcNode, 'tcPr');
  const gridSpan = childByLocalName(tcPr, 'gridSpan');
  const width = childByLocalName(tcPr, 'tcW');
  const vAlign = attr(childByLocalName(tcPr, 'vAlign'), 'w:val') ?? attr(childByLocalName(tcPr, 'vAlign'), 'val');
  const shading = childByLocalName(tcPr, 'shd');
  return {
    colSpan: Number(attr(gridSpan, 'w:val') ?? attr(gridSpan, 'val') ?? 1),
    width: twipToPx(attr(width, 'w:w') ?? attr(width, 'w')),
    verticalAlign: vAlign === 'center' ? 'middle' : vAlign === 'bottom' ? 'bottom' : 'top',
    backgroundColor: parseHexColor(attr(shading, 'w:fill') ?? attr(shading, 'fill')),
  };
}

function parseTable(tblNode: Element, id: string, context: ParseContext): DocxTableBlock {
  return {
    id,
    type: 'table',
    rows: childrenByLocalName(tblNode, 'tr').map((rowNode, rowIndex) => ({
      id: `${id}-row-${rowIndex + 1}`,
      cells: childrenByLocalName(rowNode, 'tc').map((cellNode, cellIndex): DocxTableCell => ({
        id: `${id}-cell-${rowIndex + 1}-${cellIndex + 1}`,
        ...readCellStyle(cellNode),
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

  return {
    width: Math.round(twipToPx(attr(pgSz, 'w:w') ?? attr(pgSz, 'w')) ?? DEFAULT_PAGE.width),
    minHeight: Math.round(twipToPx(attr(pgSz, 'w:h') ?? attr(pgSz, 'h')) ?? DEFAULT_PAGE.minHeight),
    marginTop: Math.round(twipToPx(attr(pgMar, 'w:top') ?? attr(pgMar, 'top')) ?? DEFAULT_PAGE.marginTop),
    marginRight: Math.round(twipToPx(attr(pgMar, 'w:right') ?? attr(pgMar, 'right')) ?? DEFAULT_PAGE.marginRight),
    marginBottom: Math.round(twipToPx(attr(pgMar, 'w:bottom') ?? attr(pgMar, 'bottom')) ?? DEFAULT_PAGE.marginBottom),
    marginLeft: Math.round(twipToPx(attr(pgMar, 'w:left') ?? attr(pgMar, 'left')) ?? DEFAULT_PAGE.marginLeft),
  };
}

function markTitle(blocks: DocxBlock[]) {
  const firstParagraph = blocks.find(
    (block): block is DocxParagraphBlock => block.type === 'paragraph' && Boolean(block.text),
  );
  if (firstParagraph) {
    firstParagraph.isTitle = true;
  }
  return firstParagraph?.text ?? 'DOCX 文档';
}

export async function parseDocx(file: File): Promise<DocxDocument> {
  const entries = await loadDocxEntries(file);
  const packageState = buildPackageState(entries);
  const documentXml = readXml(entries, 'word/document.xml');
  const documentDoc = parseXml(documentXml);
  const bodyNode = childByLocalName(documentDoc.documentElement, 'body');
  const context: ParseContext = {
    packageState,
    documentRels: packageState.relationships['word/_rels/document.xml.rels'] ?? {},
    theme: readOfficeTheme(readXml(entries, 'word/theme/theme1.xml')),
    images: [],
    imageIndex: 0,
    chartIndex: 0,
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
