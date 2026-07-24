import type {
  ChartElement,
  ImageElement,
  ShapeElement,
  SlideElement,
  TextElement,
  ThemeModel,
} from '../../presentation/types';
import {
  OFFICE_ART_RECORD,
  parseOfficeArtRecords,
  type OfficeArtRecord,
} from '../../../shared/officeart';
import { PptRecordReader } from '../binary/PptRecordReader';
import { createPptStaticPreviewCard } from '../images';
import { parsePptTextGroups } from '../text';
import type { PptParseContext } from '../types';
import { readPptOfficeArtProperties } from './readOfficeArtProperties';
import { readPptAnchor } from './readPptAnchor';

const SHAPE_NAMES: Record<number, string> = {
  1: 'rect',
  2: 'roundRect',
  3: 'ellipse',
  4: 'diamond',
  5: 'triangle',
  6: 'rtTriangle',
  7: 'parallelogram',
  8: 'trapezoid',
  20: 'line',
  32: 'line',
  33: 'bentConnector2',
  34: 'curvedConnector2',
};

function readColor(value: number) {
  const rgb = [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff];
  return `#${rgb
    .map((part) => part.toString(16).padStart(2, '0'))
    .join('')}`;
}

function findChild(record: OfficeArtRecord, type: number) {
  return record.children?.find((child) => child.type === type);
}

function isBooleanPropertyEnabled(value: number, bit: number) {
  const useMask = bit << 16;
  return value & useMask ? Boolean(value & bit) : undefined;
}

function parseShape(
  record: OfficeArtRecord,
  index: number,
  theme: ThemeModel,
  context: PptParseContext,
): SlideElement | undefined {
  const fsp = findChild(record, OFFICE_ART_RECORD.FSP);
  const anchor = readPptAnchor(
    findChild(record, OFFICE_ART_RECORD.CLIENT_ANCHOR),
  );
  if (!fsp || fsp.data.length < 8 || !anchor || !anchor.width || !anchor.height) {
    return undefined;
  }
  const fspView = new DataView(
    fsp.data.buffer,
    fsp.data.byteOffset,
    fsp.data.byteLength,
  );
  const shapeId = fspView.getUint32(0, true);
  const flags = fspView.getUint32(4, true);
  if (flags & 0x0008 || flags & 0x0400) return undefined;

  const properties = readPptOfficeArtProperties(
    findChild(record, OFFICE_ART_RECORD.FOPT),
  );
  const fillFlags = properties.get(0x01bf)?.value;
  const lineFlags = properties.get(0x01ff)?.value;
  const filled =
    fillFlags === undefined
      ? undefined
      : isBooleanPropertyEnabled(fillFlags, 0x0010);
  const lined =
    lineFlags === undefined
      ? undefined
      : isBooleanPropertyEnabled(lineFlags, 0x0008);
  const shapeType = fsp.instance;
  const shape = SHAPE_NAMES[shapeType] ?? 'rect';
  const fillColor = properties.get(0x0181)?.value;
  const lineColor = properties.get(0x01c0)?.value;
  const lineWidth = properties.get(0x01cb)?.value;
  const rotation = properties.get(0x0004)?.value;
  const common = {
    id: `ppt-shape-${shapeId}`,
    x: anchor.x,
    y: anchor.y,
    width: anchor.width,
    height: anchor.height,
    rotate: rotation === undefined ? undefined : rotation / 65536,
    flipH: Boolean(flags & 0x0040),
    flipV: Boolean(flags & 0x0080),
    zIndex: index,
    shape,
    fill:
      filled === false
        ? null
        : fillColor === undefined
        ? null
        : readColor(fillColor),
    stroke:
      lined === false
        ? null
        : lineColor === undefined
        ? null
        : readColor(lineColor),
    strokeWidth: lineWidth === undefined ? undefined : lineWidth / 12700,
  };

  const blipIndex = properties.get(0x0104)?.value;
  const clientData = findChild(record, OFFICE_ART_RECORD.CLIENT_DATA);
  let externalObjectId: number | undefined;
  if (clientData) {
    try {
      for (const clientRecord of new PptRecordReader(clientData.data).records()) {
        if (clientRecord.type === 0x0bc1 && clientRecord.data.length >= 4) {
          externalObjectId = new DataView(
            clientRecord.data.buffer,
            clientRecord.data.byteOffset,
            clientRecord.data.byteLength,
          ).getUint32(0, true);
          break;
        }
      }
    } catch {
      // ClientData 损坏时仍可继续使用同一形状的图片预览。
    }
  }
  const embeddedChart = externalObjectId
    ? context.charts.get(externalObjectId)
    : undefined;
  if (embeddedChart) {
    const chart: ChartElement = {
      id: common.id,
      type: 'chart',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      zIndex: common.zIndex,
      chart: embeddedChart.chart,
      chartId: `ppt-chart-${externalObjectId}`,
    };
    return chart;
  }
  const imageSource =
    shapeType === 75 && blipIndex
      ? context.blipUrls.get(blipIndex)
      : undefined;
  if (imageSource) {
    const image: ImageElement = {
      id: common.id,
      type: 'image',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      flipH: common.flipH,
      flipV: common.flipV,
      zIndex: common.zIndex,
      src: imageSource,
      alt: `PowerPoint 图片 ${blipIndex}`,
    };
    return image;
  }
  if (externalObjectId) {
    const image: ImageElement = {
      id: common.id,
      type: 'image',
      x: common.x,
      y: common.y,
      width: common.width,
      height: common.height,
      rotate: common.rotate,
      zIndex: common.zIndex,
      src: createPptStaticPreviewCard(
        '嵌入对象',
        `PowerPoint 对象 ${externalObjectId}`,
        context,
      ),
      alt: `PowerPoint 嵌入对象 ${externalObjectId}`,
    };
    return image;
  }

  const textbox = findChild(record, OFFICE_ART_RECORD.CLIENT_TEXTBOX);
  if (textbox) {
    const textRecords = Array.from(
      new PptRecordReader(textbox.data).records(),
    );
    const groups = parsePptTextGroups(
      textRecords,
      {
        document: {
          fontFamily: theme.fontScheme.minorLatin ?? 'Arial',
          fontSize: 18,
          color: theme.colorScheme.dk1 ?? '#000000',
        },
      },
      context,
    );
    const paragraphs = groups.flatMap((group) => group.paragraphs);
    if (paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text))) {
      const textType = groups.find((group) => group.paragraphs.length)?.textType;
      const element: TextElement = {
        ...common,
        type: 'text',
        paragraphs,
        placeholderType:
          textType === 0 || textType === 6
            ? 'title'
            : textType === 1 || textType === 5
            ? 'body'
            : undefined,
      };
      return element;
    }
  }

  const element: ShapeElement = {
    ...common,
    type: 'shape',
    fill: common.fill ?? '#ffffff',
    stroke: common.stroke ?? '#000000',
  };
  return element;
}

function collectShapeContainers(records: OfficeArtRecord[]) {
  const shapes: OfficeArtRecord[] = [];
  const visit = (items: OfficeArtRecord[]) => {
    for (const item of items) {
      if (item.type === OFFICE_ART_RECORD.SP_CONTAINER) shapes.push(item);
      else if (item.children) visit(item.children);
    }
  };
  visit(records);
  return shapes;
}

/** 将一页 PPDrawing 转换为统一的文本与基础图形元素。 */
export function parsePptDrawing(
  bytes: Uint8Array,
  theme: ThemeModel,
  context: PptParseContext,
) {
  try {
    const records = parseOfficeArtRecords(bytes, context.warnings);
    return collectShapeContainers(records)
      .map((record, index) => parseShape(record, index, theme, context))
      .filter((element): element is SlideElement => Boolean(element));
  } catch (error) {
    context.warnings.push({
      code: 'PPT_DRAWING_CORRUPT',
      message:
        error instanceof Error ? error.message : 'OfficeArt 绘图记录无法读取',
    });
    return [];
  }
}
