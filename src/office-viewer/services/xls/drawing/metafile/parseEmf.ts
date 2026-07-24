import type { SpreadsheetWarning } from '../../../spreadsheet/types';
import { XlsParseError } from '../../errors';
import type { VectorElement, VectorScene, VectorStyle } from './types';

type EmfObject =
  | { kind: 'pen'; color: string; width: number }
  | { kind: 'brush'; color?: string };

function colorRef(value: number) {
  return `#${[value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff]
    .map((component) => component.toString(16).padStart(2, '0'))
    .join('')}`;
}

function readPoint(view: DataView, offset: number): [number, number] {
  return [view.getInt32(offset, true), view.getInt32(offset + 4, true)];
}

function readRect(view: DataView, offset: number) {
  return {
    left: view.getInt32(offset, true),
    top: view.getInt32(offset + 4, true),
    right: view.getInt32(offset + 8, true),
    bottom: view.getInt32(offset + 12, true),
  };
}

function addRectangle(
  elements: VectorElement[],
  rectangle: ReturnType<typeof readRect>,
  style: VectorStyle,
  type: 'rectangle' | 'ellipse',
) {
  elements.push({
    type,
    x: Math.min(rectangle.left, rectangle.right),
    y: Math.min(rectangle.top, rectangle.bottom),
    width: Math.abs(rectangle.right - rectangle.left),
    height: Math.abs(rectangle.bottom - rectangle.top),
    style: { ...style },
  });
}

function warnUnknown(
  warnings: SpreadsheetWarning[],
  unknown: Set<number>,
  type: number,
  offset: number,
) {
  if (unknown.has(type)) return;
  unknown.add(type);
  warnings.push({
    code: 'UNKNOWN_EMF_OPCODE',
    message: `已跳过 EMF 指令 0x${type.toString(16).toUpperCase()}`,
    offset,
  });
}

/** 解释常用 EMF 图元，未知记录按其已校验长度局部跳过。 */
export function parseEmf(bytes: Uint8Array): VectorScene {
  if (bytes.length < 88) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'EMF Header 长度不足');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerSize = view.getUint32(4, true);
  if (
    view.getUint32(0, true) !== 1 ||
    headerSize < 88 ||
    headerSize > bytes.length ||
    view.getUint32(40, true) !== 0x464d4520
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'EMF Header 无效');
  }
  const bounds = readRect(view, 8);
  const width = Math.max(1, Math.abs(bounds.right - bounds.left));
  const height = Math.max(1, Math.abs(bounds.bottom - bounds.top));
  const elements: VectorElement[] = [];
  const warnings: SpreadsheetWarning[] = [];
  const unknown = new Set<number>();
  const objects = new Map<number, EmfObject>();
  const style: VectorStyle = {
    stroke: '#000000',
    fill: 'none',
    strokeWidth: 1,
    textColor: '#000000',
  };
  const savedStyles: VectorStyle[] = [];
  let current: [number, number] = [0, 0];
  let offset = headerSize;

  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || size % 4 !== 0 || offset + size > bytes.length) {
      throw new XlsParseError(
        'INVALID_RECORD_DATA',
        `EMF 记录 0x${type.toString(16)} 长度无效`,
        { offset },
      );
    }
    if (type === 14) break;
    switch (type) {
      case 33:
        savedStyles.push({ ...style });
        break;
      case 34: {
        const restored = savedStyles.pop();
        if (restored) Object.assign(style, restored);
        break;
      }
      case 24:
        style.textColor = colorRef(view.getUint32(offset + 8, true));
        break;
      case 38: {
        const handle = view.getUint32(offset + 8, true);
        objects.set(handle, {
          kind: 'pen',
          width: Math.max(1, Math.abs(view.getInt32(offset + 16, true))),
          color: colorRef(view.getUint32(offset + 28, true)),
        });
        break;
      }
      case 39: {
        const handle = view.getUint32(offset + 8, true);
        const brushStyle = view.getUint32(offset + 12, true);
        objects.set(handle, {
          kind: 'brush',
          color:
            brushStyle === 1
              ? undefined
              : colorRef(view.getUint32(offset + 16, true)),
        });
        break;
      }
      case 37: {
        const handle = view.getUint32(offset + 8, true);
        if (handle === 0x80000005) {
          style.fill = 'none';
          break;
        }
        if (handle === 0x80000008) {
          style.stroke = 'none';
          break;
        }
        const object = objects.get(handle);
        if (object?.kind === 'pen') {
          style.stroke = object.color;
          style.strokeWidth = object.width;
        } else if (object?.kind === 'brush') {
          style.fill = object.color ?? 'none';
        }
        break;
      }
      case 40:
        objects.delete(view.getUint32(offset + 8, true));
        break;
      case 27:
        current = readPoint(view, offset + 8);
        break;
      case 54: {
        const point = readPoint(view, offset + 8);
        elements.push({
          type: 'line',
          x1: current[0],
          y1: current[1],
          x2: point[0],
          y2: point[1],
          style: { ...style },
        });
        current = point;
        break;
      }
      case 2:
      case 3:
      case 4:
      case 5:
      case 6: {
        const count = view.getUint32(offset + 24, true);
        const points: Array<[number, number]> = [];
        for (let index = 0; index < count; index += 1) {
          const pointOffset = offset + 28 + index * 8;
          if (pointOffset + 8 > offset + size) {
            throw new XlsParseError(
              'TRUNCATED_RECORD',
              'EMF 多边形点数据被截断',
              { offset },
            );
          }
          points.push(readPoint(view, pointOffset));
        }
        elements.push({
          type: type === 3 ? 'polygon' : 'polyline',
          points: type === 5 || type === 6 ? [current, ...points] : points,
          style: { ...style },
        });
        if (points.length) current = points[points.length - 1];
        break;
      }
      case 42:
        addRectangle(elements, readRect(view, offset + 8), style, 'ellipse');
        break;
      case 43:
        addRectangle(elements, readRect(view, offset + 8), style, 'rectangle');
        break;
      case 44: {
        const before = elements.length;
        addRectangle(elements, readRect(view, offset + 8), style, 'rectangle');
        const element = elements[before];
        if (element?.type === 'rectangle') {
          element.radiusX = Math.abs(view.getInt32(offset + 24, true)) / 2;
          element.radiusY = Math.abs(view.getInt32(offset + 28, true)) / 2;
        }
        break;
      }
      case 45:
      case 46:
      case 47:
        // 端点仍保留在原记录中；当前以完整椭圆近似，避免整幅图消失。
        addRectangle(elements, readRect(view, offset + 8), style, 'ellipse');
        if (!unknown.has(type)) {
          unknown.add(type);
          warnings.push({
            code: 'APPROXIMATED_EMF_ARC',
            message: `EMF ${
              type === 45 ? 'Arc' : type === 46 ? 'Chord' : 'Pie'
            } 已按椭圆近似`,
            offset,
          });
        }
        break;
      case 84: {
        if (size < 76) {
          throw new XlsParseError(
            'TRUNCATED_RECORD',
            'EMF ExtTextOutW 记录长度不足',
            { offset },
          );
        }
        const characterCount = view.getUint32(offset + 44, true);
        const stringOffset = view.getUint32(offset + 48, true);
        const absoluteStringOffset = offset + stringOffset;
        if (absoluteStringOffset + characterCount * 2 > offset + size) {
          throw new XlsParseError(
            'TRUNCATED_RECORD',
            'EMF ExtTextOutW 文本数据被截断',
            { offset },
          );
        }
        let text = '';
        for (let index = 0; index < characterCount; index += 1) {
          text += String.fromCharCode(
            view.getUint16(absoluteStringOffset + index * 2, true),
          );
        }
        const [x, y] = readPoint(view, offset + 36);
        elements.push({
          type: 'text',
          x,
          y,
          text,
          style: { ...style },
        });
        break;
      }
      case 9:
      case 10:
      case 11:
      case 12:
      case 17:
      case 18:
      case 19:
      case 20:
      case 21:
      case 22:
      case 25:
      case 29:
      case 30:
      case 35:
      case 36:
      case 57:
      case 58:
      case 59:
      case 60:
      case 61:
      case 62:
      case 63:
      case 64:
      case 67:
      case 68:
        break;
      default:
        warnUnknown(warnings, unknown, type, offset);
    }
    offset += size;
  }
  return {
    width,
    height,
    viewBox: [bounds.left, bounds.top, width, height],
    elements,
    warnings,
  };
}
