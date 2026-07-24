import type { SpreadsheetWarning } from '../../../spreadsheet/types';
import { XlsParseError } from '../../errors';
import type { VectorElement, VectorScene, VectorStyle } from './types';

type WmfObject =
  | { kind: 'pen'; color: string; width: number }
  | { kind: 'brush'; color?: string }
  | { kind: 'font'; family: string; size: number; weight: number };

type WmfState = {
  current: [number, number];
  windowOrigin: [number, number];
  windowExtent: [number, number];
  viewportOrigin: [number, number];
  viewportExtent: [number, number];
  style: VectorStyle;
  textColor: string;
};

function colorRef(value: number) {
  return `#${[value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff]
    .map((component) => component.toString(16).padStart(2, '0'))
    .join('')}`;
}

function cloneState(state: WmfState): WmfState {
  return {
    ...state,
    current: [...state.current],
    windowOrigin: [...state.windowOrigin],
    windowExtent: [...state.windowExtent],
    viewportOrigin: [...state.viewportOrigin],
    viewportExtent: [...state.viewportExtent],
    style: { ...state.style },
  };
}

function mapPoint(state: WmfState, x: number, y: number): [number, number] {
  const scaleX = state.windowExtent[0]
    ? state.viewportExtent[0] / state.windowExtent[0]
    : 1;
  const scaleY = state.windowExtent[1]
    ? state.viewportExtent[1] / state.windowExtent[1]
    : 1;
  return [
    (x - state.windowOrigin[0]) * scaleX + state.viewportOrigin[0],
    (y - state.windowOrigin[1]) * scaleY + state.viewportOrigin[1],
  ];
}

function readPoint(view: DataView, offset: number): [number, number] {
  return [view.getInt16(offset + 2, true), view.getInt16(offset, true)];
}

function readRect(view: DataView, offset: number) {
  const bottom = view.getInt16(offset, true);
  const right = view.getInt16(offset + 2, true);
  const top = view.getInt16(offset + 4, true);
  const left = view.getInt16(offset + 6, true);
  return { left, top, right, bottom };
}

function addRectangle(
  elements: VectorElement[],
  state: WmfState,
  rectangle: ReturnType<typeof readRect>,
  type: 'rectangle' | 'ellipse',
) {
  const [left, top] = mapPoint(state, rectangle.left, rectangle.top);
  const [right, bottom] = mapPoint(state, rectangle.right, rectangle.bottom);
  elements.push({
    type,
    x: Math.min(left, right),
    y: Math.min(top, bottom),
    width: Math.abs(right - left),
    height: Math.abs(bottom - top),
    style: { ...state.style },
  });
}

/** 解释常用 WMF 绘图记录，未知指令仅产生去重 warning。 */
export function parseWmf(bytes: Uint8Array): VectorScene {
  if (bytes.length < 18) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'WMF Header 长度不足');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let headerOffset = 0;
  let bounds = { left: 0, top: 0, right: 1000, bottom: 1000 };
  if (view.getUint32(0, true) === 0x9ac6cdd7) {
    if (bytes.length < 40) {
      throw new XlsParseError(
        'INVALID_RECORD_DATA',
        'WMF Placeable Header 无效',
      );
    }
    bounds = {
      left: view.getInt16(6, true),
      top: view.getInt16(8, true),
      right: view.getInt16(10, true),
      bottom: view.getInt16(12, true),
    };
    headerOffset = 22;
  }
  if (
    view.getUint16(headerOffset + 2, true) !== 9 ||
    headerOffset + 18 > bytes.length
  ) {
    throw new XlsParseError('INVALID_RECORD_DATA', 'WMF 标准 Header 无效');
  }

  const width = Math.max(1, Math.abs(bounds.right - bounds.left));
  const height = Math.max(1, Math.abs(bounds.bottom - bounds.top));
  const warnings: SpreadsheetWarning[] = [];
  const unknown = new Set<number>();
  const elements: VectorElement[] = [];
  const objects: Array<WmfObject | undefined> = [];
  const state: WmfState = {
    current: [0, 0],
    windowOrigin: [bounds.left, bounds.top],
    windowExtent: [width, height],
    viewportOrigin: [0, 0],
    viewportExtent: [width, height],
    style: { stroke: '#000000', fill: 'none', strokeWidth: 1 },
    textColor: '#000000',
  };
  const savedStates: WmfState[] = [];
  let offset = headerOffset + 18;

  const addObject = (object: WmfObject) => {
    const freeIndex = objects.findIndex((item) => item === undefined);
    if (freeIndex >= 0) {
      objects[freeIndex] = object;
      return freeIndex;
    }
    objects.push(object);
    return objects.length - 1;
  };
  while (offset + 6 <= bytes.length) {
    const sizeWords = view.getUint32(offset, true);
    const functionId = view.getUint16(offset + 4, true);
    const recordSize = sizeWords * 2;
    if (sizeWords < 3 || offset + recordSize > bytes.length) {
      throw new XlsParseError(
        'INVALID_RECORD_DATA',
        `WMF 记录 0x${functionId.toString(16)} 长度越界`,
        { offset },
      );
    }
    const parameterOffset = offset + 6;
    if (functionId === 0x0000) break;
    switch (functionId) {
      case 0x001e:
        savedStates.push(cloneState(state));
        break;
      case 0x0127: {
        const restored = savedStates.pop();
        if (restored) Object.assign(state, restored);
        break;
      }
      case 0x0209:
        state.textColor = colorRef(view.getUint32(parameterOffset, true));
        break;
      case 0x020b:
        state.windowOrigin = readPoint(view, parameterOffset);
        break;
      case 0x020c:
        state.windowExtent = readPoint(view, parameterOffset);
        break;
      case 0x020d:
        state.viewportOrigin = readPoint(view, parameterOffset);
        break;
      case 0x020e:
        state.viewportExtent = readPoint(view, parameterOffset);
        break;
      case 0x02fa: {
        const penStyle = view.getUint16(parameterOffset, true) & 0x0f;
        const objectIndex = addObject({
          kind: 'pen',
          color: colorRef(view.getUint32(parameterOffset + 6, true)),
          width: Math.max(
            1,
            Math.abs(view.getInt16(parameterOffset + 2, true)),
          ),
        });
        if (penStyle === 5) {
          const object = objects[objectIndex];
          if (object?.kind === 'pen') object.color = 'none';
        }
        break;
      }
      case 0x02fc: {
        const brushStyle = view.getUint16(parameterOffset, true);
        addObject({
          kind: 'brush',
          color:
            brushStyle === 1
              ? undefined
              : colorRef(view.getUint32(parameterOffset + 2, true)),
        });
        break;
      }
      case 0x02fb: {
        const size = Math.abs(view.getInt16(parameterOffset, true)) || 12;
        const weight = view.getInt16(parameterOffset + 8, true) || 400;
        const nameBytes = bytes.subarray(
          parameterOffset + 18,
          Math.min(offset + recordSize, parameterOffset + 50),
        );
        const zero = nameBytes.indexOf(0);
        const family = new TextDecoder('windows-1252').decode(
          zero >= 0 ? nameBytes.subarray(0, zero) : nameBytes,
        );
        addObject({ kind: 'font', family, size, weight });
        break;
      }
      case 0x012d: {
        const object = objects[view.getUint16(parameterOffset, true)];
        if (object?.kind === 'pen') {
          state.style.stroke = object.color;
          state.style.strokeWidth = object.width;
        } else if (object?.kind === 'brush') {
          state.style.fill = object.color ?? 'none';
        } else if (object?.kind === 'font') {
          state.style.fontFamily = object.family;
          state.style.fontSize = object.size;
          state.style.fontWeight = object.weight;
        }
        break;
      }
      case 0x01f0:
        objects[view.getUint16(parameterOffset, true)] = undefined;
        break;
      case 0x0214:
        state.current = mapPoint(state, ...readPoint(view, parameterOffset));
        break;
      case 0x0213: {
        const point = mapPoint(state, ...readPoint(view, parameterOffset));
        elements.push({
          type: 'line',
          x1: state.current[0],
          y1: state.current[1],
          x2: point[0],
          y2: point[1],
          style: { ...state.style },
        });
        state.current = point;
        break;
      }
      case 0x0324:
      case 0x0325: {
        const count = view.getUint16(parameterOffset, true);
        const points: Array<[number, number]> = [];
        for (let index = 0; index < count; index += 1) {
          const pointOffset = parameterOffset + 2 + index * 4;
          if (pointOffset + 4 > offset + recordSize) {
            throw new XlsParseError(
              'TRUNCATED_RECORD',
              'WMF 多边形点数据被截断',
              { offset },
            );
          }
          points.push(
            mapPoint(
              state,
              view.getInt16(pointOffset, true),
              view.getInt16(pointOffset + 2, true),
            ),
          );
        }
        elements.push({
          type: functionId === 0x0324 ? 'polygon' : 'polyline',
          points,
          style: { ...state.style },
        });
        break;
      }
      case 0x041b:
        addRectangle(
          elements,
          state,
          readRect(view, parameterOffset),
          'rectangle',
        );
        break;
      case 0x0418:
        addRectangle(
          elements,
          state,
          readRect(view, parameterOffset),
          'ellipse',
        );
        break;
      case 0x061c: {
        const radiusY = Math.abs(view.getInt16(parameterOffset, true)) / 2;
        const radiusX = Math.abs(view.getInt16(parameterOffset + 2, true)) / 2;
        const rectangle = readRect(view, parameterOffset + 4);
        const before = elements.length;
        addRectangle(elements, state, rectangle, 'rectangle');
        const element = elements[before];
        if (element?.type === 'rectangle') {
          element.radiusX = radiusX;
          element.radiusY = radiusY;
        }
        break;
      }
      case 0x0521: {
        const count = view.getUint16(parameterOffset, true);
        const textOffset = parameterOffset + 2;
        const coordinateOffset = textOffset + count + (count % 2);
        if (coordinateOffset + 4 > offset + recordSize) {
          throw new XlsParseError(
            'TRUNCATED_RECORD',
            'WMF TextOut 数据被截断',
            {
              offset,
            },
          );
        }
        const text = new TextDecoder('windows-1252').decode(
          bytes.subarray(textOffset, textOffset + count),
        );
        const [x, y] = mapPoint(
          state,
          view.getInt16(coordinateOffset + 2, true),
          view.getInt16(coordinateOffset, true),
        );
        elements.push({
          type: 'text',
          x,
          y,
          text,
          style: { ...state.style, textColor: state.textColor },
        });
        break;
      }
      default:
        if (!unknown.has(functionId)) {
          unknown.add(functionId);
          warnings.push({
            code: 'UNKNOWN_WMF_OPCODE',
            message: `已跳过 WMF 指令 0x${functionId
              .toString(16)
              .toUpperCase()}`,
            offset,
          });
        }
    }
    offset += recordSize;
  }
  return {
    width,
    height,
    viewBox: [0, 0, width, height],
    elements,
    warnings,
  };
}
