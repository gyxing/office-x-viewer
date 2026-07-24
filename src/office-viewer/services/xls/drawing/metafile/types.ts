import type { SpreadsheetWarning } from '../../../spreadsheet/types';

export type VectorStyle = {
  stroke?: string;
  fill?: string;
  strokeWidth?: number;
  opacity?: number;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: number;
  textColor?: string;
};

export type VectorElement =
  | {
      type: 'line';
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      style: VectorStyle;
    }
  | {
      type: 'polyline' | 'polygon';
      points: Array<[number, number]>;
      style: VectorStyle;
    }
  | {
      type: 'rectangle' | 'ellipse';
      x: number;
      y: number;
      width: number;
      height: number;
      radiusX?: number;
      radiusY?: number;
      style: VectorStyle;
    }
  | {
      type: 'path';
      data: string;
      style: VectorStyle;
    }
  | {
      type: 'text';
      x: number;
      y: number;
      text: string;
      style: VectorStyle;
    }
  | {
      type: 'image';
      x: number;
      y: number;
      width: number;
      height: number;
      dataUrl: string;
      style: VectorStyle;
    };

export type VectorScene = {
  width: number;
  height: number;
  viewBox: [number, number, number, number];
  elements: VectorElement[];
  warnings: SpreadsheetWarning[];
};
