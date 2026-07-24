import type { TextStyle } from '../../presentation/types';

export type PptOfficeArtProperty = {
  id: number;
  value: number;
  isBlip: boolean;
  complexData?: Uint8Array;
};

export type PptShapeStyle = {
  fill?: string | null;
  stroke?: string | null;
  strokeWidth?: number;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  textStyle?: TextStyle;
};

export type PptShapeAnchor = {
  x: number;
  y: number;
  width: number;
  height: number;
};
