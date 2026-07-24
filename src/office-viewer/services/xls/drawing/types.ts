import type { SpreadsheetWarning } from '../../spreadsheet/types';
export type { OfficeArtRecord } from '../../../shared/officeart';

export type Biff8AnchorPoint = {
  row: number;
  column: number;
  rowFraction: number;
  columnFraction: number;
};

export type Biff8Anchor = {
  from: Biff8AnchorPoint;
  to: Biff8AnchorPoint;
};

export type Biff8DrawingImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

export type Biff8DrawingImage = {
  id: string;
  name?: string;
  format: Biff8DrawingImageFormat;
  bytes: Uint8Array;
  anchor: Biff8Anchor;
  alt?: string;
  compressed?: boolean;
  warnings: SpreadsheetWarning[];
};

export type Biff8DrawingShape = {
  id: string;
  shapeId?: number;
  shapeType?: number;
  name?: string;
  blipIndex?: number;
  anchor: Biff8Anchor;
};

export type DecodedBitmap = {
  width: number;
  height: number;
  rgba: Uint8ClampedArray;
};

export type ParsedBlip = {
  index: number;
  format: Biff8DrawingImageFormat;
  bytes: Uint8Array;
  compressed?: boolean;
  warnings: SpreadsheetWarning[];
};
