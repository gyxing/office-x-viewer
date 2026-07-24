export type OfficeArtWarning = {
  code: string;
  message: string;
  offset?: number;
};

export type OfficeArtRecord = {
  version: number;
  instance: number;
  type: number;
  length: number;
  offset: number;
  data: Uint8Array;
  children?: OfficeArtRecord[];
};

export type OfficeArtImageFormat =
  | 'png'
  | 'jpeg'
  | 'gif'
  | 'dib'
  | 'wmf'
  | 'emf'
  | 'pict'
  | 'unknown';

export type DecodedBitmap = {
  width: number;
  height: number;
  rgba: Uint8ClampedArray;
};

export type ParsedOfficeArtBlip = {
  index: number;
  format: OfficeArtImageFormat;
  bytes: Uint8Array;
  compressed?: boolean;
  warnings: OfficeArtWarning[];
};
