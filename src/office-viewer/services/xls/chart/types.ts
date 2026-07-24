import type {
  OfficeChartModel,
  OfficeChartType,
} from '../../../shared/ooxml/charts';
import type { SpreadsheetWarning } from '../../spreadsheet/types';
import type { Biff8Record } from '../biff8/Biff8Reader';
import type { Biff8Anchor } from '../drawing/types';
import type {
  Biff8ChartSubstream,
  Biff8SheetDescriptor,
  Biff8Workbook,
  Biff8Worksheet,
} from '../types';

export type Biff8ChartRecordNode = {
  id: number;
  offset: number;
  data: Uint8Array;
  children: Biff8ChartRecordNode[];
};

export type Biff8ChartSeries = {
  name: string;
  groupIndex: number;
  type?: OfficeChartType;
  categories: Array<string | number | null>;
  values: Array<number | null>;
  bubbleSizes: Array<number | null>;
  stacking?: 'stacked' | 'percentStacked';
  color?: string;
  marker?: { symbol?: string; size?: number };
  lineWidth?: number;
};

export type Biff8ChartModel = {
  id: string;
  sourceType: string;
  groupTypes: string[];
  is3d: boolean;
  hasSecondaryAxis: boolean;
  title?: string;
  categories: string[];
  series: Biff8ChartSeries[];
  showLegend: boolean;
  legendPosition?: 'top' | 'bottom' | 'left' | 'right';
  gapWidth?: number;
  overlap?: number;
  holeSize?: number;
  anchor: Biff8Anchor;
  previewImageSrc?: string;
  warnings: SpreadsheetWarning[];
};

export type Biff8ChartContext = {
  substream: Biff8ChartSubstream;
  workbook: Biff8Workbook;
  sourceSheet?: Biff8Worksheet;
  descriptor: Biff8SheetDescriptor;
  chartIndex: number;
  anchor?: Biff8Anchor;
  previewImageSrc?: string;
};

export type AdaptedBiff8Chart = {
  chart: OfficeChartModel;
  renderMode: 'interactive' | 'snapshot';
  degradedFrom?: string;
  warnings: SpreadsheetWarning[];
};

export type Biff8ChartCache = Map<
  number,
  Map<number, Map<number, string | number | boolean | null>>
>;

export type ChartStreamReadResult = {
  substream: Biff8ChartSubstream;
  records: Biff8Record[];
};
