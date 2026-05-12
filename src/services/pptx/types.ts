export type PptxDocument = {
  width: number;
  height: number;
  theme: ThemeModel;
  slides: SlideModel[];
};

export type ThemeModel = {
  colorScheme: Record<string, string>;
  fontScheme: Record<string, string>;
  colorMap?: Record<string, string>;
};

export type GradientStop = {
  offset: number;
  color: string;
};

export type GradientFill = {
  type: 'linear';
  angle: number;
  stops: GradientStop[];
};

export type SlideModel = {
  id: string;
  index: number;
  width: number;
  height: number;
  background?: SlideBackground;
  elements: SlideElement[];
};

export type SlideBackground = {
  fill?: string;
  fillOpacity?: number;
  imageRef?: string;
};

export type TextStyle = {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: 'none' | 'sngStrike' | 'dblStrike';
  smallCaps?: boolean;
  allCaps?: boolean;
  color?: string;
  opacity?: number;
  align?: 'left' | 'center' | 'right' | 'justify';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  writingMode?: 'horizontal-tb' | 'vertical-rl' | 'vertical-lr';
  fit?: 'none' | 'shrinkText' | 'resizeShape';
  marginLeft?: number;
  marginRight?: number;
  marginTop?: number;
  marginBottom?: number;
  lineHeight?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  textIndent?: number;
  charSpace?: number;
  baseline?: number;
  bullet?: TextBulletStyle;
};

export type TextBulletStyle = {
  char?: string;
  color?: string;
  size?: number;
  none?: boolean;
};

export type TextRun = {
  text: string;
  style?: TextStyle;
};

export type TextParagraph = {
  runs: TextRun[];
  style?: TextStyle;
  level?: number;
  bullet?: TextBulletStyle;
};

export type BaseElement = {
  id: string;
  type: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  zIndex?: number;
  opacity?: number;
  placeholderType?: string;
  placeholderIdx?: string;
};

export type TextElement = BaseElement & {
  type: 'text';
  paragraphs: TextParagraph[];
  boxStyle?: TextStyle;
  shape?: string;
  path?: string;
  viewBox?: string;
  fill?: string | GradientFill | null;
  fillOpacity?: number;
  stroke?: string | null;
  strokeOpacity?: number;
  strokeWidth?: number;
  strokeDash?: string;
  shadow?: ShadowStyle;
  borderRadius?: number;
};

export type ShapeElement = BaseElement & {
  type: 'shape';
  shape: string;
  path?: string;
  viewBox?: string;
  fill?: string | GradientFill | null;
  fillOpacity?: number;
  stroke?: string | null;
  strokeOpacity?: number;
  strokeWidth?: number;
  strokeDash?: string;
  shadow?: ShadowStyle;
  borderRadius?: number;
};

export type ImageElement = BaseElement & {
  type: 'image';
  src: string;
  alt?: string;
  crop?: ImageCrop;
};

export type TableElement = BaseElement & {
  type: 'table';
  columnWidths?: number[];
  rowHeights?: number[];
  rows: TableCell[][];
};

export type TableCell = {
  text: string;
  paragraphs?: TextParagraph[];
  style?: TextStyle;
  backgroundColor?: string | null;
  backgroundOpacity?: number;
  borderColor?: string | null;
  borderOpacity?: number;
  borderWidth?: number;
  margins?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  };
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

export type GroupElement = BaseElement & {
  type: 'group';
  children: SlideElement[];
};

export type UnsupportedElement = BaseElement & {
  type: 'unsupported';
  reason: string;
};

export type ImageCrop = {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
};

export type ShadowStyle = {
  color?: string;
  opacity?: number;
  blur?: number;
  offsetX?: number;
  offsetY?: number;
};

export type SlideElement = 
  | TextElement
  | ShapeElement
  | ImageElement
  | ChartElement
  | TableElement
  | GroupElement
  | UnsupportedElement;

export type ChartElement = BaseElement & {
  type: 'chart';
  chart: import('../office/charts').OfficeChartModel;
  chartId?: string;
  chartPath?: string;
};
