export type DocxDocument = {
  title: string;
  page: DocxPage;
  blocks: DocxBlock[];
  images: DocxImage[];
};

export type DocxPage = {
  width: number;
  minHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
};

export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxChartBlock;

export type DocxParagraphBlock = {
  id: string;
  type: 'paragraph';
  inlines: DocxInline[];
  text: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  lineHeight?: number;
  style?: DocxTextStyle;
  spacingAfter?: number;
  spacingBefore?: number;
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  backgroundColor?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  hasBorderTop?: boolean;
  hasBorderRight?: boolean;
  hasBorderBottom?: boolean;
  hasBorderLeft?: boolean;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
};

export type DocxTableBlock = {
  id: string;
  type: 'table';
  rows: DocxTableRow[];
  width?: number;
  align?: 'left' | 'center' | 'right';
  columns?: number[];
};

export type DocxChartBlock = {
  id: string;
  type: 'chart';
  chart: import('../../shared/ooxml/charts').OfficeChartModel;
  width: number;
  height: number;
};

export type DocxTableRow = {
  id: string;
  cells: DocxTableCell[];
};

export type DocxTableCell = {
  id: string;
  blocks: Array<DocxParagraphBlock | DocxChartBlock>;
  colSpan?: number;
  width?: number;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  backgroundColor?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  hasBorderTop?: boolean;
  hasBorderRight?: boolean;
  hasBorderBottom?: boolean;
  hasBorderLeft?: boolean;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  noWrap?: boolean;
};

export type DocxInline = DocxTextInline | DocxBreakInline | DocxImageInline | DocxChartInline | DocxShapeInline;

export type DocxTextInline = {
  type: 'text';
  text: string;
  style?: DocxTextStyle;
};

export type DocxBreakInline = {
  type: 'break';
};

export type DocxImageInline = {
  type: 'image';
  image: DocxImage;
};

export type DocxChartInline = {
  type: 'chart';
  chart: DocxChartBlock;
};

export type DocxShapeInline = {
  type: 'shape';
  shape: DocxShape;
};

export type DocxShape = {
  id: string;
  width: number;
  height: number;
  items: DocxShapeItem[];
};

export type DocxShapeItem = {
  id: string;
  kind: 'rect' | 'ellipse' | 'line' | 'path';
  left: number;
  top: number;
  width: number;
  height: number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  path?: string;
  viewBox?: string;
  fillColor?: string;
  border?: string;
  strokeColor?: string;
  strokeWidth?: number;
  strokeDasharray?: string;
  borderRadius?: number | string;
  textVerticalAlign?: 'top' | 'middle' | 'bottom';
  paragraphs?: DocxParagraphBlock[];
};

export type DocxImage = {
  id: string;
  name?: string;
  alt?: string;
  src: string;
  width: number;
  height: number;
};

export type DocxTextStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  lineHeight?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  backgroundColor?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
};
