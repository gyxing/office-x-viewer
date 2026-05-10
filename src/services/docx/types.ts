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
};

export type DocxBlock = DocxParagraphBlock | DocxTableBlock | DocxChartBlock;

export type DocxParagraphBlock = {
  id: string;
  type: 'paragraph';
  inlines: DocxInline[];
  text: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  style?: DocxTextStyle;
  spacingAfter?: number;
  spacingBefore?: number;
  indentLeft?: number;
  isTitle?: boolean;
};

export type DocxTableBlock = {
  id: string;
  type: 'table';
  rows: DocxTableRow[];
};

export type DocxChartBlock = {
  id: string;
  type: 'chart';
  chart: import('../office/charts').OfficeChartModel;
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
};

export type DocxInline = DocxTextInline | DocxBreakInline | DocxImageInline;

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
  color?: string;
  fontSize?: number;
};
