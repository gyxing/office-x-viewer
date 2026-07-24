export type DocDocument = {
  title: string;
  page: DocPage;
  blocks: DocBlock[];
  paragraphs: DocParagraph[];
  images: DocImage[];
  warnings: string[];
  resources?: DocResources;
};

export type DocResources = {
  objectUrls: string[];
};

/** 释放 DOC/WPS 文档创建的 Blob URL；重复调用保持幂等。 */
export function disposeDocDocument(document: DocDocument | undefined) {
  const urls = document?.resources?.objectUrls;
  if (!urls?.length) return;
  const uniqueUrls = new Set(urls);
  urls.length = 0;
  if (typeof URL === 'undefined' || typeof URL.revokeObjectURL !== 'function') {
    return;
  }
  uniqueUrls.forEach((url) => URL.revokeObjectURL(url));
}

export type DocPage = {
  width: number;
  minHeight: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
};

export type DocParagraph = {
  id: string;
  text: string;
};

export type DocBlock = DocParagraphBlock | DocTableBlock | DocListBlock;

export type DocParagraphBlock = {
  id: string;
  type: 'paragraph';
  text: string;
  inlines?: DocTextInline[];
  role?: 'title' | 'heading' | 'body';
  style?: DocTextStyle;
};

export type DocTableBlock = {
  id: string;
  type: 'table';
  rows: DocTableRow[];
  style?: DocTableStyle;
  columns?: number[];
  width?: number;
  align?: 'left' | 'center' | 'right';
};

export type DocTableRow = {
  id: string;
  cells: DocTableCell[];
};

export type DocTableCell = {
  id: string;
  text: string;
  inlines?: DocTextInline[];
  style?: DocTextStyle;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  width?: number;
  colSpan?: number;
  verticalAlign?: 'top' | 'middle' | 'bottom';
};

export type DocListBlock = {
  id: string;
  type: 'list';
  ordered: boolean;
  items: DocListItem[];
  style?: DocTextStyle;
};

export type DocListItem = {
  id: string;
  text: string;
  inlines?: DocTextInline[];
};

export type DocTextInline = DocTextRunInline | DocImageInline;

export type DocTextRunInline = {
  type: 'text';
  text: string;
  style?: DocTextStyle;
};

export type DocImageInline = {
  type: 'image';
  image: DocImage;
};

export type DocTextStyle = {
  color?: string;
  backgroundColor?: string;
  fontSize?: number;
  fontWeight?: number;
  fontStyle?: 'normal' | 'italic';
  textDecoration?: string;
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  lineHeight?: number;
  fontFamily?: string;
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  paddingTop?: number;
  paddingRight?: number;
  paddingBottom?: number;
  paddingLeft?: number;
};

export type DocTableStyle = {
  headerBackgroundColor?: string;
  headerTextColor?: string;
  borderColor?: string;
  cellBackgroundColor?: string;
  stripedRowBackgroundColor?: string;
};

export type DocImage = {
  id: string;
  src: string;
  mimeType: string;
  width?: number;
  height?: number;
  caption?: string;
  offset?: number;
  anchored?: boolean;
};
