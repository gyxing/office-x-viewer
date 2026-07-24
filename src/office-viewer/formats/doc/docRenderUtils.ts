// docRenderUtils 提供 DOC 降级渲染所需的样式转换和图片排版辅助方法。
import type { CSSProperties } from 'react';
import type {
  DocBlock,
  DocImage,
  DocPage,
  DocParagraphBlock,
  DocTextInline,
  DocTextStyle,
} from '../../services/doc/types';

export const DOC_IMAGE_ROW_GAP = 6;

const DOC_PAGE_HEIGHT_BUFFER = 24;
const DOC_IMAGE_LAYOUT_ROW_GAP = 12;
const DOC_IMAGE_LAYOUT_VERTICAL_MARGIN = 22;

// DOC 没有 OOXML 的显式页面模型，这里记录前端估算分页后的页面块集合。
export type PaginatedDocPage = {
  id: string;
  blocks: DocBlock[];
};

// DOC 解析出的文本样式字段和 React CSS 字段基本一一对应，集中转换便于后续补充新属性。
export function docTextStyleToCss(style?: DocTextStyle): CSSProperties {
  if (!style) return {};

  return {
    color: style.color,
    background: style.backgroundColor,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    textDecoration: style.textDecoration,
    textAlign: style.textAlign,
    lineHeight: style.lineHeight,
    fontFamily: style.fontFamily,
    marginLeft: style.indentLeft,
    marginRight: style.indentRight,
    textIndent: style.firstLineIndent,
    marginTop: style.spacingBefore,
    marginBottom: style.spacingAfter,
    paddingTop: style.paddingTop,
    paddingRight: style.paddingRight,
    paddingBottom: style.paddingBottom,
    paddingLeft: style.paddingLeft,
  };
}

export function inlineStyleToCss(
  style?: DocTextStyle,
  options?: { preserveBlockTypography?: boolean },
): CSSProperties {
  const css = docTextStyleToCss(style);
  // 行内片段不能继承段落级缩进/间距，否则会把整段排版撑乱。
  delete css.textAlign;
  delete css.marginLeft;
  delete css.marginRight;
  delete css.textIndent;
  delete css.marginTop;
  delete css.marginBottom;
  delete css.paddingTop;
  delete css.paddingRight;
  delete css.paddingBottom;
  delete css.paddingLeft;

  if (options?.preserveBlockTypography) {
    delete css.fontSize;
    delete css.fontWeight;
    delete css.lineHeight;
  }

  return css;
}

export function imagesFromImageOnlyParagraph(block: DocBlock) {
  // 二进制 DOC 没有稳定的图片锚点模型，这里用“无文字且全是图片”的段落作为图片布局信号。
  if (block.type !== 'paragraph' || block.text.trim()) return [];
  const inlines = block.inlines ?? [];
  const hasVisibleText = inlines.some(
    (inline) => inline.type === 'text' && inline.text.trim(),
  );
  const images = inlines.flatMap((inline) =>
    inline.type === 'image' ? [inline.image] : [],
  );
  if (!images.length || hasVisibleText) return [];
  return images;
}

export function canShareImageRow(
  left: DocImage,
  right: DocImage,
  contentWidth: number,
) {
  if (!left.width || !right.width) return false;
  // 只让较小图片并排，避免大图被压缩后影响文档可读性。
  const maxSmallImageWidth = Math.min(300, contentWidth * 0.55);
  return (
    left.width <= maxSmallImageWidth &&
    right.width <= maxSmallImageWidth &&
    left.width + right.width + DOC_IMAGE_ROW_GAP <= contentWidth
  );
}

export function imageRows(images: DocImage[], contentWidth: number) {
  const rows: DocImage[][] = [];
  let index = 0;

  while (index < images.length) {
    const current = images[index];
    const next = images[index + 1];
    if (next && canShareImageRow(current, next, contentWidth)) {
      rows.push([current, next]);
      index += 2;
      continue;
    }
    rows.push([current]);
    index += 1;
  }

  return rows;
}

function weightedTextLength(text: string) {
  return Array.from(text).reduce(
    (sum, char) => sum + (/[\u4e00-\u9fa5]/.test(char) ? 1 : 0.55),
    0,
  );
}

function estimateLineCount(text: string, width: number, fontSize: number) {
  const weightedLength = weightedTextLength(text);
  const charsPerLine = Math.max(8, Math.floor(width / (fontSize * 0.95)));
  return Math.max(1, Math.ceil(weightedLength / charsPerLine));
}

function estimateParagraphTextHeight(
  block: DocParagraphBlock,
  contentWidth: number,
) {
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const fontSize =
    block.style?.fontSize ?? (isTitle ? 22 : isHeading ? 16 : 14);
  const lineHeight =
    block.style?.lineHeight ?? (isTitle ? 1.45 : isHeading ? 1.65 : 1.8);
  const defaultSpacingAfter = isTitle ? 18 : isHeading ? 14 : 12;
  const spacingBefore = block.style?.spacingBefore ?? 0;
  const spacingAfter = block.style?.spacingAfter ?? defaultSpacingAfter;
  const padding =
    (block.style?.paddingTop ?? 0) + (block.style?.paddingBottom ?? 0);
  const lines = estimateLineCount(block.text || ' ', contentWidth, fontSize);

  return Math.max(
    18,
    lines * fontSize * lineHeight + spacingBefore + spacingAfter + padding,
  );
}

function estimateImageHeight(
  image: DocImage,
  contentWidth: number,
  rowLength: number,
) {
  const preferredWidth = image.width
    ? Math.min(image.width, contentWidth)
    : contentWidth;
  const renderedWidth =
    rowLength > 1 && image.width
      ? Math.min(image.width, (contentWidth - DOC_IMAGE_ROW_GAP) / rowLength)
      : preferredWidth;

  if (image.width && image.height) {
    return (image.height * renderedWidth) / image.width;
  }

  return Math.min(240, contentWidth * 0.55);
}

function estimateImageRowHeight(row: DocImage[], contentWidth: number) {
  return Math.max(
    ...row.map((image) => estimateImageHeight(image, contentWidth, row.length)),
    0,
  );
}

function estimateTableHeight(
  block: Extract<DocBlock, { type: 'table' }>,
  contentWidth: number,
) {
  const columnCount = Math.max(...block.rows.map((row) => row.cells.length), 1);
  const columns = block.columns?.length
    ? block.columns
    : Array.from({ length: columnCount }, () => contentWidth / columnCount);
  const totalColumns =
    columns.reduce((sum, width) => sum + width, 0) || contentWidth;

  const rowHeights = block.rows.map((row) =>
    Math.max(
      28,
      ...row.cells.map((cell, cellIndex) => {
        const fontSize = cell.style?.fontSize ?? 13;
        const lineHeight = cell.style?.lineHeight ?? 1.65;
        const padding =
          (cell.style?.paddingTop ?? 5) + (cell.style?.paddingBottom ?? 5);
        const width =
          cell.width ?? (columns[cellIndex] / totalColumns) * contentWidth;
        return (
          estimateLineCount(
            cell.text || ' ',
            Math.max(48, width - 16),
            fontSize,
          ) *
            fontSize *
            lineHeight +
          padding
        );
      }),
    ),
  );

  return rowHeights.reduce((sum, height) => sum + height, 0) + 16;
}

function estimateListHeight(
  block: Extract<DocBlock, { type: 'list' }>,
  contentWidth: number,
) {
  const fontSize = block.style?.fontSize ?? 14;
  const lineHeight = block.style?.lineHeight ?? 1.7;
  const itemHeight = block.items.reduce(
    (sum, item) =>
      sum +
      estimateLineCount(item.text || ' ', contentWidth - 24, fontSize) *
        fontSize *
        lineHeight +
      8,
    0,
  );
  return itemHeight + 8;
}

function estimateBlockHeight(block: DocBlock, contentWidth: number) {
  if (block.type === 'table') return estimateTableHeight(block, contentWidth);
  if (block.type === 'list') return estimateListHeight(block, contentWidth);

  const images =
    block.inlines?.flatMap((inline) =>
      inline.type === 'image' ? [inline.image] : [],
    ) ?? [];
  const imageHeight = images.reduce(
    (sum, image) => sum + estimateImageHeight(image, contentWidth, 1) + 6,
    0,
  );
  return estimateParagraphTextHeight(block, contentWidth) + imageHeight;
}

function createImageParagraphBlock(
  id: string,
  images: DocImage[],
): DocParagraphBlock {
  const inlines: DocTextInline[] = images.map((image) => ({
    type: 'image',
    image,
  }));
  return {
    id,
    type: 'paragraph',
    text: '',
    inlines,
    role: 'body',
  };
}

/**
 * 按 DOC 页面可用高度估算分页，并把连续图片段落拆到多个页面中。
 * 旧版 DOC 缺少稳定的前端分页信息，该方法只保证内容不再挤成一个长页面。
 */
export function paginateDocBlocks(
  blocks: DocBlock[],
  page: DocPage,
  contentWidth: number,
): PaginatedDocPage[] {
  const contentHeight = Math.max(
    240,
    page.minHeight -
      page.marginTop -
      page.marginBottom -
      DOC_PAGE_HEIGHT_BUFFER,
  );
  const pages: PaginatedDocPage[] = [];
  let currentBlocks: DocBlock[] = [];
  let currentHeight = 0;
  let syntheticImageIndex = 0;

  const flushPage = () => {
    if (!currentBlocks.length) return;
    pages.push({ id: `doc-page-${pages.length + 1}`, blocks: currentBlocks });
    currentBlocks = [];
    currentHeight = 0;
  };

  const appendBlock = (block: DocBlock, estimatedHeight: number) => {
    if (
      currentBlocks.length &&
      currentHeight + estimatedHeight > contentHeight
    ) {
      flushPage();
    }
    currentBlocks.push(block);
    currentHeight += estimatedHeight;
  };

  const appendImageRows = (images: DocImage[]) => {
    const rows = imageRows(images, contentWidth);
    let pendingImages: DocImage[] = [];
    let pendingHeight = DOC_IMAGE_LAYOUT_VERTICAL_MARGIN;

    const flushImages = () => {
      if (!pendingImages.length) return;
      syntheticImageIndex += 1;
      appendBlock(
        createImageParagraphBlock(
          `doc-image-page-group-${syntheticImageIndex}`,
          pendingImages,
        ),
        pendingHeight,
      );
      pendingImages = [];
      pendingHeight = DOC_IMAGE_LAYOUT_VERTICAL_MARGIN;
    };

    rows.forEach((row) => {
      const rowHeight =
        estimateImageRowHeight(row, contentWidth) +
        (pendingImages.length ? DOC_IMAGE_LAYOUT_ROW_GAP : 0);
      if (
        pendingImages.length &&
        currentHeight + pendingHeight + rowHeight > contentHeight
      ) {
        flushImages();
      }
      pendingImages.push(...row);
      pendingHeight += rowHeight;
    });

    flushImages();
  };

  let index = 0;
  while (index < blocks.length) {
    const imageOnlyParagraphImages = imagesFromImageOnlyParagraph(
      blocks[index],
    );

    if (imageOnlyParagraphImages.length) {
      const imageGroup = [...imageOnlyParagraphImages];
      let nextIndex = index + 1;
      while (nextIndex < blocks.length) {
        const nextImages = imagesFromImageOnlyParagraph(blocks[nextIndex]);
        if (!nextImages.length) break;
        imageGroup.push(...nextImages);
        nextIndex += 1;
      }

      appendImageRows(imageGroup);
      index = nextIndex;
      continue;
    }

    appendBlock(
      blocks[index],
      estimateBlockHeight(blocks[index], contentWidth),
    );
    index += 1;
  }

  flushPage();
  return pages.length ? pages : [{ id: 'doc-page-1', blocks }];
}
