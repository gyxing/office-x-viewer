import type { CSSProperties } from 'react';
import type { DocBlock, DocImage, DocTextStyle } from '../../services/doc/types';

export const DOC_IMAGE_ROW_GAP = 6;

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

export function inlineStyleToCss(style?: DocTextStyle, options?: { preserveBlockTypography?: boolean }): CSSProperties {
  const css = docTextStyleToCss(style);
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
  if (block.type !== 'paragraph' || block.text.trim()) return [];
  const inlines = block.inlines ?? [];
  if (!inlines.length || inlines.some((inline) => inline.type !== 'image')) return [];
  return inlines.map((inline) => inline.image);
}

export function canShareImageRow(left: DocImage, right: DocImage, contentWidth: number) {
  if (!left.width || !right.width) return false;
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
