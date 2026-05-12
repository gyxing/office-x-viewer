// docRenderUtils 提供 DOC 降级渲染所需的样式转换和图片排版辅助方法。
import type { CSSProperties } from 'react';
import type { DocBlock, DocImage, DocTextStyle } from '../../services/doc/types';

export const DOC_IMAGE_ROW_GAP = 6;

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

export function inlineStyleToCss(style?: DocTextStyle, options?: { preserveBlockTypography?: boolean }): CSSProperties {
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
  if (!inlines.length || inlines.some((inline) => inline.type !== 'image')) return [];
  return inlines.flatMap((inline) => (inline.type === 'image' ? [inline.image] : []));
}

export function canShareImageRow(left: DocImage, right: DocImage, contentWidth: number) {
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
