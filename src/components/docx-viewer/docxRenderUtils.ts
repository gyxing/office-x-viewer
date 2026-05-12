// docxRenderUtils 提供 DOCX 渲染阶段的样式转换和空段落高度计算。
import type { CSSProperties } from 'react';
import type { DocxParagraphBlock, DocxTextStyle } from '../../services/docx/types';

// DOCX 的样式已经在解析阶段完成继承合并，这里只负责把最终样式映射到 React CSS。
export function buildDocxTextStyle(style?: DocxTextStyle, options?: { includeBackground?: boolean }): CSSProperties {
  const css: CSSProperties = {
    fontWeight: style?.bold === true ? 700 : style?.bold === false ? 400 : undefined,
    fontStyle: style?.italic === true ? 'italic' : style?.italic === false ? 'normal' : undefined,
    textDecoration: [style?.underline ? 'underline' : '', style?.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ') || undefined,
    color: style?.color,
    fontSize: style?.fontSize,
    fontFamily: style?.fontFamily,
    textTransform: style?.allCaps ? 'uppercase' : undefined,
    fontVariant: style?.smallCaps ? 'small-caps' : undefined,
    background: options?.includeBackground ? style?.backgroundColor : undefined,
  };
  return Object.fromEntries(Object.entries(css).filter(([, value]) => value !== undefined)) as CSSProperties;
}

export function getDocxEmptyParagraphHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  // 空段落没有文字撑高，需要按行高补出可见高度，否则换行会丢失。
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}
