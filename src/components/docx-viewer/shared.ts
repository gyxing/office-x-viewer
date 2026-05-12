import type { CSSProperties } from 'react';
import type { DocxParagraphBlock, DocxTextStyle } from '../../services/docx/types';

export function textStyleToCss(style?: DocxTextStyle, options?: { includeBackground?: boolean }): CSSProperties {
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

export function emptyParagraphHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}

