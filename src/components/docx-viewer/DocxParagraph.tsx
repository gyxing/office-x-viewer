import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxParagraphBlock } from '../../services/docx/types';
import { emptyParagraphHeight, textStyleToCss } from './shared';
import { DocxInlineContent } from './DocxInlineContent';

type DocxParagraphProps = {
  block: DocxParagraphBlock;
  compact?: boolean;
};

function DocxParagraphComponent({ block, compact = false }: DocxParagraphProps) {
  const hasContent = block.inlines.length > 0;
  const paragraphStyle = useMemo<CSSProperties>(
    () => ({
      margin: 0,
      marginTop: compact ? 0 : block.spacingBefore,
      marginRight: block.indentRight,
      marginBottom: block.spacingAfter ?? 0,
      marginLeft: block.indentLeft,
      paddingLeft: block.paddingLeft,
      paddingRight: block.paddingRight,
      minHeight: hasContent ? undefined : emptyParagraphHeight(block),
      textAlign: block.align,
      lineHeight: block.lineHeight,
      color: block.style?.color ?? '#000',
      fontSize: block.style?.fontSize ?? 14,
      fontWeight: block.style?.bold ? 700 : 400,
      background: block.backgroundColor,
      borderTop: block.borderTop,
      borderRight: block.borderRight,
      borderBottom: block.borderBottom,
      borderLeft: block.borderLeft,
      textIndent: block.firstLineIndent,
      paddingTop: block.paddingTop,
      paddingBottom: block.paddingBottom,
      ...textStyleToCss(block.style),
    }),
    [block, compact, hasContent],
  );

  return (
    <p style={paragraphStyle}>
      {block.inlines.map((inline, index) => (
        <DocxInlineContent key={`${block.id}-inline-${index}`} inline={inline} />
      ))}
    </p>
  );
}

export const DocxParagraph = memo(DocxParagraphComponent);

