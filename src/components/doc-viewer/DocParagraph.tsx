import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocParagraphBlock } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './shared';

type DocParagraphProps = {
  block: DocParagraphBlock;
};

function DocParagraphComponent({ block }: DocParagraphProps) {
  const isTitle = block.role === 'title';
  const isHeading = block.role === 'heading';
  const paragraphStyle = useMemo<CSSProperties>(
    () => ({
      marginBottom: isTitle ? 18 : isHeading ? 14 : 12,
      fontSize: isTitle ? 22 : isHeading ? 16 : 14,
      lineHeight: isTitle ? 1.45 : isHeading ? 1.65 : 1.8,
      fontWeight: isTitle || isHeading ? 700 : 400,
      ...docTextStyleToCss(block.style),
    }),
    [block.style, isHeading, isTitle],
  );

  return (
    <p className="oxv-doc-paragraph" style={paragraphStyle}>
      <DocInlineContent inlines={block.inlines} fallback={block.text} preserveBlockTypography={isTitle || isHeading} />
    </p>
  );
}

export const DocParagraph = memo(DocParagraphComponent);
