// DocParagraphBlock 渲染 DOC 段落块，并应用推断出的标题、正文和文字样式。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocParagraphBlock as DocParagraphBlockModel } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './docRenderUtils';

type DocParagraphBlockProps = {
  block: DocParagraphBlockModel;
};

function DocParagraphBlockComponent({ block }: DocParagraphBlockProps) {
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

export const DocParagraphBlock = memo(DocParagraphBlockComponent);
