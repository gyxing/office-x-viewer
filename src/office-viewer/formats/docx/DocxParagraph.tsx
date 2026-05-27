// DocxParagraph 渲染 DOCX 段落块，并应用段落级缩进、间距、边框和文字样式。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxParagraphBlock } from '../../services/docx/types';
import { buildDocxTextStyle, getDocxEmptyParagraphHeight } from './docxRenderUtils';
import { DocxInlineContent } from './DocxInlineContent';
import { calculatePositionStyle } from './positionUtils';

type DocxParagraphProps = {
  block: DocxParagraphBlock;
  compact?: boolean;
  asDiv?: boolean; // 强制使用 div 而不是 p,用于避免嵌套问题
};

function DocxParagraphComponent({ block, compact = false, asDiv = false }: DocxParagraphProps) {
  const hasContent = block.inlines.length > 0;
  const hasFlowContent = block.inlines.some((inline) => {
    if (inline.type === 'text') return inline.text.length > 0;
    if (inline.type === 'break') return true;
    if (inline.type === 'image') return !inline.image.position;
    if (inline.type === 'shape') return !inline.shape.position;
    if (inline.type === 'chart') return !inline.chart.position;
    return false;
  });

  // 检查是否包含定位元素,如果包含则使用 div 而不是 p 以避免 DOM 嵌套警告
  const hasPositionedElements = block.inlines.some((inline) => {
    if (inline.type === 'image') return Boolean(inline.image.position);
    if (inline.type === 'shape') return Boolean(inline.shape.position);
    if (inline.type === 'chart') return Boolean(inline.chart.position);
    return false;
  });

  const positionStyle = calculatePositionStyle(block.position);

  const paragraphStyle = useMemo<CSSProperties>(
    () => {
      // 纯浮动锚点段落在 Word 中流高度为 0，所有浮动共享页面顶部坐标系；
      // 无需撑开高度，否则会造成段落级联扩张、浮动元素偏离预期位置。
      const baseMinHeight = hasFlowContent
        ? undefined
        : hasContent
          ? 0
          : getDocxEmptyParagraphHeight(block);
      return {
        ...positionStyle,
        position: block.position ? positionStyle.position : 'relative',
        margin: block.position ? 0 : undefined,
        marginTop: block.position ? undefined : compact ? 0 : block.spacingBefore,
        marginRight: block.position ? undefined : block.indentRight,
        marginBottom: block.position ? undefined : block.spacingAfter ?? 0,
        marginLeft: block.position ? undefined : block.indentLeft,
        paddingLeft: block.paddingLeft,
        paddingRight: block.paddingRight,
        minHeight: baseMinHeight,
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
        maxWidth: block.position ? 'none' : undefined,
        ...buildDocxTextStyle(block.style),
      };
    },
    [block, compact, hasContent, hasFlowContent, positionStyle],
  );

  // 使用 div 而不是 p 来避免 DOM 嵌套警告(当包含定位元素或强制使用 div 时)
  const Container = (hasPositionedElements || asDiv) ? 'div' : 'p';

  return (
    <Container style={paragraphStyle}>
      {block.inlines.map((inline, index) => (
        <DocxInlineContent key={`${block.id}-inline-${index}`} inline={inline} />
      ))}
    </Container>
  );
}

export const DocxParagraph = memo(DocxParagraphComponent);
