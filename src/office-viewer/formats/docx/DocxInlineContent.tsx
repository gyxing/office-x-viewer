// DocxInlineContent 渲染 DOCX 段落内的文本、换行、图片、图表和形状。
import { memo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxImage } from './DocxImage';
import { DocxInlineChart } from './DocxInlineChart';
import { DocxShape } from './DocxShape';
import { buildDocxTextStyle } from './docxRenderUtils';

type DocxInlineContentProps = {
  inline: DocxInline;
};

function DocxInlineContentComponent({ inline }: DocxInlineContentProps) {
  if (inline.type === 'break') return <br />;
  if (inline.type === 'image') return <DocxImage inline={inline} />;
  if (inline.type === 'chart') return <DocxInlineChart inline={inline} />;
  if (inline.type === 'shape') return <DocxShape inline={inline} />;
  return <span style={buildDocxTextStyle(inline.style, { includeBackground: true })}>{inline.text}</span>;
}

export const DocxInlineContent = memo(DocxInlineContentComponent);
