import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocListBlock } from '../../services/doc/types';
import { DocInlineContent } from './DocInlineContent';
import { docTextStyleToCss } from './shared';

type DocListProps = {
  block: DocListBlock;
};

function DocListComponent({ block }: DocListProps) {
  const itemStyle = useMemo<CSSProperties>(
    () => ({
      ...docTextStyleToCss(block.style),
    }),
    [block.style],
  );
  const Tag = block.ordered ? 'ol' : 'ul';

  return (
    <Tag className="oxv-doc-list">
      {block.items.map((item) => (
        <li key={item.id} className="oxv-doc-list__item" style={itemStyle}>
          <DocInlineContent inlines={item.inlines} fallback={item.text} />
        </li>
      ))}
    </Tag>
  );
}

export const DocList = memo(DocListComponent);
