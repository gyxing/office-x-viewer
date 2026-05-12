import { memo } from 'react';
import type { DocTextInline } from '../../services/doc/types';
import { inlineStyleToCss } from './shared';

type DocInlineContentProps = {
  inlines?: DocTextInline[];
  fallback: string;
  preserveBlockTypography?: boolean;
};

function DocInlineContentComponent({ inlines, fallback, preserveBlockTypography }: DocInlineContentProps) {
  if (!inlines?.length) return <>{fallback}</>;

  return (
    <>
      {inlines.map((inline, index) =>
        inline.type === 'image' ? (
          <span key={`${inline.image.id}-${index}`} className="oxv-doc-inline-image">
            <img
              className="oxv-doc-inline-image__img"
              src={inline.image.src}
              alt={inline.image.caption ?? inline.image.id}
              style={{
                width: inline.image.width && inline.image.width <= 520 ? inline.image.width : undefined,
                height: inline.image.height && inline.image.width && inline.image.width <= 520 ? inline.image.height : undefined,
              }}
            />
          </span>
        ) : (
          <span key={`${inline.text}-${index}`} style={inlineStyleToCss(inline.style, { preserveBlockTypography })}>
            {inline.text}
          </span>
        ),
      )}
    </>
  );
}

export const DocInlineContent = memo(DocInlineContentComponent);
