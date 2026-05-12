import { memo } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxParagraph } from './DocxParagraph';

type DocxShapeProps = {
  inline: Extract<DocxInline, { type: 'shape' }>;
};

function DocxShapeComponent({ inline }: DocxShapeProps) {
  const shape = inline.shape;
  const justifyContent = (align?: 'top' | 'middle' | 'bottom') =>
    align === 'middle' ? 'center' : align === 'bottom' ? 'flex-end' : 'flex-start';
  const shapePath = (item: typeof shape.items[number]) => {
    if (item.path) return item.path;
    if (item.kind === 'ellipse') {
      return `M ${item.width / 2} 0 A ${item.width / 2} ${item.height / 2} 0 1 0 ${item.width / 2} ${item.height} A ${item.width / 2} ${item.height / 2} 0 1 0 ${item.width / 2} 0`;
    }
    return undefined;
  };

  return (
    <span
      style={{
        display: 'inline-block',
        position: 'relative',
        width: shape.width,
        height: shape.height,
        maxWidth: '100%',
        verticalAlign: 'middle',
        margin: '8px 0',
      }}
    >
      {shape.items.map((item) => {
        const path = shapePath(item);
        const drawAsSvg = Boolean(path) || item.kind === 'line';
        return (
          <div
            key={item.id}
            style={{
              position: 'absolute',
              left: item.left,
              top: item.top,
              width: item.width,
              height: item.height,
              boxSizing: 'border-box',
              display: 'flex',
              flexDirection: 'column',
              justifyContent: justifyContent(item.textVerticalAlign),
              overflow: 'visible',
              background: drawAsSvg ? undefined : item.fillColor,
              border: drawAsSvg ? undefined : item.border,
              borderRadius: item.borderRadius,
              paddingTop: item.paddingTop,
              paddingRight: item.paddingRight,
              paddingBottom: item.paddingBottom,
              paddingLeft: item.paddingLeft,
            }}
          >
            {path ? (
              <svg
                viewBox={item.viewBox ?? `0 0 ${Math.max(1, item.width)} ${Math.max(1, item.height)}`}
                preserveAspectRatio="none"
                style={{
                  position: 'absolute',
                  inset: 0,
                  width: '100%',
                  height: '100%',
                  overflow: 'visible',
                }}
              >
                <path
                  d={path}
                  fill={item.fillColor ?? 'none'}
                  stroke={item.strokeColor ?? 'none'}
                  strokeWidth={item.strokeWidth}
                  strokeDasharray={item.strokeDasharray}
                  vectorEffect="non-scaling-stroke"
                />
              </svg>
            ) : null}
            {item.paragraphs?.map((paragraph) => (
              <DocxParagraph key={paragraph.id} block={paragraph} compact />
            ))}
          </div>
        );
      })}
    </span>
  );
}

export const DocxShape = memo(DocxShapeComponent);

