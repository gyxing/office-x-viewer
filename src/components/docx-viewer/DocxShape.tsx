import { memo, useCallback, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { DocxParagraph } from './DocxParagraph';

type DocxShapeProps = {
  inline: Extract<DocxInline, { type: 'shape' }>;
};

function DocxShapeComponent({ inline }: DocxShapeProps) {
  const shape = inline.shape;
  const shapeStyle = useMemo<CSSProperties>(
    () =>
      ({
        '--oxv-docx-shape-width': `${shape.width}px`,
        '--oxv-docx-shape-height': `${shape.height}px`,
      }) as CSSProperties,
    [shape.height, shape.width],
  );
  const justifyContent = useCallback(
    (align?: 'top' | 'middle' | 'bottom') =>
      align === 'middle' ? 'center' : align === 'bottom' ? 'flex-end' : 'flex-start',
    [],
  );
  const shapePath = useCallback((item: typeof shape.items[number]) => {
    if (item.path) return item.path;
    if (item.kind === 'ellipse') {
      return `M ${item.width / 2} 0 A ${item.width / 2} ${item.height / 2} 0 1 0 ${item.width / 2} ${item.height} A ${item.width / 2} ${item.height / 2} 0 1 0 ${item.width / 2} 0`;
    }
    return undefined;
  }, []);

  return (
    <span className="oxv-docx-shape" style={shapeStyle}>
      {shape.items.map((item) => {
        const path = shapePath(item);
        const drawAsSvg = Boolean(path) || item.kind === 'line';
        return (
          <div
            key={item.id}
            className="oxv-docx-shape__item"
            style={{
              left: item.left,
              top: item.top,
              width: item.width,
              height: item.height,
              justifyContent: justifyContent(item.textVerticalAlign),
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
                className="oxv-docx-shape__svg"
                viewBox={item.viewBox ?? `0 0 ${Math.max(1, item.width)} ${Math.max(1, item.height)}`}
                preserveAspectRatio="none"
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
