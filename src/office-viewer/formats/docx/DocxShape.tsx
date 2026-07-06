// DocxShape 渲染 DOCX 行内形状，支持矩形、椭圆、线条、自定义路径和形状内文字。
import { memo, useCallback, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxInline } from '../../services/docx/types';
import { OfficeChartView } from '../../shared/chart/OfficeChartView';
import { DocxParagraph } from './DocxParagraph';
import { DocxTableBlock } from './DocxTableBlock';
import { calculatePositionStyle } from './positionUtils';

type DocxShapeProps = {
  inline: Extract<DocxInline, { type: 'shape' }>;
};

function DocxShapeComponent({ inline }: DocxShapeProps) {
  const shape = inline.shape;
  const positionStyle = calculatePositionStyle(shape.position);

  const shapeStyle = useMemo<CSSProperties>(
    () => {
      // 当 Shape 有定位时,给 z-index 添加一个小的偏移量,确保文本在图片上方
      const adjustedZIndex = positionStyle.zIndex !== undefined
        ? positionStyle.zIndex + 1
        : undefined;

      return {
        '--oxv-docx-shape-width': `${shape.width}px`,
        '--oxv-docx-shape-height': `${shape.height}px`,
        ...positionStyle,
        zIndex: adjustedZIndex,
        maxWidth: shape.position ? 'none' : undefined,
        margin: shape.position ? 0 : undefined,
      } as CSSProperties;
    },
    [positionStyle, shape.height, shape.position, shape.width],
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

        // 调试输出
        if (item.blocks && item.blocks.length > 0) {
          const firstText = item.blocks[0].type === 'paragraph' ? item.blocks[0].text : '';
          if (firstText && (firstText.includes('班级') || firstText.includes('姓名'))) {
            console.log('Shape item:', {
              text: firstText,
              fitShapeToText: item.fitShapeToText,
              width: item.width,
              height: item.height
            });
          }
        }

        return (
          <div
            key={item.id}
            className="oxv-docx-shape__item"
            style={{
              left: item.left,
              top: item.top,
              ...(item.fitShapeToText
                ? { minWidth: item.width, minHeight: item.height }
                : { width: item.width, height: item.height }),
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
            {(item.blocks ?? item.paragraphs)?.map((block) =>
              block.type === 'table' ? (
                <DocxTableBlock key={block.id} block={block} availableWidth={item.width} />
              ) : block.type === 'chart' ? (
                <div key={block.id} className="oxv-docx-table-block__chart">
                  <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={100} />
                </div>
              ) : (
                <DocxParagraph key={block.id} block={block} compact asDiv />
              ),
            )}
          </div>
        );
      })}
    </span>
  );
}

export const DocxShape = memo(DocxShapeComponent);
