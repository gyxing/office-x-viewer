import { Empty, Typography } from 'antd';
import type { CSSProperties } from 'react';
import { memo, useMemo } from 'react';
import type {
  DocxBlock,
  DocxChartBlock,
  DocxDocument,
  DocxImageInline,
  DocxInline,
  DocxParagraphBlock,
  DocxTableBlock,
  DocxTextStyle,
} from '../../services/docx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';

type DocxViewerProps = {
  document?: DocxDocument;
  zoom: number;
};

function textStyleToCss(style?: DocxTextStyle, options?: { includeBackground?: boolean }): CSSProperties {
  const css: CSSProperties = {
    fontWeight: style?.bold === true ? 700 : style?.bold === false ? 400 : undefined,
    fontStyle: style?.italic === true ? 'italic' : style?.italic === false ? 'normal' : undefined,
    textDecoration: [style?.underline ? 'underline' : '', style?.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' ') || undefined,
    color: style?.color,
    fontSize: style?.fontSize,
    fontFamily: style?.fontFamily,
    textTransform: style?.allCaps ? 'uppercase' : undefined,
    fontVariant: style?.smallCaps ? 'small-caps' : undefined,
    background: options?.includeBackground ? style?.backgroundColor : undefined,
  };
  return Object.fromEntries(Object.entries(css).filter(([, value]) => value !== undefined)) as CSSProperties;
}

function emptyParagraphHeight(block: DocxParagraphBlock) {
  const fontSize = block.style?.fontSize ?? 14;
  if (block.lineHeight === undefined) return fontSize * 1.2;
  return block.lineHeight > 4 ? block.lineHeight : fontSize * block.lineHeight;
}

function DocxImage({ inline }: { inline: DocxImageInline }) {
  const image = inline.image;
  return (
    <img
      src={image.src}
      alt={image.alt ?? ''}
      title={image.name}
      style={{
        display: 'inline-block',
        width: image.width,
        maxWidth: '100%',
        height: 'auto',
        verticalAlign: 'middle',
      }}
    />
  );
}

function DocxInlineChart({ inline }: { inline: Extract<DocxInline, { type: 'chart' }> }) {
  const chart = inline.chart;
  return (
    <span style={{ display: 'inline-block', width: chart.width, height: chart.height, verticalAlign: 'middle' }}>
      <OfficeChartView chart={chart.chart} width={chart.width} height={chart.height} zoom={100} />
    </span>
  );
}

function InlineContent({ inline }: { inline: DocxInline }) {
  if (inline.type === 'break') return <br />;
  if (inline.type === 'image') return <DocxImage inline={inline} />;
  if (inline.type === 'chart') return <DocxInlineChart inline={inline} />;
  if (inline.type === 'shape') return <DocxShape inline={inline} />;
  return <span style={textStyleToCss(inline.style, { includeBackground: true })}>{inline.text}</span>;
}

function DocxShape({ inline }: { inline: Extract<DocxInline, { type: 'shape' }> }) {
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
              <Paragraph key={paragraph.id} block={paragraph} compact />
            ))}
          </div>
        );
      })}
    </span>
  );
}

function ParagraphComponent({ block, compact = false }: { block: DocxParagraphBlock; compact?: boolean }) {
  const hasContent = block.inlines.length > 0;
  const paragraphStyle = useMemo<CSSProperties>(
    () => ({
      margin: 0,
      marginTop: compact ? 0 : block.spacingBefore,
      marginRight: block.indentRight,
      marginBottom: compact ? block.spacingAfter ?? 0 : block.spacingAfter ?? 0,
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
        <InlineContent key={`${block.id}-inline-${index}`} inline={inline} />
      ))}
    </p>
  );
}

const Paragraph = memo(ParagraphComponent);

function TableBlockComponent({ block, availableWidth }: { block: DocxTableBlock; availableWidth?: number }) {
  const marginLeft = block.align === 'center' ? 'auto' : block.align === 'right' ? 'auto' : 0;
  const marginRight = block.align === 'center' ? 'auto' : block.align === 'right' ? 0 : 'auto';
  const totalColumns = block.columns?.reduce((sum, width) => sum + width, 0) ?? block.width ?? 0;
  const shouldFit = Boolean(availableWidth && block.width && block.width > availableWidth);
  const tableWidth = shouldFit ? '100%' : block.width ?? availableWidth ?? '100%';
  return (
    <div style={{ margin: 0 }}>
      <table
        style={{
          borderCollapse: 'collapse',
          width: tableWidth,
          marginLeft,
          marginRight,
          tableLayout: 'fixed',
          fontSize: 13,
          color: '#000',
          fontFamily: '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif',
        }}
      >
        {block.columns?.length ? (
          <colgroup>
            {block.columns.map((width, index) => (
              <col
                key={`${block.id}-col-${index}`}
                style={{
                  width:
                    shouldFit && totalColumns > 0
                      ? `${(width / totalColumns) * 100}%`
                      : width,
                }}
              />
            ))}
          </colgroup>
        ) : null}
        <tbody>
          {block.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  colSpan={cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined}
                  style={{
                    borderTop: cell.borderTop ?? (cell.hasBorderTop ? 'none' : '1px solid #cfd7e3'),
                    borderRight: cell.borderRight ?? (cell.hasBorderRight ? 'none' : '1px solid #cfd7e3'),
                    borderBottom: cell.borderBottom ?? (cell.hasBorderBottom ? 'none' : '1px solid #cfd7e3'),
                    borderLeft: cell.borderLeft ?? (cell.hasBorderLeft ? 'none' : '1px solid #cfd7e3'),
                    paddingTop: cell.paddingTop ?? 0,
                    paddingRight: cell.paddingRight ?? 7,
                    paddingBottom: cell.paddingBottom ?? 0,
                    paddingLeft: cell.paddingLeft ?? 7,
                    width: shouldFit ? undefined : cell.width,
                    verticalAlign: cell.verticalAlign,
                    background: cell.backgroundColor ?? '#fff',
                    wordBreak: cell.noWrap ? 'normal' : 'break-word',
                    overflowWrap: cell.noWrap ? 'normal' : 'anywhere',
                    whiteSpace: cell.noWrap ? 'nowrap' : undefined,
                    color: '#000',
                    fontFamily: '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif',
                  }}
                >
                  {cell.blocks.map((item) =>
                    item.type === 'chart' ? (
                      <div key={item.id} style={{ margin: '8px 0' }}>
                        <OfficeChartView chart={item.chart} width={item.width} height={item.height} zoom={100} />
                      </div>
                    ) : (
                      <Paragraph key={item.id} block={item} compact />
                    ),
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const TableBlock = memo(TableBlockComponent);

function ChartBlock({ block, zoom }: { block: DocxChartBlock; zoom: number }) {
  return (
    <div style={{ margin: 0 }}>
      <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={zoom} />
    </div>
  );
}

function BlockRenderer({ block, availableWidth }: { block: DocxBlock; availableWidth?: number }) {
  if (block.type === 'table') return <TableBlock block={block} availableWidth={availableWidth} />;
  if (block.type === 'chart') return <ChartBlock block={block} zoom={100} />;
  return <Paragraph block={block} />;
}

export function DocxViewer({ document, zoom }: DocxViewerProps) {
  const scale = zoom / 100;
  const page = document?.page;
  const contentWidth = page ? page.width - page.marginLeft - page.marginRight : undefined;
  const summaryText = useMemo(
    () => (document ? `${document.blocks.length} 个内容块 / ${document.images.length} 张图片` : ''),
    [document],
  );
  const pageShellStyle = useMemo<CSSProperties>(
    () =>
      page
        ? {
            width: page.width * scale,
            minHeight: page.minHeight * scale,
            margin: '0 auto',
          }
        : {},
    [page, scale],
  );
  const articleStyle = useMemo<CSSProperties>(
    () =>
      page
        ? {
            width: page.width,
            minHeight: page.minHeight,
            padding: `${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px`,
            background: '#fff',
            boxShadow: '0 14px 30px rgba(15, 23, 42, 0.14)',
            boxSizing: 'border-box',
            borderTop: page.borderTop,
            borderRight: page.borderRight,
            borderBottom: page.borderBottom,
            borderLeft: page.borderLeft,
            transform: `scale(${scale})`,
            transformOrigin: 'top left',
            fontFamily: '"Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", Arial, sans-serif',
            letterSpacing: 0,
          }
        : {},
    [page, scale],
  );

  if (!document?.blocks.length || !page) {
    return <Empty description="请先上传 DOCX 文件开始预览" />;
  }

  return (
    <div
      style={{
        height: 'calc(100vh - 56px)',
        display: 'flex',
        flexDirection: 'column',
        background: '#eef1f6',
        overflow: 'hidden',
      }}
    >
      <div
        style={{
          flex: '0 0 auto',
          height: 40,
          padding: '0 18px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          background: '#fff',
          borderBottom: '1px solid #dde3ec',
        }}
      >
        <Typography.Text strong ellipsis style={{ maxWidth: 520 }}>
          {document.title}
        </Typography.Text>
        <Typography.Text type="secondary" style={{ fontSize: 12 }}>
          {summaryText}
        </Typography.Text>
      </div>
      <div
        style={{
          flex: '1 1 auto',
          minHeight: 0,
          overflow: 'auto',
          padding: 24,
          scrollbarGutter: 'stable both-edges',
        }}
      >
        <div style={pageShellStyle}>
          <article style={articleStyle}>
            {document.blocks.map((block) => (
              <BlockRenderer key={block.id} block={block} availableWidth={contentWidth} />
            ))}
          </article>
        </div>
      </div>
    </div>
  );
}
