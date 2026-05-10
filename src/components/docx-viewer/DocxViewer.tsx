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

function textStyleToCss(style?: DocxTextStyle): CSSProperties {
  return {
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? 'italic' : undefined,
    textDecoration: style?.underline ? 'underline' : undefined,
    color: style?.color,
    fontSize: style?.fontSize,
  };
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
        margin: '8px 0',
      }}
    />
  );
}

function InlineContent({ inline }: { inline: DocxInline }) {
  if (inline.type === 'break') {
    return <br />;
  }
  if (inline.type === 'image') {
    return <DocxImage inline={inline} />;
  }
  return <span style={textStyleToCss(inline.style)}>{inline.text}</span>;
}

function ParagraphComponent({ block, compact = false }: { block: DocxParagraphBlock; compact?: boolean }) {
  const hasContent = block.inlines.length > 0;
  const paragraphStyle = useMemo<CSSProperties>(
    () => ({
      margin: 0,
      marginTop: compact ? 0 : block.spacingBefore,
      marginBottom: compact ? 4 : block.spacingAfter ?? 12,
      paddingLeft: block.indentLeft,
      minHeight: hasContent ? undefined : compact ? 16 : 20,
      textAlign: block.align,
      lineHeight: 1.65,
      color: '#1f2937',
      fontSize: block.isTitle ? 20 : block.style?.fontSize ?? 14,
      fontWeight: block.isTitle ? 700 : block.style?.bold ? 700 : 400,
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

function TableBlockComponent({ block }: { block: DocxTableBlock }) {
  return (
    <div style={{ overflowX: 'auto', margin: '12px 0 16px' }}>
      <table
        style={{
          borderCollapse: 'collapse',
          width: '100%',
          tableLayout: 'fixed',
          fontSize: 13,
          color: '#1f2937',
        }}
      >
        <tbody>
          {block.rows.map((row) => (
            <tr key={row.id}>
              {row.cells.map((cell) => (
                <td
                  key={cell.id}
                  colSpan={cell.colSpan && cell.colSpan > 1 ? cell.colSpan : undefined}
                  style={{
                    border: '1px solid #cfd7e3',
                    padding: '8px 10px',
                    width: cell.width,
                    verticalAlign: cell.verticalAlign,
                    background: cell.backgroundColor ?? '#fff',
                    wordBreak: 'break-word',
                  }}
                >
                  {cell.blocks.map((block) =>
                    block.type === 'chart' ? (
                      <div key={block.id} style={{ margin: '8px 0' }}>
                        <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={100} />
                      </div>
                    ) : (
                      <Paragraph key={block.id} block={block} compact />
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
    <div style={{ margin: '16px 0' }}>
      <OfficeChartView chart={block.chart} width={block.width} height={block.height} zoom={zoom} />
    </div>
  );
}

function BlockRenderer({ block }: { block: DocxBlock }) {
  if (block.type === 'table') {
    return <TableBlock block={block} />;
  }
  if (block.type === 'chart') {
    return <ChartBlock block={block} zoom={100} />;
  }
  return <Paragraph block={block} />;
}

export function DocxViewer({ document, zoom }: DocxViewerProps) {
  const scale = zoom / 100;
  const page = document?.page;
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
              <BlockRenderer key={block.id} block={block} />
            ))}
          </article>
        </div>
      </div>
    </div>
  );
}
