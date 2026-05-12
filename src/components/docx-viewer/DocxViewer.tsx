import { Typography } from 'antd';
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxDocument } from '../../services/docx/types';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';
import { DocxBlockRenderer } from './DocxBlockRenderer';

type DocxViewerProps = {
  document?: DocxDocument;
  zoom: number;
};

function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
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
    return <OfficeEmpty kind="docx" />;
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
              <DocxBlockRenderer key={block.id} block={block} availableWidth={contentWidth} />
            ))}
          </article>
        </div>
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);

