// DocxViewer 负责 DOCX 文档预览整体布局，包括顶部摘要栏和页面滚动区。
import { Typography } from 'antd';
import { memo, useMemo } from 'react';
import type { DocxDocument } from '../../services/docx/types';
import { OfficeEmpty } from '../../shell/Empty';
import { DocxBlockRenderer } from './DocxBlockRenderer';
import { DocxPageFrame } from './DocxPageFrame';
import './index.less';

type DocxViewerProps = {
  document?: DocxDocument;
  zoom: number;
};

function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
  const page = document?.page;
  const contentWidth = page ? page.width - page.marginLeft - page.marginRight : undefined;
  const summaryText = useMemo(
    () => (document ? `${document.blocks.length} 个内容块 / ${document.images.length} 张图片` : ''),
    [document],
  );

  if (!document?.blocks.length || !page) {
    return <OfficeEmpty kind="docx" />;
  }

  return (
    <div className="oxv-docx-viewer">
      <div className="oxv-docx-viewer__header">
        <Typography.Text strong ellipsis className="oxv-docx-viewer__title">
          {document.title}
        </Typography.Text>
        <Typography.Text type="secondary" className="oxv-docx-viewer__summary">
          {summaryText}
        </Typography.Text>
      </div>
      <div className="oxv-docx-viewer__scroller">
        <DocxPageFrame page={page} zoom={zoom}>
          {document.blocks.map((block) => (
            <DocxBlockRenderer key={block.id} block={block} availableWidth={contentWidth} />
          ))}
        </DocxPageFrame>
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
