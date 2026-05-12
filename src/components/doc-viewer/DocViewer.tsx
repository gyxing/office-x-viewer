// DocViewer 负责旧版 DOC 降级预览的整体布局、警告说明和页面滚动区。
import { Alert, Typography } from 'antd';
import { memo, useMemo } from 'react';
import type { DocDocument } from '../../services/doc/types';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
import { DocPageFrame } from './DocPageFrame';
import './index.less';

type DocViewerProps = {
  document?: DocDocument;
  zoom: number;
};

function collectAnchoredImageIds(document?: DocDocument) {
  const ids = new Set<string>();
  document?.blocks.forEach((block) => {
    if (block.type === 'paragraph') {
      block.inlines?.forEach((inline) => {
        if (inline.type === 'image') ids.add(inline.image.id);
      });
    } else if (block.type === 'table') {
      block.rows.forEach((row) =>
        row.cells.forEach((cell) =>
          cell.inlines?.forEach((inline) => {
            if (inline.type === 'image') ids.add(inline.image.id);
          }),
        ),
      );
    } else {
      block.items.forEach((item) =>
        item.inlines?.forEach((inline) => {
          if (inline.type === 'image') ids.add(inline.image.id);
        }),
      );
    }
  });
  return ids;
}

function DocViewerComponent({ document, zoom }: DocViewerProps) {
  const page = document?.page;
  const contentWidth = page ? page.width - page.marginLeft - page.marginRight : 0;
  const summaryText = useMemo(
    () =>
      document
        ? `${document.paragraphs.length} 个文本段 / ${document.blocks.length} 个内容块 / ${document.images.length} 张图片`
        : '',
    [document],
  );
  const anchoredImageIds = useMemo(() => collectAnchoredImageIds(document), [document]);
  const unanchoredImages = useMemo(
    () => document?.images.filter((image) => !anchoredImageIds.has(image.id)) ?? [],
    [anchoredImageIds, document],
  );

  if (!document?.blocks.length || !page) {
    return <OfficeEmpty kind="doc" />;
  }

  return (
    <div className="oxv-doc-viewer">
      <div className="oxv-doc-viewer__header">
        <Typography.Text strong ellipsis className="oxv-doc-viewer__title">
          {document.title}
        </Typography.Text>
        <Typography.Text type="secondary" className="oxv-doc-viewer__summary">
          {summaryText}
        </Typography.Text>
      </div>
      <div className="oxv-doc-viewer__scroller">
        <DocPageFrame page={page} zoom={zoom}>
          {document.warnings.length ? (
            <div className="oxv-doc-viewer__warning">
              <Alert type="warning" showIcon message="DOC 预览说明" description={document.warnings.join(' ')} />
            </div>
          ) : null}
          <DocContentRenderer blocks={document.blocks} contentWidth={contentWidth} />
          <DocImageGallery images={unanchoredImages} />
        </DocPageFrame>
      </div>
    </div>
  );
}

export const DocViewer = memo(DocViewerComponent);
