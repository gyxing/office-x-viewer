import React, { memo, useMemo } from 'react';
import type { DocxDocument } from '../../services/docx/types';
import { OfficeEmpty } from '../../shell/Empty';
import { DocxBlockRenderer } from './DocxBlockRenderer';
import { DocxPageFrame } from './DocxPageFrame';
import './index.less';

type DocxViewerProps = {
  document?: DocxDocument;
  zoom: number;
};

// DocxViewer 负责 DOCX 页面内容的缩放渲染和滚动布局。
function DocxViewerComponent({ document, zoom }: DocxViewerProps) {
  const pages = useMemo(
    () =>
      document
        ? document.pages?.length
          ? document.pages
          : [
              {
                id: 'docx-page-1',
                page: document.page,
                blocks: document.blocks,
              },
            ]
        : [],
    [document],
  );
  if (!document?.blocks.length || !pages.length) {
    return <OfficeEmpty kind="docx" />;
  }

  return (
    <div className="oxv-docx-viewer">
      <div className="oxv-docx-viewer__scroller">
        {pages.map((pageItem) => {
          const contentWidth =
            pageItem.page.width -
            pageItem.page.marginLeft -
            pageItem.page.marginRight;
          return (
            <DocxPageFrame key={pageItem.id} page={pageItem.page} zoom={zoom}>
              {pageItem.blocks.map((block) => (
                <DocxBlockRenderer
                  key={block.id}
                  block={block}
                  availableWidth={contentWidth}
                />
              ))}
            </DocxPageFrame>
          );
        })}
      </div>
    </div>
  );
}

export const DocxViewer = memo(DocxViewerComponent);
