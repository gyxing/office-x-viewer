import React, { memo, useMemo } from 'react';
import type { DocDocument } from '../../services/doc/types';
import { OfficeEmpty } from '../../shell/Empty';
import { DocContentRenderer } from './DocContentRenderer';
import { DocImageGallery } from './DocImageGallery';
import { DocPageFrame } from './DocPageFrame';
import { paginateDocBlocks } from './docRenderUtils';
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

// DocViewer 负责旧版 DOC/WPS 降级预览的固定警告和页面滚动区。
function DocViewerComponent({ document, zoom }: DocViewerProps) {
  const page = document?.page;
  const contentWidth = page
    ? page.width - page.marginLeft - page.marginRight
    : 0;
  const pages = useMemo(
    () =>
      document && page
        ? paginateDocBlocks(document.blocks, page, contentWidth)
        : [],
    [contentWidth, document, page],
  );
  const anchoredImageIds = useMemo(
    () => collectAnchoredImageIds(document),
    [document],
  );
  const unanchoredImages = useMemo(
    () =>
      document?.images.filter((image) => !anchoredImageIds.has(image.id)) ?? [],
    [anchoredImageIds, document],
  );

  if (!document?.blocks.length || !page || !pages.length) {
    return <OfficeEmpty kind="doc" />;
  }

  return (
    <div className="oxv-doc-viewer">
      {document.warnings.length ? (
        <div className="oxv-doc-viewer__notice" role="alert">
          {document.warnings.join(' ')}
        </div>
      ) : null}
      <div className="oxv-doc-viewer__scroller">
        {pages.map((docPage, pageIndex) => (
          <DocPageFrame key={docPage.id} page={page} zoom={zoom}>
            <DocContentRenderer
              blocks={docPage.blocks}
              contentWidth={contentWidth}
            />
            {pageIndex === pages.length - 1 ? (
              <DocImageGallery images={unanchoredImages} />
            ) : null}
          </DocPageFrame>
        ))}
      </div>
    </div>
  );
}

export const DocViewer = memo(DocViewerComponent);
