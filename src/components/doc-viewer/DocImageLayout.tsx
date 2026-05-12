// DocImageLayout 渲染 DOC 中连续图片段落形成的图片组，并按内容宽度决定排布。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocImage } from '../../services/doc/types';
import { DOC_IMAGE_ROW_GAP, imageRows } from './docRenderUtils';

type DocImageLayoutProps = {
  images: DocImage[];
  contentWidth: number;
};

function DocImageLayoutComponent({ images, contentWidth }: DocImageLayoutProps) {
  const rows = useMemo(() => imageRows(images, contentWidth), [contentWidth, images]);

  if (!images.length) return null;

  return (
    <div className="oxv-doc-image-layout">
      {rows.map((row) => (
        <div key={row.map((image) => image.id).join('-')} className="oxv-doc-image-layout__row" style={{ maxWidth: contentWidth }}>
          {row.map((image) => {
            const preferredWidth = image.width ? Math.min(image.width, contentWidth) : contentWidth;
            const rowWidth =
              row.length > 1 && image.width
                ? Math.min(image.width, (contentWidth - DOC_IMAGE_ROW_GAP) / row.length)
                : preferredWidth;
            const figureStyle: CSSProperties = {
              width: rowWidth,
              maxWidth: row.length > 1 ? `calc((100% - ${DOC_IMAGE_ROW_GAP}px) / ${row.length})` : '100%',
            };

            return (
              <figure key={image.id} className="oxv-doc-image-layout__figure" style={figureStyle}>
                <img className="oxv-doc-image-layout__img" src={image.src} alt={image.caption ?? image.id} />
              </figure>
            );
          })}
        </div>
      ))}
    </div>
  );
}

export const DocImageLayout = memo(DocImageLayoutComponent);
