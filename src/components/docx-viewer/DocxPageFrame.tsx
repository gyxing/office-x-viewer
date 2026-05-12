import { memo, useMemo } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { DocxPage } from '../../services/docx/types';

type DocxPageFrameProps = {
  page: DocxPage;
  zoom: number;
  children: ReactNode;
};

function DocxPageFrameComponent({ page, zoom, children }: DocxPageFrameProps) {
  const scale = zoom / 100;
  const shellStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width * scale,
      minHeight: page.minHeight * scale,
    }),
    [page.minHeight, page.width, scale],
  );
  const articleStyle = useMemo<CSSProperties>(
    () => ({
      width: page.width,
      minHeight: page.minHeight,
      padding: `${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px`,
      borderTop: page.borderTop,
      borderRight: page.borderRight,
      borderBottom: page.borderBottom,
      borderLeft: page.borderLeft,
      transform: `scale(${scale})`,
    }),
    [
      page.borderBottom,
      page.borderLeft,
      page.borderRight,
      page.borderTop,
      page.marginBottom,
      page.marginLeft,
      page.marginRight,
      page.marginTop,
      page.minHeight,
      page.width,
      scale,
    ],
  );

  return (
    <div className="oxv-docx-page-frame" style={shellStyle}>
      <article className="oxv-docx-page-frame__article" style={articleStyle}>
        {children}
      </article>
    </div>
  );
}

export const DocxPageFrame = memo(DocxPageFrameComponent);
