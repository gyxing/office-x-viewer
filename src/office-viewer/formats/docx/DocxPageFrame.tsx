// DocxPageFrame 提供 DOCX 页面框架，负责页宽、页高、页边距、边框和缩放。
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
  // DOCX 的边框和页边距来自文档本身，放在 article 上才能随页面坐标系一起缩放。
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
