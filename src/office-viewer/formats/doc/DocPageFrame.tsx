// DocPageFrame 提供 DOC 降级预览的页面框架，负责页面尺寸、页边距和缩放。
import { memo, useMemo } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { DocPage } from '../../services/doc/types';

type DocPageFrameProps = {
  page: DocPage;
  zoom: number;
  children: ReactNode;
};

function DocPageFrameComponent({ page, zoom, children }: DocPageFrameProps) {
  const scale = zoom / 100;
  // 外层占位使用缩放后的尺寸，内层 article 保留原始 Word 坐标系并用 transform 缩放。
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
      transform: `scale(${scale})`,
    }),
    [page.marginBottom, page.marginLeft, page.marginRight, page.marginTop, page.minHeight, page.width, scale],
  );

  return (
    <div className="oxv-doc-page-frame" style={shellStyle}>
      <article className="oxv-doc-page-frame__article" style={articleStyle}>
        {children}
      </article>
    </div>
  );
}

export const DocPageFrame = memo(DocPageFrameComponent);
