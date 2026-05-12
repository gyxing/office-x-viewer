// DocxImage 渲染 DOCX 行内图片，并保留文档解析出的图片宽度。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { DocxImageInline } from '../../services/docx/types';

type DocxImageProps = {
  inline: DocxImageInline;
};

function DocxImageComponent({ inline }: DocxImageProps) {
  const image = inline.image;
  const imageStyle = useMemo<CSSProperties>(
    () =>
      ({
        '--oxv-docx-inline-image-width': `${image.width}px`,
      }) as CSSProperties,
    [image.width],
  );

  return <img className="oxv-docx-inline-image" src={image.src} alt={image.alt ?? ''} title={image.name} style={imageStyle} />;
}

export const DocxImage = memo(DocxImageComponent);
