// DocxImage 渲染 DOCX 行内图片，并保留文档解析出的图片宽度。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { DocxImageInline } from '../../services/docx/types';
import { calculatePositionStyle } from './positionUtils';

type DocxImageProps = {
  inline: DocxImageInline;
};

// 自定义变量把解析后的图片尺寸交给 Less，避免静态样式散落在 JSX 中。
type DocxImageStyle = CSSProperties & {
  '--oxv-docx-inline-image-width': string;
  '--oxv-docx-inline-image-height': string;
};

function DocxImageComponent({ inline }: DocxImageProps) {
  const image = inline.image;
  const positionStyle = calculatePositionStyle(image.position);

  const imageStyle = useMemo<DocxImageStyle>(
    () => ({
      '--oxv-docx-inline-image-width': `${image.width}px`,
      '--oxv-docx-inline-image-height': `${image.height}px`,
      ...positionStyle,
      maxWidth: image.position ? 'none' : undefined,
    }),
    [image.height, image.position, image.width, positionStyle],
  );

  return (
    <img
      className="oxv-docx-inline-image"
      src={image.src}
      alt={image.alt ?? ''}
      title={image.name}
      style={imageStyle}
    />
  );
}

export const DocxImage = memo(DocxImageComponent);
