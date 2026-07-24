// XlsxFloatingImages 渲染锚定在工作表画布上的浮动图片。
import type { CSSProperties } from 'react';
import React, { memo, useMemo } from 'react';
import type { XlsxImage } from '../../services/xlsx/types';

type XlsxFloatingImagesProps = {
  images: XlsxImage[];
};

function XlsxFloatingImage({ image }: { image: XlsxImage }) {
  const imageStyle = useMemo<CSSProperties>(
    () => ({
      left: 48 + image.x,
      top: 28 + image.y,
      width: image.width,
      height: image.height,
    }),
    [image.height, image.width, image.x, image.y],
  );

  return (
    <img
      className="oxv-xlsx-sheet-grid__floating-image"
      src={image.src}
      alt={image.alt ?? ''}
      title={image.name}
      style={imageStyle}
      onError={(event) => {
        event.currentTarget.setAttribute(
          'aria-label',
          image.alt ? `${image.alt}（图片加载失败）` : '图片加载失败',
        );
        event.currentTarget.setAttribute('data-load-error', 'true');
      }}
    />
  );
}

const MemoXlsxFloatingImage = memo(XlsxFloatingImage);

function XlsxFloatingImagesComponent({ images }: XlsxFloatingImagesProps) {
  return (
    <>
      {images.map((image) => (
        <MemoXlsxFloatingImage key={image.id} image={image} />
      ))}
    </>
  );
}

export const XlsxFloatingImages = memo(XlsxFloatingImagesComponent);
