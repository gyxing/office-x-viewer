import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
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
