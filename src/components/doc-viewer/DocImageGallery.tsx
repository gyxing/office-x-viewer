// DocImageGallery 展示未能稳定锚定到正文位置的 DOC 图片。
import { Typography } from 'antd';
import { memo } from 'react';
import type { DocImage } from '../../services/doc/types';

type DocImageGalleryProps = {
  images: DocImage[];
};

function DocImageGalleryComponent({ images }: DocImageGalleryProps) {
  if (!images.length) return null;

  return (
    <section className="oxv-doc-image-gallery">
      <Typography.Text strong className="oxv-doc-image-gallery__title">
        文档图片
      </Typography.Text>
      <div className="oxv-doc-image-gallery__grid">
        {images.map((image) => (
          <figure key={image.id} className="oxv-doc-image-gallery__figure">
            <img className="oxv-doc-image-gallery__img" src={image.src} alt={image.caption ?? image.id} />
          </figure>
        ))}
      </div>
    </section>
  );
}

export const DocImageGallery = memo(DocImageGalleryComponent);
