import { memo } from 'react';
import type { DocxImageInline } from '../../services/docx/types';

type DocxImageProps = {
  inline: DocxImageInline;
};

function DocxImageComponent({ inline }: DocxImageProps) {
  const image = inline.image;
  return (
    <img
      src={image.src}
      alt={image.alt ?? ''}
      title={image.name}
      style={{
        display: 'inline-block',
        width: image.width,
        maxWidth: '100%',
        height: 'auto',
        verticalAlign: 'middle',
      }}
    />
  );
}

export const DocxImage = memo(DocxImageComponent);

