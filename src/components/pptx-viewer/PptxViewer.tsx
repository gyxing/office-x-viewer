import { memo, useEffect, useRef } from 'react';
import type { PptxDocument } from '../../services/pptx/types';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';
import { PptxThumbnail } from './PptxThumbnail';
import { PptxSlide } from './PptxSlide';

type PptxViewerProps = {
  document?: PptxDocument;
  activeIndex: number;
  zoom: number;
  onSelectSlide: (index: number) => void;
};

function PptxViewerComponent({ document, activeIndex, zoom, onSelectSlide }: PptxViewerProps) {
  const viewportRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    viewportRef.current?.scrollTo({ left: 0, top: 0 });
  }, [activeIndex, zoom]);

  if (!document?.slides.length) {
    return <OfficeEmpty kind="pptx" />;
  }

  const currentSlide = document.slides[activeIndex];

  return (
    <div
      style={{
        display: 'flex',
        height: 'calc(100vh - 56px)',
        background: '#eef1f6',
        overflow: 'hidden',
      }}
    >
      <aside
        style={{
          width: 292,
          flex: '0 0 292px',
          background: '#fff',
          borderRight: '1px solid #dde3ec',
          display: 'flex',
          flexDirection: 'column',
          height: 'calc(100vh - 56px)',
          position: 'sticky',
          top: 56,
          overflow: 'hidden',
        }}
      >
        <div style={{ padding: 16, borderBottom: '1px solid #eef2f7', flex: '0 0 auto' }}>
          <div style={{ fontSize: 12, color: '#667085' }}>共 {document.slides.length} 页</div>
        </div>
        <div style={{ flex: '1 1 auto', overflowY: 'auto', padding: 16 }}>
          {document.slides.map((slide, index) => (
            <div
              key={slide.id}
              style={{ marginBottom: 12, cursor: 'pointer' }}
              onClick={() => onSelectSlide(index)}
            >
              <PptxThumbnail slide={slide} active={index === activeIndex} />
            </div>
          ))}
        </div>
      </aside>
      <section
        ref={viewportRef}
        style={{
          flex: '1 1 auto',
          minWidth: 0,
          minHeight: 0,
          display: 'flex',
          justifyContent: 'flex-start',
          alignItems: 'flex-start',
          overflow: 'auto',
          padding: 32,
        }}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'flex-start',
            alignItems: 'flex-start',
            width: '100%',
            minHeight: '100%',
          }}
        >
          {currentSlide ? <PptxSlide slide={currentSlide} zoom={zoom} renderKey={`slide-${currentSlide.id}`} /> : null}
        </div>
      </section>
    </div>
  );
}

export const PptxViewer = memo(PptxViewerComponent);

