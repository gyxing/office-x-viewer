import { Empty } from 'antd';
import type { PptxDocument } from '../../services/pptx/types';
import { PptxThumbnail } from './PptxThumbnail';
import { PptxSlide } from './PptxSlide';

type PptxViewerProps = {
  document?: PptxDocument;
  activeIndex: number;
  zoom: number;
  onSelectSlide: (index: number) => void;
};

export function PptxViewer({ document, activeIndex, zoom, onSelectSlide }: PptxViewerProps) {
  if (!document?.slides.length) {
    return <Empty description="请先上传 PPTX 文件开始预览" />;
  }

  const currentSlide = document.slides[activeIndex];

  return (
    <div
      style={{
        display: 'flex',
        minHeight: 'calc(100vh - 56px)',
        background: '#eef1f6',
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
        style={{
          flex: '1 1 auto',
          minWidth: 0,
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          overflow: 'auto',
          padding: 32,
        }}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            width: '100%',
            minHeight: '100%',
          }}
        >
          {currentSlide ? <PptxSlide slide={currentSlide} zoom={zoom} /> : null}
        </div>
      </section>
    </div>
  );
}
