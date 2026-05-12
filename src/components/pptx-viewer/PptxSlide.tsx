// PptxSlide 按 PPTX 解析模型渲染单页幻灯片背景、图形、文本、图片、表格和图表。
import { memo, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { SlideElement, SlideModel } from '../../services/pptx/types';
import { OfficeChartView } from '../office-chart/OfficeChartView';
import { ImageRenderer } from './renderers/ImageRenderer';
import { colorWithOpacity } from './renderers/paint';
import { ShapeRenderer } from './renderers/ShapeRenderer';
import { TableRenderer } from './renderers/TableRenderer';
import { TextRenderer } from './renderers/TextRenderer';
import { UnsupportedRenderer } from './renderers/UnsupportedRenderer';

type PptxSlideProps = {
  slide: SlideModel;
  zoom: number;
  renderKey?: string;
};

const ChartFrame = memo(function ChartFrame({
  element,
  zoom,
}: {
  element: Extract<SlideElement, { type: 'chart' }>;
  zoom: number;
}) {
  const frameStyle = useMemo<CSSProperties>(
    () => ({
      left: element.x,
      top: element.y,
      width: element.width,
      height: element.height,
    }),
    [element.height, element.width, element.x, element.y],
  );

  return (
    <div className="oxv-pptx-chart-frame" style={frameStyle}>
      <OfficeChartView chart={element.chart} width={element.width} height={element.height} zoom={zoom} />
    </div>
  );
});

function PptxSlideComponent({ slide, zoom, renderKey }: PptxSlideProps) {
  const scale = zoom / 100;
  // renderKey 会参与 SVG 渐变 id，缩略图和主画布同时渲染同一页时必须保持 id 隔离。
  const slideRenderKey = renderKey ?? `slide-${slide.id}`;
  const slideStyle = useMemo<CSSProperties>(
    () => ({
      width: slide.width,
      height: slide.height,
      minWidth: slide.width,
      minHeight: slide.height,
      transform: `scale(${scale})`,
    }),
    [scale, slide.height, slide.width],
  );
  const backgroundStyle = useMemo<CSSProperties>(
    () => ({
      background: colorWithOpacity(slide.background?.fill ?? '#fff', slide.background?.fillOpacity),
      backgroundImage: slide.background?.imageRef ? `url(${slide.background.imageRef})` : undefined,
    }),
    [slide.background?.fill, slide.background?.fillOpacity, slide.background?.imageRef],
  );

  return (
    <div className="oxv-pptx-slide" style={slideStyle}>
      <div className="oxv-pptx-slide__background" style={backgroundStyle} />
      <div className="oxv-pptx-slide__elements">
        {slide.elements.map((element) => {
          switch (element.type) {
            case 'text':
              return <TextRenderer key={element.id} element={element} renderKey={slideRenderKey} />;
            case 'shape':
              return <ShapeRenderer key={element.id} element={element} renderKey={slideRenderKey} />;
            case 'image':
              return <ImageRenderer key={element.id} element={element} />;
            case 'table':
              return <TableRenderer key={element.id} element={element} />;
            case 'chart':
              return <ChartFrame key={element.id} element={element} zoom={zoom} />;
            case 'unsupported':
              return <UnsupportedRenderer key={element.id} element={element} />;
            default:
              return null;
          }
        })}
      </div>
    </div>
  );
}

export const PptxSlide = memo(PptxSlideComponent);
