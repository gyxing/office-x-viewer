import type {
  PresentationDocument,
  SlideElement,
} from '../presentation/types';
import type { PptBinaryDocument, PptParseContext } from './types';

function cloneMasterElement(
  element: SlideElement,
  slideId: string,
  index: number,
): SlideElement {
  return {
    ...element,
    id: `${slideId}-master-${index}-${element.id}`,
    zIndex: index,
  };
}

/** 将 PPT 私有结构适配为 PPTX 渲染器复用的统一演示文稿模型。 */
export function adaptPptDocument(
  source: PptBinaryDocument,
  context: PptParseContext,
): PresentationDocument {
  return {
    width: source.width,
    height: source.height,
    theme: source.theme,
    slides: source.slides.map((slide) => {
      const masterElements =
        source.masters.get(slide.masterId ?? Number.NaN)?.elements ?? [];
      const inherited = masterElements.map((element, index) =>
        cloneMasterElement(element, slide.id, index),
      );
      return {
        id: slide.id,
        index: slide.index,
        width: slide.width,
        height: slide.height,
        hidden: slide.hidden,
        background: slide.background,
        elements: [
          ...inherited,
          ...slide.elements.map((element, index) => ({
            ...element,
            zIndex: inherited.length + index,
          })),
        ],
      };
    }),
    warnings: source.warnings,
    resources: { objectUrls: [...context.objectUrls] },
  };
}
