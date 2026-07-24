import { PptDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import type { PresentationDocument } from '../presentation/types';
import { parsePptCore } from './parsePptCore';

function createYieldIfNeeded() {
  let deadline = Date.now() + 12;
  return async () => {
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    deadline = Date.now() + 12;
  };
}

/** 在纯浏览器中解析未加密的 PowerPoint 97–2003 PPT 文件。 */
export async function parsePpt(file: File): Promise<PresentationDocument> {
  const assembler = new PptDocumentAssembler(new ResourceRegistry());
  const yieldIfNeeded = createYieldIfNeeded();
  try {
    const result = await parsePptCore(await file.arrayBuffer(), {
      checkpoint: yieldIfNeeded,
    });
    for (const resource of result.resources) {
      await assembler.addResource(resource);
    }
    const { width, height, theme, warnings, slides } = result.document;
    assembler.setMetadata({ width, height, theme, warnings });
    slides.forEach((slide, index) => assembler.addSlide(index, slide));
    return assembler.complete();
  } catch (error) {
    assembler.dispose();
    throw error;
  }
}
