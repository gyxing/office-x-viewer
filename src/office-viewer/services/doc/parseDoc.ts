import { DocDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import {
  chunkDocBlocks,
  documentMetadataFromDoc,
} from './chunkDocBlocks';
import { parseDocCore } from './parseDocCore';
import type { DocDocument } from './types';

function createYieldIfNeeded() {
  let deadline = Date.now() + 12;
  return async () => {
    if (Date.now() < deadline) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    deadline = Date.now() + 12;
  };
}

/** 保留原有 DOC/WPS 主线程解析入口，并复用跨运行时组装流程。 */
export async function parseDoc(file: File): Promise<DocDocument> {
  const assembler = new DocDocumentAssembler(new ResourceRegistry());
  const yieldIfNeeded = createYieldIfNeeded();
  try {
    const result = await parseDocCore(await file.arrayBuffer(), {
      fileName: file.name,
      checkpoint: yieldIfNeeded,
    });
    assembler.setMetadata(documentMetadataFromDoc(result.document));
    for (const chunk of chunkDocBlocks(result.document.blocks)) {
      assembler.addBlocks(chunk.startIndex, chunk.blocks);
    }
    for (const resource of result.resources) {
      await assembler.addResource(resource);
    }
    return assembler.complete();
  } catch (error) {
    assembler.dispose();
    throw error;
  }
}
