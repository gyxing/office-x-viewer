import { CfbParseError, parseCfb } from '../../shared/binary/cfb';
import { disposePresentationDocument } from '../presentation/dispose';
import type { PresentationDocument } from '../presentation/types';
import { adaptPptDocument } from './adapter';
import { readPptEmbeddedCharts } from './chart';
import { readPptBinaryDocument } from './document';
import { PptParseError } from './errors';
import { readPptPictures } from './images';
import { buildPptEditChain } from './persistence';
import { createPptParseContext } from './types';

function createYieldIfNeeded() {
  let deadline = performance.now() + 12;
  return async () => {
    if (performance.now() < deadline) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    deadline = performance.now() + 12;
  };
}

function hasVbaStorage(entries: Awaited<ReturnType<typeof parseCfb>>['entries']) {
  return entries.some((entry) => {
    const value = `${entry.path}/${entry.name}`.toLowerCase();
    return value.includes('/vba/') || value.endsWith('/vba');
  });
}

/** 在纯浏览器中解析未加密的 PowerPoint 97–2003 PPT 文件。 */
export async function parsePpt(file: File): Promise<PresentationDocument> {
  const context = createPptParseContext(createYieldIfNeeded());
  let target: PresentationDocument | undefined;
  try {
    const cfb = await parseCfb(await file.arrayBuffer(), {
      yieldIfNeeded: context.yieldIfNeeded,
    });
    const documentStream = cfb.getStream('PowerPoint Document');
    const currentUserStream = cfb.getStream('Current User');
    if (!documentStream || !currentUserStream) {
      throw new PptParseError(
        'PPT_REQUIRED_STREAM_MISSING',
        'PPT 文件缺少 PowerPoint Document 或 Current User 数据流',
      );
    }
    if (hasVbaStorage(cfb.entries)) {
      context.warnings.push({
        code: 'PPT_VBA_DETECTED',
        message: '检测到 VBA 工程；预览器不会读取或执行宏代码',
      });
    }

    const editChain = await buildPptEditChain(
      documentStream,
      currentUserStream,
      context,
    );
    await readPptPictures(cfb.getStream('Pictures'), context);
    await readPptEmbeddedCharts(documentStream, editChain, context);
    const binary = await readPptBinaryDocument(
      documentStream,
      editChain,
      context,
    );
    target = adaptPptDocument(binary, context);
    return target;
  } catch (error) {
    if (target) disposePresentationDocument(target);
    else {
      for (const url of context.objectUrls) URL.revokeObjectURL(url);
      context.objectUrls.clear();
    }
    if (error instanceof CfbParseError) {
      throw new PptParseError(
        'PPT_INVALID_RECORD',
        `PPT 复合文档容器损坏：${error.message}`,
      );
    }
    throw error;
  }
}
