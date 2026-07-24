import { CfbParseError, parseCfb } from '../../shared/binary/cfb';
import type {
  PortablePresentationMetadata,
  PortableResource,
} from '../parsing/protocol/messages';
import type { ParseProgress } from '../parsing/types';
import type {
  PresentationDocument,
  SlideModel,
} from '../presentation/types';
import { adaptPptDocument, adaptPptSlide } from './adapter';
import { readPptEmbeddedCharts } from './chart';
import { readPptBinaryDocument } from './document';
import { PptParseError } from './errors';
import { readPptPictures } from './images';
import { buildPptEditChain } from './persistence';
import {
  createPptParseContext,
  type PptMasterModel,
  type PptParseContext,
} from './types';

export type PptCoreOutput = {
  resource(resource: PortableResource): Promise<void>;
  presentationMetadata(
    metadata: PortablePresentationMetadata,
  ): Promise<void>;
  slide(index: number, slide: SlideModel): Promise<void>;
};

export type PptCoreContext = {
  checkpoint(progress?: ParseProgress): Promise<void>;
  output?: PptCoreOutput;
};

export type PptCoreResult = {
  document: PresentationDocument;
  resources: PortableResource[];
};

/** 消费已生成资源，确保引用它们的幻灯片不会先到达主线程。 */
async function flushPptResources(
  context: PptParseContext,
  output: PptCoreOutput | undefined,
) {
  if (!output || !context.resources.length) return;
  const resources = context.resources.splice(0);
  for (const resource of resources) {
    await output.resource(resource);
  }
}

function hasVbaStorage(
  entries: Awaited<ReturnType<typeof parseCfb>>['entries'],
) {
  // 宏只做目录级检测，绝不读取、反编译或执行 VBA 字节码。
  return entries.some((entry) => {
    const value = `${entry.path}/${entry.name}`.toLowerCase();
    return value.includes('/vba/') || value.endsWith('/vba');
  });
}

/** 解析 PPT 二进制并返回环境无关的演示文稿与资源。 */
export async function parsePptCore(
  input: ArrayBuffer | Uint8Array,
  coreContext: PptCoreContext,
): Promise<PptCoreResult> {
  const context = createPptParseContext(() => coreContext.checkpoint());
  await coreContext.checkpoint({
    stage: 'container',
    percent: 0.05,
    message: '正在读取 PPT 复合文档',
  });

  let cfb: Awaited<ReturnType<typeof parseCfb>>;
  try {
    cfb = await parseCfb(input, {
      yieldIfNeeded: context.yieldIfNeeded,
    });
  } catch (error) {
    if (error instanceof CfbParseError) {
      throw new PptParseError(
        'PPT_INVALID_RECORD',
        `PPT 复合文档容器损坏：${error.message}`,
      );
    }
    throw error;
  }

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

  await coreContext.checkpoint({
    stage: 'structure',
    percent: 0.2,
    message: '正在解析 PPT 文档结构',
  });
  const editChain = await buildPptEditChain(
    documentStream,
    currentUserStream,
    context,
  );

  await coreContext.checkpoint({
    stage: 'resources',
    percent: 0.35,
    message: '正在解析 PPT 图片资源',
  });
  await readPptPictures(cfb.getStream('Pictures'), context);

  await coreContext.checkpoint({
    stage: 'resources',
    percent: 0.5,
    message: '正在解析 PPT 嵌入图表',
  });
  await readPptEmbeddedCharts(documentStream, editChain, context);

  await coreContext.checkpoint({
    stage: 'content',
    percent: 0.65,
    message: '正在解析 PPT 幻灯片',
  });
  let currentMasters = new Map<number, PptMasterModel>();
  const binary = await readPptBinaryDocument(
    documentStream,
    editChain,
    context,
    coreContext.output
      ? {
          structure: async ({ width, height, theme, masters }) => {
            currentMasters = masters;
            await flushPptResources(context, coreContext.output);
            await coreContext.output!.presentationMetadata({
              width,
              height,
              theme,
              warnings: [...context.warnings],
            });
          },
          slide: async (index, slide) => {
            await flushPptResources(context, coreContext.output);
            await coreContext.output!.slide(
              index,
              adaptPptSlide(slide, currentMasters),
            );
          },
        }
      : undefined,
  );
  if (coreContext.output) {
    await flushPptResources(context, coreContext.output);
    await coreContext.output.presentationMetadata({
      width: binary.width,
      height: binary.height,
      theme: binary.theme,
      warnings: [...binary.warnings],
    });
  }

  await coreContext.checkpoint({
    stage: 'assembling',
    percent: 0.95,
    message: '正在组装演示文稿',
  });
  return {
    document: adaptPptDocument(binary),
    resources: context.resources,
  };
}
