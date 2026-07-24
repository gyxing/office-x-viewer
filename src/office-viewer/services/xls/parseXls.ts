import { XlsDocumentAssembler } from '../parsing/assembly/DocumentAssembler';
import { ResourceRegistry } from '../parsing/assembly/ResourceRegistry';
import { MainThreadRuntime } from '../parsing/runtime/MainThreadRuntime';
import type { SpreadsheetWorkbook } from '../spreadsheet/types';

/** 在纯浏览器中解析未加密的 Excel 97–2003 BIFF8 工作簿。 */
export async function parseXls(file: File): Promise<SpreadsheetWorkbook> {
  const controller = new AbortController();
  const runtime = new MainThreadRuntime();
  const assembler = new XlsDocumentAssembler(new ResourceRegistry());
  let target: SpreadsheetWorkbook | undefined;
  try {
    await runtime.run(file, 'xls', controller.signal, {
      progress: () => undefined,
      resource: (resource) => assembler.addResource(resource),
      sheet: async (index, revision, sheet) =>
        assembler.addSheet(index, revision, sheet),
      presentationMetadata: async () => {
        throw new Error('XLS 主线程运行时返回了错误的演示文稿元数据');
      },
      slide: async () => {
        throw new Error('XLS 主线程运行时返回了错误的幻灯片分块');
      },
      documentMetadata: async () => {
        throw new Error('XLS 主线程运行时返回了错误的文档元数据');
      },
      documentBlocks: async () => {
        throw new Error('XLS 主线程运行时返回了错误的正文分块');
      },
      parsed: async () => {
        throw new Error('XLS 主线程运行时返回了错误的完整文档消息');
      },
      complete: (warnings) => {
        assembler.setWarnings(warnings);
        target = assembler.complete();
      },
      error: () => undefined,
    });
    if (!target) throw new Error('XLS 解析未返回完整工作簿');
    return target;
  } catch (error) {
    assembler.dispose();
    throw error;
  } finally {
    runtime.dispose();
  }
}
