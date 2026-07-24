import type { ParsedOfficeFile } from '../../preview';
import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import type { SlideModel } from '../../presentation/types';
import type { DocBlock } from '../../doc/types';
import type {
  PortableDocMetadata,
  PortablePresentationMetadata,
  PortableResource,
} from '../protocol/messages';
import type { ParseProgress } from '../types';

export type RuntimeSink = {
  progress(progress: ParseProgress): void;
  resource(resource: PortableResource): Promise<void>;
  sheet(
    index: number,
    revision: number,
    sheet: SpreadsheetSheet,
  ): Promise<void>;
  presentationMetadata(
    metadata: PortablePresentationMetadata,
  ): Promise<void>;
  slide(index: number, slide: SlideModel): Promise<void>;
  documentMetadata(metadata: PortableDocMetadata): Promise<void>;
  documentBlocks(startIndex: number, blocks: DocBlock[]): Promise<void>;
  parsed(parsed: ParsedOfficeFile): Promise<void>;
  complete(warnings?: SpreadsheetWarning[]): void;
  error(error: unknown): void;
};

/** 创建跨运行时一致的取消错误。 */
export function createParseAbortError() {
  const error = new Error('文件解析已取消');
  error.name = 'AbortError';
  return error;
}
