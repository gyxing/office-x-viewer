import type { ParsedOfficeFile } from '../preview';
import type { OfficeParseSession } from './types';

/** OfficeViewer 内部会话可订阅非拥有型快照，公开解析 API 不暴露该能力。 */
export type OfficeViewerParseSession =
  OfficeParseSession<ParsedOfficeFile> & {
    readonly partialResult: ParsedOfficeFile | undefined;
    subscribePartial(
      listener: (parsed: ParsedOfficeFile) => void,
    ): () => void;
  };
