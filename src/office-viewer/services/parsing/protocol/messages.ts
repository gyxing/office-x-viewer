import type {
  SpreadsheetSheet,
  SpreadsheetWarning,
} from '../../spreadsheet/types';
import type {
  PresentationDocument,
  SlideModel,
} from '../../presentation/types';
import type { DocBlock, DocDocument } from '../../doc/types';
import type { ParseProgress } from '../types';
import type { SerializedParseError } from './errors';

export type PortablePresentationMetadata = Omit<
  PresentationDocument,
  'slides' | 'resources'
>;

export type PortableDocMetadata = Omit<
  DocDocument,
  'blocks' | 'paragraphs' | 'resources'
>;

export type PortableResource =
  | {
      id: string;
      encoding: 'binary';
      mimeType: string;
      buffer: ArrayBuffer;
    }
  | {
      id: string;
      encoding: 'text';
      mimeType: 'image/svg+xml';
      text: string;
    }
  | {
      id: string;
      encoding: 'rgba';
      mimeType: 'image/png';
      width: number;
      height: number;
      buffer: ArrayBuffer;
    };

export type MainToWorkerMessage =
  | {
      type: 'parse-start';
      version: number;
      taskId: string;
      kind: 'xls' | 'ppt' | 'doc';
      fileName: string;
      buffer: ArrayBuffer;
    }
  | {
      type: 'parse-cancel';
      version: number;
      taskId: string;
    }
  | {
      type: 'parse-ack';
      version: number;
      taskId: string;
      sequence: number;
    };

export type WorkerToMainMessage =
  | {
      type: 'worker-ready';
      version: number;
    }
  | {
      type: 'parse-progress';
      version: number;
      taskId: string;
      progress: ParseProgress;
    }
  | {
      type: 'parse-resource';
      version: number;
      taskId: string;
      sequence: number;
      resource: PortableResource;
    }
  | {
      type: 'parse-sheet';
      version: number;
      taskId: string;
      sequence: number;
      sheetIndex: number;
      revision: number;
      sheet: SpreadsheetSheet;
    }
  | {
      type: 'parse-presentation-meta';
      version: number;
      taskId: string;
      sequence: number;
      metadata: PortablePresentationMetadata;
    }
  | {
      type: 'parse-slide';
      version: number;
      taskId: string;
      sequence: number;
      slideIndex: number;
      slide: SlideModel;
    }
  | {
      type: 'parse-document-meta';
      version: number;
      taskId: string;
      sequence: number;
      metadata: PortableDocMetadata;
    }
  | {
      type: 'parse-document-blocks';
      version: number;
      taskId: string;
      sequence: number;
      startIndex: number;
      blocks: DocBlock[];
    }
  | {
      type: 'parse-complete';
      version: number;
      taskId: string;
      warnings?: SpreadsheetWarning[];
    }
  | {
      type: 'parse-error';
      version: number;
      taskId: string;
      error: SerializedParseError;
    }
  | {
      type: 'parse-cancelled';
      version: number;
      taskId: string;
    };
