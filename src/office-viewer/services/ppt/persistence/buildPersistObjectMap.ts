import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { PptParseError } from '../errors';
import type { PptEditChain, PptParseContext } from '../types';
import { readPptCurrentUser } from './readCurrentUser';
import { readPptPersistDirectory } from './readPersistDirectory';
import { readPptUserEdit } from './readUserEdit';

/** 合并增量保存链，构建最终有效的 persist 对象目录。 */
export async function buildPptEditChain(
  documentStream: Uint8Array,
  currentUserStream: Uint8Array,
  context: PptParseContext,
): Promise<PptEditChain> {
  const currentUser = readPptCurrentUser(currentUserStream, context);
  const visited = new Set<number>();
  const editOffsets: number[] = [];
  const persistOffsets = new Map<number, number>();
  let editOffset = currentUser.offsetToCurrentEdit;
  let documentPersistId = 0;
  let persistIdSeed = 0;

  while (editOffset) {
    if (visited.has(editOffset)) {
      throw new PptParseError(
        'PPT_EDIT_CHAIN_CYCLE',
        'PowerPoint 增量保存链存在循环',
        { offset: editOffset },
      );
    }
    visited.add(editOffset);
    editOffsets.push(editOffset);

    const edit = readPptUserEdit(documentStream, editOffset);
    if (edit.encryptSessionPersistId) {
      throw new PptParseError('PPT_ENCRYPTED', '暂不支持加密的 PPT 文件', {
        offset: editOffset,
        recordType: PPT_RECORD.USER_EDIT_ATOM,
      });
    }
    if (!documentPersistId) documentPersistId = edit.documentPersistId;
    persistIdSeed = Math.max(persistIdSeed, edit.persistIdSeed);

    const directory = readPptPersistDirectory(
      documentStream,
      edit.offsetPersistDirectory,
      edit.offsetLastEdit,
    );
    directory.forEach((offset, persistId) => {
      if (!persistOffsets.has(persistId)) {
        persistOffsets.set(persistId, offset);
      }
    });
    editOffset = edit.offsetLastEdit;
    await context.yieldIfNeeded();
  }

  const documentOffset = persistOffsets.get(documentPersistId);
  if (documentOffset === undefined) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '持久化目录中缺少根文档对象',
    );
  }
  const documentRecord = new PptRecordReader(
    documentStream,
    documentOffset,
    documentStream.length,
  ).readRecord();
  if (documentRecord?.type !== PPT_RECORD.DOCUMENT) {
    throw new PptParseError(
      'PPT_DOCUMENT_MISSING',
      '根持久化对象不是 DocumentContainer',
      { offset: documentOffset, recordType: documentRecord?.type },
    );
  }

  return {
    documentPersistId,
    persistIdSeed,
    persistOffsets,
    editOffsets,
  };
}
