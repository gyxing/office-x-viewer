import type { ThemeModel } from '../../presentation/types';
import { PPT_RECORD } from '../binary/constants';
import { PptRecordReader } from '../binary/PptRecordReader';
import { parsePptDrawing } from '../drawing';
import type {
  PptEditChain,
  PptMasterModel,
  PptParseContext,
} from '../types';
import type { PptMasterDescriptor } from './readSlideLists';

function readRgbColor(bytes: Uint8Array, offset: number) {
  return `#${[bytes[offset + 2], bytes[offset + 1], bytes[offset]]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

function applyColorScheme(theme: ThemeModel, bytes: Uint8Array) {
  if (bytes.length < 32) return;
  const colors = Array.from({ length: 8 }, (_, index) =>
    readRgbColor(bytes, index * 4),
  );
  theme.colorScheme = {
    ...theme.colorScheme,
    lt1: colors[0],
    dk1: colors[1],
    dk2: colors[2],
    lt2: colors[3],
    accent1: colors[4],
    accent2: colors[5],
    hlink: colors[6],
    folHlink: colors[7],
  };
}

/** 读取一个主母版的主题、背景与可继承绘图元素。 */
export function readPptMaster(
  documentStream: Uint8Array,
  editChain: PptEditChain,
  descriptor: PptMasterDescriptor,
  theme: ThemeModel,
  context: PptParseContext,
): PptMasterModel | undefined {
  const offset = editChain.persistOffsets.get(descriptor.persistId);
  if (offset === undefined) {
    context.warnings.push({
      code: 'PPT_MASTER_MISSING',
      message: `持久化目录中缺少母版 ${descriptor.masterId}`,
    });
    return undefined;
  }
  const record = new PptRecordReader(
    documentStream,
    offset,
    documentStream.length,
  ).readRecord();
  if (
    !record ||
    (record.type !== PPT_RECORD.MAIN_MASTER &&
      record.type !== PPT_RECORD.SLIDE)
  ) {
    context.warnings.push({
      code: 'PPT_MASTER_INVALID',
      message: `母版 ${descriptor.masterId} 的记录类型无效`,
      offset,
    });
    return undefined;
  }

  const children = new PptRecordReader(
    documentStream,
    record.dataOffset,
    record.endOffset,
  );
  let drawing: Uint8Array | undefined;
  for (const child of children.records()) {
    if (child.type === PPT_RECORD.COLOR_SCHEME_ATOM) {
      applyColorScheme(theme, child.data);
    }
    if (child.type === PPT_RECORD.PP_DRAWING) drawing = child.data;
  }

  return {
    id: descriptor.masterId,
    persistId: descriptor.persistId,
    background: { fill: theme.colorScheme.lt1 ?? '#ffffff' },
    elements: drawing ? parsePptDrawing(drawing, theme, context) : [],
  };
}
