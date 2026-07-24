import { CfbParseError, parseCfb } from '../../shared/binary/cfb';
import type { SpreadsheetWorkbook } from '../spreadsheet/types';
import { disposeSpreadsheetWorkbook } from '../spreadsheet/types';
import {
  adaptBiff8Workbook,
  attachBiff8Charts,
  attachBiff8DrawingImages,
} from './adapter';
import {
  createParseYieldState,
  yieldToBrowserIfNeeded,
} from './biff8/Biff8Reader';
import { parseBiff8Globals } from './biff8/parseGlobals';
import { parseBiff8Worksheet } from './biff8/parseWorksheet';
import { readBiff8ChartSubstream } from './chart/parseCharts';
import { XlsParseError } from './errors';
import type { Biff8Workbook, Biff8Worksheet } from './types';

function mapCfbError(error: CfbParseError) {
  const corruptedChain =
    error.code === 'CHAIN_CYCLE' ||
    error.code === 'CHAIN_TRUNCATED' ||
    error.code === 'SECTOR_OUT_OF_RANGE';
  return new XlsParseError(
    corruptedChain ? 'CORRUPTED_SECTOR_CHAIN' : 'INVALID_CFB',
    `XLS 容器解析失败：${error.message}`,
  );
}

function hasVbaStorage(
  entries: Awaited<ReturnType<typeof parseCfb>>['entries'],
) {
  // 宏只做目录级检测，绝不读取、反编译或执行 VBA 字节码。
  return entries.some((entry) => {
    const normalized = entry.path.toLowerCase();
    return (
      normalized.includes('/vba/') ||
      entry.name.toLowerCase() === '_vba_project_cur' ||
      entry.name.toLowerCase() === 'vba'
    );
  });
}

/** 在纯浏览器中解析未加密的 Excel 97–2003 BIFF8 工作簿。 */
export async function parseXls(file: File): Promise<SpreadsheetWorkbook> {
  const yieldState = createParseYieldState();
  let cfb: Awaited<ReturnType<typeof parseCfb>>;
  try {
    cfb = await parseCfb(await file.arrayBuffer(), {
      yieldIfNeeded: () => yieldToBrowserIfNeeded(yieldState),
    });
  } catch (error) {
    if (error instanceof CfbParseError) throw mapCfbError(error);
    throw error;
  }

  const workbookStream = cfb.getStream('Workbook', 'Book');
  if (!workbookStream) {
    throw new XlsParseError(
      'MISSING_WORKBOOK_STREAM',
      'XLS 文件缺少 Workbook 数据流',
    );
  }
  const globals = await parseBiff8Globals(workbookStream, yieldState);
  const warnings = [...globals.warnings];
  globals.hasVba = hasVbaStorage(cfb.entries);
  if (globals.hasVba) {
    warnings.push({
      code: 'VBA_DETECTED',
      message: '检测到 VBA 工程；预览器不会读取或执行宏代码',
    });
  }

  const descriptorsByOffset = [...globals.sheets].sort(
    (left, right) => left.streamOffset - right.streamOffset,
  );
  const parsedById = new Map<string, Biff8Worksheet>();
  const chartSheets: Biff8Workbook['chartSheets'] = [];
  for (let index = 0; index < descriptorsByOffset.length; index += 1) {
    const descriptor = descriptorsByOffset[index];
    const endOffset =
      descriptorsByOffset[index + 1]?.streamOffset ?? workbookStream.length;
    if (descriptor.type === 'worksheet') {
      const sheet = await parseBiff8Worksheet(
        workbookStream,
        descriptor,
        globals,
        endOffset,
        yieldState,
      );
      parsedById.set(descriptor.id, sheet);
      warnings.push(...sheet.warnings);
    } else if (descriptor.type === 'chart') {
      chartSheets.push({
        descriptor,
        substream: readBiff8ChartSubstream(
          workbookStream,
          descriptor,
          endOffset,
        ),
      });
    }
    await yieldToBrowserIfNeeded(yieldState);
  }

  const worksheets = globals.sheets
    .map((descriptor) => parsedById.get(descriptor.id))
    .filter((sheet): sheet is Biff8Worksheet => Boolean(sheet));
  const workbook: Biff8Workbook = {
    globals,
    worksheets,
    chartSheets,
    warnings,
  };
  const target = adaptBiff8Workbook(workbook);
  try {
    await attachBiff8DrawingImages(target, workbook);
    attachBiff8Charts(target, workbook);
    return target;
  } catch (error) {
    disposeSpreadsheetWorkbook(target);
    throw error;
  }
}
