import type { Biff8Record } from '../biff8/Biff8Reader';
import { BIFF8_RECORD } from '../biff8/constants';
import { XlsParseError } from '../errors';
import type { Biff8ChartRecordNode } from './types';

/** 将 BIFF Begin/End 作用域转换为便于局部容错的记录树。 */
export function buildChartRecordTree(
  records: Biff8Record[],
): Biff8ChartRecordNode[] {
  const roots: Biff8ChartRecordNode[] = [];
  const collections: Biff8ChartRecordNode[][] = [roots];

  for (const record of records) {
    if (record.id === BIFF8_RECORD.BOF || record.id === BIFF8_RECORD.EOF) {
      continue;
    }
    if (record.id === BIFF8_RECORD.BEGIN) {
      const siblings = collections[collections.length - 1];
      const owner = siblings[siblings.length - 1];
      if (!owner) {
        throw new XlsParseError(
          'INVALID_RECORD_DATA',
          '图表 Begin 记录缺少所属对象',
          { offset: record.offset, recordId: record.id },
        );
      }
      collections.push(owner.children);
      continue;
    }
    if (record.id === BIFF8_RECORD.END) {
      if (collections.length === 1) {
        throw new XlsParseError(
          'INVALID_RECORD_DATA',
          '图表 End 记录没有匹配的 Begin',
          { offset: record.offset, recordId: record.id },
        );
      }
      collections.pop();
      continue;
    }
    collections[collections.length - 1].push({
      id: record.id,
      offset: record.offset,
      data: record.data,
      children: [],
    });
  }

  if (collections.length !== 1) {
    throw new XlsParseError(
      'INVALID_RECORD_DATA',
      '图表记录存在未闭合的 Begin 作用域',
    );
  }
  return roots;
}

/** 深度优先收集指定类型的图表节点。 */
export function collectChartNodes(
  nodes: Biff8ChartRecordNode[],
  id: number,
  result: Biff8ChartRecordNode[] = [],
) {
  for (const node of nodes) {
    if (node.id === id) result.push(node);
    collectChartNodes(node.children, id, result);
  }
  return result;
}
