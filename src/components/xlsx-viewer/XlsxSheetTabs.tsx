// XlsxSheetTabs 渲染工作表标签栏，并展示当前工作表范围或行列数量。
import { Tabs, Typography } from 'antd';
import { memo, useMemo } from 'react';
import type { XlsxSheet, XlsxWorkbook } from '../../services/xlsx/types';

type XlsxSheetTabsProps = {
  workbook?: XlsxWorkbook;
  activeSheet: XlsxSheet;
  onSelectSheet: (sheetId: string) => void;
};

const EMPTY_TABS: Array<{ key: string; label: string }> = [];

function XlsxSheetTabsComponent({ workbook, activeSheet, onSelectSheet }: XlsxSheetTabsProps) {
  const tabItems = useMemo(
    () =>
      workbook?.sheets.map((sheet) => ({
        key: sheet.id,
        label: sheet.name,
      })) ?? EMPTY_TABS,
    [workbook],
  );
  const rangeText = activeSheet.range ?? `${activeSheet.rowCount} 行 x ${activeSheet.columnCount} 列`;

  return (
    <div className="oxv-xlsx-sheet-tabs">
      <Tabs
        activeKey={activeSheet.id}
        onChange={onSelectSheet}
        items={tabItems}
        tabBarExtraContent={
          <Typography.Text type="secondary" className="oxv-xlsx-sheet-tabs__range">
            {rangeText}
          </Typography.Text>
        }
      />
    </div>
  );
}

export const XlsxSheetTabs = memo(XlsxSheetTabsComponent);
