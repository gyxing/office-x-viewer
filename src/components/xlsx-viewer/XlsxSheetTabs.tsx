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

  return (
    <div
      style={{
        flex: '0 0 auto',
        background: '#fff',
        borderBottom: '1px solid #dde3ec',
        padding: '0 16px',
        boxShadow: '0 1px 0 rgba(15, 23, 42, 0.04)',
      }}
    >
      <Tabs
        activeKey={activeSheet.id}
        onChange={onSelectSheet}
        items={tabItems}
        tabBarExtraContent={
          <Typography.Text type="secondary" style={{ fontSize: 12 }}>
            {activeSheet.range ?? `${activeSheet.rowCount} 行 x ${activeSheet.columnCount} 列`}
          </Typography.Text>
        }
      />
    </div>
  );
}

export const XlsxSheetTabs = memo(XlsxSheetTabsComponent);

