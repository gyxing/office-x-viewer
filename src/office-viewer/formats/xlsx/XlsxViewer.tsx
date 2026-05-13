// XlsxViewer 负责 XLSX 工作簿预览的整体布局，包括工作表选择和当前工作表内容区。
import { memo, useMemo } from 'react';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { OfficeEmpty } from '../../shell/Empty';
import './index.less';
import { XlsxSheetGrid } from './XlsxSheetGrid';
import { XlsxSheetTabs } from './XlsxSheetTabs';

type XlsxViewerProps = {
  workbook?: XlsxWorkbook;
  activeSheetId?: string;
  zoom: number;
  onSelectSheet: (sheetId: string) => void;
};

function XlsxViewerComponent({ workbook, activeSheetId, zoom, onSelectSheet }: XlsxViewerProps) {
  const activeSheet = useMemo(
    () => workbook?.sheets.find((sheet) => sheet.id === activeSheetId) ?? workbook?.sheets[0],
    [activeSheetId, workbook],
  );

  if (!activeSheet) {
    return <OfficeEmpty kind="xlsx" />;
  }

  return (
    <div className="oxv-xlsx-viewer">
      <XlsxSheetTabs workbook={workbook} activeSheet={activeSheet} onSelectSheet={onSelectSheet} />
      <XlsxSheetGrid sheet={activeSheet} zoom={zoom} />
    </div>
  );
}

export const XlsxViewer = memo(XlsxViewerComponent);
