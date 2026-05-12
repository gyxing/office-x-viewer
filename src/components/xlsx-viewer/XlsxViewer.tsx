import { memo, useMemo } from 'react';
import type { XlsxWorkbook } from '../../services/xlsx/types';
import { OfficeEmpty } from '../office-viewer/OfficeEmpty';
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
    <div
      style={{
        minHeight: 'calc(100vh - 56px)',
        height: 'calc(100vh - 56px)',
        display: 'flex',
        flexDirection: 'column',
        background: '#f2f5f9',
        overflow: 'hidden',
      }}
    >
      <XlsxSheetTabs workbook={workbook} activeSheet={activeSheet} onSelectSheet={onSelectSheet} />
      <XlsxSheetGrid sheet={activeSheet} zoom={zoom} />
    </div>
  );
}

export const XlsxViewer = memo(XlsxViewerComponent);

