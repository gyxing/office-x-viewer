import { useCallback, useState } from 'react';
import type { DocxDocument } from '../services/docx/types';
import type { PptxDocument } from '../services/pptx/types';
import type { XlsxWorkbook } from '../services/xlsx/types';
import { OfficeViewer } from '../components/office-viewer';
import { detectPreviewKind, parseOfficeFile, type PreviewKind } from '../services/officePreview';

export default function HomePage() {
  const [fileName, setFileName] = useState('未加载文件');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [previewKind, setPreviewKind] = useState<PreviewKind>('pptx');
  const [pptxDocument, setPptxDocument] = useState<PptxDocument>();
  const [xlsxWorkbook, setXlsxWorkbook] = useState<XlsxWorkbook>();
  const [docxDocument, setDocxDocument] = useState<DocxDocument>();
  const [activeIndex, setActiveIndex] = useState(0);
  const [activeSheetId, setActiveSheetId] = useState<string>();
  const [zoom, setZoom] = useState(100);

  const handleUpload = useCallback(async (file: File) => {
    setLoading(true);
    setError(undefined);

    try {
      const fileKind = detectPreviewKind(file.name);
      setPreviewKind(fileKind);
      setFileName(file.name);
      setActiveIndex(0);
      setZoom(100);

      const parsed = await parseOfficeFile(file);
      setPptxDocument(parsed.kind === 'pptx' ? parsed.document : undefined);
      setXlsxWorkbook(parsed.kind === 'xlsx' ? parsed.workbook : undefined);
      setDocxDocument(parsed.kind === 'docx' ? parsed.document : undefined);
      setActiveSheetId(parsed.kind === 'xlsx' ? parsed.workbook.sheets[0]?.id : undefined);
    } catch (nextError) {
      setError(nextError instanceof Error ? nextError.message : '文件解析失败');
    } finally {
      setLoading(false);
    }
  }, []);

  return (
    <OfficeViewer
      fileName={fileName}
      loading={loading}
      error={error}
      previewKind={previewKind}
      pptxDocument={pptxDocument}
      xlsxWorkbook={xlsxWorkbook}
      docxDocument={docxDocument}
      activeIndex={activeIndex}
      activeSheetId={activeSheetId}
      zoom={zoom}
      onSelectSlide={setActiveIndex}
      onSelectSheet={setActiveSheetId}
      onZoomChange={setZoom}
      onUpload={handleUpload}
    />
  );
}

