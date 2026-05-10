import {
  Alert,
  Button,
  Layout,
  Select,
  Space,
  Spin,
  Tooltip,
  Typography,
  Upload,
} from 'antd';
import {
  FileExcelOutlined,
  FilePptOutlined,
  FileWordOutlined,
  FullscreenOutlined,
  LeftOutlined,
  RightOutlined,
  ZoomInOutlined,
  ZoomOutOutlined,
} from '@ant-design/icons';
import { useCallback, useMemo, useState } from 'react';
import type { DocxDocument } from '../services/docx/types';
import type { PptxDocument } from '../services/pptx/types';
import type { XlsxWorkbook } from '../services/xlsx/types';
import { DocxViewer } from '../components/docx-viewer/DocxViewer';
import { PptxViewer } from '../components/pptx-viewer/PptxViewer';
import { XlsxViewer } from '../components/xlsx-viewer/XlsxViewer';
import { detectPreviewKind, parseOfficeFile, type PreviewKind } from '../services/officePreview';

const { Header, Content } = Layout;

function fileIcon(kind: PreviewKind) {
  if (kind === 'xlsx') return <FileExcelOutlined />;
  if (kind === 'docx') return <FileWordOutlined />;
  return <FilePptOutlined />;
}

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

  const hasDocument = useMemo(
    () =>
      previewKind === 'pptx'
        ? Boolean(pptxDocument)
        : previewKind === 'xlsx'
          ? Boolean(xlsxWorkbook)
          : Boolean(docxDocument),
    [docxDocument, pptxDocument, previewKind, xlsxWorkbook],
  );

  const uploadIcon = useMemo(() => fileIcon(previewKind), [previewKind]);

  return (
    <Layout style={{ minHeight: '100vh', background: '#eef1f6' }}>
      <Header
        style={{
          height: 56,
          background: '#fff',
          borderBottom: '1px solid #dde3ec',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 16,
          padding: '0 16px',
          position: 'sticky',
          top: 0,
          zIndex: 20,
        }}
      >
        <Typography.Text strong ellipsis style={{ maxWidth: 360 }}>
          {fileName}
        </Typography.Text>
        <Space size={8} wrap>
          <Upload
            accept=".pptx,.xlsx,.docx"
            showUploadList={false}
            beforeUpload={(file) => {
              void handleUpload(file);
              return false;
            }}
          >
            <Button icon={uploadIcon}>上传文件</Button>
          </Upload>
          <Tooltip title="上一页">
            <Button
              icon={<LeftOutlined />}
              disabled={previewKind !== 'pptx' || !pptxDocument || activeIndex === 0}
              onClick={() => setActiveIndex((value) => Math.max(value - 1, 0))}
            />
          </Tooltip>
          <Tooltip title="下一页">
            <Button
              icon={<RightOutlined />}
              disabled={
                previewKind !== 'pptx' ||
                !pptxDocument ||
                activeIndex >= (pptxDocument?.slides.length ?? 1) - 1
              }
              onClick={() =>
                setActiveIndex((value) => Math.min(value + 1, (pptxDocument?.slides.length ?? 1) - 1))
              }
            />
          </Tooltip>
          <Select
            value={zoom}
            style={{ width: 104 }}
            onChange={setZoom}
            options={[
              { value: 50, label: '50%' },
              { value: 75, label: '75%' },
              { value: 100, label: '100%' },
              { value: 125, label: '125%' },
              { value: 150, label: '150%' },
              { value: 200, label: '200%' },
            ]}
          />
          <Tooltip title="缩小">
            <Button
              icon={<ZoomOutOutlined />}
              disabled={!hasDocument}
              onClick={() => setZoom((value) => Math.max(25, value - 25))}
            />
          </Tooltip>
          <Tooltip title="放大">
            <Button
              icon={<ZoomInOutlined />}
              disabled={!hasDocument}
              onClick={() => setZoom((value) => Math.min(300, value + 25))}
            />
          </Tooltip>
          <Button disabled={!hasDocument} onClick={() => setZoom(100)}>
            100%
          </Button>
          <Button icon={<FullscreenOutlined />} disabled={!hasDocument}>
            全屏
          </Button>
        </Space>
      </Header>
      <Content style={{ background: '#eef1f6', height: 'calc(100vh - 56px)', overflow: 'hidden' }}>
        {error ? (
          <div style={{ padding: 24 }}>
            <Alert type="error" showIcon message="预览失败" description={error} />
          </div>
        ) : loading ? (
          <div style={{ height: 'calc(100vh - 56px)', display: 'grid', placeItems: 'center' }}>
            <Spin size="large" tip="正在解析文件" />
          </div>
        ) : previewKind === 'xlsx' ? (
          <XlsxViewer
            workbook={xlsxWorkbook}
            activeSheetId={activeSheetId}
            zoom={zoom}
            onSelectSheet={setActiveSheetId}
          />
        ) : previewKind === 'docx' ? (
          <DocxViewer document={docxDocument} zoom={zoom} />
        ) : (
          <PptxViewer
            document={pptxDocument}
            activeIndex={activeIndex}
            zoom={zoom}
            onSelectSlide={setActiveIndex}
          />
        )}
      </Content>
    </Layout>
  );
}
