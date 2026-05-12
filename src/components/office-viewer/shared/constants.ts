// OfficeViewer 的工具栏尺寸、缩放范围和空状态文案等共享常量。
export const OFFICE_VIEWER_HEADER_HEIGHT = 56;

export const OFFICE_ZOOM_LEVELS = [50, 75, 100, 125, 150, 200];
export const OFFICE_ZOOM_STEP = 25;
export const OFFICE_MIN_ZOOM = 25;
export const OFFICE_MAX_ZOOM = 300;
export const OFFICE_DEFAULT_ZOOM = 100;

export const OFFICE_EMPTY_DESCRIPTIONS = {
  pptx: '请先上传 PPTX 文件开始预览',
  xlsx: '请先上传 XLSX 文件开始预览',
  docx: '请先上传 DOCX 文件开始预览',
  doc: '请先上传 DOC 文件开始预览',
} as const;
