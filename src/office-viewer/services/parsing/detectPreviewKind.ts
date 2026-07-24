// 文件类型检测属于解析运行时公共能力，避免会话与 preview 入口形成循环依赖。
export type PreviewKind = 'pptx' | 'ppt' | 'xlsx' | 'xls' | 'docx' | 'doc';

export const SUPPORTED_OFFICE_EXTENSIONS = [
  '.pptx',
  '.ppt',
  '.xlsx',
  '.xls',
  '.docx',
  '.doc',
  '.wps',
] as const;

export function isSupportedOfficeFileName(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return SUPPORTED_OFFICE_EXTENSIONS.some((extension) =>
    lower.endsWith(extension),
  );
}

/** 判断当前格式是否复用电子表格预览链路。 */
export function isSpreadsheetPreviewKind(
  kind: PreviewKind,
): kind is 'xlsx' | 'xls' {
  return kind === 'xlsx' || kind === 'xls';
}

/** 判断当前格式是否复用统一演示文稿渲染链路。 */
export function isPresentationPreviewKind(
  kind: PreviewKind,
): kind is 'pptx' | 'ppt' {
  return kind === 'pptx' || kind === 'ppt';
}

/** 根据文件名推断 Office 预览格式。 */
export function detectPreviewKind(fileName: string): PreviewKind {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.pptx')) return 'pptx';
  if (lower.endsWith('.ppt')) return 'ppt';
  if (lower.endsWith('.xlsx')) return 'xlsx';
  if (lower.endsWith('.xls')) return 'xls';
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.doc') || lower.endsWith('.wps')) return 'doc';
  // 保留历史行为：无可识别扩展名时按 PPTX 尝试解析。
  return 'pptx';
}
