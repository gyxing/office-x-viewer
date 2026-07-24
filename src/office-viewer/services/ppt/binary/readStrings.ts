import type { PptParseContext } from '../types';

/** 解码 PowerPoint UTF-16LE 文本，并移除末尾的空字符。 */
export function readPptUnicodeString(bytes: Uint8Array) {
  const evenLength = bytes.length - (bytes.length % 2);
  return new TextDecoder('utf-16le')
    .decode(bytes.subarray(0, evenLength))
    .replace(/\u0000+$/g, '');
}

/** 解码旧式单字节文本，浏览器不认识代码页时安全回退到 Windows-1252。 */
export function readPptByteString(
  bytes: Uint8Array,
  codePage: number | undefined,
  context: PptParseContext,
) {
  const label = codePage ? `windows-${codePage}` : 'windows-1252';
  try {
    return new TextDecoder(label).decode(bytes).replace(/\u0000+$/g, '');
  } catch {
    context.warnings.push({
      code: 'PPT_CODEPAGE_FALLBACK',
      message: `浏览器无法识别代码页 ${codePage ?? 1252}，已回退到 Windows-1252`,
    });
    return new TextDecoder('windows-1252')
      .decode(bytes)
      .replace(/\u0000+$/g, '');
  }
}
