import type { Biff8DefinedName, Biff8SheetDescriptor } from '../types';
import { Biff8Reader } from './Biff8Reader';
import { readBiff8UnicodeString } from './strings';

export type FormulaDecodeContext = {
  row: number;
  column: number;
  definedNames: Biff8DefinedName[];
  sheets: Biff8SheetDescriptor[];
};

export type DecodedFormula = {
  formula?: string;
  formulaTokens: string;
  unsupported: boolean;
};

type FunctionInfo = {
  name: string;
  arity: number;
};

// BIFF8 固定参数函数表来自 MS-XLS 的内建函数编号。
const FIXED_FUNCTIONS: Record<number, FunctionInfo> = {
  0: { name: 'COUNT', arity: 1 },
  1: { name: 'IF', arity: 3 },
  2: { name: 'ISNA', arity: 1 },
  3: { name: 'ISERROR', arity: 1 },
  4: { name: 'SUM', arity: 1 },
  5: { name: 'AVERAGE', arity: 1 },
  6: { name: 'MIN', arity: 1 },
  7: { name: 'MAX', arity: 1 },
  8: { name: 'ROW', arity: 1 },
  9: { name: 'COLUMN', arity: 1 },
  10: { name: 'NA', arity: 0 },
  11: { name: 'NPV', arity: 2 },
  12: { name: 'STDEV', arity: 1 },
  13: { name: 'DOLLAR', arity: 2 },
  14: { name: 'FIXED', arity: 2 },
  15: { name: 'SIN', arity: 1 },
  16: { name: 'COS', arity: 1 },
  17: { name: 'TAN', arity: 1 },
  18: { name: 'ATAN', arity: 1 },
  19: { name: 'PI', arity: 0 },
  20: { name: 'SQRT', arity: 1 },
  21: { name: 'EXP', arity: 1 },
  22: { name: 'LN', arity: 1 },
  23: { name: 'LOG10', arity: 1 },
  24: { name: 'ABS', arity: 1 },
  25: { name: 'INT', arity: 1 },
  26: { name: 'SIGN', arity: 1 },
  27: { name: 'ROUND', arity: 2 },
  30: { name: 'REPT', arity: 2 },
  31: { name: 'MID', arity: 3 },
  32: { name: 'LEN', arity: 1 },
  33: { name: 'VALUE', arity: 1 },
  34: { name: 'TRUE', arity: 0 },
  35: { name: 'FALSE', arity: 0 },
  36: { name: 'AND', arity: 1 },
  37: { name: 'OR', arity: 1 },
  38: { name: 'NOT', arity: 1 },
  39: { name: 'MOD', arity: 2 },
  63: { name: 'RAND', arity: 0 },
  65: { name: 'DATE', arity: 3 },
  66: { name: 'TIME', arity: 3 },
  67: { name: 'DAY', arity: 1 },
  68: { name: 'MONTH', arity: 1 },
  69: { name: 'YEAR', arity: 1 },
  70: { name: 'WEEKDAY', arity: 1 },
  71: { name: 'HOUR', arity: 1 },
  72: { name: 'MINUTE', arity: 1 },
  73: { name: 'SECOND', arity: 1 },
  74: { name: 'NOW', arity: 0 },
  75: { name: 'AREAS', arity: 1 },
  76: { name: 'ROWS', arity: 1 },
  77: { name: 'COLUMNS', arity: 1 },
  97: { name: 'ATAN2', arity: 2 },
  98: { name: 'ASIN', arity: 1 },
  99: { name: 'ACOS', arity: 1 },
  109: { name: 'LOG', arity: 2 },
  111: { name: 'CHAR', arity: 1 },
  112: { name: 'LOWER', arity: 1 },
  113: { name: 'UPPER', arity: 1 },
  114: { name: 'PROPER', arity: 1 },
  115: { name: 'LEFT', arity: 2 },
  116: { name: 'RIGHT', arity: 2 },
  117: { name: 'EXACT', arity: 2 },
  118: { name: 'TRIM', arity: 1 },
  119: { name: 'REPLACE', arity: 4 },
  120: { name: 'SUBSTITUTE', arity: 3 },
  121: { name: 'CODE', arity: 1 },
  124: { name: 'FIND', arity: 2 },
  125: { name: 'CELL', arity: 2 },
  126: { name: 'ISERR', arity: 1 },
  127: { name: 'ISTEXT', arity: 1 },
  128: { name: 'ISNUMBER', arity: 1 },
  129: { name: 'ISBLANK', arity: 1 },
  130: { name: 'T', arity: 1 },
  131: { name: 'N', arity: 1 },
  169: { name: 'COUNTA', arity: 1 },
  183: { name: 'PRODUCT', arity: 1 },
  184: { name: 'FACT', arity: 1 },
  189: { name: 'DPRODUCT', arity: 3 },
  190: { name: 'ISNONTEXT', arity: 1 },
  197: { name: 'TRUNC', arity: 2 },
  198: { name: 'ISLOGICAL', arity: 1 },
  212: { name: 'ROUNDUP', arity: 2 },
  213: { name: 'ROUNDDOWN', arity: 2 },
  221: { name: 'TODAY', arity: 0 },
  229: { name: 'SINH', arity: 1 },
  230: { name: 'COSH', arity: 1 },
  231: { name: 'TANH', arity: 1 },
  232: { name: 'ASINH', arity: 1 },
  233: { name: 'ACOSH', arity: 1 },
  234: { name: 'ATANH', arity: 1 },
  279: { name: 'EVEN', arity: 1 },
  280: { name: 'ODD', arity: 1 },
};

const ERROR_VALUES: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

function columnLabel(index: number) {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function signed14(value: number) {
  const normalized = value & 0x3fff;
  return normalized & 0x2000 ? normalized - 0x4000 : normalized;
}

function formatCellReference(
  rawRow: number,
  rawColumn: number,
  context: FormulaDecodeContext,
) {
  const rowRelative = Boolean(rawColumn & 0x8000);
  const columnRelative = Boolean(rawColumn & 0x4000);
  const row = rowRelative
    ? context.row + (rawRow & 0x8000 ? rawRow - 0x10000 : rawRow)
    : rawRow;
  const column = columnRelative
    ? context.column + signed14(rawColumn)
    : rawColumn & 0x3fff;
  const rowPrefix = rowRelative ? '' : '$';
  const columnPrefix = columnRelative ? '' : '$';
  return `${columnPrefix}${columnLabel(Math.max(0, column))}${rowPrefix}${
    Math.max(0, row) + 1
  }`;
}

function readCellReference(reader: Biff8Reader, context: FormulaDecodeContext) {
  return formatCellReference(reader.readUint16(), reader.readUint16(), context);
}

function readAreaReference(reader: Biff8Reader, context: FormulaDecodeContext) {
  const firstRow = reader.readUint16();
  const lastRow = reader.readUint16();
  const firstColumn = reader.readUint16();
  const lastColumn = reader.readUint16();
  return `${formatCellReference(
    firstRow,
    firstColumn,
    context,
  )}:${formatCellReference(lastRow, lastColumn, context)}`;
}

function quoteSheetName(name: string) {
  return `'${name.replace(/'/g, "''")}'`;
}

function popArguments(stack: string[], count: number) {
  if (stack.length < count) return undefined;
  return stack.splice(stack.length - count, count);
}

function applyBinary(stack: string[], operator: string) {
  const arguments_ = popArguments(stack, 2);
  if (!arguments_) return false;
  stack.push(`(${arguments_[0]}${operator}${arguments_[1]})`);
  return true;
}

function applyFunction(
  stack: string[],
  functionId: number,
  argumentCount?: number,
) {
  const info = FIXED_FUNCTIONS[functionId];
  const count = argumentCount ?? info?.arity;
  if (count === undefined) return false;
  const arguments_ = popArguments(stack, count);
  if (!arguments_) return false;
  const name = info?.name ?? `_xlfn.BIFF_${functionId}`;
  stack.push(`${name}(${arguments_.join(',')})`);
  return true;
}

function normalizeToken(token: number) {
  return token >= 0x20 ? (token & 0x1f) | 0x20 : token;
}

function bytesToHex(bytes: Uint8Array) {
  return Array.from(bytes, (value) => value.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/** 将常见 BIFF8 RPN token 反编译为 A1 公式，未知 token 保留原始十六进制。 */
export function decodeBiff8Formula(
  tokens: Uint8Array,
  context: FormulaDecodeContext,
): DecodedFormula {
  const reader = new Biff8Reader(tokens);
  const stack: string[] = [];
  let unsupported = false;

  try {
    while (reader.remaining > 0 && !unsupported) {
      const rawToken = reader.readUint8();
      const token = normalizeToken(rawToken);
      switch (token) {
        case 0x03:
          unsupported = !applyBinary(stack, '+');
          break;
        case 0x04:
          unsupported = !applyBinary(stack, '-');
          break;
        case 0x05:
          unsupported = !applyBinary(stack, '*');
          break;
        case 0x06:
          unsupported = !applyBinary(stack, '/');
          break;
        case 0x07:
          unsupported = !applyBinary(stack, '^');
          break;
        case 0x08:
          unsupported = !applyBinary(stack, '&');
          break;
        case 0x09:
          unsupported = !applyBinary(stack, '<');
          break;
        case 0x0a:
          unsupported = !applyBinary(stack, '<=');
          break;
        case 0x0b:
          unsupported = !applyBinary(stack, '=');
          break;
        case 0x0c:
          unsupported = !applyBinary(stack, '>=');
          break;
        case 0x0d:
          unsupported = !applyBinary(stack, '>');
          break;
        case 0x0e:
          unsupported = !applyBinary(stack, '<>');
          break;
        case 0x0f:
          unsupported = !applyBinary(stack, ' ');
          break;
        case 0x10:
          unsupported = !applyBinary(stack, ',');
          break;
        case 0x11:
          unsupported = !applyBinary(stack, ':');
          break;
        case 0x12:
          if (!stack.length) unsupported = true;
          break;
        case 0x13:
          if (!stack.length) unsupported = true;
          else stack.push(`(-${stack.pop()})`);
          break;
        case 0x14:
          if (!stack.length) unsupported = true;
          else stack.push(`(${stack.pop()}%)`);
          break;
        case 0x15:
          if (!stack.length) unsupported = true;
          else stack.push(`(${stack.pop()})`);
          break;
        case 0x16:
          stack.push('');
          break;
        case 0x17:
          stack.push(
            `"${readBiff8UnicodeString(reader, 1).value.replace(/"/g, '""')}"`,
          );
          break;
        case 0x19: {
          const flags = reader.readUint8();
          const data = reader.readUint16();
          if (flags & 0x04) reader.readBytes((data + 1) * 2);
          if (flags & 0x10) unsupported = !applyFunction(stack, 4, 1);
          break;
        }
        case 0x1c:
          stack.push(ERROR_VALUES[reader.readUint8()] ?? '#ERROR!');
          break;
        case 0x1d:
          stack.push(reader.readUint8() ? 'TRUE' : 'FALSE');
          break;
        case 0x1e:
          stack.push(String(reader.readUint16()));
          break;
        case 0x1f:
          stack.push(String(reader.readFloat64()));
          break;
        case 0x21:
          unsupported = !applyFunction(stack, reader.readUint16());
          break;
        case 0x22: {
          const argumentCount = reader.readUint8() & 0x7f;
          unsupported = !applyFunction(
            stack,
            reader.readUint16(),
            argumentCount,
          );
          break;
        }
        case 0x23: {
          const nameId = reader.readUint32();
          const name = context.definedNames[nameId - 1]?.name;
          stack.push(name ?? `_Name${nameId}`);
          break;
        }
        case 0x24:
        case 0x2c:
          stack.push(readCellReference(reader, context));
          break;
        case 0x25:
        case 0x2d:
          stack.push(readAreaReference(reader, context));
          break;
        case 0x26:
        case 0x27:
        case 0x28:
          reader.readBytes(6);
          break;
        case 0x29:
          reader.readBytes(2);
          break;
        case 0x2a:
          reader.readBytes(4);
          stack.push('#REF!');
          break;
        case 0x2b:
          reader.readBytes(8);
          stack.push('#REF!');
          break;
        case 0x39:
          reader.readBytes(2);
          stack.push(`_ExternalName${reader.readUint32()}`);
          break;
        case 0x3a: {
          const sheetIndex = reader.readUint16();
          const reference = readCellReference(reader, context);
          const sheet =
            context.sheets[sheetIndex]?.name ?? `Sheet${sheetIndex + 1}`;
          stack.push(`${quoteSheetName(sheet)}!${reference}`);
          break;
        }
        case 0x3b: {
          const sheetIndex = reader.readUint16();
          const reference = readAreaReference(reader, context);
          const sheet =
            context.sheets[sheetIndex]?.name ?? `Sheet${sheetIndex + 1}`;
          stack.push(`${quoteSheetName(sheet)}!${reference}`);
          break;
        }
        case 0x3c:
          reader.readBytes(6);
          stack.push('#REF!');
          break;
        case 0x3d:
          reader.readBytes(10);
          stack.push('#REF!');
          break;
        case 0x01:
        case 0x02:
          reader.readBytes(4);
          unsupported = true;
          break;
        case 0x20:
          reader.readBytes(7);
          unsupported = true;
          break;
        default:
          unsupported = true;
      }
    }
  } catch {
    unsupported = true;
  }

  return {
    formula: !unsupported && stack.length === 1 ? `=${stack[0]}` : undefined,
    formulaTokens: bytesToHex(tokens),
    unsupported,
  };
}
