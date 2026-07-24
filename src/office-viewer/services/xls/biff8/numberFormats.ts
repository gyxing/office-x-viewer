export const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

function pad(value: number, length = 2) {
  return String(value).padStart(length, '0');
}

function stripFormatLiterals(format: string) {
  return format
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[[^\]]+\]/g, '')
    .replace(/_.|\\*./g, '');
}

/** 判断格式是否包含 Excel 日期或时间标记。 */
export function isDateFormat(format: string) {
  const normalized = stripFormatLiterals(format).toLowerCase();
  return /(^|[^a-z])[ymdhis]+([^a-z]|$)/.test(normalized);
}

function excelSerialToDate(serial: number, date1904: boolean) {
  const wholeDays = Math.floor(serial);
  const fraction = serial - wholeDays;
  const adjustedDays = date1904
    ? wholeDays
    : wholeDays >= 60
    ? wholeDays - 1
    : wholeDays;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
  return new Date(epoch + (adjustedDays + fraction) * 86400000);
}

function formatDate(serial: number, format: string, date1904: boolean) {
  if (/\[h\]/i.test(format)) {
    const totalSeconds = Math.round(serial * 86400);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    return `${hours}:${pad(minutes)}:${pad(totalSeconds % 60)}`;
  }

  const date = excelSerialToDate(serial, date1904);
  const hasDate = /[yd]/i.test(stripFormatLiterals(format));
  const hasTime = /[hs]/i.test(stripFormatLiterals(format));
  const usesAmPm = /AM\/PM/i.test(format);
  let result = '';
  if (hasDate) {
    if (/d-mmm-yy/i.test(format)) {
      const month = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec',
      ][date.getUTCMonth()];
      result = `${date.getUTCDate()}-${month}-${pad(
        date.getUTCFullYear() % 100,
      )}`;
    } else {
      result = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(
        date.getUTCDate(),
      )}`;
    }
  }
  if (hasTime) {
    const rawHours = date.getUTCHours();
    const hours = usesAmPm ? rawHours % 12 || 12 : rawHours;
    const showSeconds = /s/i.test(stripFormatLiterals(format));
    const time = `${pad(hours)}:${pad(date.getUTCMinutes())}${
      showSeconds ? `:${pad(date.getUTCSeconds())}` : ''
    }${usesAmPm ? (rawHours >= 12 ? ' PM' : ' AM') : ''}`;
    result = result ? `${result} ${time}` : time;
  }
  return result || String(serial);
}

function formatFraction(value: number, format: string) {
  const denominatorLimit = format.includes('??') ? 99 : 9;
  const whole = Math.trunc(value);
  const fraction = Math.abs(value - whole);
  let bestNumerator = 0;
  let bestDenominator = 1;
  let bestError = Number.POSITIVE_INFINITY;
  for (let denominator = 1; denominator <= denominatorLimit; denominator += 1) {
    const numerator = Math.round(fraction * denominator);
    const error = Math.abs(fraction - numerator / denominator);
    if (error < bestError) {
      bestError = error;
      bestNumerator = numerator;
      bestDenominator = denominator;
    }
  }
  if (!bestNumerator) return String(whole);
  return `${whole || ''}${whole ? ' ' : ''}${bestNumerator}/${bestDenominator}`;
}

/** 按常用 BIFF8 数字格式生成展示文本，不执行任何公式计算。 */
export function formatBiff8Value(
  value: string | number | boolean | null,
  format: string | undefined,
  date1904: boolean,
) {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string' || !format || format === 'General') {
    return String(value);
  }
  if (isDateFormat(format)) return formatDate(value, format, date1904);
  if (format.includes('%')) {
    const decimals = /\.([0]+)/.exec(format)?.[1].length ?? 0;
    return `${(value * 100).toFixed(decimals)}%`;
  }
  if (format.includes('?/?') || format.includes('??/??')) {
    return formatFraction(value, format);
  }
  if (/E[+-]0+/i.test(format)) {
    const decimals = /\.([0#]+)/.exec(format)?.[1].length ?? 0;
    return value.toExponential(decimals).replace('e', 'E');
  }

  const decimals = /\.([0#]+)/.exec(format)?.[1].length ?? 0;
  const useGrouping = format.includes(',');
  const absolute = Math.abs(value);
  const formatted = absolute.toLocaleString('en-US', {
    useGrouping,
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
  const currency = /[$£¥€]/.exec(format)?.[0] ?? '';
  if (value < 0 && format.includes('(')) return `(${currency}${formatted})`;
  return `${value < 0 ? '-' : ''}${currency}${formatted}`;
}
