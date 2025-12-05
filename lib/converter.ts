// Modern Excel to Text converter

type CellValue = string | number | Date | boolean | null | undefined;

export function convertExcelToText(data: CellValue[][]): string {
  // Skip header row and filter empty rows
  return data
    .slice(1) // Skip first row (header)
    .filter(row => row.some(cell => cell !== null && cell !== undefined && cell !== ''))
    .map(row => processRow(row))
    .join('\n') + '\n'; // Add newline after each row, including the last one
}

function processRow(row: CellValue[]): string {
  return row
    .map((cell, index) => formatCell(cell, index))
    .join(';');
}

function formatCell(value: CellValue, columnIndex: number): string {
  if (value === null || value === undefined) return '';

  switch (columnIndex) {
    case 0: // Column 1: Integer (no decimals)
      return formatInteger(value);
    case 1: // Column 2: Decimal with 2 places
      return formatDecimal(value, 2);
    case 5: // Column 6: Date
      return formatDate(value);
    case 6: // Column 7: Invoice 13 digits
      return formatInvoice(value, 13);
    case 8: // Column 9: Decimal with 2 places
      return formatDecimal(value, 2);
    case 9: // Column 10: Date
      return formatDate(value);
    case 10: // Column 11: Invoice 12 digits
      return formatInvoice(value, 12);
    default:
      return String(value);
  }
}

function formatInteger(value: CellValue): string {
  if (value === null || value === undefined || value === '') return '';

  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(',', '.'));

  return isNaN(num) ? String(value) : Math.round(num).toString();
}

function formatDecimal(value: CellValue, decimals: number): string {
  if (value === null || value === undefined || value === '') return '';

  const num = typeof value === 'number' ? value : parseFloat(String(value).replace(',', '.'));

  return isNaN(num) ? String(value) : num.toFixed(decimals);
}

function formatDate(value: CellValue): string {
  if (!value) return '';

  let date: Date;

  if (value instanceof Date) {
    date = value;
  } else if (typeof value === 'number') {
    // Excel serial date
    date = excelSerialToDate(value);
  } else {
    const str = String(value).replace(/-/g, '/');
    date = new Date(str);
  }

  if (isNaN(date.getTime())) return String(value);

  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();

  return `${day}/${month}/${year}`;
}

function formatInvoice(value: CellValue, length: number): string {
  if (!value) return '';

  const num = typeof value === 'number' ? value : parseInt(String(value));

  return isNaN(num) ? String(value) : String(num).padStart(length, '0');
}

function excelSerialToDate(serial: number): Date {
  // Excel stores dates as days since 1900-01-01 (with bugs for dates before 1900-03-01)
  const epoch = new Date(1899, 11, 30);
  return new Date(epoch.getTime() + serial * 86400000);
}
