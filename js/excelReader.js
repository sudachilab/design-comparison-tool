/** XLSX is loaded from the SheetJS browser bundle. */
export async function readExcel(file) {
  if (!globalThis.XLSX) {
    throw new Error(
      'Excel解析ライブラリを読み込めません。インターネット接続を確認してください。',
    );
  }

  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, {
    type: 'array',
    cellText: true,
    cellDates: true,
    cellStyles: true,
  });
}

export function getSheet(workbook, name) {
  return workbook.Sheets[name];
}
