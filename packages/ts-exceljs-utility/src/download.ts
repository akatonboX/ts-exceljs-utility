import ExcelJS from "exceljs";
/**
 * ExcelJSのworkbookをダウンロードする。※２回目以降の実行では、chromeにおいて、「このサイトで複数ファイルの自動ダウンロードが試行されました」という警告と許可を求めるダイアログが表示される。
 * @param workbook ダウンロードさせるworkbook
 * @param fileName ダウンロードするファイル名
 */
export async function downloadWorkbook(workbook: ExcelJS.Workbook, fileName: string){
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = window.URL.createObjectURL(blob);
  try{
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
  }
  finally{
    window.URL.revokeObjectURL(url);
  }
}