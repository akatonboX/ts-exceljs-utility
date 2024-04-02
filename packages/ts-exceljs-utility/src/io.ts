import ExcelJS from "exceljs";
import { writeWithColBreaks } from "./colBreaks";
/**
 * URLからxlsxをダウンロードしてworkbookに読み込みます。
 * @param url 
 * @returns 
 */
export async function loadWorkbook(url: string): Promise<ExcelJS.Workbook | undefined>{
  //■指定されたURLから、ダウンロード
  const response = await fetch(url);
  if (!response.ok) {
    return undefined;
  }
  //■workbookに読み来い
  const workbook = new ExcelJS.Workbook();
  try{
  await workbook.xlsx.load(await response.arrayBuffer());
  }
  catch(e){
    console.warn("Workbook.xlsx.load method failed. ", e);
    return undefined;
  }

  return workbook;
}

/**
 * ExcelJSのworkbookをダウンロードする。※２回目以降の実行では、chromeにおいて、「このサイトで複数ファイルの自動ダウンロードが試行されました」という警告と許可を求めるダイアログが表示される。
 * @param workbook ダウンロードさせるworkbook
 * @param fileName ダウンロードするファイル名
 */
export async function downloadWorkbook(workbook: ExcelJS.Workbook, fileName: string){
  const blob = await writeWithColBreaks(workbook);
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
