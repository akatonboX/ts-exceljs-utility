import ExcelJS from "exceljs";
/**
 * ExcelJSのworkbookをダウンロードする。※２回目以降の実行では、chromeにおいて、「このサイトで複数ファイルの自動ダウンロードが試行されました」という警告と許可を求めるダイアログが表示される。
 * @param workbook ダウンロードさせるworkbook
 * @param fileName ダウンロードするファイル名
 */
export declare function downloadWorkbook(workbook: ExcelJS.Workbook, fileName: string): Promise<void>;
