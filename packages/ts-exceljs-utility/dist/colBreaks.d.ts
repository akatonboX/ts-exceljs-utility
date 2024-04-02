import ExcelJS from "exceljs";
/**
 * シートに設定されているcolBreaksを取得する。
 * @param worksheet
 * @returns
 */
export declare function getColBreaks(worksheet: ExcelJS.Worksheet): number[];
/**
 * シートにcolBreaksを設定する。
 */
export declare function setColBreaks(worksheet: ExcelJS.Worksheet, colBreaks: number[]): void;
/**
 * シートにcolBreakを追加する。
 * @param worksheet
 * @param colIndex 列を表す0から始まるindex。
 */
export declare function addColBreak(worksheet: ExcelJS.Worksheet, colIndex: number): void;
export declare function writeWithColBreaks(workbook: ExcelJS.Workbook): Promise<Blob>;
