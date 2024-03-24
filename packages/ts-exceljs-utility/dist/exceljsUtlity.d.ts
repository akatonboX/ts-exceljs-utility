import ExcelJS from "exceljs";
export interface CellRange {
    rangeStr: string;
    top: number;
    left: number;
    bottom: number;
    right: number;
}
/** "A1"や"A1:B2"のような文字列のセルおよびのセル範囲の表現から、Rangeを返します。 */
export declare function getCellRange(rangeStr: string): CellRange;
export declare function getCellRange(top: number, left: number): CellRange;
export declare function getCellRange(top: number, left: number, bottom: number, right: number): CellRange;
export declare namespace xcelJsUtility {
    /** "A1:B2"のような文字表現のセル範囲から、Rangeを返します。 */
    function getRange(range: string): ExcelJS.Range;
}
