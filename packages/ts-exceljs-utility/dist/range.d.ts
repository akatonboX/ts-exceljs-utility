import ExcelJS from "exceljs";
import { CellAddress, CellAddressRange } from "./core";
export interface CellRange extends CellAddressRange {
    rangeStr: string;
}
/** "A1"や"A1:B2"のような文字列のセルおよびのセル範囲の表現から、Rangeを返します。 */
export declare function getCellRange(address: CellAddress): CellRange;
export declare namespace xcelJsUtility {
    /** "A1:B2"のような文字表現のセル範囲から、Rangeを返します。 */
    function getRange(range: string): ExcelJS.Range;
}
