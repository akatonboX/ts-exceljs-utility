import ExcelJS from "exceljs";
import { CellAddressRange } from "./core";
export declare const DEFAULT_STANDAED_FONT_WIDTH = 8;
/**
 * px単位をEMU単位に変換する
 */
export declare function px2Emu(pixcel: number): number;
/**
 * セルの高さをpxに変換する
 * @param point
 */
export declare function height2Px(worksheet: ExcelJS.Worksheet, height: number): number;
/**
 * セルの幅をpxに変換する
 * @param width
 * @param standardFontWidth
 * @returns
 */
export declare function width2Px(worksheet: ExcelJS.Worksheet, width?: number, standardFontWidth?: number): number;
/**
 * セルまたはセル範囲を指定して、その高さと幅をpixcelで得る。
 * ※高さについて、元の単位はpoint(1/72インチ)であり、96dpiとして変換します。
 * ※幅について、https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html の「<width> (Column Width)」の記述に従い、pixcelに変換します。
 */
export declare function getCellPixcelSize(worksheet: ExcelJS.Worksheet, range: CellAddressRange, standardFontWidth?: number): {
    width: number;
    height: number;
    widths: number[];
    heights: number[];
};
