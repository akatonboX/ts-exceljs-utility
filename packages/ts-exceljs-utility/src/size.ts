import ExcelJS from "exceljs";
import { CellAddressRange } from "./core";

export const DEFAULT_STANDAED_FONT_WIDTH = 8;
/**
 * px単位をEMU単位に変換する
 */
export function px2Emu(pixcel: number): number{
  return Math.round(pixcel * 9525);
}

/**
 * セルの高さをpxに変換する
 * @param point 
 */
export function height2Px(worksheet: ExcelJS.Worksheet, height: number): number{
  return (height * 96) / 72;
}
/**
 * セルの幅をpxに変換する
 * @param width 
 * @param standardFontWidth 
 * @returns 
 */
export function width2Px(worksheet: ExcelJS.Worksheet, width?: number, standardFontWidth: number = DEFAULT_STANDAED_FONT_WIDTH): number{
  const w = width ?? worksheet.properties.defaultColWidth;
  if(w == null)return NaN;
  return Math.floor((((256 * w) + Math.floor(128 / standardFontWidth)) / 256) * standardFontWidth);
}


/**
 * セルまたはセル範囲を指定して、その高さと幅をpixcelで得る。
 * ※高さについて、元の単位はpoint(1/72インチ)であり、96dpiとして変換します。
 * ※幅について、https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html の「<width> (Column Width)」の記述に従い、pixcelに変換します。
 */
export function getCellPixcelSize(worksheet: ExcelJS.Worksheet, range: CellAddressRange, standardFontWidth: number = DEFAULT_STANDAED_FONT_WIDTH): {width: number, height: number, widths: number[], heights: number[]}{

  //■対象の範囲に含まれるrowとcolumnの配列を得る
  const rows = new Array(range.bottom - range.top + 1).fill(1).map((item, index) => worksheet.getRow(range.top + index));
  const columns = new Array(range.right - range.left + 1).fill(1).map((item, index) => worksheet.getColumn(range.left + index));
 
  //■pxに変換
  const heights = rows.map(item => height2Px(worksheet, item.height));
  const widths = columns.map(item => width2Px(worksheet, item.width, standardFontWidth));

  //■合計
  const height = heights.reduce((previouseValue, currentValue) => previouseValue + currentValue, 0);
  const width = widths.reduce((previouseValue, currentValue) => previouseValue + currentValue, 0);

  return {width: width, height: height, widths: widths, heights: heights};
}
