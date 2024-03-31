import ExcelJS from "exceljs";
import { CellAddress, CellAddressRange, isCellAddressOne, isCellAddressRange, isCellAddressString } from "./core";


export interface CellRange extends CellAddressRange{
  rangeStr: string;
}

/** "A1"や"A1:B2"のような文字列のセルおよびのセル範囲の表現から、Rangeを返します。 */
// export function getCellRange(rangeStr: string): CellRange;
// export function getCellRange(top: number, left: number): CellRange;
// export function getCellRange(top: number, left: number, bottom: number, right: number): CellRange;
// export function getCellRange(...args: any[]): CellRange{
//   const Range = require('exceljs/lib/doc/range');

//   const range = (() => {
//     if (args.length === 1 && typeof args[0] === 'string') {//rangeStr: string
//       return new Range(args) as ExcelJS.Range;
//     }
//     else if(args.length === 2 && typeof args[0] === 'number' && typeof args[1] === 'number') {//top: number, left: number
//       return new Range({
//         top: args[0],
//         left: args[1],
//         bottom: args[0],
//         right: args[1],
//       }) as ExcelJS.Range;
//     }
//     else if(args.length === 4 && typeof args[0] === 'number' && typeof args[1] === 'number' && typeof args[2] === 'number' && typeof args[3] === 'number') {//op: number, left: number, bottom: number, right: number)
//       return new Range({
//         top: args[0],
//         left: args[1],
//         bottom: args[2],
//         right: args[3],
//       }) as ExcelJS.Range;
//     }
//     else{
//       console.log("args", args)
//       throw new Error("invalid argments");
//     }
//   })();

//   return {
//     rangeStr: range.range,
//     top: range.top,
//     left: range.left,
//     bottom: range.bottom,
//     right: range.right,
//   }
// }
export function getCellRange(address: CellAddress): CellRange{
  const Range = require('exceljs/lib/doc/range');

  const range = (() => {
    if (isCellAddressString(address)) {//文字列表現
      return new Range(address) as ExcelJS.Range;
    }
    else if(isCellAddressRange(address)) {//top, left, bottom, right
      return new Range({
        top: address.top,
        left: address.left,
        bottom: address.bottom,
        right: address.right,
      }) as ExcelJS.Range;
    }
    else if(isCellAddressOne(address)) {//top, left
      return new Range({
        top: address.top,
        left: address.left,
        bottom: address.top,
        right: address.left,
      }) as ExcelJS.Range;
    }
    else{
      throw new Error("invalid argments");
    }
  })();

  return {
    rangeStr: range.range,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
  }
}
/**
 * px単位をEMU単位に変換する
 */
function convertPxToEMU(pixcel: number): number{
  return Math.round(pixcel * 9525);
}
  


/**
 * セルの位置を表す。
 * ただし、exceljsと同様に、1から始まるインデックス。
 */
interface CellPosition{
  row: number;
  column: number
}
export namespace xcelJsUtility{

  /** "A1:B2"のような文字表現のセル範囲から、Rangeを返します。 */
  export function getRange(range: string): ExcelJS.Range{
    const Range = require('exceljs/lib/doc/range');
    return new Range(range);
  }
  /**
   * px単位をEMU単位に変換する
   */
  function convertPxToEMU(pixcel: number): number{
    return Math.round(pixcel * 9525);
  }
  

  /**
   * 列幅をpixcelに変換する
   * @param worksheet 
   * @param cellIndex 
   * @param rightBottomCell 
   */
  function convertColumnWidthToPx(worksheet: ExcelJS.Worksheet, cellIndex: CellPosition, rightBottomCell?: CellPosition): number{
    return 0;
  }

}
// /**
//  * デフォルトの列サイズ
//  */
// export const EXCEL_DEDAULT_COLUMN_WIDTH = 8.38;


// /**
//  * セルまたはセル範囲を指定して、その高さと幅をpixcelで得る。
//  * ※高さについて、元の単位はpoint(1/72インチ)であり、96dipとして規定します。
//  * ※幅について、convertColumnWidthToPixcelに移譲します。
//  * @param worksheet 
//  * @param leftTopCell 
//  * @param rightBottomCell 
//  */
// export function getCellPixcelSize(worksheet: ExcelJS.Worksheet, leftTopCell: {row: number, column: number}, rightBottomCell?: {row: number, column: number}): {width: number, height: number}{
//   //■rightBottomCellが指定されなかった場合、leftTopCell～leftTopCellを範囲とする
//   const _rightBottomCell = rightBottomCell ?? leftTopCell;

//   //■対象のrowとcolumnを得る
//   const columns = new Array(_rightBottomCell.column - leftTopCell.column + 1).fill(1).map((item, index) => worksheet.getColumn(leftTopCell.column + index));
//   const rows = new Array(_rightBottomCell.row - leftTopCell.row + 1).fill(1).map((item, index) => worksheet.getRow(leftTopCell.row + index));

//   //■Exceljsの単位で長さを算出する
//   const targetWidth = columns.reduce((previouseValue, currentValue) => previouseValue + convertColumnWidthToPixcel(currentValue.width ?? worksheet.properties.defaultColWidth ?? EXCEL_DEDAULT_COLUMN_WIDTH, 8), 0);
//   const targetHeight = rows.reduce((previouseValue, currentValue) => previouseValue + (currentValue.height * 96) / 72, 0);

//   return {width: targetWidth, height: targetHeight};
// }

// /**
//  * カラムの幅をpixcelに変換します。
//  * xlxsデータ上でのカラムサイズ(=ExcelJsのカラムサイズ)は、要約すると「基準フォントが11ptの時に格納できる文字数」です。(※実際の計算は複雑です)
//  * https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html の「<width> (Column Width)」の記述に従い、pixcelに変換します。
//  * @param columnWidth 返還対象のカラムのwidth
//  * @param standardFontWidth 基準ほんとの幅。指定しない場合、「MS Pゴシック」を前提に、8pxになります。
//  * @returns 変換後のpx値
//  */
// function convertColumnWidthToPixcel(columnWidth: number, standardFontWidth: number = 8){
//   return Math.floor((((256 * columnWidth) + Math.floor(128 / standardFontWidth)) / 256) * standardFontWidth);
// }

// // function convertColumnWidthToPixcel2(columnWidth: number, standardFontWidth: number){
// //   const standardColumnWidth = Math.floor(((((((standardFontWidth - 1) / 4) + 3) * 5) / 4) * 3) + standardFontWidth - 5);
// //   if(columnWidth > 1){
// //     return Math.round((columnWidth - 1) * standardFontWidth) + standardColumnWidth;
// //   }
// //   else{
// //     return Math.round(columnWidth * standardColumnWidth);
// //   }
// //   // return Math.floor((((256 * columnWidth) + Math.floor(128 / standardFontWidth)) / 256) * standardFontWidth);
// // }

// /**
//  * セルあるいはセル範囲の中央に、画像の比率を変えずに画像を挿入する
//  * @param worksheet 対象のworksheet
//  * @param imageId workbook.addImageで得たimageId
//  * @param imageWidth 画像の幅
//  * @param imageHeight 画像の高さ
//  * @param leftTopCell 挿入先のセル
//  * @param rightBottomCell 挿入先のセル範囲の右下のセル
//  */
// export function addImageToCell(worksheet: ExcelJS.Worksheet, imageId: number, imageWidth: number, imageHeight: number, leftTopCell: {row: number, column: number}, rightBottomCell?: {row: number, column: number}): void{
//   //■rightBottomCellが指定されなかった場合、leftTopCell～leftTopCellを範囲とする
//   const _rightBottomCell = rightBottomCell ?? leftTopCell;

//   //■配置先のサイズを算出
//   const targetSize = getCellPixcelSize(worksheet, leftTopCell, _rightBottomCell);

//   //■上と下の余白の決定(px)し、targitSizeから除外(ボーダーとかぶるため)
//   const globalMargin = 2;
//   const globalMarginEmu = convertPixcelToEMU(globalMargin);
//   targetSize.width -= (globalMargin * 2);
//   targetSize.height -= (globalMargin * 2);

//   //■セル範囲に対して、EMU単位で余白を決定
//   const [left, right, top, botom] = (() => {
//     if((imageHeight / imageWidth) * targetSize.width > targetSize.height){//高さを基準にする(※縦横比を変えずに幅を名いっぱいとると高さが超える)
//       const width = (imageWidth * targetSize.height) / imageHeight;
//       const margin = convertPixcelToEMU((targetSize.width - width) / 2);

//       return [margin + globalMarginEmu, (globalMarginEmu + margin) * -1, globalMarginEmu, globalMarginEmu * -1];
//     }
//     else{//幅を基準
//       const height = (imageHeight * targetSize.width) / imageWidth;
//       const margin = convertPixcelToEMU((targetSize.height - height) / 2);
//       return [globalMarginEmu, globalMarginEmu * -1, globalMarginEmu + margin, (globalMarginEmu + margin) * -1];
//     }
//   })();

//   //■イメージの追加
//   worksheet.addImage(imageId, {
//     tl: { nativeCol: leftTopCell.column - 1, nativeRow: leftTopCell.row - 1, nativeColOff: left, nativeRowOff: top } as Anchor,
//     br: { nativeCol: _rightBottomCell.column, nativeRow: _rightBottomCell.row, nativeColOff: right, nativeRowOff: botom } as Anchor,
//     editAs: "absolute"
//   });
// }

// /** Excelを描画するためのテンプレートデータ */
// export interface ExcelTemplate{
//   rows: ExcelHeaderTemplate[];
//   columns: ExcelHeaderTemplate[];
//   rowCount: number;
//   columnCount: number;
//   cells: ExcelCellTemplate[];
// }

// interface ExcelHeaderTemplate{
//   index: number,
//   size: number,
// }

// interface ExcelCellTemplate{
//   rowIndex: number;
//   columnIndex: number;
//   colSpan: number;
//   rowSpan: number;
//   value: ExcelJS.CellValue;
//   style: Partial<ExcelJS.Style>;
// }

// /**
//  * ArrayBufferに格納されたxlsxデータを、WijmoのWorkbookとしてロードする
//  * @param data 
//  * @returns 
//  */
// async function loadExcelToWjWorkBook(data: ArrayBuffer): Promise<wjXlsx.Workbook> {
//   const book = new wjXlsx.Workbook();
//   return new Promise<wjXlsx.Workbook>((resolve, reject) => {
//     book.loadAsync(data, (book => {
//       resolve(book);
//     }), error => {
//       reject(error);
//     });
//   });
// }

// /**
//  * Excelファイルをロードし、指定されたシート、行数、列数の範囲をExcelTemplate型のデータとして読み取る
//  * @param templateData Excelデータ
//  * @param sheetName ロード対象のシート名
//  * @param rowCount テンプレートとしてロードする行数
//  * @param columnCount テンプレートとしてロードする列数
//  * @returns 
//  */
// export async function loadExcelTemplate(templateData: Blob, sheetName: string, rowCount: number, columnCount: number): Promise<ExcelTemplate>{
//     //■wjimoのworkbookをロード
//     const arrayBuffer = await templateData.arrayBuffer();
//     const wjBook = await loadExcelToWjWorkBook(arrayBuffer);
//     const wjSheet = wjBook.sheets.find(sheet => sheet.name === sheetName);
//     if(wjSheet == null)throw new Error(`workbook何に指定されたworksheetがありません。sheetName=${sheetName}`);

//     //■ExcelJsのworkbookをロード
//     const excelJsBook = new ExcelJS.Workbook();
//     await excelJsBook.xlsx.load(arrayBuffer);
//     const excelJsSheet = excelJsBook.getWorksheet(sheetName);
//     if(excelJsSheet == null)throw new Error(`workbook何に指定されたworksheetがありません。sheetName=${sheetName}`);

//   //■rowsの生成
//   const rows: ExcelHeaderTemplate[] = [];
//   for(let rowIndex = 0; rowIndex < rowCount; rowIndex++){
//     rows.push({
//       index: rowIndex + 1,
//       size: excelJsSheet.getRow(rowIndex + 1).height ?? excelJsSheet.properties.defaultRowHeight,
//     });
//   }

//   //■columnsの生成
//   const columns: ExcelHeaderTemplate[] = [];
//   for(let columnIndex = 0; columnIndex < columnCount; columnIndex++){
//     columns.push({
//       index: columnIndex + 1,
//       size: excelJsSheet.getColumn(columnIndex + 1).width ?? excelJsSheet.properties.defaultColWidth ?? EXCEL_DEDAULT_COLUMN_WIDTH,
//     });
//   }

//   //■cellsの生成
//   const cells: ExcelCellTemplate[] = [];
//   for(let rowIndex = 0; rowIndex < rowCount; rowIndex++){
//     const wjRow = wjSheet.rows[rowIndex]
//     for(let columnIndex = 0; columnIndex < columnCount; columnIndex++){
//       const excelJsCell = excelJsSheet.getCell(rowIndex + 1, columnIndex + 1);
//       const wjCell = wjRow.cells[columnIndex];       
      
//       //■CellTemplateの生成
//       const cellTemplate: ExcelCellTemplate = {
//         rowIndex: rowIndex + 1,
//         columnIndex: columnIndex + 1,
//         colSpan: wjCell != null && wjCell.colSpan != null ? wjCell.colSpan : 1,
//         rowSpan: wjCell != null && wjCell.rowSpan != null ? wjCell.rowSpan : 1,
//         value: lodash.cloneDeep(excelJsCell.value),
//         style: lodash.cloneDeep(excelJsCell.style),
//       }
//       cells.push(cellTemplate);
//     }
//   }

//   return {
//     rows: rows,
//     columns: columns,
//     rowCount: rowCount,
//     columnCount: columnCount,
//     cells: cells,
//   }
// }

// /**
//  * ExcelTemplateから、指定した行へ行サイズを設定する
//  * @param sheet 描画対象のシート
//  * @param row 描画先の行。1から始まるインデックス。
//  * @param excelTemplate 描画データを格納するExcelTemplate
//  */
// export function drawExcelTemplateRowHeaderSize(sheet: ExcelJS.Worksheet, row: number, excelTemplate: ExcelTemplate): void{
//   if(row <= 0)throw new Error("rowが0以下です。");
//   const startRow = row - 1;
//   //■行の設定
//   excelTemplate.rows.forEach(item => {
//     sheet.getRow(item.index + startRow).height = item.size;
//   });
// }
// /**
//  * ExcelTemplateから、指定した列へ列サイズを設定する
//  * @param sheet 描画対象のシート
//  * @param column 描画先の列。1から始まるインデックス。
//  * @param excelTemplate 描画データを格納するExcelTemplate
//  */
// export function drawExcelTemplateColumnHeaderSize(sheet: ExcelJS.Worksheet, column: number, excelTemplate: ExcelTemplate): void{
//   if(column <= 0)throw new Error("columnが0以下です。")
//   const startColumn = column - 1;
//   //■列の設定
//   excelTemplate.columns.forEach(item => {
//     sheet.getColumn(item.index + startColumn).width = item.size;
//   })
// }


// const drawExcelTemplateBodyRegex = /^\$\{([^}]+)\}$/;//差し込み式({[プロパティのパス]})かどうかの判断
// /**
//  * ExcelTemplateから、指定した行、列へ、セルを描画する。
//  * @param sheet 描画対象のシート
//  * @param row 描画先の行。1から始まるインデックス。
//  * @param column 描画先の列。1から始まるインデックス。
//  * @param excelTemplate 描画データを格納するExcelTemplate
//  * @param value データの差し込みに利用するオブジェクト。undefinedの場合は差し込みしない。
//  * @param valueCustomizer valueをカスタマイズして
//  */
// export function drawExcelTemplateBody(sheet: ExcelJS.Worksheet, row: number, column: number, excelTemplate: ExcelTemplate, value?: object, valueCustomizer?: (value: object) => void): void{
//   if(row <= 0)throw new Error("rowが0以下です。");
//   if(column <= 0)throw new Error("columnが0以下です。")
//   const [startRow, startColumn] = [row - 1, column - 1];


//   //■値の差し替えのためのデータ作成
//   const _value = value == null ? undefined : valueCustomizer == null ? value : (() => {
//     const clonedValue = lodash.cloneDeep(value);
//     valueCustomizer(clonedValue);
//     return clonedValue;
//   })();

//   //■セルの描画
//   excelTemplate.cells.forEach(cellTemplate => {
//     //■cellの取得
//     const cell = sheet.getCell(cellTemplate.rowIndex + startRow, cellTemplate.columnIndex + startColumn);

//     //■マージの反映
//     if(cellTemplate.rowSpan > 1 || cellTemplate.colSpan > 1){
//       sheet.mergeCells(cellTemplate.rowIndex + startRow, cellTemplate.columnIndex + startColumn, cellTemplate.rowIndex + startRow + cellTemplate.rowSpan - 1, cellTemplate.columnIndex + startColumn + cellTemplate.colSpan - 1);
//     }
//     //■スタイルの反映
//     cell.style = cellTemplate.style;

//     //■値の反映
//     if(_value != null && lodash.isString(cellTemplate.value)){//値の差し替えがされない条件を除外。1)describerが指定されていない。2)値が文字列ではない。(差し込み式ではない）
//       //■値の差し込み
//       const match = cellTemplate.value.match(drawExcelTemplateBodyRegex);
//       if (match != null && match.length >= 0 && match[0] != null) {//差し込み式が指定がされている
//         cell.value = lodash.get(_value, match[1], null);
//       } else {
//         cell.value = cellTemplate.value;
//       }
//     }
//     else{
//       //■テンプレートの値をそのまま転記
//       cell.value = cellTemplate.value;
//     }
//   });
// }


// /**
//  * ExcelJSのworkbookをダウンロードする。※２回目以降の実行では、chromeにおいて、「このサイトで複数ファイルの自動ダウンロードが試行されました」という警告と許可を求めるダイアログが表示される。
//  * @param workbook ダウンロードさせるworkbook
//  * @param fileName ダウンロードするファイル名
//  */
// export async function downloadWorkbook(workbook: ExcelJS.Workbook, fileName: string){
//   const buffer = await workbook.xlsx.writeBuffer();
//   const blob = new Blob([buffer], {
//     type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//   });
//   const url = window.URL.createObjectURL(blob);
//   try{
//     const anchor = document.createElement('a');
//     anchor.href = url;
//     anchor.download = fileName;
//     anchor.click();
//   }
//   finally{
//     window.URL.revokeObjectURL(url);
//   }
  
// }
