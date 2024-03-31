'use strict';

var ExcelJS = require('exceljs');

/**
 * CellAddressが文字列表現かどうかを調べる
 * @param address
 * @returns
 */
function isCellAddressString(address) {
    return typeof address === 'string';
}
/**
 * CellAddressが単一表現かどうかを調べる
 * @param address
 * @returns
 */
function isCellAddressOne(address) {
    return typeof address === 'object' && 'top' in address && 'left' in address;
}
/**
 * CellAddressが範囲表現かどうかを調べる
 * @param address
 * @returns
 */
function isCellAddressRange(address) {
    return typeof address === 'object' && 'top' in address && 'left' in address && 'bottom' in address && 'right' in address;
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
function getCellRange(address) {
    const Range = require('exceljs/lib/doc/range');
    const range = (() => {
        if (isCellAddressString(address)) { //文字列表現
            return new Range(address);
        }
        else if (isCellAddressRange(address)) { //top, left, bottom, right
            return new Range({
                top: address.top,
                left: address.left,
                bottom: address.bottom,
                right: address.right,
            });
        }
        else if (isCellAddressOne(address)) { //top, left
            return new Range({
                top: address.top,
                left: address.left,
                bottom: address.top,
                right: address.left,
            });
        }
        else {
            throw new Error("invalid argments");
        }
    })();
    return {
        rangeStr: range.range,
        top: range.top,
        left: range.left,
        bottom: range.bottom,
        right: range.right,
    };
}
var xcelJsUtility;
(function (xcelJsUtility) {
    /** "A1:B2"のような文字表現のセル範囲から、Rangeを返します。 */
    function getRange(range) {
        const Range = require('exceljs/lib/doc/range');
        return new Range(range);
    }
    xcelJsUtility.getRange = getRange;
})(xcelJsUtility || (xcelJsUtility = {}));
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

/******************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */
/* global Reflect, Promise, SuppressedError, Symbol */


function __awaiter(thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}

typeof SuppressedError === "function" ? SuppressedError : function (error, suppressed, message) {
    var e = new Error(message);
    return e.name = "SuppressedError", e.error = error, e.suppressed = suppressed, e;
};

const DEFAULT_STANDAED_FONT_WIDTH = 8;
/**
 * px単位をEMU単位に変換する
 */
function px2Emu(pixcel) {
    return Math.round(pixcel * 9525);
}
/**
 * セルの高さをpxに変換する
 * @param point
 */
function height2Px(worksheet, height) {
    return (height * 96) / 72;
}
/**
 * セルの幅をpxに変換する
 * @param width
 * @param standardFontWidth
 * @returns
 */
function width2Px(worksheet, width, standardFontWidth = DEFAULT_STANDAED_FONT_WIDTH) {
    const w = width !== null && width !== void 0 ? width : worksheet.properties.defaultColWidth;
    if (w == null)
        return NaN;
    return Math.floor((((256 * w) + Math.floor(128 / standardFontWidth)) / 256) * standardFontWidth);
}
/**
 * セルまたはセル範囲を指定して、その高さと幅をpixcelで得る。
 * ※高さについて、元の単位はpoint(1/72インチ)であり、96dpiとして変換します。
 * ※幅について、https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html の「<width> (Column Width)」の記述に従い、pixcelに変換します。
 */
function getCellPixcelSize(worksheet, range, standardFontWidth = DEFAULT_STANDAED_FONT_WIDTH) {
    //■対象の範囲に含まれるrowとcolumnの配列を得る
    const rows = new Array(range.bottom - range.top + 1).fill(1).map((item, index) => worksheet.getRow(range.top + index));
    const columns = new Array(range.right - range.left + 1).fill(1).map((item, index) => worksheet.getColumn(range.left + index));
    //■pxに変換
    const heights = rows.map(item => height2Px(worksheet, item.height));
    const widths = columns.map(item => width2Px(worksheet, item.width, standardFontWidth));
    //■合計
    const height = heights.reduce((previouseValue, currentValue) => previouseValue + currentValue, 0);
    const width = widths.reduce((previouseValue, currentValue) => previouseValue + currentValue, 0);
    return { width: width, height: height, widths: widths, heights: heights };
}

function blob2Base64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}
function createAddImageParam(data) {
    return __awaiter(this, void 0, void 0, function* () {
        //■blobを生成
        const blob = yield (() => __awaiter(this, void 0, void 0, function* () {
            if (typeof data === 'string') {
                const response = yield fetch(data);
                return yield response.blob();
            }
            else if (typeof data === "object" && data instanceof Blob) {
                return data;
            }
            else {
                throw new Error("invalid argments");
            }
        }))();
        //■extensionの決定
        const extension = (() => {
            switch (blob.type) {
                case "image/jpeg": return "jpeg";
                case "image/png": return "png";
                case "image/gif": return "gif";
                default: throw new Error(`invalid mimetype.minetype=${blob.type}`);
            }
        })();
        //■base64文字列化
        const base64 = yield blob2Base64(blob);
        //■戻り値の生成
        return {
            base64: base64,
            extension: extension,
        };
    });
}
// export function addImageToCell2(worksheet: ExcelJS.Worksheet, imageId: number, imageWidth: number, imageHeight: number, targetAddress: CellAddress, standardFontWidth: number = DEFAULT_STANDAED_FONT_WIDTH): void{
//   //■targetAddressをCellRangeに変換
//   const targetRange = getCellRange(targetAddress);
//   //■配置先のサイズを算出
//   const targetPxSize = getCellPixcelSize(worksheet, targetRange, standardFontWidth);
//   //■paddingを2pxにする
//   const globalMargin = 2;
//   const globalMarginEmu = px2Emu(globalMargin);
//   targetPxSize.width -= (globalMargin * 2);
//   targetPxSize.height -= (globalMargin * 2);
//   //■セル範囲に対して、EMU単位で余白を決定
//   const [left, right, top, bottom] = (() => {
//     if((imageHeight / imageWidth) * targetPxSize.width > targetPxSize.height){//高さを基準にする
//       console.log("★高さ基準")
//       //■マージンを計算
//       const width = (imageWidth * targetPxSize.height) / imageHeight;
//       const margin = px2Emu((targetPxSize.width - width) / 2);
//       return [margin + globalMarginEmu, (globalMarginEmu + margin) * -1, globalMarginEmu, globalMarginEmu * -1];
//     }
//     else{//幅を基準にする
//       console.log("★幅基準", targetPxSize)
//       const height = (imageHeight * targetPxSize.width) / imageWidth;
//       const margin = px2Emu((targetPxSize.height - height) / 2);
//       return [globalMarginEmu, globalMarginEmu * -1, globalMarginEmu + margin, (globalMarginEmu + margin) * -1];
//     }
//   })();
//   //■イメージの追加
//   worksheet.addImage(imageId, {
//     tl: { nativeCol: targetRange.left - 1, nativeRow: targetRange.top - 1, nativeColOff: left, nativeRowOff: top } as ExcelJS.Anchor,
//     br: { nativeCol: targetRange.right, nativeRow: targetRange.bottom, nativeColOff: right, nativeRowOff: bottom } as ExcelJS.Anchor,
//     // tl: { nativeCol: targetRange.left - 1, nativeRow: targetRange.top - 1} as ExcelJS.Anchor,
//     // br: { nativeCol: targetRange.right, nativeRow: targetRange.bottom} as ExcelJS.Anchor,
//     editAs: "absolute"
//   });
// }
function addImageToCell(worksheet, imageId, imageWidth, imageHeight, targetAddress, standardFontWidth = DEFAULT_STANDAED_FONT_WIDTH) {
    //■targetAddressをCellRangeに変換
    const targetRange = getCellRange(targetAddress);
    //■配置先のサイズを算出
    const targetPxSize = getCellPixcelSize(worksheet, targetRange, standardFontWidth);
    console.log("★", targetPxSize);
    if (isNaN(targetPxSize.height)) {
        throw new Error("Some rows have no height set.");
    }
    if (isNaN(targetPxSize.width)) {
        throw new Error("Some columns have no width set.");
    }
    //■paddingを2pxにする
    const globalMargin = 2;
    targetPxSize.width -= (globalMargin * 2);
    targetPxSize.height -= (globalMargin * 2);
    //■セル範囲に対して、EMU単位で余白を決定
    const [left, right, top, bottom, leftMargin, rightMargin, topMargin, botomMargin] = (() => {
        if ((imageHeight / imageWidth) * targetPxSize.width > targetPxSize.height) { //高さを基準にする
            //■マージンを計算
            const width = (imageWidth * targetPxSize.height) / imageHeight;
            const margin = ((targetPxSize.width - width) / 2) + globalMargin;
            //■左マージンの調整(マージンに全体が含まれるセルを除外)
            const [left, leftMargin] = (() => {
                let calcMargin = margin;
                let index = 0;
                for (let i = 0; i < targetPxSize.widths.length; i++) {
                    const colWidth = targetPxSize.widths[i];
                    if (calcMargin < colWidth) {
                        index = i;
                        break;
                    }
                    calcMargin -= colWidth;
                }
                return [index, calcMargin];
            })();
            //■右マージンの調整(マージンに全体が含まれるセルを除外)
            const [right, rightMargin] = (() => {
                let calcMargin = margin;
                let index = 0;
                for (let i = targetPxSize.widths.length - 1; i >= 0; i--) {
                    const colWidth = targetPxSize.widths[i];
                    if (calcMargin < colWidth) {
                        index = i;
                        break;
                    }
                    calcMargin -= colWidth;
                }
                return [index, calcMargin];
            })();
            return [targetRange.left + left, targetRange.left + right, targetRange.top, targetRange.bottom, leftMargin, rightMargin * -1, globalMargin, globalMargin * -1];
        }
        else { //幅を基準にする
            const height = (imageHeight * targetPxSize.width) / imageWidth;
            const margin = ((targetPxSize.height - height) / 2) + globalMargin;
            //■上マージンの調整(マージンに全体が含まれるセルを除外)
            const [top, topMargin] = (() => {
                let calcMargin = margin;
                let index = 0;
                for (let i = 0; i < targetPxSize.heights.length; i++) {
                    const rowHeight = targetPxSize.heights[i];
                    if (calcMargin < rowHeight) {
                        index = i;
                        break;
                    }
                    calcMargin -= rowHeight;
                }
                return [index, calcMargin];
            })();
            //■下マージンの調整(マージンに全体が含まれるセルを除外)
            const [bottom, bottomMargin] = (() => {
                let calcMargin = margin;
                let index = 0;
                for (let i = targetPxSize.heights.length - 1; i >= 0; i--) {
                    const rowHeight = targetPxSize.heights[i];
                    if (calcMargin < rowHeight) {
                        index = i;
                        break;
                    }
                    calcMargin -= rowHeight;
                }
                return [index, calcMargin];
            })();
            return [targetRange.left, targetRange.right, targetRange.top + top, targetRange.top + bottom, globalMargin, globalMargin * -1, topMargin, bottomMargin * -1];
        }
    })();
    //■イメージの追加
    worksheet.addImage(imageId, {
        tl: { nativeCol: left - 1, nativeRow: top - 1, nativeColOff: px2Emu(leftMargin), nativeRowOff: px2Emu(topMargin) },
        br: { nativeCol: right, nativeRow: bottom, nativeColOff: px2Emu(rightMargin), nativeRowOff: px2Emu(botomMargin) },
        // editAs: "absolute"
    });
}

/**
 * URLからxlsxをダウンロードしてworkbookに読み込みます。
 * @param url
 * @returns
 */
function loadWorkbook(url) {
    return __awaiter(this, void 0, void 0, function* () {
        //■指定されたURLから、ダウンロード
        const response = yield fetch(url);
        if (!response.ok) {
            return undefined;
        }
        //■workbookに読み来い
        const workbook = new ExcelJS.Workbook();
        try {
            yield workbook.xlsx.load(yield response.arrayBuffer());
        }
        catch (e) {
            console.warn("Workbook.xlsx.load method failed. ", e);
            return undefined;
        }
        return workbook;
    });
}
/**
 * ExcelJSのworkbookをダウンロードする。※２回目以降の実行では、chromeにおいて、「このサイトで複数ファイルの自動ダウンロードが試行されました」という警告と許可を求めるダイアログが表示される。
 * @param workbook ダウンロードさせるworkbook
 * @param fileName ダウンロードするファイル名
 */
function downloadWorkbook(workbook, fileName) {
    return __awaiter(this, void 0, void 0, function* () {
        const buffer = yield workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = window.URL.createObjectURL(blob);
        try {
            const anchor = document.createElement('a');
            anchor.href = url;
            anchor.download = fileName;
            anchor.click();
        }
        finally {
            window.URL.revokeObjectURL(url);
        }
    });
}

exports.addImageToCell = addImageToCell;
exports.createAddImageParam = createAddImageParam;
exports.downloadWorkbook = downloadWorkbook;
exports.getCellRange = getCellRange;
exports.loadWorkbook = loadWorkbook;
//# sourceMappingURL=index.js.map
