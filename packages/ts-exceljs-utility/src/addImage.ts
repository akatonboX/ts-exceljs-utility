import ExcelJS from "exceljs";
import { DEFAULT_STANDAED_FONT_WIDTH, getCellPixcelSize, px2Emu } from "./size";
import { CellAddress } from "./core";
import { getCellRange } from "./range";



function blob2Base64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
     const reader = new FileReader();
     reader.onloadend = () => resolve(reader.result as string);
     reader.onerror = reject;
     reader.readAsDataURL(blob);
  });
 }

export async function createAddImageParam(url: string): Promise<{extension: 'jpeg' | 'png' | 'gif', base64: string}>;
export async function createAddImageParam(blob: Blob): Promise<{extension: 'jpeg' | 'png' | 'gif', base64: string}>;
export async function createAddImageParam(data: string | Blob): Promise<{extension: 'jpeg' | 'png' | 'gif', base64: string}>{
  //■blobを生成
  const blob = await (async () => {
    if( typeof data === 'string'){
      const response = await fetch(data);
      return await response.blob();
    }
    else if(typeof data === "object" && data instanceof Blob){
      return data;
    }
    else{
      throw new Error("invalid argments");
    }
  })();

  //■extensionの決定
  const extension: 'jpeg' | 'png' | 'gif' = (() => {
    switch(blob.type){
      case "image/jpeg": return "jpeg";
      case "image/png": return "png";
      case "image/gif": return "gif";
      default: throw new Error(`invalid mimetype.minetype=${blob.type}`);
    }
  })();

  //■base64文字列化
  const base64 = await blob2Base64(blob);

  //■戻り値の生成
  return {
    base64: base64,
    extension: extension,
  };
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
export function addImageToCell(worksheet: ExcelJS.Worksheet, imageId: number, imageWidth: number, imageHeight: number, targetAddress: CellAddress, standardFontWidth: number = DEFAULT_STANDAED_FONT_WIDTH): void{
  //■targetAddressをCellRangeに変換
  const targetRange = getCellRange(targetAddress);

  //■配置先のサイズを算出
  const targetPxSize = getCellPixcelSize(worksheet, targetRange, standardFontWidth);
  console.log("★", targetPxSize)
  if(isNaN(targetPxSize.height)){
    throw new Error("Some rows have no height set.");
  }
  if(isNaN(targetPxSize.width)){
    throw new Error("Some columns have no width set.");
  }
  //■paddingを2pxにする
  const globalMargin = 2;
  targetPxSize.width -= (globalMargin * 2);
  targetPxSize.height -= (globalMargin * 2);

  //■セル範囲に対して、EMU単位で余白を決定
  const [left, right, top, bottom, leftMargin, rightMargin, topMargin, botomMargin] = (() => {
    if((imageHeight / imageWidth) * targetPxSize.width > targetPxSize.height){//高さを基準にする
      //■マージンを計算
      const width = (imageWidth * targetPxSize.height) / imageHeight;
      const margin = ((targetPxSize.width - width) / 2) + globalMargin;

      //■左マージンの調整(マージンに全体が含まれるセルを除外)
      const [left, leftMargin] = (() => {
        let calcMargin = margin;
        let index = 0;
        for(let i = 0; i < targetPxSize.widths.length; i++){
          const colWidth = targetPxSize.widths[i];
          if(calcMargin < colWidth){
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
        for(let i = targetPxSize.widths.length - 1; i >= 0; i--){
          const colWidth = targetPxSize.widths[i];
          if(calcMargin < colWidth){
            index = i;
            break;
          }
          calcMargin -= colWidth;
        }
        return [index, calcMargin];
      })();
      return [targetRange.left + left, targetRange.left + right, targetRange.top, targetRange.bottom, leftMargin, rightMargin * -1, globalMargin, globalMargin * -1];
    }
    else{//幅を基準にする
      
      const height = (imageHeight * targetPxSize.width) / imageWidth;
      const margin = ((targetPxSize.height - height) / 2) + globalMargin;
      //■上マージンの調整(マージンに全体が含まれるセルを除外)
      const [top, topMargin] = (() => {
        let calcMargin = margin;
        let index = 0;
        for(let i = 0; i < targetPxSize.heights.length; i++){
          const rowHeight = targetPxSize.heights[i];
          if(calcMargin < rowHeight){
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
        for(let i = targetPxSize.heights.length - 1; i >= 0; i--){
          const rowHeight = targetPxSize.heights[i];
          if(calcMargin < rowHeight){
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
    tl: { nativeCol: left - 1, nativeRow: top - 1, nativeColOff: px2Emu(leftMargin), nativeRowOff: px2Emu(topMargin) } as ExcelJS.Anchor,
    br: { nativeCol: right, nativeRow: bottom, nativeColOff: px2Emu(rightMargin), nativeRowOff: px2Emu(botomMargin) } as ExcelJS.Anchor,
    // editAs: "absolute"
  });
}