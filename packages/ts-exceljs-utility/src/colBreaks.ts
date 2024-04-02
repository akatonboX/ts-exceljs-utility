import ExcelJS from "exceljs";
import * as JSZip from 'jszip';

const COLBREAKS_PROPERTY_NAME = "_colBreaks";

/**
 * シートに設定されているcolBreaksを取得する。
 * @param worksheet 
 * @returns 
 */
export function getColBreaks(worksheet: ExcelJS.Worksheet): number[]{
  const colBreaks = (worksheet as any)[COLBREAKS_PROPERTY_NAME];
  if(colBreaks == null){
    return [];
  }
  else{
    return colBreaks as number[];
  }
}

/**
 * シートにcolBreaksを設定する。
 */
export function setColBreaks(worksheet: ExcelJS.Worksheet, colBreaks: number[]): void{
  if(colBreaks == null || colBreaks.length === 0){
    delete (worksheet as any)[COLBREAKS_PROPERTY_NAME];
  }
  else{
    (worksheet as any)[COLBREAKS_PROPERTY_NAME] = colBreaks.sort((a, b) => a - b);
  }
}
/**
 * シートにcolBreakを追加する。
 * @param worksheet 
 * @param colIndex 列を表す0から始まるindex。
 */
export function addColBreak(worksheet: ExcelJS.Worksheet, colIndex: number): void{
  setColBreaks(worksheet, [...getColBreaks(worksheet), colIndex]);
}

const colBreaksRegex = /<colBreaks>[\s\S]*?<\/colBreaks>/g;
export async function writeWithColBreaks(workbook: ExcelJS.Workbook): Promise<Blob>{

  //■workbookをbufferに出力し、zipで解析。
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);

  //■sheet毎に、colBreaksの追加処理
  for(var i = 0; i < workbook.worksheets.length; i++){
    //■該当のworksheetからcolBreaksの設定を取得
    const worksheet = workbook.worksheets[i];
    const colBreaks = getColBreaks(worksheet);

    //■worksheetにcolbreaksを追加
    //※DOMParser, XMLSerializerで処理しようとすると、xmlnsを持つ<worksheet>に追加した<colBreaks>が、xmlns=""を持ってしまうので、文字列操作で行う。
    if(colBreaks.length > 0){//colBreaksの設定がある場合
      //■zipから該当のworksheetのxmlを取り出す。
      const xmlFilePath = `xl/worksheets/${workbook.worksheets[0].name.toLowerCase()}.xml`;
      const xmlFile = zip.file(xmlFilePath);
      if(xmlFile == null)throw Error(`file is not found. file=${xmlFilePath}`)
      const sheetFileText = (await xmlFile.async("string")).replace(colBreaksRegex, "");//すでに<colBreaks>があれば削除
      console.log("★1", sheetFileText)
      //■<brk>を作成
      const brksText = colBreaks.map(colBreak => `<brk id="${colBreak}" man="1" max="1048576" />`).reduce((previouseValue, currentValue) => previouseValue + currentValue, "");
      console.log("★2", brksText)
      //■<colBreaks>を作成
      const colBreaksText = `<colBreaks count="${colBreaks.length}" manualBreakCount="${colBreaks.length}">${brksText}</colBreaks>`;
      console.log("★3", colBreaksText)

      //■<colBreaks>を挿入
      const generatedSheetFileText = sheetFileText.replace("</worksheet>", colBreaksText + "</worksheet>");
      console.log("★4", generatedSheetFileText)
      //■修正したxmlをzipに上書き
      await zip.file(xmlFilePath, generatedSheetFileText);
    }
  }

  //■zipをblobに変換して返却
  return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}



/* DOMParser, XMLSerializerで処理しようとすると、xmlnsを持つ<worksheet>に追加した<colBreaks>が、xmlns=""を持ってしまうので使用できない。*/
// export async function writeWithColBreaks(workbook: ExcelJS.Workbook): Promise<Blob>{

//   //■workbookをbufferに出力し、zipで解析。
//   const buffer = await workbook.xlsx.writeBuffer();
//   const zip = await JSZip.loadAsync(buffer);

//   //■xmlをdomで解析するためのオブジェクトを構築
//   const parser = new DOMParser();
//   const serializer = new XMLSerializer();

//   //■sheet毎に、colBreaksの追加処理
//   for(var i = 0; i < workbook.worksheets.length; i++){
//     //■該当のworksheetからcolBreaksの設定を取得
//     const worksheet = workbook.worksheets[i];
//     const colBreaks = getColBreaks(worksheet);
//     // console.log("★1", serializer.serializeToString(doc))
//     //■worksheetにcolbreaksを追加
//     if(colBreaks.length > 0){//colBreaksの設定がある場合
//       //■zipから該当のworksheetのxmlを取り出す。
//       const xmlFilePath = `xl/worksheets/${workbook.worksheets[0].name.toLowerCase()}.xml`;
//       const xmlFile = zip.file(xmlFilePath);
//       if(xmlFile == null)throw Error(`file is not found. file=${xmlFilePath}`)

//       //■domで解析
//       const doc = parser.parseFromString(await xmlFile.async("string"), 'application/xml');
//       const worksheetElement = doc.getElementsByTagName("worksheet")[0];
//       if(worksheetElement == null)throw Error("root node is not found.");

//       //■<colBreaks>を作成
//       const colBreaksElement = doc.createElement("colBreaks");
//       worksheetElement.appendChild(colBreaksElement);
//       colBreaksElement.setAttribute("count", String(colBreaks.length));
//       colBreaksElement.setAttribute("manualBreakCount", String(colBreaks.length));
//       colBreaksElement.removeAttribute("xmlns");

//       //■<brk>を作成
//       colBreaks.forEach(colBreak => {
//         console.log("★1", colBreak)
//         const brkElement = doc.createElement("brk");
//         brkElement.setAttribute("id", String(colBreak));
//         brkElement.setAttribute("man", "1");
//         brkElement.setAttribute("max", "1048576");
//         colBreaksElement.appendChild(brkElement);
//       });
// console.log("★", serializer.serializeToString(colBreaksElement))
//       //■修正したxmlをzipに上書き
//       await zip.file(xmlFilePath, serializer.serializeToString(doc));
//     }
//   }

//   //■zipをblobに変換して返却
//   return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
// }
