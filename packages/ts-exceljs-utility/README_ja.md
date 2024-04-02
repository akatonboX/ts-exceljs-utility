# ts-exceljs-utility
* ExcelJS(https://www.npmjs.com/package/exceljs)を利用するにあたって、便利であろう機能を集めたパッケージです。
1. "A1:B1"と[row, col]の相互変換
1. 画像の比率を変えずに、指定したセル範囲に画像を挿入する。
1. workbookをダウンロードさせる。

# Note
* CRAで作成したReactアプリケーションから、chromeで動作確認しています。

# Installation
```shell
yarn add ts-exceljs-utility
```

# release
* [2024/3/31]v1.0.1 released.
* [2024/3/24]v1.0.0 released.

# CellAddress
* セル、またはセルの範囲を表現するインターフェースです。下記の表現を許容します。
1. 文字列による表現。```"A1", "A1:B2"```
2. 一つのセルを、行と列のインデックスで表現。```{top: 1, left: 1} //="A1"```
3. セルの範囲を、左上の行と列のインデックスと右下の行と列のインデックスで表現。```{top: 1, left: 1, bottom: 2, right: 2} //="A1:B2"```
* 下記のように定義されます。
``` typescript
export type CellAddress = 
  string 
  | {
    top: number;
    left: number;
  } 
  | {   
    top: number;
    left: number;
    bottom: number;
    right: number;
  }
```

# セルまたはセル範囲の、文字表現とインデックス表現の相互変換
<a href="https://app.archive-gp.com/ts-exceljs-utility/example1">demo</a>
* 下記のメソッドを利用します。
```typescript
function getCellRange(address: CellAddress): CellRange

type CellRange = {
  rangeStr: string;//"A1:B2"のような文字表現
  top: number;//左上のセルの行の1から始まるインデックス。
  left: number;//左上のセルの列の1から始まるインデックス。
  bottom: number;//右下のセルの行の1から始まるインデックス。
  right: number;//右↓のセルの列の1から始まるインデックス
}
```
* 下記のように使用します。
``` typescript
import { getCellRange } from "ts-exceljs-utility";

const range1 = getCellRange("A1:B2");
// range1 = {rangeStr: "A1:B2", top: 1, left: 1, bottom: 2, right: 2}

const range2 = getCellRange({top: 1, left: 1});
// range2 = {rangeStr: "A1:A1", top: 1, left: 1, bottom: 1, right: 1}

const range3 = getCellRange({top: 1, left: 1, bottom: 2, right: 2});
// range3 = {rangeStr: "A1:B2", top: 1, left: 1, bottom: 2, right: 2}
```

# 画像の比率を変えずに、指定したセル範囲に画像を挿入
<a href="https://app.archive-gp.com/ts-exceljs-utility/example2">demo</a>

## ```createAddImageParam```メソッド
* 画像のURL(string)あるいは、Blobから、ExcelJSの```workbook.addImage```の引数を構築できます。
## ```addImageToCell```メソッド
* 画像のURL(string)あるいはBlobから、画像を挿入します。
* ```CellAddress```で挿入先のセル、またはセル範囲を指定できます。
* 画像の縦横比を変えずに、かつ収まるように拡大あるいは縮小して、画像を挿入します。
* 画像のサイズを指定する必要があります。
* 上下左右に2pxのマージンが付きます。

```typescript
import { createAddImageParam, addImageToCell} from "ts-exceljs-utility";

//■workbookとworkshettを作成
const book = new ExcelJS.Workbook();
const sheet = book.addWorksheet("sample");

//■画像の挿入対象の行列には、明示的に幅と高さを設定
for(let i = 1; i <= 100; i++){
  sheet.getRow(i).height = 13.5;
  sheet.getColumn(i).width = 8.38;
}

//■workbookにimageを挿入
const imageId = book.addImage(await createAddImageParam(horizonImage));

//■worksheetにimageを挿入
//サイズが712x357の画像を"B1:K10"に挿入します。
addImageToCell(sheet, imageId, 712, 357, "B1:K10");

```

# xlsxをassetとして利用する[v1.0.1]
<a href="https://app.archive-gp.com/ts-exceljs-utility/example3">demo</a>
* xlsxをpng画像などのように、assetとして管理し、workbookにロードする機能を提供します。
* これは、CRAで作成されたReactアプリケーションを前提とします。

## cracoの導入と設定
* xlsxをassetとして利用するためには、craco(https://www.npmjs.com/package/@craco/craco)の導入が必要です。
```javascript
[craco.config.js]
module.exports = {
  webpack: {
    configure: (webpackConfig, { env, paths }) => {
      webpackConfig.module.rules.push({
        test: /\.xlsx$/,
        type: 'asset',
        parser: {
          dataUrlCondition: {
            maxSize: 0,
          },
        },
      });
      return webpackConfig;
    },
  },
};
```
## xlsx.d.tsを用意する
```typescript
[xlsx.d.ts]
declare module '*.xlsx' {
  const content: string; 
  export default content;
}
```

## assetとして、```src```ディレクトリにxlxsを格納
* 例えば、```src/assets/resource.xlxs```を格納します。

## assetをworkbookにロードする
* ```import assetXlsx from "../assets/example.xlsx";```でURLを取得し、```loadWorkbook```でworkbookを得ることができます。
* 存在しないURLや、ダウンロードされたデータをwokrbookが読めなかった場合、```loadWorkbook```の戻り値はundefiendになります。
```typescript
import assetXlsx from "../assets/example.xlsx";
.
.
.
const workbook = await loadWorkbook(assetXlsx);
if(workbook != null){
  downloadWorkbook(workbook, "example.xlsx");
}
```

# colBreaks(列方向の改行)を挿入する。[v1.0.2]
<a href="https://app.archive-gp.com/ts-exceljs-utility/example4">demo</a>
* ```addColBreak(worksheet: ExcelJS.Worksheet, colIndex: number): void```を使用して、worksheetにcolBreak(列方向の改行)を追加します。
* ```colIndex```は、0から始まる列の位置です。
* ```downloadWorkbook```メソッド、もしくは```writeWithColBreaks```メソッドで、colBreaksを反映したxlsxファイルを取得できます。
```javascript
  //■リソースからworkbookをロード
  const workbook = await loadWorkbook(assetXlsx);
  if(workbook == null)return;

  //■入力から改行する列の配列に変換([5,10,15])
  const breaks = input.split(",").map(item => item.trim()).filter(item => item.length > 0).map(item => Number(item)).filter(item => !isNaN(item));

  //■一つ目のシートに列方向の改行を追加
  const sheet = workbook.worksheets[0];
  breaks.forEach(colBreak => {
    addColBreak(sheet, colBreak);
  });

  //■ダウンロードさせる
  downloadWorkbook(workbook, "example.xlsx");
  //const blob = await writeWithColBreaks(workbook);
```