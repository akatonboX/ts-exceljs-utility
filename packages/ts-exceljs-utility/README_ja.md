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
* [2024/3/24]v1.0.0 released 

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
