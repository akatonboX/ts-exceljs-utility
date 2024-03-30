# ts-exceljs-utility
* This is a package that gathers useful functionalities for using ExcelJS (https://www.npmjs.com/package/exceljs).
1. Mutual conversion between "A1:B1" and [row, col].
1. Inserting images into specified cell ranges without changing the aspect ratio of the images.
1. Allowing downloading of workbooks.
# Note
* Tested on Chrome from a React application created with CRA.
# Note
* CRAで作成したReactアプリケーションから、chromeで動作確認しています。

# Installation
```shell
yarn add ts-exceljs-utility
```

# release
* [2024/3/24]v1.0.0 released 

# CellAddress
* An interface representing a cell or a range of cells. It allows the following representations:
1. Representation by string. ```"A1"```, ```"A1:B2"```
1. Representation of a single cell by row and column indices. ```{top: 1, left: 1} //="A1"```
1. Representation of a cell range by indices of the top-left and bottom-right cells. ```{top: 1, left: 1, bottom: 2, right: 2} //="A1:B2"```
* Defined as follows:
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

# Mutual conversion between cell or cell range's string representation and index representation
<a href="https://app.archive-gp.com/ts-exceljs-utility/example1">demo</a>
* Use the method below.
```typescript
function getCellRange(address: CellAddress): CellRange

type CellRange = {
  rangeStr: string;//String representation like "A1:B2"
  top: number;//Index of the top-left cell starting from 1 for rows.
  left: number;//Index of the top-left cell starting from 1 for columns.
  bottom: number;//Index of the bottom-right cell starting from 1 for rows.
  right: number;//Index of the bottom-right cell starting from 1 for columns.
}
```
* Usage:
``` typescript
import { getCellRange } from "ts-exceljs-utility";

const range1 = getCellRange("A1:B2");
// range1 = {rangeStr: "A1:B2", top: 1, left: 1, bottom: 2, right: 2}

const range2 = getCellRange({top: 1, left: 1});
// range2 = {rangeStr: "A1:A1", top: 1, left: 1, bottom: 1, right: 1}

const range3 = getCellRange({top: 1, left: 1, bottom: 2, right: 2});
// range3 = {rangeStr: "A1:B2", top: 1, left: 1, bottom: 2, right: 2}
```

# Inserting images into specified cell ranges without changing the aspect ratio
<a href="https://app.archive-gp.com/ts-exceljs-utility/example2">demo</a>

## ```createAddImageParam```method
* Constructs arguments for ExcelJS's ```workbook.addImage``` from a URL (string) or Blob.
## ```addImageToCell```method
* Inserts an image from a URL (string) or Blob.
* Allows specifying the destination cell or cell range with ```CellAddress```.
* Inserts the image without changing its aspect ratio, enlarging or reducing it to fit, with a specified size.
* Requires specifying the image size.
* Adds a margin of 2px on all sides.

```typescript
import { createAddImageParam, addImageToCell} from "ts-exceljs-utility";

//■Create workbook and worksheet
const book = new ExcelJS.Workbook();
const sheet = book.addWorksheet("sample");

//■Explicitly set width and height for the rows and columns where images will be inserted
for(let i = 1; i <= 100; i++){
  sheet.getRow(i).height = 13.5;
  sheet.getColumn(i).width = 8.38;
}

//■Insert image into workbook
const imageId = book.addImage(await createAddImageParam(horizonImage));

//■Insert image into worksheet
//Insert an image with a size of 712x357 into "B1:K10".
addImageToCell(sheet, imageId, 712, 357, "B1:K10");
```
