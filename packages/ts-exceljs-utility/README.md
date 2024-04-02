# ts-exceljs-utility
* This is a package that gathers useful functionalities for using ExcelJS (https://www.npmjs.com/package/exceljs).
1. Mutual conversion between "A1:B1" and [row, col].
1. Inserting images into specified cell ranges without changing the aspect ratio of the images.
1. Allowing downloading of workbooks.
# Note
* Tested on Chrome from a React application created with CRA.
# Note
* Tested on Chrome from a React application created with CRA.

# Installation
```shell
yarn add ts-exceljs-utility
```

# release
* [2024/3/31]v1.0.1 released.
* [2024/3/24]v1.0.0 released.

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

# Using xlsx as an Asset[v1.0.1]
<a href="https://app.archive-gp.com/ts-exceljs-utility/example3">demo</a>
* Provides functionality to manage xlsx files as assets, similar to PNG images, and load them into workbooks.
* Assumes a React application created with CRA.

## Installing and Configuring Craco
* To use xlsx as an asset, you need to install Craco (https://www.npmjs.com/package/@craco/craco).
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
## Prepare xlsx.d.ts
```typescript
[xlsx.d.ts]
declare module '*.xlsx' {
  const content: string; 
  export default content;
}
```

## Store xlsx as an Asset in the ```src``` Directory
* Store, for example, ```src/assets/resource.xlxs```.

## Load Asset into Workbook
* Obtain the URL with ```import assetXlsx from "../assets/example.xlsx";``` and get the workbook using ```loadWorkbook```.
If the URL does not exist or the workbook fails to read the downloaded data, the return value of ```loadWorkbook``` will be undefined.
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

# Inserting colBreaks[v1.0.2]
<a href="https://app.archive-gp.com/ts-exceljs-utility/example4">demo</a>
* Use ```addColBreak(worksheet: ExcelJS.Worksheet, colIndex: number): void``` to add colBreaks (line breaks in columns) to the worksheet.
* ```colIndex``` represents the position of the column starting from 0.
* You can obtain an xlsx file that reflects colBreaks using the ``downloadWorkbook`` method or the ``writeWithColBreaks`` method.
```typescript
   //■Load workbook from resources
  const workbook = await loadWorkbook(assetXlsx);
  if(workbook == null) return;

  //■Convert input to an array of columns to break ([5,10,15])
  const breaks = input.split(",").map(item => item.trim()).filter(item => item.length > 0).map(item => Number(item)).filter(item => !isNaN(item));

  //■Add column breaks to the first sheet
  const sheet = workbook.worksheets[0];
  breaks.forEach(colBreak => {
    addColBreak(sheet, colBreak);
  });

  //■Trigger download
  downloadWorkbook(workbook, "example.xlsx");
  //const blob = await writeWithColBreaks(workbook);

```