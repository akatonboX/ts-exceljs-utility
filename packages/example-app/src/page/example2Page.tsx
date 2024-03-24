import React from 'react';
import ExcelJS from "exceljs";
import { PageLayout } from '../layout/pageLayout';
import styles from "./example2Page.module.scss";
import ExcelJSUtility, { addImageToCell, createAddImageParam, downloadWorkbook } from "ts-exceljs-utility";
import horizonImage from "./example2Page.horizon.png";
import verticalImage from "./example2Page.vertical.png";

export function Example2Page(
  props: {
  }
) 
{
  const [input1, setInput1] = React.useState("A1:B2");
  const [input2, setInput2] = React.useState("A1:B2");
  return (
    <PageLayout title="Example2(addImageToCell)">
      <div className={styles.root}>
        <div>
          <div>
            <img src={horizonImage} style={{width: 200}}/>
            <span>
              (712x357)
            </span>
          </div>
          <div>
            <input value={input1} onChange={e => {setInput1(e.target.value);}}/>
            <button onClick={async e => {
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
              addImageToCell(sheet, imageId, 712, 357, input1);
              
              //■ダウンロードさせる
              downloadWorkbook(book, `horizon(${input1.replace(":", "：")}).xlsx`);
            }}>set image and download</button>
          </div>
        </div>
        <div>
          <div>
            <img src={verticalImage}  style={{height: 200}} />
            <span>
              (357x712)
            </span>
          </div>
          <div>
            <input value={input2} onChange={e => {setInput2(e.target.value);}}/>
            <button onClick={async e => {
              //■workbookとworkshettを作成
              const book = new ExcelJS.Workbook();
              const sheet = book.addWorksheet("sample");

              //■画像の挿入対象の行列には、明示的に幅と高さを設定
              for(let i = 1; i <= 100; i++){
                sheet.getRow(i).height = 13.5;
                sheet.getColumn(i).width = 8.38;
              }

              //■workbookにimageを挿入
              const imageId = book.addImage(await createAddImageParam(verticalImage));

              //■worksheetにimageを挿入
              addImageToCell(sheet, imageId, 357, 712, input2);
              
              //■ダウンロードさせる
              downloadWorkbook(book, `vertical(${input2.replace(":", "：")}).xlsx`);
            }}>set image and download</button>
          </div>
        </div>
      </div>
    </PageLayout>
  );
}