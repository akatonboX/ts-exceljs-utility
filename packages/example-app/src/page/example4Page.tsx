import React from 'react';
import ExcelJS from "exceljs";
import { PageLayout } from '../layout/pageLayout';
import styles from "./example4Page.module.scss";
import { downloadWorkbook, loadWorkbook } from 'ts-exceljs-utility';
import assetXlsx from "../assets/example.xlsx";
import { inputLabelClasses } from '@mui/material';
import { addColBreak } from 'ts-exceljs-utility';


export function Example4Page(
  props: {
  }
) 
{
  const [input, setInput] = React.useState("5, 10, 15");

  return (
    <PageLayout title="Example4(Column Breaks)">
      <div className={styles.root}>
        <input value={input} onChange={e => {setInput(e.target.value);}} />

        <button onClick={async e => {
          //■リソースからworkbookをロード
          const workbook = await loadWorkbook(assetXlsx);
          if(workbook == null)return;

          //■入力から改行する列の配列に変換
          const breaks = input.split(",").map(item => item.trim()).filter(item => item.length > 0).map(item => Number(item)).filter(item => !isNaN(item));

          //■一つ目のシートに列方向の改行を追加
          const sheet = workbook.worksheets[0];
          breaks.forEach(colBreak => {
            addColBreak(sheet, colBreak);
          });

          //■ダウンロードさせる
          downloadWorkbook(workbook, "example.xlsx");
        }}>download</button>
      </div>
    </PageLayout>
  );
}