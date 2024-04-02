import React from 'react';
import ExcelJS from "exceljs";
import { PageLayout } from '../layout/pageLayout';
import styles from "./example3Page.module.scss";
import { downloadWorkbook,  loadWorkbook } from 'ts-exceljs-utility';
import assetXlsx from "../assets/example.xlsx";


export function Example3Page(
  props: {
  }
) 
{
  return (
    <PageLayout title="Example3(assets)">
      <div className={styles.root}>
        <button onClick={async e => {
          const workbook = await loadWorkbook(assetXlsx);
          if(workbook != null){
            //■ダウンロードさせる
            downloadWorkbook(workbook, "example.xlsx");
          }

        }}>download asset xlsx file</button>
      </div>
    </PageLayout>
  );
}