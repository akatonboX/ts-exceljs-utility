import React from 'react';
import ExcelJS from "exceljs";
import { PageLayout } from '../layout/pageLayout';
import styles from "./example1Page.module.scss";
import ExcelJSUtility from "ts-exceljs-utility";

export function Example1Page(
  props: {
  }
) 
{
  const [input1, setInput1] = React.useState("A1:B2");
  const [result1, setResult1] = React.useState<ExcelJSUtility.CellRange | undefined>(undefined);
  const [input2Top, setInput2Top] = React.useState("1");
  const [input2Left, setInput2Left] = React.useState("1");
  const [result2, setResult2] = React.useState<ExcelJSUtility.CellRange | undefined>(undefined);
  const [input3Top, setInput3Top] = React.useState("1");
  const [input3Left, setInput3Left] = React.useState("1");
  const [input3Bottom, setInput3Bottom] = React.useState("1");
  const [input3Right, setInput3Right] = React.useState("1");
  const [result3, setResult3] = React.useState<ExcelJSUtility.CellRange | undefined>(undefined);
  return (
    <PageLayout title="Example1">
      <div className={styles.root}>
        <input onChange={e => {setInput1(e.target.value);}} value={input1} />
        <button onClick={e => {
          const range = ExcelJSUtility.getCellRange(input1);
          setResult1(range);
        }}>getCellRange</button>
        {"->"}
        top: {result1?.top ?? "?"}, left: {result1?.left ?? "?"}, bottom: {result1?.bottom ?? "?"}, right: {result1?.right ?? "?"}, rangeStr: {result1?.rangeStr ?? "?"}
        <hr />

        top:<input onChange={e => {setInput2Top(e.target.value);}} value={input2Top} />
        left:<input onChange={e => {setInput2Left(e.target.value);}} value={input2Left} />
        <button onClick={e => {
          const range = ExcelJSUtility.getCellRange({
            top: Number(input2Top), 
            left: Number(input2Left)
          });
          setResult2(range);
        }}>getCellRange</button>
        {"->"}
        top: {result2?.top ?? "?"}, left: {result2?.left ?? "?"}, bottom: {result2?.bottom ?? "?"}, right: {result2?.right ?? "?"}, rangeStr: {result2?.rangeStr ?? "?"}
        <hr />

        top:<input onChange={e => {setInput3Top(e.target.value);}} value={input3Top}/>
        left:<input onChange={e => {setInput3Left(e.target.value);}} value={input3Left} />
        bottom:<input onChange={e => {setInput3Bottom(e.target.value);}} value={input3Bottom} />
        left:<input onChange={e => {setInput3Right(e.target.value);}} value={input3Right} />
        <button onClick={e => {
          const range = ExcelJSUtility.getCellRange({
            top: Number(input3Top), 
            left: Number(input3Left),
            bottom: Number(input3Bottom), 
            right: Number(input3Right),
          });
          setResult3(range);
        }}>getCellRange</button>
        {"->"}
        top: {result3?.top ?? "?"}, left: {result3?.left ?? "?"}, bottom: {result3?.bottom ?? "?"}, right: {result3?.right ?? "?"}, rangeStr: {result3?.rangeStr ?? "?"}
        <hr />
      </div>
    </PageLayout>
  );
}