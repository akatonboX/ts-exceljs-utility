# ts-react-listview
* react用のListViewコンポーネントです。
* 固定の行ヘッダ、列ヘッダ、列の幅の変更をサポートします。


# Note
* クライアントは、Windowsを想定しています。
* chromeで動作確認されています。
* cssに、-webkit-scrollbarを使用しています。
# Installation
```shell
npm install ts-react-listview
```

# release
* [2024/3/17]v1.0.0 released 

# 簡単な例
<a href="https://app.archive-gp.com/ts-react-listview/example1">demo</a>

``` typescript
import React from 'react';
import { PageLayout } from '../layout/pageLayout';
import styles from "./example1Page.module.scss";
import { ListView, ListViewItem, ListViewRow } from 'ts-react-listview';
export function Example1Page(
  props: {
  }
) 
{
  const data : {a: string, b: string, c: string}[] = new Array(100).fill(1).map((item, index) => {
    return {
      a: `A${index}`,
      b: `B${index}`,
      c: `C${index}`,
    };
  });
  return (
    <PageLayout title="Example1">
      <div>
        <ListView headers={[
          {
            name: "a",
            label: <div className={styles.listViewCloumn}>A</div>,
            defaultWidth: 70,
          },
          {
            name: "b",
            label: <div className={styles.listViewCloumn}>B</div>,
            defaultWidth: 100,
          },
          {
            name: "c",
            label: <div className={styles.listViewCloumn}>C</div>,
            defaultWidth: 120,
          },
        ]}>
          {(data ?? []).map((item, index) => (
            <ListViewRow key={index} >
              <ListViewItem name="a">
                <div className={styles.listViewCell}>{item.a}</div>
              </ListViewItem>
              <ListViewItem name="b">
                <div className={styles.listViewCell}>{item.b}</div>
              </ListViewItem>
              <ListViewItem name="c">
                <div className={styles.listViewCell}>{item.c}</div>
              </ListViewItem>
            </ListViewRow>
          ))}
        </ListView>
      </div>
    </PageLayout>
  );
}
```

# 行ヘッダを表示する
<a href="https://app.archive-gp.com/ts-react-listview/example2">demo</a>
* ```hasRowHeader={true}```に設定すると、行ヘッダが表示されます。
* 行ヘッダを表示した場合、行の幅と、高さを指定する必要があります。```rowHeaderWidth={300} rowHeight={100} ```
* 
``` typescript
import React from 'react';
import { PageLayout } from '../layout/pageLayout';
import styles from "./example2Page.module.scss";
import { ListView, ListViewItem, ListViewRow } from 'ts-react-listview';
export function Example2Page(
  props: {
  }
) 
{
  const data : {a: string, b: string, c: string}[] = new Array(100).fill(1).map((item, index) => {
    return {
      a: `A${index}`,
      b: `B${index}`,
      c: `C${index}`,
    };
  });
  return (
    <PageLayout title="Example1">
      <div>
        <ListView hasRowHeader={true} rowHeaderWidth={300} rowHeight={100}headers={[
          {
            name: "a",
            label: <div className={styles.listViewCloumn}>A</div>,
            defaultWidth: 70,
          },
          {
            name: "b",
            label: <div className={styles.listViewCloumn}>B</div>,
            defaultWidth: 100,
          },
          {
            name: "c",
            label: <div className={styles.listViewCloumn}>C</div>,
            defaultWidth: 120,
          },
        ]}>
          {(data ?? []).map((item, index) => (
            <ListViewRow key={index}  header={<div style={{height: "100%", display: "flex", justifyContent: "center", alignItems: "center"}}>{index}</div>} >
              <ListViewItem name="a">
                <div className={styles.listViewCell}>{item.a}</div>
              </ListViewItem>
              <ListViewItem name="b">
                <div className={styles.listViewCell}>{item.b}</div>
              </ListViewItem>
              <ListViewItem name="c">
                <div className={styles.listViewCell}>{item.c}</div>
              </ListViewItem>
            </ListViewRow>
          ))}
        </ListView>
      </div>
    </PageLayout>
  );
}
```

# 罫線を変更する
<a href="https://app.archive-gp.com/ts-react-listview/example3">demo</a>
* ```horizontalLineStyle="1px solid black" verticvalLineStyle="1px solid black"```のように、罫線のスタイルを指定できます。

# 列の幅が変更されたときの処理
<a href="https://app.archive-gp.com/ts-react-listview/example3">demo</a>

* ```onResizeColumn```プロパティを使用します。
* 下記は、ローカルストレージで変更された列の幅を保持し、リロードで適用する例です。
```typescript
 <ListView horizontalLineStyle="1px solid black" verticvalLineStyle="1px solid black" onResizeColumn={(name, width) => {
          localStorage.setItem(`Example3Page_${name}`, String(width));
        }} headers={[
          {
            name: "a",
            label: <div className={styles.listViewCloumn}>A</div>,
            defaultWidth: Number(localStorage.getItem(`Example3Page_a`) ?? 100),
          },
          {
            name: "b",
            label: <div className={styles.listViewCloumn}>B</div>,
            defaultWidth: Number(localStorage.getItem(`Example3Page_b`) ?? 100),
          },
          {
            name: "c",
            label: <div className={styles.listViewCloumn}>C</div>,
            defaultWidth: Number(localStorage.getItem(`Example3Page_c`) ?? 100),
          },
        ]}>
```
# カラムサイズの変更を禁止する
<a href="https://app.archive-gp.com/ts-react-listview/example4">demo</a>
* ```resizeColumnEnabled={false}```を設定することで、列の幅を変更することができなくなります。

# 大きさ
<a href="https://app.archive-gp.com/ts-react-listview/example5">demo</a>
* ListViewは、それを配置した要素のサイズになります。