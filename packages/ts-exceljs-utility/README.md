# ts-react-listview
* ListView component for react.
* Supports fixed row headers, column headers, and changing column widths.


# Note
* The client is assumed to be Windows.
* Confirmed to work with chrome.
* -webkit-scrollbar is used for css.
# Installation
```shell
yarn install ts-react-listview
```

# release
* [2024/3/17]v1.0.0 released 

# simple example
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

# Show row headers
<a href="https://app.archive-gp.com/ts-react-listview/example2">demo</a>
* Set ```hasRowHeader={true}``` to display row headers.
* If you display row headers, you need to specify the row width and height. ```rowHeaderWidth={300} rowHeight={100} ```

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

# Change borders
<a href="https://app.archive-gp.com/ts-react-listview/example3">demo</a>
* You can specify the border style, such as ```horizontalLineStyle="1px solid black" verticvalLineStyle="1px solid black"```.

# What happens when column width changes
<a href="https://app.archive-gp.com/ts-react-listview/example3">demo</a>

* Use the ```onResizeColumn``` property.
* Below is an example of retaining the changed column width in local storage and applying it on reload.
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
# Prohibit changing column size
<a href="https://app.archive-gp.com/ts-react-listview/example4">demo</a>
* By setting ```resizeColumnEnabled={false}```, you will not be able to change the column width.

# Size
<a href="https://app.archive-gp.com/ts-react-listview/example5" target=">demo</a>
* The ListView will be the size of the element you place it on.