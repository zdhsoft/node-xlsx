<!-- markdownlint-disable no-inline-html -->

# @zdhsoft/nodexlsx

<p align="center">
  <a href="https://www.npmjs.com/package/@zdhsoft/nodexlsx">
    <img src="https://img.shields.io/npm/v/@zdhsoft/nodexlsx?style=for-the-badge" alt="npm version" />
  </a>
  <a href="https://www.npmjs.com/package/@zdhsoft/nodexlsx">
    <img src="https://img.shields.io/npm/dt/@zdhsoft/nodexlsx?style=for-the-badge" alt="npm total downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@zdhsoft/nodexlsx">
    <img src="https://img.shields.io/npm/dm/@zdhsoft/nodexlsx?style=for-the-badge" alt="npm monthly downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@zdhsoft/nodexlsx">
    <img src="https://img.shields.io/npm/l/@zdhsoft/nodexlsx?style=for-the-badge" alt="npm license" />
  </a>
</p>

## 申明
- 原node-xlsx不支持样式，所以这里将复fork过来后，更新xlsx库和增加样式的功能。
- 原来默认使用的是xlsx这个包，现在使用xlsx-js-style替换它，但是这个增加了对样式的支持
- 增加了能单元格类型的导出。

## 变更记录
- v0.0.4
  - 修复没有导出原有xlsx-js-style接口函数的功能
  - 这个nodexlsx在原有的xlsx-js-style上面，增加了parse, parseMetadata, build三个函数
  - 增加了nodexlsx属性

## Features

Straightforward excel file parser and builder.

- Relies on [SheetJS xlsx](https://github.com/SheetJS/sheetjs) module to parse/build excel sheets.
- Built with [TypeScript](https://www.typescriptlang.org/) for static type checking with exported types along the
  library.

## Install

```bash
npm i @zdhsoft/nodexlsx
```

## Quickstart

### Parse an xlsx file

```js
import xlsx from '@zdhsoft/nodexlsx';
// Or var xlsx = require('@zdhsoft/nodexlsx').default;

// Parse a buffer
const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${__dirname}/myFile.xlsx`));
// Parse a file
const workSheetsFromFile = xlsx.parse(`${__dirname}/myFile.xlsx`);
```

### Build an xlsx file

```js
import xlsx from '@zdhsoft/nodexlsx';
// Or var xlsx = require('node-xlsx').default;

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
var buffer = xlsx.build([{name: 'mySheetName', data: data}]); // Returns a buffer
```

### Custom column width

```js
import xlsx from '@zdhsoft/nodexlsx';
// Or var xlsx = require('@zdhsoft/nodexlsx').default;

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
const sheetOptions = {'!cols': [{wch: 6}, {wch: 7}, {wch: 10}, {wch: 20}]};

var buffer = xlsx.build([{name: 'mySheetName', data: data}], {sheetOptions}); // Returns a buffer
```

### Spanning multiple rows `A1:A4` in every sheets

```js
import xlsx from '@zdhsoft/nodexlsx';
// Or var xlsx = require('@zdhsoft/nodexlsx').default;

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
const range = {s: {c: 0, r: 0}, e: {c: 0, r: 3}}; // A1:A4
const sheetOptions = {'!merges': [range]};

var buffer = xlsx.build([{name: 'mySheetName', data: data}], {sheetOptions}); // Returns a buffer
```

### Spanning multiple rows `A1:A4` in second sheet

```js
import xlsx from '@zdhsoft/nodexlsx';
// Or var xlsx = require('@zdhsoft/nodexlsx').default;

const dataSheet1 = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
const dataSheet2 = [
  [4, 5, 6],
  [7, 8, 9, 10],
  [11, 12, 13, 14],
  ['baz', null, 'qux'],
];
const range = {s: {c: 0, r: 0}, e: {c: 0, r: 3}}; // A1:A4
const sheetOptions = {'!merges': [range]};

var buffer = xlsx.build([
  {name: 'myFirstSheet', data: dataSheet1},
  {name: 'mySecondSheet', data: dataSheet2, options: sheetOptions},
]); // Returns a buffer
```

### 增加使用样式的功能示例
```typescript
// typescript的代码
import {TCellType, nodexlsx, CellStyle, WorkSheetOptions} from '@zdhsoft/nodexlsx';
import fs from 'fs';

function simpleExcel() {
    console.info('---->开始执行');
    const sheetOptions: WorkSheetOptions = {
        '!cols': [{wch: 6}, {wch: 30}, {wch: 10}, {wch: 15}, {wch: 10}, {wch: 10}],
        '!merges': [
            {s: {c: 0, r: 0}, e: {c: 2, r: 0}},
            {s: {c: 3, r: 0}, e: {c: 5, r: 0}},
        ],
    };
    const s: CellStyle = {
        alignment: {
            horizontal: 'center', // 水平居中
            vertical: 'center', // 垂直居中
        },
    };
    const contentCellStyle: CellStyle = {
        border: {
            top: {
                style: 'dashDot',
                color: {rgb: '00ff00'},
            },
            bottom: {
                style: 'medium',
                color: {rgb: '0000ff'},
            },
            left: {
                style: 'medium',
                color: {rgb: 'ff0000'},
            },
            right: {
                style: 'medium',
                color: {rgb: '770066'},
            },
        },
    };
    // 指定标题单元格样式：加粗居中
    const headerStyle: CellStyle = {
        font: {
            bold: true,
        },
        alignment: {
            horizontal: 'center',
        },
    };
    const title: TCellType[] = ['序号', '名称', '年级', {v: '任课老师', s}, '学生数量', '已报名数量'];
    const title1: TCellType[] = [{v: '天下无难事', s: headerStyle}, null, null, '右边', null];
    const value: TCellType[] = [1, 'test', 9, {v: '张老师', s: contentCellStyle}, 10, 99];

    const rows: TCellType[][] = [];
    rows.push(title1);
    rows.push(title);
    rows.push(value);
    const data = nodexlsx.build([{name: '社团列表', data: rows, options: sheetOptions}]);
    const outFileName = 'd:\\temp\\a.xlsx';
    fs.writeFileSync(outFileName, data);
    console.info('生成' + outFileName + ' ok!');
    return true;
}
```

_Beware that if you try to merge several times the same cell, your xlsx file will be seen as corrupted._

- Using Primitive Object Notation Data values can also be specified in a non-abstracted representation.

Examples:

```js
const rowAverage = [[{t: 'n', z: 10, f: '=AVERAGE(2:2)'}], [1, 2, 3]];
var buffer = xlsx.build([{name: 'Average Formula', data: rowAverage}]);
```

Refer to [xlsx](https://sheetjs.gitbooks.io) documentation for valid structure and values:

- [cell object]: (https://sheetjs.gitbooks.io/docs/#cell-object)
- [data types]: (https://sheetjs.gitbooks.io/docs/#data-types)
- [Format](https://sheetjs.gitbooks.io/docs/#number-formats)

### Troubleshooting

This library requires at least node.js v10. For legacy versions, you can use this workaround before using the lib.

```sh
npm i --save object-assign
```

```js
Object.prototype.assign = require('object-assign');
```

### Contributing

Please submit all pull requests the against master branch. If your unit test contains javascript patches or features,
you should include relevant unit tests. Thanks!

### Available scripts

| **Script**    | **Description**                |
| ------------- | ------------------------------ |
| start         | Alias of test:watch            |
| test          | Run mocha unit tests           |
| test:watch    | Run and watch mocha unit tests |
| lint          | Run eslint static tests        |
| compile       | Compile the library            |
| compile:watch | Compile and watch the library  |

## Authors

**Olivier Louvignes**

- http://olouv.com
- http://github.com/mgcrea
- http://github.com/zdhsoft
## Copyright and license

[Apache License 2.0](https://spdx.org/licenses/Apache-2.0.html)

```
Copyright (C) 2012-2014  Olivier Louvignes

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.

Except where noted, this license applies to any and all software programs and associated documentation files created by the Original Author and distributed with the Software:

Inspired by SheetJS gist examples, Copyright (c) SheetJS.
```
