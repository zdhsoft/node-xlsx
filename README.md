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

## 申明(statement)
- 原node-xlsx不支持样式，所以这里将复fork过来后，更新xlsx库和增加样式的功能。(The original node-xlsx does not support styles, so after the fork, it will update the xlsx library and add styles.)
- 原来默认使用的是xlsx这个包，现在使用xlsx-js-style替换它，但是这个增加了对样式的支持(Originally, the xlsx package was used by default, and now it is replaced by xlsx-js-style, but this adds support for styles)
- 增加了能单元格类型的导出。(Added ability to export cell types.)

## 变更记录(change log)
- v0.1.0
  - readme增加style api和单元格的样式示例（Readme adds style api and cell style examples）
- v0.0.5
  - 增加所有xlsx-js-style中的type与interface的定义(Add the definition of type and interface in all xlsx-js-style)
- v0.0.4
  - 修复没有导出原有xlsx-js-style接口函数的功能(Fixed the function of not exporting the original xlsx-js-style interface function)
  - 这个nodexlsx在原有的xlsx-js-style上面，增加了parse, parseMetadata, build三个函数(his nodexlsx adds three functions: parse, parseMetadata, and build to the original xlsx-js-style)
  - 增加了nodexlsx属性(Added nodexlsx attribute)

## Features

Straightforward excel file parser and builder.

- Relies on [SheetJS xlsx](https://github.com/SheetJS/sheetjs) module to parse/build excel sheets.
- Built with [TypeScript](https://www.typescriptlang.org/) for static type checking with exported types along the
  library.

## Install

```bash
npm i @zdhsoft/nodexlsx
```

## Style API

### Cell Style Example

```js
// STEP 1: Create a new workbook
const wb = XLSX.utils.book_new();

// STEP 2: Create data rows and styles
let row = [
	{ v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
	{ v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
	{ v: "fill: color", t: "s", s: { fill: { fgColor: { rgb: "E9E9E9" } } } },
	{ v: "line\nbreak", t: "s", s: { alignment: { wrapText: true } } },
];

// STEP 3: Create worksheet with rows; Add worksheet to workbook
const ws = XLSX.utils.aoa_to_sheet([row]);
XLSX.utils.book_append_sheet(wb, ws, "readme demo");

// STEP 4: Write Excel file to browser
XLSX.writeFile(wb, "xlsx-js-style-demo.xlsx");
```

### Cell Style Properties

-   Cell styles are specified by a style object that roughly parallels the OpenXML structure.
-   Style properties currently supported are: `alignment`, `border`, `fill`, `font`, `numFmt`.

| Style Prop  | Sub Prop       | Default     | Description/Values                                                                                |
| :---------- | :------------- | :---------- | ------------------------------------------------------------------------------------------------- |
| `alignment` | `vertical`     | `bottom`    | `"top"` or `"center"` or `"bottom"`                                                               |
|             | `horizontal`   | `left`      | `"left"` or `"center"` or `"right"`                                                               |
|             | `wrapText`     | `false`     | `true` or `false`                                                                                 |
|             | `textRotation` | `0`         | `0` to `180`, or `255` // `180` is rotated down 180 degrees, `255` is special, aligned vertically |
| `border`    | `top`          |             | `{ style: BORDER_STYLE, color: COLOR_STYLE }`                                                     |
|             | `bottom`       |             | `{ style: BORDER_STYLE, color: COLOR_STYLE }`                                                     |
|             | `left`         |             | `{ style: BORDER_STYLE, color: COLOR_STYLE }`                                                     |
|             | `right`        |             | `{ style: BORDER_STYLE, color: COLOR_STYLE }`                                                     |
|             | `diagonal`     |             | `{ style: BORDER_STYLE, color: COLOR_STYLE, diagonalUp: true/false, diagonalDown: true/false }`   |
| `fill`      | `patternType`  | `"none"`    | `"solid"` or `"none"`                                                                             |
|             | `fgColor`      |             | foreground color: see `COLOR_STYLE`                                                               |
|             | `bgColor`      |             | background color: see `COLOR_STYLE`                                                               |
| `font`      | `bold`         | `false`     | font bold `true` or `false`                                                                       |
|             | `color`        |             | font color `COLOR_STYLE`                                                                          |
|             | `italic`       | `false`     | font italic `true` or `false`                                                                     |
|             | `name`         | `"Calibri"` | font name                                                                                         |
|             | `strike`       | `false`     | font strikethrough `true` or `false`                                                              |
|             | `sz`           | `"11"`      | font size (points)                                                                                |
|             | `underline`    | `false`     | font underline `true` or `false`                                                                  |
|             | `vertAlign`    |             | `"superscript"` or `"subscript"`                                                                  |
| `numFmt`    |                | `0`         | Ex: `"0"` // integer index to built in formats, see StyleBuilder.SSF property                     |
|             |                |             | Ex: `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF                          |
|             |                |             | Ex: `"0.0%"` // string specifying a custom format                                                 |
|             |                |             | Ex: `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|             |                |             | Ex: `"m/dd/yy"` // string a date format using Excel's format notation                             |

### `COLOR_STYLE` {object} Properties

Colors for `border`, `fill`, `font` are specified as an name/value object - use one of the following:

| Color Prop | Description       | Example                                                         |
| :--------- | ----------------- | --------------------------------------------------------------- |
| `rgb`      | hex RGB value     | `{rgb: "FFCC00"}`                                               |
| `theme`    | theme color index | `{theme: 4}` // (0-n) // Theme color index 4 ("Blue, Accent 1") |
| `tint`     | tint by percent   | `{theme: 1, tint: 0.4}` // ("Blue, Accent 1, Lighter 40%")      |

### `BORDER_STYLE` {string} Properties

Border style property is one of the following values:

-   `dashDotDot`
-   `dashDot`
-   `dashed`
-   `dotted`
-   `hair`
-   `mediumDashDotDot`
-   `mediumDashDot`
-   `mediumDashed`
-   `medium`
-   `slantDashDot`
-   `thick`
-   `thin`

**Border Notes**

Borders for merged areas are specified for each cell within the merged area. For example, to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:

-   left borders (for the three cells on the left)
-   right borders (for the cells on the right)
-   top borders (for the cells on the top)
-   bottom borders (for the cells on the left)

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
