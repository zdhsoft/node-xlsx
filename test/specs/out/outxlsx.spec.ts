import {TCellType, nodexlsx, CellStyle} from '../../../src';
import fs from 'fs';

function simpleExcel() {
    console.info('---->开始执行');
    const sheetOptions = {
        '!cols': [{wch: 6}, {wch: 30}, {wch: 10}, {wch: 15}, {wch: 10}, {wch: 10}],
        '!merges': [
            {s: {c: 0, r: 0}, e: {c: 2, r: 0}},
            {s: {c: 3, r: 0}, e: {c: 5, r: 0}},
        ],
    };
    const s = {
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
    const title: TCellType[] = ['序号', '名称', '年级', '任课老师', '学生数量', '已报名数量'];
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
    // process.exit(0);
}

describe('测试样式', () => {
    it('测试基本的样式生成', () => {
        // eslint-disable-next-line jest/valid-expect
        expect(() => simpleExcel());
    });
});
