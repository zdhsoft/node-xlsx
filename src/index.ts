import {
    AOA2SheetOpts,
    AutoFilterInfo,
    ColInfo,
    ParsingOptions,
    ProtectInfo,
    Range,
    read,
    readFile,
    RowInfo,
    Sheet2JSONOpts,
    utils,
    write,
    WritingOptions,
    ExcelDataType,
    Comments,
    NumberFormat,
    Hyperlink,
    CellStyle,
} from 'xlsx-js-style';
import {isString} from './helpers';
import {WorkBook} from './workbook';

export * from 'xlsx-js-style';

export const parse = (mixed: unknown, options: Sheet2JSONOpts & ParsingOptions = {}) => {
    const {dateNF, header = 1, range, blankrows, defval, raw = true, rawNumbers, ...otherOptions} = options;
    const workBook = isString(mixed)
        ? readFile(mixed, {dateNF, raw, ...otherOptions})
        : read(mixed, {dateNF, raw, ...otherOptions});
    return Object.keys(workBook.Sheets).map((name) => {
        const sheet = workBook.Sheets[name];
        return {
            name,
            data: utils.sheet_to_json(sheet, {
                dateNF,
                header,
                range: typeof range === 'function' ? range(sheet) : range,
                blankrows,
                defval,
                raw,
                rawNumbers,
            }),
        };
    });
};

export const parseMetadata = (mixed: unknown, options: ParsingOptions = {}) => {
    const workBook = isString(mixed) ? readFile(mixed, options) : read(mixed, options);
    return Object.keys(workBook.Sheets).map((name) => {
        const sheet = workBook.Sheets[name];
        return {name, data: sheet['!ref'] ? utils.decode_range(sheet['!ref']) : null};
    });
};

export type WorkSheetOptions = {
    /** Column Info */
    '!cols'?: ColInfo[];

    /** Row Info */
    '!rows'?: RowInfo[];

    /** Merge Ranges */
    '!merges'?: Range[];

    /** Worksheet Protection info */
    '!protect'?: ProtectInfo;

    /** AutoFilter info */
    '!autofilter'?: AutoFilterInfo;
};

export type CellBaseType = string | number | boolean | Date;

/** 单元对齐方式 */
export interface ICellAlignment {
    /** 垂直对齐，默认: bottom */
    vertical?: 'top' | 'center' | 'bottom';
    /** 水平对齐：默认: left */
    horizontal?: 'left' | 'right' | 'center';
    /** 自动换行：默认false */
    wrapText?: boolean;
}
export interface ICellObject {
    /** 单元格原始值.  如果指定了公式，则可以省略 */
    v?: CellBaseType;

    /** 格式化文本 */
    w?: string;

    /**
     * 格式元的数据格式.
     * b Boolean, n Number, e Error, s String, d Date, z Empty
     */
    t?: ExcelDataType;

    /** 单元格公式 */
    f?: string;

    /** 如果公式是数组公式，则封闭数组的范围 */
    F?: string;

    /** 使用HTML格式的富文本 (if applicable) */
    h?: string;

    /** 单元格注释 */
    c?: Comments;

    /** Number format string associated with the cell (if requested) */
    z?: NumberFormat;

    /** Cell hyperlink object (.Target holds link, .tooltip is tooltip) */
    l?: Hyperlink;

    /** The style/theme of the cell (if applicable) */
    s?: CellStyle;
}

export type TCellType = CellBaseType | ICellObject | null;

export type WorkSheet = {
    name: string;
    data: TCellType[][];
    options: WorkSheetOptions;
};

export type BuildOptions = WorkSheetOptions & {
    parseOptions?: AOA2SheetOpts;
    writeOptions?: WritingOptions;
    sheetOptions?: WorkSheetOptions;
};

export const build = (
    worksheets: WorkSheet[],
    {parseOptions = {}, writeOptions = {}, sheetOptions = {}, ...otherOptions}: BuildOptions = {}
): Buffer => {
    const {bookType = 'xlsx', bookSST = false, type = 'buffer', ...otherWriteOptions} = writeOptions;
    const legacyOptions = Object.keys(otherOptions).filter((key) => {
        if (['!cols', '!rows', '!merges', '!protect', '!autofilter'].includes(key)) {
            console.debug(`Deprecated options['${key}'], please use options.sheetOptions['${key}'] instead.`);
            return true;
        }
        console.debug(`Unknown options['${key}'], please use options.parseOptions / options.writeOptions`);
        return false;
    });
    const workBook = worksheets.reduce<WorkBook>((soFar, {name, data, options = {}}, index) => {
        const sheetName = name || `Sheet_${index}`;
        const sheetData = utils.aoa_to_sheet(data, parseOptions);
        soFar.SheetNames.push(sheetName);
        soFar.Sheets[sheetName] = sheetData;
        Object.assign(soFar.Sheets[sheetName], legacyOptions, sheetOptions, options);
        return soFar;
    }, new WorkBook());
    return write(workBook, {bookType, bookSST, type, ...otherWriteOptions});
};

export const nodexlsx = {parse, parseMetadata, build};
export default nodexlsx;
