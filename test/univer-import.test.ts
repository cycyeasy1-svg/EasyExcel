// M1 导入层单测：ExcelJS → IWorkbookData 的保真断言。
// node 环境（jsdom 会干扰 exceljs 内部 jszip 的环境探测）；无 DOMParser 时
// theme 解析回退 Office 默认调色板，与 openpyxl fixtures 的主题一致。
// theme1.xml 的 DOMParser 解析路径由 univer-theme.test.ts（jsdom）覆盖。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { BorderStyleTypes, type IStyleData, type ICellData, type IWorkbookData } from '@univerjs/core';
import { loadForUniver } from '../src/react/view/excel/univer/loader';
import { applyTint, parseThemePalette, resolveExcelColor, OFFICE_DEFAULT_THEME } from '../src/react/view/excel/univer/theme_colors';
import { dateToExcelSerial } from '../src/react/view/excel/univer/import';

const load = (name: string) => {
  const buf = readFileSync(`test/fixtures/${name}`);
  return loadForUniver(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), name.split('.').pop()!, name);
};

const firstSheet = (wb: IWorkbookData) => wb.sheets[wb.sheetOrder[0]];

const styleOf = (wb: IWorkbookData, cell: ICellData | undefined): IStyleData | undefined => {
  if (!cell?.s) return undefined;
  return typeof cell.s === 'string' ? (wb.styles[cell.s] as IStyleData) : (cell.s as IStyleData);
};

describe('theme_colors', () => {
  it('applyTint matches OOXML behavior', () => {
    expect(applyTint('4472C4', 0)).toBe('4472C4');
    // 提亮后每通道不小于原值，且与原色不同
    const lighter = applyTint('4472C4', 0.4);
    expect(lighter).not.toBe('4472C4');
    const darker = applyTint('4472C4', -0.25);
    expect(parseInt(darker.slice(0, 2), 16)).toBeLessThan(0x44 + 1);
  });

  it('parseThemePalette falls back to Office default on garbage', () => {
    expect(parseThemePalette(undefined)).toEqual(OFFICE_DEFAULT_THEME);
    expect(parseThemePalette('<not-xml')).toEqual(OFFICE_DEFAULT_THEME);
  });

  it('resolveExcelColor handles argb/theme/indexed', () => {
    const palette = OFFICE_DEFAULT_THEME;
    expect(resolveExcelColor({ argb: 'FF112233' }, palette)).toBe('#112233');
    expect(resolveExcelColor({ theme: 4 }, palette)).toBe('#4472c4');
    expect(resolveExcelColor({ indexed: 2 }, palette)).toBe('#ff0000');
    expect(resolveExcelColor(undefined, palette)).toBeUndefined();
  });

  it('dateToExcelSerial anchors correctly', () => {
    // 2023-06-15 = 45092 (Excel 1900 体系)
    expect(dateToExcelSerial(new Date(Date.UTC(2023, 5, 15)))).toBe(45092);
  });
});

describe('import: theme-colors.xlsx', () => {
  it('resolves theme fonts, fills, tints, indexed and keeps white fill', async () => {
    const { workbookData } = await load('theme-colors.xlsx');
    const sheet = firstSheet(workbookData);

    const fontColor = styleOf(workbookData, sheet.cellData[0][0])?.cl?.rgb?.toLowerCase();
    expect(fontColor).toBe('#4472c4'); // accent1

    const fill = styleOf(workbookData, sheet.cellData[0][1])?.bg?.rgb?.toLowerCase();
    expect(fill).toBe('#4472c4');

    const tinted = styleOf(workbookData, sheet.cellData[7][0])?.cl?.rgb?.toLowerCase();
    expect(tinted).toBeTruthy();
    expect(tinted).not.toBe('#4472c4'); // tint 0.4 应与基色不同

    const indexed = styleOf(workbookData, sheet.cellData[10][0])?.cl?.rgb?.toLowerCase();
    expect(indexed).toBe('#008000'); // indexed 17

    const whiteFill = styleOf(workbookData, sheet.cellData[11][0])?.bg?.rgb?.toLowerCase();
    expect(whiteFill).toBe('#ffffff'); // 旧 reader 会丢弃
  });
});

describe('import: numfmt-exotic.xlsx', () => {
  it('passes number format patterns through verbatim', async () => {
    const { workbookData } = await load('numfmt-exotic.xlsx');
    const sheet = firstSheet(workbookData);
    for (let ri = 0; ri < 10; ri += 1) {
      const expected = sheet.cellData[ri][1].v as string;
      const pattern = styleOf(workbookData, sheet.cellData[ri][0])?.n?.pattern;
      expect(pattern, `row ${ri + 1}`).toBe(expected);
    }
    // 数字保留为 number 类型
    expect(typeof sheet.cellData[0][0].v).toBe('number');
  });
});

describe('import: richtext.xlsx', () => {
  it('converts rich text to IDocumentData with runs', async () => {
    const { workbookData } = await load('richtext.xlsx');
    const sheet = firstSheet(workbookData);
    const cell = sheet.cellData[0][0];
    expect(cell.v).toBe('Bold red plain big italic');
    expect(cell.p).toBeTruthy();
    const body = (cell.p as { body: { dataStream: string; textRuns: { st: number; ed: number; ts: IStyleData }[] } }).body;
    expect(body.dataStream).toBe('Bold red plain big italic\r\n');
    expect(body.textRuns.length).toBeGreaterThanOrEqual(2);
    const boldRun = body.textRuns.find(r => r.st === 0);
    expect(boldRun?.ts.bl).toBe(1);
    expect(boldRun?.ts.cl?.rgb?.toLowerCase()).toBe('#ff0000');
    // 普通文本不带 p
    expect(sheet.cellData[1][0].p).toBeUndefined();
  });
});

describe('import: borders.xlsx', () => {
  it('maps all 13 border styles distinctly', async () => {
    const { workbookData } = await load('borders.xlsx');
    const sheet = firstSheet(workbookData);
    const seen = new Set<BorderStyleTypes>();
    for (const row of Object.values(sheet.cellData)) {
      for (const cell of Object.values(row as Record<number, ICellData>)) {
        const bd = styleOf(workbookData, cell)?.bd;
        if (bd?.t) seen.add(bd.t.s);
      }
    }
    expect(seen.size).toBe(13);
    const borderCell = Object.values(sheet.cellData)[0];
    const bd = styleOf(workbookData, Object.values(borderCell as Record<number, ICellData>)[0])?.bd;
    expect(bd?.t?.cl.rgb?.toLowerCase()).toBe('#3366cc');
  });
});

describe('import: freeze-merge.xlsx', () => {
  it('imports freeze, merges, row heights, col widths, multi-sheet', async () => {
    const { workbookData } = await load('freeze-merge.xlsx');
    expect(workbookData.sheetOrder.length).toBe(2);
    const sheet = firstSheet(workbookData);

    expect(sheet.freeze).toMatchObject({ xSplit: 2, ySplit: 2, startRow: 2, startColumn: 2 });

    expect(sheet.mergeData).toContainEqual({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
    expect(sheet.mergeData).toContainEqual({ startRow: 3, startColumn: 3, endRow: 3, endColumn: 5 });

    expect(sheet.rowData[4]?.h).toBe(Math.round(40 * 96 / 72));
    expect(sheet.columnData[3]?.w).toBe(Math.round(30 * 7 + 5));

    const second = workbookData.sheets[workbookData.sheetOrder[1]];
    expect(second.name).toBe('Second');
    expect(second.cellData[0][0].v).toBe('second sheet');
  });
});

describe('import: chart-bearing file (sanitize)', () => {
  it('loads r1-chart.xlsx without crashing and reports lossy features', async () => {
    const { workbookData, lossy } = await load('r1-chart.xlsx');
    expect(lossy?.charts).toBe(true);
    const sheet = firstSheet(workbookData);
    expect(sheet.cellData[1][0].v).toBe('Jan'); // 数据完好
    expect(sheet.cellData[1][1].v).toBe(10); // 数字类型保留
    // 第二个 sheet（Untouched）也完好
    const second = workbookData.sheets[workbookData.sheetOrder[1]];
    expect(second.cellData[0][0].v).toBe('this sheet is never edited');
  });

  it('no-chart file reports no lossy charts', async () => {
    const { lossy } = await load('r1-features.xlsx');
    expect(lossy?.charts).toBe(false);
    expect(lossy?.vba).toBe(false);
  });
});

describe('import: csv path', () => {
  it('parses csv into workbook data', async () => {
    const csv = 'name,score\nalice,90\nbob,85\n';
    const buf = new TextEncoder().encode(csv).buffer;
    const { workbookData, csvDelimiter } = await loadForUniver(buf as ArrayBuffer, 'csv', 'test.csv');
    const sheet = firstSheet(workbookData);
    expect(csvDelimiter).toBe(',');
    expect(sheet.cellData[0][0].v).toBe('name');
    expect(sheet.cellData[1][1].v).toBe(90); // 数字识别
  });
});
