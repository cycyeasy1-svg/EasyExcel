// M2 验收测试：增量导出的往返保真。
// 核心断言：「只改 A1」后保存，批注 / 条件格式 / definedNames / 自定义
// numFmt / 未编辑 sheet 全部原样保留 —— 旧引擎（从零重建）全部丢失。
import { describe, it, expect } from 'vitest';
import { readFileSync as readFileSyncNode } from 'node:fs';
import JSZip from 'jszip';
import ExcelJS from '@cweijan/exceljs';
import type { ICellData, IWorkbookData } from '@univerjs/core';
import { loadForUniver } from '../src/react/view/excel/univer/loader';
import { diffWorkbook, hasFormattingChangedUniver } from '../src/react/view/excel/univer/diff';
import { applyDiffToWorkbook } from '../src/react/view/excel/univer/apply';

const readFixture = (name: string) => {
  const buf = readFileSyncNode(`test/fixtures/${name}`);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

const listParts = async (buf: Uint8Array) => {
  const zip = await JSZip.loadAsync(buf);
  return Object.keys(zip.files).filter(n => !zip.files[n].dir).sort();
};

const readPart = async (buf: Uint8Array, name: string) => {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(name);
  return f ? await f.async('string') : null;
};

async function loadAndEdit(fixture: string, edit: (current: IWorkbookData) => void, structural: string[] = []) {
  const result = await loadForUniver(readFixture(fixture), 'xlsx', fixture);
  const baseline = clone(result.workbookData);
  const current = clone(baseline);
  edit(current);
  const diff = diffWorkbook(baseline, current, new Set(structural));
  applyDiffToWorkbook(result.originalWorkbook!, current, diff, { sheetIdMap: result.sheetIdMap! });
  const out = new Uint8Array(await result.originalWorkbook!.xlsx.writeBuffer());
  return { baseline, current, diff, out };
}

describe('M2 round-trip: edit A1 only', () => {
  it('preserves comments, CF, definedNames, custom numFmt, untouched sheet', async () => {
    const { diff, out } = await loadAndEdit('r1-features.xlsx', (current) => {
      const sheet = current.sheets[current.sheetOrder[0]];
      (sheet.cellData as Record<number, Record<number, ICellData>>)[0][0] = { v: 'CHANGED' };
    });

    expect(diff.isEmpty).toBe(false);
    expect(diff.sheets[0].cellChanges).toHaveLength(1);
    expect(diff.sheets[0].cellChanges[0]).toMatchObject({ row: 0, col: 0 });

    // 批注部件仍在
    const parts = await listParts(out);
    expect(parts).toContain('xl/comments1.xml');

    // definedName / 条件格式 / 自定义 numFmt 原样保留
    expect(await readPart(out, 'xl/workbook.xml')).toMatch(/SalesRange/);
    expect(await readPart(out, 'xl/worksheets/sheet1.xml')).toMatch(/conditionalFormatting/);
    expect(await readPart(out, 'xl/styles.xml')).toMatch(/0\.00(&quot;|")m(&quot;|")/);

    // 修改生效 + 未触碰内容不变
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    expect(wb.worksheets[0].getCell('A1').value).toBe('CHANGED');
    expect(wb.worksheets[0].getCell('A2').value).toBe('Jan');
    expect(wb.worksheets[1].getCell('A1').value).toBe('this sheet is never edited');
    // 批注对象级验证
    expect(wb.worksheets[0].getCell('A1').note).toBeTruthy();
  });

  it('zero-edit diff is empty and save keeps all parts', async () => {
    const { diff, out } = await loadAndEdit('r1-features.xlsx', () => { });
    expect(diff.isEmpty).toBe(true);
    const parts = await listParts(out);
    expect(parts).toContain('xl/comments1.xml');
    expect(await readPart(out, 'xl/worksheets/sheet1.xml')).toMatch(/conditionalFormatting/);
  });
});

describe('M2 diff granularity', () => {
  it('detects style-only change without value change', async () => {
    const { diff } = await loadAndEdit('r1-features.xlsx', (current) => {
      const sheet = current.sheets[current.sheetOrder[0]];
      const cell = (sheet.cellData as Record<number, Record<number, ICellData>>)[1][0];
      cell.s = { bl: 1 } as never; // 加粗 A2
    });
    const change = diff.sheets[0].cellChanges[0];
    expect(change.valueChanged).toBe(false);
    expect(change.styleChanged).toBe(true);
  });

  it('detects cleared cell', async () => {
    const { diff, out } = await loadAndEdit('r1-features.xlsx', (current) => {
      const sheet = current.sheets[current.sheetOrder[0]];
      delete (sheet.cellData as Record<number, Record<number, ICellData>>)[1][0];
    });
    expect(diff.sheets[0].cellChanges[0].cell).toBeNull();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    expect(wb.worksheets[0].getCell('A2').value).toBeNull();
  });

  it('detects rename, merge add and row height change', async () => {
    const { diff, out } = await loadAndEdit('freeze-merge.xlsx', (current) => {
      const sheetId = current.sheetOrder[0];
      const sheet = current.sheets[sheetId];
      sheet.name = 'Renamed';
      (sheet.mergeData as unknown[]).push({ startRow: 7, startColumn: 0, endRow: 8, endColumn: 1 });
      (sheet.rowData as Record<number, { h?: number }>)[0] = { h: 60 };
    });
    const sd = diff.sheets[0];
    expect(sd.renamed).toBe(true);
    expect(sd.mergesAdded).toHaveLength(1);
    expect(sd.rowChanges).toContainEqual(expect.objectContaining({ index: 0 }));

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    expect(wb.worksheets[0].name).toBe('Renamed');
    // 原有合并 A1:B2 仍在（未触碰）
    expect((wb.worksheets[0].getCell('A1') as never as { isMerged: boolean }).isMerged).toBe(true);
    expect(Math.round(wb.worksheets[0].getRow(1).height)).toBe(45); // 60px -> 45pt
  });

  it('structural sheet is rebuilt with current data', async () => {
    const { diff, out } = await loadAndEdit('r1-features.xlsx', (current) => {
      const sheet = current.sheets[current.sheetOrder[0]];
      // 模拟「在第 1 行前插入一行」后的快照：所有行下移
      const cells = sheet.cellData as Record<number, Record<number, ICellData>>;
      const shifted: Record<number, Record<number, ICellData>> = {};
      for (const [r, row] of Object.entries(cells)) shifted[Number(r) + 1] = row as Record<number, ICellData>;
      shifted[0] = { 0: { v: 'inserted' } };
      sheet.cellData = shifted as never;
    }, ['sheet-1']);
    expect(diff.sheets[0].status).toBe('rebuilt');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    expect(wb.worksheets[0].getCell('A1').value).toBe('inserted');
    expect(wb.worksheets[0].getCell('A2').value).toBe('Month');
    // 未结构性编辑的第二个 sheet 完好
    expect(wb.worksheets[1].getCell('A1').value).toBe('this sheet is never edited');
  });

  it('added and removed sheets are applied', async () => {
    const { out } = await loadAndEdit('freeze-merge.xlsx', (current) => {
      // 删除 Second，加一个 NewSheet
      const removed = current.sheetOrder[1];
      current.sheetOrder = [current.sheetOrder[0], 'sheet-new'];
      delete current.sheets[removed];
      current.sheets['sheet-new'] = {
        id: 'sheet-new',
        name: 'NewSheet',
        rowCount: 40,
        columnCount: 26,
        cellData: { 0: { 0: { v: 'hello new' } } },
      } as never;
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    expect(wb.worksheets).toHaveLength(2);
    expect(wb.worksheets.map(w => w.name)).toContain('NewSheet');
    expect(wb.worksheets.map(w => w.name)).not.toContain('Second');
    const ns = wb.worksheets.find(w => w.name === 'NewSheet')!;
    expect(ns.getCell('A1').value).toBe('hello new');
  });
});

describe('M2 formatting-change detection (non-xlsx confirm)', () => {
  it('no change -> false; style change -> true', async () => {
    const result = await loadForUniver(readFixture('freeze-merge.xlsx'), 'xlsx', 'x');
    const baseline = clone(result.workbookData);
    const same = clone(baseline);
    expect(hasFormattingChangedUniver(baseline, same)).toBe(false);

    const changed = clone(baseline);
    const sheet = changed.sheets[changed.sheetOrder[0]];
    (sheet.cellData as Record<number, Record<number, ICellData>>)[2][2] = { v: 100, s: { bl: 1 } as never };
    expect(hasFormattingChangedUniver(baseline, changed)).toBe(true);
  });
});

describe('M2 numfmt preservation on value-only edit', () => {
  it('editing a value keeps the original exotic numFmt', async () => {
    const { out } = await loadAndEdit('numfmt-exotic.xlsx', (current) => {
      const sheet = current.sheets[current.sheetOrder[0]];
      const cell = (sheet.cellData as Record<number, Record<number, ICellData>>)[5][0];
      cell.v = 9.99; // 改 3.14159 -> 9.99，样式不动
    });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    const cell = wb.worksheets[0].getCell('A6');
    expect(cell.value).toBe(9.99);
    expect(cell.numFmt).toBe('0.00"m";[Red]-0.00"m"');
  });
});
