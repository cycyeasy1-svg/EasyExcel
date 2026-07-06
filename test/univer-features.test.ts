// M3 单测：sheetFeatures 导入收集（DV/图片/保护）+ DV 导出重写 + 超链接写回。
import { describe, it, expect } from 'vitest';
import { readFileSync as readFileSyncNode } from 'node:fs';
import ExcelJS from '@cweijan/exceljs';
import type { ICellData, IWorkbookData } from '@univerjs/core';
import { loadForUniver } from '../src/react/view/excel/univer/loader';
import { diffWorkbook } from '../src/react/view/excel/univer/diff';
import { applyDiffToWorkbook } from '../src/react/view/excel/univer/apply';
import { parseRef } from '../src/react/view/excel/univer/features';

const readFixture = (name: string) => {
  const buf = readFileSyncNode(`test/fixtures/${name}`);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

describe('M3: parseRef', () => {
  it('parses cell and range refs', () => {
    expect(parseRef('A1')).toEqual({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 });
    expect(parseRef('B2:D5')).toEqual({ startRow: 1, startColumn: 1, endRow: 4, endColumn: 3 });
    expect(parseRef('$AA$10')).toEqual({ startRow: 9, startColumn: 26, endRow: 9, endColumn: 26 });
    expect(parseRef('not-a-ref')).toBeNull();
  });
});

describe('M3: sheetFeatures import collection', () => {
  it('collects validations, images and protection flags', async () => {
    const result = await loadForUniver(readFixture('m3-features.xlsx'), 'xlsx', 'm3');
    const features = result.sheetFeatures!;
    const first = features['sheet-1'];
    expect(first.protected).toBe(false);
    expect(first.validations.length).toBeGreaterThanOrEqual(2);
    const listDv = first.validations.find(v => v.type === 'list');
    expect(listDv?.value).toBe('red,green,blue');
    const numDv = first.validations.find(v => v.type === 'number');
    expect(numDv?.operator).toBe('be');
    expect(first.images).toHaveLength(1);
    expect(first.images[0].extension).toBe('png');
    expect(first.images[0].base64.length).toBeGreaterThan(10);

    const second = features['sheet-2'];
    expect(second.protected).toBe(true);
  });
});

describe('M3: DV export via resource diff', () => {
  const withDvResource = (wb: IWorkbookData, sheetId: string, rules: unknown[]): IWorkbookData => {
    const next = clone(wb);
    (next as { resources?: { name: string; data: string }[] }).resources = [
      { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: JSON.stringify({ [sheetId]: rules }) },
    ];
    return next;
  };

  it('writes changed DV rules to the worksheet', async () => {
    const result = await loadForUniver(readFixture('freeze-merge.xlsx'), 'xlsx', 'x');
    const baseline = clone(result.workbookData);
    const sheetId = baseline.sheetOrder[0];
    const current = withDvResource(baseline, sheetId, [
      {
        uid: 'dv1',
        type: 'list',
        ranges: [{ startRow: 0, startColumn: 4, endRow: 4, endColumn: 4 }],
        formula1: 'yes,no',
        allowBlank: true,
      },
      {
        uid: 'dv2',
        type: 'decimal',
        operator: 'between',
        ranges: [{ startRow: 0, startColumn: 5, endRow: 0, endColumn: 5 }],
        formula1: '1',
        formula2: '9',
      },
    ]);

    const diff = diffWorkbook(baseline, current, new Set());
    expect(diff.dvChangedSheetIds).toEqual([sheetId]);

    applyDiffToWorkbook(result.originalWorkbook!, current, diff, { sheetIdMap: result.sheetIdMap! });
    const out = new Uint8Array(await result.originalWorkbook!.xlsx.writeBuffer());

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    const dvModel = (wb.worksheets[0] as never as { dataValidations: { model: Record<string, { type: string; formulae: unknown[] }> } })
      .dataValidations.model;
    const refs = Object.keys(dvModel);
    expect(refs.length).toBeGreaterThanOrEqual(2);
    const listRef = refs.find(r => dvModel[r].type === 'list');
    expect(listRef).toBeTruthy();
    expect(String(dvModel[listRef!].formulae[0])).toBe('"yes,no"');
  });

  it('unchanged DV keeps original file validations untouched', async () => {
    const result = await loadForUniver(readFixture('m3-features.xlsx'), 'xlsx', 'm3');
    const baseline = clone(result.workbookData);
    const current = clone(baseline);
    // 只改一个值，不动 DV
    const sheet = current.sheets[current.sheetOrder[0]];
    (sheet.cellData as Record<number, Record<number, ICellData>>)[0][0] = { v: 'edited' };

    const diff = diffWorkbook(baseline, current, new Set());
    expect(diff.dvChangedSheetIds).toEqual([]);

    applyDiffToWorkbook(result.originalWorkbook!, current, diff, { sheetIdMap: result.sheetIdMap! });
    const out = new Uint8Array(await result.originalWorkbook!.xlsx.writeBuffer());
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    const dvModel = (wb.worksheets[0] as never as { dataValidations: { model: Record<string, unknown> } })
      .dataValidations.model;
    // 原文件的两条 DV 仍在
    expect(Object.keys(dvModel).length).toBeGreaterThanOrEqual(2);
    // 保护也未受影响
    const prot = (wb.worksheets[1] as never as { sheetProtection?: { sheet?: boolean } }).sheetProtection;
    expect(prot?.sheet).toBe(true);
  });
});

describe('M3: hyperlink write-back from cell.p customRanges', () => {
  it('writes a whole-cell hyperlink as ExcelJS hyperlink value', async () => {
    const result = await loadForUniver(readFixture('freeze-merge.xlsx'), 'xlsx', 'x');
    const baseline = clone(result.workbookData);
    const current = clone(baseline);
    const sheet = current.sheets[current.sheetOrder[0]];
    const text = 'Univer';
    (sheet.cellData as Record<number, Record<number, ICellData>>)[9] = {
      0: {
        v: text,
        p: {
          id: 'link-cell',
          body: {
            dataStream: `${text}\r\n`,
            textRuns: [],
            paragraphs: [{ startIndex: text.length }],
            customRanges: [{
              startIndex: 0,
              endIndex: text.length - 1,
              rangeId: 'r1',
              rangeType: 0,
              properties: { url: 'https://univer.ai' },
            }],
          },
          documentStyle: {},
        } as never,
      },
    };

    const diff = diffWorkbook(baseline, current, new Set());
    applyDiffToWorkbook(result.originalWorkbook!, current, diff, { sheetIdMap: result.sheetIdMap! });
    const out = new Uint8Array(await result.originalWorkbook!.xlsx.writeBuffer());
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    const cell = wb.worksheets[0].getCell('A10');
    const value = cell.value as { text?: string; hyperlink?: string };
    expect(value?.hyperlink).toBe('https://univer.ai');
    expect(value?.text).toBe('Univer');
  });
});
