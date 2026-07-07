// M4 单测：条件格式导入（ExcelJS → Univer 资源）与写回（资源 diff → ExcelJS）。
import { describe, it, expect } from 'vitest';
import { readFileSync as readFileSyncNode } from 'node:fs';
import ExcelJS from '@cweijan/exceljs';
import type { IWorkbookData } from '@univerjs/core';
import { loadForUniver } from '../src/react/view/excel/univer/loader';
import { diffWorkbook } from '../src/react/view/excel/univer/diff';
import { applyDiffToWorkbook } from '../src/react/view/excel/univer/apply';
import { CF_PLUGIN, type UniverCfRule } from '../src/react/view/excel/univer/cf';

const readFixture = (name: string) => {
  const buf = readFileSyncNode(`test/fixtures/${name}`);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

const cfResourceOf = (wb: IWorkbookData): Record<string, UniverCfRule[]> => {
  const entry = (wb as { resources?: { name: string; data: string }[] }).resources
    ?.find(r => r.name === CF_PLUGIN);
  return entry ? JSON.parse(entry.data) : {};
};

describe('M4: CF import', () => {
  it('maps cellIs greaterThan rule from r1-features.xlsx into Univer resource', async () => {
    const { workbookData } = await loadForUniver(readFixture('r1-features.xlsx'), 'xlsx', 'r1');
    const cf = cfResourceOf(workbookData);
    const rules = cf['sheet-1'];
    expect(rules?.length).toBeGreaterThanOrEqual(1);
    const rule = rules[0];
    expect(rule.rule.type).toBe('highlightCell');
    expect(rule.rule.subType).toBe('number');
    expect(rule.rule.operator).toBe('greaterThan');
    expect(rule.rule.value).toBe(20);
    expect((rule.rule.style as { bg?: { rgb?: string } })?.bg?.rgb?.toLowerCase()).toBe('#ffc7ce');
    expect(rule.ranges[0]).toMatchObject({ startRow: 1, startColumn: 1, endRow: 4, endColumn: 1 });
  });
});

describe('M4: CF export via resource diff', () => {
  it('rewrites changed CF rules into the worksheet', async () => {
    const result = await loadForUniver(readFixture('r1-features.xlsx'), 'xlsx', 'r1');
    const baseline = clone(result.workbookData);
    const current = clone(baseline);
    // 修改 CF：阈值 20 → 15，加一条 colorScale
    const resources = (current as { resources: { name: string; data: string }[] }).resources;
    const cfEntry = resources.find(r => r.name === CF_PLUGIN)!;
    const cfMap = JSON.parse(cfEntry.data) as Record<string, UniverCfRule[]>;
    (cfMap['sheet-1'][0].rule as { value: number }).value = 15;
    cfMap['sheet-1'].push({
      cfId: 'new-cs',
      ranges: [{ startRow: 1, startColumn: 1, endRow: 4, endColumn: 1 }],
      stopIfTrue: false,
      rule: {
        type: 'colorScale',
        config: [
          { index: 0, color: '#ffffff', value: { type: 'min' } },
          { index: 1, color: '#00ff00', value: { type: 'max' } },
        ],
      },
    });
    cfEntry.data = JSON.stringify(cfMap);

    const diff = diffWorkbook(baseline, current, new Set());
    expect(diff.cfChangedSheetIds).toEqual(['sheet-1']);

    applyDiffToWorkbook(result.originalWorkbook!, current, diff, { sheetIdMap: result.sheetIdMap! });
    const out = new Uint8Array(await result.originalWorkbook!.xlsx.writeBuffer());
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out);
    const cfs = (wb.worksheets[0] as never as { conditionalFormattings: { ref: string; rules: { type: string; formulae?: unknown[] }[] }[] })
      .conditionalFormattings;
    expect(cfs.length).toBe(2);
    const cellIs = cfs.find(c => c.rules[0]?.type === 'cellIs');
    expect(Number(cellIs?.rules[0]?.formulae?.[0])).toBe(15);
    expect(cfs.some(c => c.rules[0]?.type === 'colorScale')).toBe(true);
  });

  it('unchanged CF keeps the original rule untouched', async () => {
    const result = await loadForUniver(readFixture('r1-features.xlsx'), 'xlsx', 'r1');
    const baseline = clone(result.workbookData);
    const current = clone(baseline);
    const diff = diffWorkbook(baseline, current, new Set());
    expect(diff.cfChangedSheetIds).toEqual([]);
    expect(diff.isEmpty).toBe(true);
  });
});
