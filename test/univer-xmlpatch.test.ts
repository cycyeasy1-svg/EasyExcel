// M5 验收测试：XML 级差分补丁的字节级无损保存。
// 核心断言：改 cell 后，除被改 sheet 的 XML（与必要时的 styles/calcChain/
// workbook.xml calcPr）外，其余所有 zip 部件与原文件逐字节相同 ——
// 图形（xdr:sp textbox）/图表/主题等 ExcelJS 未建模部件物理透传。
import { describe, it, expect } from 'vitest';
import { readFileSync as readFileSyncNode } from 'node:fs';
import JSZip from 'jszip';
import ExcelJS from '@cweijan/exceljs';
import type { ICellData, IWorkbookData } from '@univerjs/core';
import { loadForUniver } from '../src/react/view/excel/univer/loader';
import { diffWorkbook } from '../src/react/view/excel/univer/diff';
import { assertPatchableDiff, patchXlsxBytes, XmlPatchBlockedError } from '../src/react/view/excel/univer/xml_patch';

const readFixture = (name: string) => {
  const buf = readFileSyncNode(`test/fixtures/${name}`);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
};

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

type Cells = Record<number, Record<number, ICellData>>;
const cellsOf = (wb: IWorkbookData, sheetIndex = 0) =>
  wb.sheets[wb.sheetOrder[sheetIndex]].cellData as Cells;

async function partsOf(buf: ArrayBuffer | Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(buf);
  const out = new Map<string, Uint8Array>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (!file.dir) out.set(name, await file.async('uint8array'));
  }
  return out;
}

const bytesEqual = (a?: Uint8Array, b?: Uint8Array) =>
  !!a && !!b && a.length === b.length && a.every((v, i) => v === b[i]);

/** 输出与原文件的部件差异清单（解压后内容级比较） */
async function diffParts(original: ArrayBuffer, patched: Uint8Array) {
  const before = await partsOf(original);
  const after = await partsOf(patched);
  const changed: string[] = [];
  for (const [name, data] of before) {
    if (!after.has(name)) continue;
    if (!bytesEqual(data, after.get(name))) changed.push(name);
  }
  return {
    changed: changed.sort(),
    added: [...after.keys()].filter(n => !before.has(n)).sort(),
    removed: [...before.keys()].filter(n => !after.has(n)).sort(),
  };
}

async function patchEdit(fixture: string, edit: (current: IWorkbookData) => void) {
  const original = readFixture(fixture);
  const result = await loadForUniver(original, 'xlsx', fixture);
  const baseline = clone(result.workbookData);
  const current = clone(baseline);
  edit(current);
  const diff = diffWorkbook(baseline, current, new Set());
  const sheetNames = Object.fromEntries(
    baseline.sheetOrder.map(id => [id, baseline.sheets[id].name ?? '']));
  const patched = await patchXlsxBytes(result.originalBuffer!, diff, sheetNames);
  return { original, baseline, current, diff, ...patched, parts: await diffParts(original, patched.bytes) };
}

const readPart = async (buf: Uint8Array, name: string) => {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(name);
  return f ? await f.async('string') : null;
};

const reopen = async (bytes: Uint8Array) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  return wb;
};

describe('M5 byte-level pass-through', () => {
  it('edit one cell: only that sheet XML changes, shape drawing survives byte-identical', async () => {
    const { parts, bytes } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[1][0] = { ...cells[1][0], v: 'CHANGED' };
    });

    expect(parts.changed).toEqual(['xl/worksheets/sheet1.xml']);
    expect(parts.added).toEqual([]);
    expect(parts.removed).toEqual([]);

    const sheetXml = await readPart(bytes, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toContain('t="inlineStr"');
    expect(sheetXml).toContain('CHANGED');

    // textbox 图形部件逐字节保留（M2 的 ExcelJS 重写会整个丢掉）
    const drawing = await readPart(bytes, 'xl/drawings/drawing1.xml');
    expect(drawing).toContain('design note');

    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('A2').value).toBe('CHANGED');
    expect(wb.worksheets[0].getCell('A3').value).toBe('Feb');
    expect(wb.worksheets[1].getCell('A1').value).toBe('this sheet is never edited');
  });

  it('zero-edit save leaves every part byte-identical', async () => {
    const { parts } = await patchEdit('m5-shapes.xlsx', () => { });
    expect(parts.changed).toEqual([]);
    expect(parts.added).toEqual([]);
    expect(parts.removed).toEqual([]);
  });

  it('chart file: chart XML byte-identical after cell edit (M2 dropped it)', async () => {
    const { parts, bytes } = await patchEdit('r1-chart.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[0][0] = { v: 'PATCHED' };
    });

    expect(parts.changed).toContain('xl/worksheets/sheet1.xml');
    expect(parts.changed).not.toContain('xl/charts/chart1.xml');
    expect(parts.changed).not.toContain('xl/drawings/drawing1.xml');
    expect(parts.removed).toEqual([]);

    // 补丁后仍可通过完整导入管线打开（sanitize + ExcelJS），值已更新
    const re = await loadForUniver(
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
      'xlsx', 'r1-chart.xlsx');
    const cells = cellsOf(re.workbookData);
    expect(cells[0][0].v).toBe('PATCHED');
  });
});

describe('M5 value types', () => {
  it('handles string escaping, whitespace preserve, number, boolean, formula, clear, new cell/row', async () => {
    const special = 'a & b < c > "d"';
    const padded = '  pad  ';
    const { parts, bytes } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[1][0] = { v: special };            // A2 特殊字符
      cells[2][0] = { v: padded };             // A3 前后空白
      cells[1][1] = { v: 99.25 };              // B2 数字（原有 s= 保留）
      cells[0][2] = { v: 0, t: 3 };            // C1 布尔 false（Univer BOOLEAN）
      delete cells[0][3];                      // D1 清空
      cells[3][1] = { f: '=SUM(B2:B3)*2', v: 248.5 }; // B4 改公式
      cells[0][5] = { v: '中文字符' };          // F1 行内新 cell
      (cells[9] ??= {})[0] = { v: 'new row' }; // A10 新行
    });

    expect(parts.changed).toContain('xl/worksheets/sheet1.xml');
    expect(parts.added).toEqual([]);

    const sheetXml = (await readPart(bytes, 'xl/worksheets/sheet1.xml'))!;
    expect(sheetXml).toContain('a &amp; b &lt; c &gt; "d"');
    expect(sheetXml).toContain('xml:space="preserve"');
    expect(sheetXml).toContain('<row r="10"><c r="A10" t="inlineStr"><is><t>new row</t></is></c></row>');

    const wb = await reopen(bytes);
    const ws = wb.worksheets[0];
    expect(ws.getCell('A2').value).toBe(special);
    expect(ws.getCell('A3').value).toBe(padded);
    expect(ws.getCell('B2').value).toBe(99.25);
    expect(ws.getCell('C1').value).toBe(false);
    expect(ws.getCell('D1').value).toBeNull();
    expect((ws.getCell('B4').value as { formula?: string })?.formula).toBe('SUM(B2:B3)*2');
    expect(ws.getCell('F1').value).toBe('中文字符');
    expect(ws.getCell('A10').value).toBe('new row');
  });

  it('keeps original style index when only the value changes', async () => {
    const { bytes, parts } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      // B2 原有货币格式（s= 指向 fmt_money xf）：只改值
      cells[1][1] = { ...cells[1][1], v: 777 };
    });
    expect(parts.changed).toEqual(['xl/worksheets/sheet1.xml']);
    const sheetXml = (await readPart(bytes, 'xl/worksheets/sheet1.xml'))!;
    const b2 = sheetXml.match(/<c r="B2"[^>]*>/)?.[0];
    expect(b2).toMatch(/\bs="\d+"/);
    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('B2').numFmt).toBe('0.00"m"');
  });

  it('downgrades rich-text cells to plain inlineStr and reports it', async () => {
    const { bytes, richTextDowngraded } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[1][0] = {
        p: { body: { dataStream: 'RICH text\r\n' } },
      } as never;
    });
    expect(richTextDowngraded).toBe(true);
    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('A2').value).toBe('RICH text');
  });
});

describe('M5 style patch (append-only)', () => {
  it('appends new font/fill/xf, keeps every original styles.xml entry, dedupes same style', async () => {
    const originalStyles = (await readPart(new Uint8Array(readFixture('m5-shapes.xlsx')), 'xl/styles.xml'))!;
    const { bytes, parts } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      const style = { bl: 1, cl: { rgb: '#1ABC9C' }, bg: { rgb: '#222222' } };
      cells[1][0] = { ...cells[1][0], s: style };
      cells[2][0] = { ...cells[2][0], s: style }; // 同样式 → 复用同一 xf
    });
    expect(parts.changed.sort()).toEqual(['xl/styles.xml', 'xl/worksheets/sheet1.xml']);

    const styles = (await readPart(bytes, 'xl/styles.xml'))!;
    // 原有每个 xf/font 片段仍原样存在（append-only：前缀内容不动）
    for (const frag of originalStyles.match(/<xf [^>]*\/>/g) ?? []) {
      expect(styles).toContain(frag);
    }
    expect(styles).toContain('<color rgb="FF1ABC9C"/>');
    expect(styles).toContain('<fgColor rgb="FF222222"/>');

    const sheetXml = (await readPart(bytes, 'xl/worksheets/sheet1.xml'))!;
    const a2s = sheetXml.match(/<c r="A2" s="(\d+)"/)?.[1];
    const a3s = sheetXml.match(/<c r="A3" s="(\d+)"/)?.[1];
    expect(a2s).toBeDefined();
    expect(a2s).toBe(a3s);

    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('A2').font?.bold).toBe(true);
  });

  it('appends custom numFmt with id above existing max', async () => {
    const { bytes } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[1][0] = { ...cells[1][0], s: { n: { pattern: 'yyyy/mm/dd' } } };
    });
    const styles = (await readPart(bytes, 'xl/styles.xml'))!;
    const m = styles.match(/<numFmt numFmtId="(\d+)" formatCode="yyyy\/mm\/dd"\/>/);
    expect(m).toBeTruthy();
    expect(Number(m![1])).toBeGreaterThanOrEqual(164);
    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('A2').numFmt).toBe('yyyy/mm/dd');
  });
});

describe('M5 calcChain & recalc', () => {
  it('drops calcChain (with content-types/rels cleanup) and forces full recalc', async () => {
    const { parts, bytes } = await patchEdit('m5-calcchain.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[1][1] = { ...cells[1][1], v: 999 }; // B2：被 B4 公式引用
    });
    expect(parts.removed).toEqual(['xl/calcChain.xml']);
    expect(await readPart(bytes, '[Content_Types].xml')).not.toContain('calcChain');
    expect(await readPart(bytes, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
    expect(await readPart(bytes, 'xl/workbook.xml')).toMatch(/<calcPr[^>]*fullCalcOnLoad="1"/);

    const wb = await reopen(bytes);
    expect((wb.worksheets[0].getCell('B4').value as { formula?: string })?.formula).toBe('SUM(B2:B3)');
  });

  it('openpyxl-produced file (no calcChain, different XML shapes) round-trips', async () => {
    const { parts, bytes } = await patchEdit('m5-openpyxl.xlsx', (cur) => {
      const cells = cellsOf(cur);
      cells[0][0] = { v: 'patched' };
      (cells[2] ??= {})[2] = { v: 3.5 }; // C3 新行新 cell
    });
    expect(parts.changed).toContain('xl/worksheets/sheet1.xml');
    expect(parts.changed).not.toContain('xl/styles.xml');
    expect(parts.removed).toEqual([]);
    const wb = await reopen(bytes);
    expect(wb.worksheets[0].getCell('A1').value).toBe('patched');
    expect(wb.worksheets[0].getCell('C3').value).toBe(3.5);
    expect(wb.worksheets[0].getCell('B2').font?.bold).toBe(true); // 原样式保留
  });
});

describe('M5 defensive rejection', () => {
  const load = async (fixture: string) => {
    const result = await loadForUniver(readFixture(fixture), 'xlsx', fixture);
    const baseline = clone(result.workbookData);
    return { result, baseline, current: clone(baseline) };
  };

  it('rejects merge changes as structural', async () => {
    const { baseline, current } = await load('m5-shapes.xlsx');
    const sheet = current.sheets[current.sheetOrder[0]];
    sheet.mergeData = [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }];
    const diff = diffWorkbook(baseline, current, new Set());
    expect(() => assertPatchableDiff(diff)).toThrowError(XmlPatchBlockedError);
    expect(() => assertPatchableDiff(diff)).toThrowError(/structural/);
  });

  it('rejects structural sheets, row size changes and renames', async () => {
    const { baseline, current } = await load('m5-shapes.xlsx');
    const diffStructural = diffWorkbook(baseline, current, new Set([current.sheetOrder[0]]));
    expect(() => assertPatchableDiff(diffStructural)).toThrowError(/structural/);

    const { baseline: b2, current: c2 } = await load('m5-shapes.xlsx');
    (c2.sheets[c2.sheetOrder[0]].rowData as Record<number, { h?: number }>)[3] = { h: 60 };
    expect(() => assertPatchableDiff(diffWorkbook(b2, c2, new Set()))).toThrowError(/structural/);

    const { baseline: b3, current: c3 } = await load('m5-shapes.xlsx');
    c3.sheets[c3.sheetOrder[0]].name = 'Renamed';
    expect(() => assertPatchableDiff(diffWorkbook(b3, c3, new Set()))).toThrowError(/structural/);
  });

  it('rejects DV/CF changes as dvcf', async () => {
    const { baseline, current } = await load('m5-shapes.xlsx');
    (current as { resources?: { name: string; data: string }[] }).resources = [{
      name: 'SHEET_DATA_VALIDATION_PLUGIN',
      data: JSON.stringify({ [current.sheetOrder[0]]: [{ uid: 'x', type: 'list', formula1: 'a,b' }] }),
    }];
    const diff = diffWorkbook(baseline, current, new Set());
    expect(() => assertPatchableDiff(diff)).toThrowError(/dvcf/);
  });
});

describe('M5 sheet name resolution', () => {
  it('patches the right part when sheet names need XML escaping', async () => {
    const { parts, bytes } = await patchEdit('m5-shapes.xlsx', (cur) => {
      const cells = cellsOf(cur, 1); // "Notes & <Specs>"
      cells[0][0] = { v: 'edited on escaped sheet' };
    });
    expect(parts.changed).toEqual(['xl/worksheets/sheet2.xml']);
    const wb = await reopen(bytes);
    expect(wb.worksheets[1].getCell('A1').value).toBe('edited on escaped sheet');
  });
});
