// R1 probe: does @cweijan/exceljs preserve modeled and unmodeled xlsx parts
// across load -> writeBuffer? Decides the round-trip promise level for the
// Univer migration plan (charts/pivots/vba passthrough vs warn-on-save).
// Also probes: can chart-bearing files be LOADED at all (fork + upstream)?
import { describe, it, expect } from 'vitest';
import { readFileSync as readFileSyncNode, writeFileSync, mkdirSync } from 'node:fs';

const readFileSync = (p: string) => readFileSyncNode(p) as unknown as Buffer;

mkdirSync('test/output', { recursive: true });
import ExcelJSFork from '@cweijan/exceljs';
import ExcelJSUpstream from 'exceljs';
import JSZip from 'jszip';

const NO_CHART = 'test/fixtures/r1-features.xlsx';
const WITH_CHART = 'test/fixtures/r1-chart.xlsx';

const listParts = async (buf: Buffer | Uint8Array) => {
  const zip = await JSZip.loadAsync(buf);
  return Object.keys(zip.files).filter((n) => !zip.files[n].dir).sort();
};

const readPart = async (buf: Buffer | Uint8Array, name: string) => {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(name);
  return f ? await f.async('string') : null;
};

describe('R1: chart-file load behavior', () => {
  it('fork: loading a chart-bearing xlsx', async () => {
    const wb = new ExcelJSFork.Workbook();
    let error: unknown = null;
    try {
      await wb.xlsx.load(readFileSync(WITH_CHART));
    } catch (e) {
      error = e;
    }
    console.log('FORK chart load:', error ? `CRASH: ${(error as Error).message}` : 'OK');
  });

  it('upstream: loading a chart-bearing xlsx', async () => {
    const wb = new ExcelJSUpstream.Workbook();
    let error: unknown = null;
    try {
      await wb.xlsx.load(readFileSync(WITH_CHART));
    } catch (e) {
      error = e;
    }
    console.log('UPSTREAM chart load:', error ? `CRASH: ${(error as Error).message}` : 'OK');
  });
});

describe('R1: fork load->write passthrough (no chart)', () => {
  it('reports part-level and feature-level survival', async () => {
    const src = readFileSync(NO_CHART);
    const wb = new ExcelJSFork.Workbook();
    await wb.xlsx.load(src);

    const out = Buffer.from(await wb.xlsx.writeBuffer());
    writeFileSync('test/output/r1-features.roundtrip.xlsx', out);

    const before = await listParts(src);
    const after = await listParts(out);

    const wbXmlAfter = (await readPart(out, 'xl/workbook.xml')) ?? '';
    const sheet1After = (await readPart(out, 'xl/worksheets/sheet1.xml')) ?? '';
    const stylesAfter = (await readPart(out, 'xl/styles.xml')) ?? '';

    const report = {
      lostParts: before.filter((p) => !after.includes(p)),
      addedParts: after.filter((p) => !before.includes(p)),
      definedNameSurvived: /SalesRange/.test(wbXmlAfter),
      conditionalFormattingSurvived: /conditionalFormatting/.test(sheet1After),
      customNumFmtSurvived: /0\.00(&quot;|")m(&quot;|")/.test(stylesAfter),
      commentPartsAfter: after.filter((p) => /comment/i.test(p)),
    };
    console.log('R1 REPORT (fork, no chart)\n' + JSON.stringify(report, null, 2));

    expect(report.definedNameSurvived).toBe(true);
    expect(report.conditionalFormattingSurvived).toBe(true);
  });
});
