// @vitest-environment jsdom
// P0 smoke test: @zwight/luckyexcel as a dev-time comparison baseline for the
// future self-built ExcelJS -> IWorkbookData importer. Not a runtime dep.
// jsdom env: the lib is browser-only (File/FileReader based).
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import LuckyExcel from '@zwight/luckyexcel';

describe('P0: luckyexcel converts fixture to Univer workbook data', () => {
  it('transforms r1-features.xlsx', async () => {
    const buf = readFileSync('test/fixtures/r1-features.xlsx');
    const file = new File([new Uint8Array(buf)], 'r1-features.xlsx');

    const result = await new Promise<Record<string, unknown>>((resolve, reject) => {
      try {
        (LuckyExcel as unknown as {
          transformExcelToUniver(
            f: File,
            ok: (d: Record<string, unknown>) => void,
            err: (e: unknown) => void,
          ): void;
        }).transformExcelToUniver(file, resolve, reject);
      } catch (e) {
        reject(e);
      }
    });

    console.log(
      'LUCKYEXCEL REPORT\n' +
        JSON.stringify(
          {
            topLevelKeys: Object.keys(result ?? {}),
            sheetCount: Object.keys((result as { sheets?: object }).sheets ?? {}).length,
            styleCount: Object.keys((result as { styles?: object }).styles ?? {}).length,
          },
          null,
          2,
        ),
    );
    expect(result).toBeTruthy();
  });
});
