// R1 probe: does @cweijan/exceljs preserve unmodeled/modeled parts across
// load -> writeBuffer? Compares zip part lists and spot-checks key XML content.
import { createRequire } from 'node:module';
import { readFileSync, writeFileSync } from 'node:fs';
import zlib from 'node:zlib';

const require = createRequire(import.meta.url);
const ExcelJS = require('@cweijan/exceljs');
const JSZip = require('jszip');

const SRC = 'test/fixtures/r1-features.xlsx';
const OUT = 'test/fixtures/r1-features.roundtrip.xlsx';

const listParts = async (buf) => {
  const zip = await JSZip.loadAsync(buf);
  return Object.keys(zip.files).filter((n) => !zip.files[n].dir).sort();
};

const readPart = async (buf, name) => {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file(name);
  return f ? await f.async('string') : null;
};

const src = readFileSync(SRC);
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(src);

const out = Buffer.from(await wb.xlsx.writeBuffer());
writeFileSync(OUT, out);

const before = await listParts(src);
const after = await listParts(out);

console.log('=== parts BEFORE ===\n' + before.join('\n'));
console.log('\n=== parts AFTER ===\n' + after.join('\n'));
console.log('\n=== LOST parts ===\n' + (before.filter((p) => !after.includes(p)).join('\n') || '(none)'));
console.log('\n=== ADDED parts ===\n' + (after.filter((p) => !before.includes(p)).join('\n') || '(none)'));

// Spot checks on modeled features
const wbXmlAfter = await readPart(out, 'xl/workbook.xml');
console.log('\ndefinedName survived:', /SalesRange/.test(wbXmlAfter));

const sheet1After = await readPart(out, 'xl/worksheets/sheet1.xml');
console.log('conditionalFormatting survived:', /conditionalFormatting/.test(sheet1After ?? ''));

const stylesAfter = await readPart(out, 'xl/styles.xml');
console.log('custom numFmt survived:', /0\.00&quot;m&quot;|0\.00"m"/.test(stylesAfter ?? ''));

const commentParts = after.filter((p) => /comment/i.test(p));
console.log('comment parts after:', commentParts.length ? commentParts.join(', ') : '(none)');

const chartParts = after.filter((p) => /chart/i.test(p));
console.log('chart parts after:', chartParts.length ? chartParts.join(', ') : '(NONE — charts dropped)');

// Does the fork expose any raw-part/media passthrough API?
console.log('\nworkbook model keys:', Object.keys(wb.model ?? {}).join(', '));
