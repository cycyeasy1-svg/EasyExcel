// @vitest-environment jsdom
// theme1.xml 解析路径（需要 DOMParser）：钉死 clrScheme 文档顺序 → theme 索引的互换。
import { describe, it, expect } from 'vitest';
import { parseThemePalette } from '../src/react/view/excel/univer/theme_colors';

const THEME_XML = `<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="EEEEEE"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>
      <a:accent1><a:srgbClr val="AA0001"/></a:accent1>
      <a:accent2><a:srgbClr val="AA0002"/></a:accent2>
      <a:accent3><a:srgbClr val="AA0003"/></a:accent3>
      <a:accent4><a:srgbClr val="AA0004"/></a:accent4>
      <a:accent5><a:srgbClr val="AA0005"/></a:accent5>
      <a:accent6><a:srgbClr val="AA0006"/></a:accent6>
      <a:hlink><a:srgbClr val="BB0001"/></a:hlink>
      <a:folHlink><a:srgbClr val="BB0002"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;

describe('parseThemePalette (DOMParser path)', () => {
  it('maps clrScheme document order to theme indices with 0/1, 2/3 swap', () => {
    const palette = parseThemePalette(THEME_XML);
    expect(palette[0]).toBe('EEEEEE'); // lt1 → theme 0
    expect(palette[1]).toBe('111111'); // dk1 → theme 1 (默认文字色)
    expect(palette[2]).toBe('DDDDDD'); // lt2 → theme 2
    expect(palette[3]).toBe('222222'); // dk2 → theme 3
    expect(palette[4]).toBe('AA0001'); // accent1 → theme 4
    expect(palette[9]).toBe('AA0006'); // accent6 → theme 9
    expect(palette[10]).toBe('BB0001'); // hlink
    expect(palette[11]).toBe('BB0002'); // folHlink
  });
});
