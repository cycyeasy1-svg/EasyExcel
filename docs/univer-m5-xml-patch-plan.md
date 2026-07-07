# M5 实施方案：XML 级差分补丁（字节级无损保存）

> 新会话接手用。前置状态：分支 `spike/univer-phase0`（收官 commit fba39cc + docs 98c9739），
> Univer 已是唯一引擎，迁移全记录见 `docs/univer-phase0-report.md`。
> 本方案已获用户确认，产品取舍已拍板（见下）。

## 背景与目标

插件定位：**在 VSCode 里辅助预览/编辑 Excel 设计书，不是 Excel 替代品**。
核心要求：编辑保存后，除了被改的内容，**原文件的任何内容和格式一个字节都不能变**（"变更内容的单纯差分取入"）。

当前 M2 的增量导出是"模型级差分"：diff 只应用被改的 cell，但最终经
`ExcelJS.writeBuffer()` 重新序列化**整个文件**——ExcelJS 模型外的特性丢失。

**已实测的关键缺口**（本方案的直接动机）：
- 含图形/文本框（`xdr:sp`）的文件能打开，但 ExcelJS 重写后 **shape drawing 整个丢失**
  （实测脚本 `test/output/shape_load_test.mjs`，fixture 生成逻辑在其上方的 python 块，
  可参考重建为正式 fixture）。设计书里的矢印/枠/テキストボックス全灭 → 不可接受。
- 同理丢失：图表、透视表、宏、打印设置、行列分组等一切 ExcelJS 未建模部件。

## 方案概述

保存到原文件时**不再走 ExcelJS**，改为在**原始文件字节**上打 XML 补丁：

1. 打开时保留**原始 buffer**（注意：loader 现在只保留了 ExcelJS workbook，
   而且喂给 ExcelJS 的是 sanitize 过的副本——M5 必须在 `loader.ts` 增加
   `originalBuffer: ArrayBuffer` 字段，存**未经 sanitize 的真原始字节**）。
2. 保存时用现有 `univer/diff.ts` 算 cell 级变更清单（已有 valueChanged/styleChanged 区分，经 26+ 测试验证）。
3. 用 jszip 打开原始 buffer，只修改：
   - 被编辑 cell 所在的 `xl/worksheets/sheetN.xml` 中对应的 `<c>` 节点
   - 样式变更：`xl/styles.xml` **只追加**新 xf/font/fill/border/numFmt（原索引全不动）
   - 删除 `xl/calcChain.xml`（若存在且有公式变更；Excel 自动重建，避免陈旧引用）
4. 其余所有 zip 部件**原样透传**——图形/图表/透视表/宏/打印设置物理上不被触碰。

**净效果**：图表/宏文件的"保存前丢失警告"在保存到原文件的路径上可以**删除**
（这些部件现在会存活）；警告仅保留在另存为路径（仍走 ExcelJS 重建，导出副本）。

## 产品取舍（用户已确认）

**结构性操作在 Univer 里完全禁用**（不是"允许但警告"）：
插删行/列、增删/重命名 sheet、移动行列。理由：XML 级正确平移引用（公式/CF 范围/
图形锚点/definedNames）≈ 重新实现 Excel 语义；辅助工具定位下这些操作应回 Excel 做。
被禁用操作触发时弹 toast："此操作请在 Excel 中进行"（i18n：excel_i18n.ts 加 key）。

**M5 允许的编辑**（全部走 XML 补丁）：
- 改单元格值 / 公式 / 清空单元格
- 单元格样式（字体/颜色/填充/边框/对齐/数字格式）
- 建议二期再开（初版也禁用，实现简单但先收窄承诺面）：行高列宽、合并/取消合并、冻结

## 实施细节

### 1. `univer/xml_patch.ts`（新，核心 ~核心 3-4 天）

输入：`originalBuffer` + `WorkbookDiff`（diff.ts 产物）+ sheetId→sheet XML 部件映射。

**sheetId → 部件路径映射**：workbook.xml 的 `<sheet name= sheetId= r:id=>` +
workbook.xml.rels 的 rId→target。导入时 sheet 顺序 = workbook.worksheets 顺序 =
我们的 `sheet-{i+1}`，但**不要靠顺序**——用 loader 里 ExcelJS worksheet 的真名匹配
workbook.xml 的 name 属性（注意 XML 转义）。

**cell 补丁**（正则/字符串操作足够，见风险节；或用 DOMParser+XMLSerializer——
但序列化器可能重排属性/自闭合形式导致整个 part 差异变大，**推荐字符串手术**）：
- 定位 `<row r="N" ...>` → 行内定位 `<c r="A1"` 节点（可能自闭合 `<c r="A1" s="3"/>`）。
- 值写入类型策略：
  - 字符串 → **inlineStr**（`<c r=.. t="inlineStr"><is><t>text</t></is></c>`）——
    **关键简化**：不动 sharedStrings.xml（不追加不重排），Excel 完全接受 inlineStr。
    xml:space="preserve" 处理前后空白。
  - 数字/布尔 → `<v>` （t 省略 / t="b"）。
  - 公式 → `<f>公式</f>`（去掉 `=` 前缀）+ 可选缓存 `<v>`；有公式变更则删 calcChain.xml。
  - 清空 → 移除 `<v>/<f>/<is>`，若无样式则整个删掉 `<c>` 节点。
- **保留原有属性**：只改值时，原 `s=`（样式索引）等属性原样保留。
- cell 不存在时：在 row 内按列序插入新 `<c>`；row 不存在时按行序插入
  `<row r="N">`（注意 spans 属性可以不写，Excel 容忍）。
- 转义：`& < >` 必须转义；比较/写入统一处理。

**样式补丁**（append-only，~1-2 天）：
- 解析 styles.xml 的 cellXfs/fonts/fills/borders/numFmts 的 count。
- 对 styleChanged 的 cell：把 IStyleData（复用 `export_styles.ts` 的映射逻辑改造为
  产 XML 片段）转成 font/fill/border/numFmt XML，**追加**到对应集合尾部，
  新建 xf 引用这些新索引，cell 的 s= 指向新 xf 索引。同 IStyleData 去重复用同一 xf。
- 自定义 numFmt id 从 164 起，取现有 max+1。
- 注意：Univer 的 style 是"完整样式"（非增量），直接整体映射为新 xf 即可，
  语义与 M2 的 applyUniverStyleToCell 一致。

### 2. `loader.ts` 改动（~0.5 天）
- `UniverLoadResult` 增加 `originalBuffer?: ArrayBuffer`（xlsx/xlsm 时为真原始字节）。
- 沿用现有 sanitize 逻辑喂 ExcelJS（渲染用副本），二者并存。

### 3. `export.ts` 改动（~0.5 天）
- 保存到原文件（非 saveAs）且 ext 为 xlsx/xlsm 且有 originalBuffer：
  走 `xml_patch.ts`；若 diff 含结构变更（理论上已被禁用，防御性检查
  `diff.sheets.some(s => s.status !== 'incremental') || removedSheetIds.length ||
  dvChangedSheetIds.length || cfChangedSheetIds.length || merges/rowChanges/colChanges/
  freeze/renamed 非空`）→ 报错提示（不应发生）。
- DV/CF 编辑：M5 初版**也禁用**（同结构操作；它们的资源级重写只在 saveAs 路径保留），
  或允许但走 ExcelJS 路径需警告——推荐禁用，承诺面干净。
  禁用方式见下节；若 Univer 无法细粒度禁 DV/CF 编辑 UI，兜底：保存时检测到
  dv/cfChangedSheetIds 非空 → 弹"DV/CF 修改无法无损保存，请在 Excel 中修改"并中止。
- saveAs 路径完全不动（仍 ExcelJS 重建导出副本，保留 lossy 警告）。

### 4. 结构操作禁用（~1 天，需运行时验证）
两个候选机制（新会话先各花 30 分钟验证哪个可行）：
- (a) 权限点：`fWorksheet.getWorksheetPermission().setPoint(point, false)`——
  查 `WorksheetPermissionPoint` 枚举有无 InsertRow/DeleteRow/InsertColumn/DeleteColumn/
  Sort/MoveRows 等点位（d.ts 在 @univerjs/sheets/lib/types/facade/permission/）。
  注意坑：**必须先 `protect()` 再 setPoint**（M3 已踩过）。且要确认 protect 后
  普通编辑（setCellEdit point）仍为 true。
- (b) 命令拦截：`univerAPI.addEvent(univerAPI.Event.BeforeCommandExecute, cb)` 里
  对 `sheet.command.insert-row` 等结构命令 `cancel`（查 f-univer.d.ts 确认事件名与
  cancel 机制），同时隐藏右键菜单项（UniverSheetsCorePreset 有 menu 配置项可隐藏，
  查 preset 工厂参数）。
- sheet 增删/重命名：拦截 `sheet.command.insert-sheet` / `remove-sheet` /
  `set-worksheet-name`（bottombar 的 UI 隐藏可查 preset 配置 `footer` 选项）。
- adapter 的结构 mutation 日志（STRUCTURAL_MUTATION）保留作为最后防线：
  保存时若 structuralSheetIds 非空 → 中止并提示（防拦截遗漏）。

### 5. 测试（~1.5 天）
新建 `test/univer-xmlpatch.test.ts`（node 环境）：
- **字节级断言**（核心验收）：改 A1 → 输出 zip 中除被改 sheet XML（与可能的
  calcChain 删除）外，**其余所有部件与原文件逐字节相同**（比较 zip entry 内容）。
- shape fixture（把本方案背景节的构造脚本落成 `test/fixtures/make_m5_fixtures.py`）：
  改值保存 → drawing1.xml 逐字节保留、textbox 文字还在。
- r1-chart.xlsx：改值保存 → `xl/charts/chart1.xml` 逐字节保留（**图表存活**，
  对比 M2 是质变）；用 openpyxl 重新打开验证图表对象存在。
- 各值类型：字符串（含中文/特殊字符 &<>/空白保留）、数字、布尔、公式、清空、
  新 cell/新 row 插入、只改样式不改值（s= 更新且原 xf 不动）、numFmt 追加。
- styles.xml 追加后原有 xf 数量与内容不变（前缀逐字节相同）。
- 往返：patch 后文件用 ExcelJS/openpyxl 能正常打开、值正确。
- e2e（沿用 UniverSpike + saveProbe 模式）：真实 Univer 编辑 → patch 字节 →
  python 端 openpyxl 验证 + 部件 diff 清单打印。

### 6. 文案与收尾（~0.5 天）
- excel_i18n.ts 加 key：结构操作禁用提示、DV/CF 禁用提示（en/zh-CN/zh-TW）。
- 保存路径的 lossy 图表警告移除（saveAs 保留）。
- README/报告更新："在 VSCode 里改的只有你改的格子，其余一个字节都不会变"。

## 风险与注意

- **字符串手术 vs XML parser**：选字符串手术是为了让未触碰的行/节点保持原字节。
  风险是 sheet XML 的形态多样（自闭合、命名空间前缀、属性顺序）——真实 Excel 输出
  形态稳定，fixture 需覆盖 Excel 与 openpyxl 两种产出。回退方案：对"被改的 row"
  整行重建（行内其他 cell 从原 XML 提取原样拼回），行外零触碰。
- **jszip 重压缩**：未修改的 entry 内容不变但压缩字节可能不同——承诺口径是
  "**部件内容**逐字节不变"，zip 容器层压缩差异不影响 Excel 打开与内容一致性。
  测试断言按解压后内容比较。
- **富文本 cell 的编辑**：值变更若涉及 cell.p（富文本），初版降级为纯文本 inlineStr
  写入（用户在 Univer 里改了富文本格的文字 → 保存为纯文本 + 弹提示）或直接禁编辑
  富文本格——新会话酌情，推荐前者+提示。
- **公式重算**：Univer 里改了被公式引用的值，公式 cell 的缓存 `<v>` 会陈旧——
  写入时对"公式引用受影响"不做分析，统一删 calcChain 并可给 workbook.xml 的
  calcPr 加 fullCalcOnLoad="1"（一行属性补丁，Excel 打开自动重算；验证此属性补丁
  不破坏其他内容）。
- 沙箱外注意：`@cweijan/exceljs` 仅浏览器构建，vitest 需 `server.deps.inline`（已配）。

## 工作量合计：约 6-8 人日

## 验收标准（一句话）

打开任意真实设计书（含图形/图表/打印设置），改三个格子的文字保存，
用 zip diff 工具对比前后文件：除对应 sheet XML（与 calcChain）外全部部件逐字节一致，
Excel 打开后图形/图表/格式完好、三个格子已更新。
