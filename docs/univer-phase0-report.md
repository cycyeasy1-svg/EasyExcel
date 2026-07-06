# Univer 迁移 Phase 0 技术验证报告

日期：2026-07-06 ｜ 分支：`spike/univer-phase0` ｜ 计划：`~/.claude/plans/pure-weaving-oasis.md`

## Gate 判定：**通过（一项带条件）**

| Gate 条件 | 阈值 | 实测 | 判定 |
|---|---|---|---|
| bundle gzip 增量 | < 2MB | **1,799 KB**（Univer 全 preset 懒加载 chunk，含 8 个 preset + 双语 locale + CSS 87KB） | ✅ |
| 冷启动增量 | < 800ms | 导航→表格就绪 **0.7~1.1s**（生产构建，Playwright headless ×3）；现有应用外壳基线 0.15s | ⚠️ 边界（见下） |
| webview 无阻断性问题 | — | 浏览器环境零控制台错误；`vscode-webview://` 生产环境需 F5 手测确认 | ✅（留一项手测） |

**冷启动说明**：0.15s 基线只是 React 外壳，不含旧引擎解析 xlsx + canvas 首绘的时间，真实增量小于表面数字；且 Univer chunk 加载可与文件字节 fetch 并行。M1 优化空间：univer-core/extras 分包、`<link rel=modulepreload>`。判定为可接受，M1 结束时用真实文件复测。

## 实测数据汇总

### 1. 渲染与功能（Playwright，dev + 生产构建各一轮）

- 完整工具栏/右键菜单/sheet bar/缩放控件，中文 locale 正常。
- **自定义数字格式直通渲染**：`0.00"m";[Red]-0.00"m"` 正确显示为 `10.00m`（现有引擎必丢的场景）。
- **公式引擎**：`=SUM(B2:B5)` 实时计算，改值后自动重算（85→175→1149）。
- **暗色模式**：`univerAPI.toggleDarkMode(true)` 一行生效，且自动处理显式填充色对比度（黄底自动压暗、红字自动提亮）——旧引擎路线需专门开发的能力，这里免费。
- **dispose/重建**：5 次循环堆内存稳定在 202MB 无增长，重建后功能正常 → `fileChange` 重载路径可行。
- **CommandExecuted 事件**：编辑触发 `sheet.mutation.set-range-values` 等 mutation id → dirty 检测姿势确认。

### 2. API 核实清单（全部在免费包中确认，含 d.ts 位置）

| API | 位置 | 状态 |
|---|---|---|
| `univerAPI.toggleDarkMode(bool)` | @univerjs/core facade | ✅ |
| `fWorkbook.setEditable(bool)`（只读模式） | @univerjs/sheets facade | ✅ |
| `fWorkbook.getSnapshot(): IWorkbookData` | 同上 | ✅ |
| workbook/worksheet/range 三级 permission facade | @univerjs/sheets/facade/permission | ✅ 免费（R6 解除） |
| `fWorksheet.scrollToCell(row,col)` / `fWorkbook.getScrollStateBySheetId(id)` | @univerjs/sheets-ui facade | ✅ |
| `fWorkbook.setActiveSheet()` / `fWorksheet.activate()` | @univerjs/sheets facade | ✅ |
| `univerAPI.addEvent(Event.CommandExecuted, cb)` | @univerjs/core facade（`onCommandExecuted` 已废弃） | ✅ |
| range protection UI 选项（`protectedRangeShadow` 等） | UniverSheetsCorePreset 工厂参数 | ✅ 在免费 preset |

### 3. R1 实测：exceljs 往返保真（影响最大的发现）

fixtures：openpyxl 生成 + 规范化为 Excel 标准部件命名（`test/fixtures/make_r1_fixture.py` + `normalize_fixture.py`）。

| 场景 | 结果 |
|---|---|
| **含图表文件 load** | **fork 和上游 exceljs 双双崩溃**（`drawing.anchors` undefined，worksheet-xform.js reconcile）→ **现有插件本来就打不开任何含图表的 xlsx**（此前未知的生产 bug） |
| 无图表文件 fork 往返 | **零丢失**：definedNames、条件格式、自定义 numFmt、批注全部存活（part 级 + 内容 spot check） |
| 非标准部件命名（openpyxl 原生输出） | fork 崩溃（批注路径硬编码 `xl/comments\d+.xml`）；真实 Excel 文件不受影响 |

**对计划的修正**：
- M1 导入层必须加"jszip 预清洗"：检测到 `xl/drawings/*` 含 graphicFrame 时剥离 drawing 引用再喂 ExcelJS——这使**含图表文件从"打不开"变为"可查看"**（相对现状是净改进）。
- M2 保存安全网照旧（图表文件保存前警告）；M4 的 zip 级图表回填从可选升为真正的完全保真路径。
- 增量导出模型对"ExcelJS 已建模特性"层（批注/CF/numFmt/definedNames）**完全成立**，R1 主要风险解除。

### 4. luckyexcel 对照基准

`@zwight/luckyexcel@1.1.6` 冒烟通过（jsdom 环境，浏览器专用库）：fixture → 合法 IWorkbookData（sheets/styles/sheetOrder/resources）。可作 M1 自研导入的查漏对照，不进运行时依赖。

### 5. 交付物

- `src/react/view/excel/univer/UniverSpike.tsx` — 全 preset spike 组件（`?univer-spike` 进入，懒加载独立 chunk）
- `test/r1-passthrough.test.ts`、`test/luckyexcel-baseline.test.ts` + `vitest.config.ts`（注意：exceljs fork 只有 browser ESM 构建，vitest 需 `server.deps.inline`）
- `test/fixtures/` 生成脚本（python/openpyxl，可重现）
- 新 devDependencies：vitest、jszip、jsdom、exceljs（上游，仅对照）、@zwight/luckyexcel（仅对照）；dependencies：@univerjs/presets@0.25.1

## 遗留手测项（需 F5 扩展宿主，无法自动化）

1. 生产 `vscode-webview://` 环境：`<base href>` 替换后 Univer 懒加载 chunk 与字体/CSS 是否正常加载。
2. webview 内剪贴板复制粘贴（Univer 用 clipboard API，webview 有权限模型）。
3. 中文 IME 在 Univer 单元格编辑器内的表现。
4. Ctrl+S 是否被 Univer 占用（浏览器环境未见拦截，webview 内需确认与现有 window capture 拦截的相容性）。

## 结论

Gate 通过，建议进入 M1（只读预览）。体积与冷启动在预估区间内且有明确优化手段；所有关键 API 均在免费包中核实；R1 的结论把「图表处理」从风险项变成了修复现有 bug 的机会。

---

# M1 完成记录（2026-07-06，commit a3672a6）

**验收结论：M1 全部达成。** `easyExcel.engine: "univer"`（或 URL `?engine=univer`）启用只读预览。

- **导入层**（`src/react/view/excel/univer/`）：theme_colors（主题色/tint/索引色）、import（ExcelJS→IWorkbookData 全保真装配）、sanitize（图表预清洗）、loader（xlsx/xls/ods/csv/tsv 统一入口）、adapter（生命周期/只读/暗色/缩放/视图恢复/超链接）。
- **渲染验收**（Playwright 逐 fixture，全部零控制台错误）：accent1-6 主题色字体+填充、tint 深浅、indexed 色、白色填充保留；10 种奇异数字格式原样渲染（中文日期、分数、科学计数、自定义 `0.00"m"`）；富文本逐 run（粗红+16pt 斜体）；13 种边框样式各自可辨（double 真双线）；freeze {2,2} + 合并 + 行高列宽 + 多 sheet。
- **含图表文件从「打不开」变为「可查看」**（sanitize 剥离 graphicFrame drawing part 后 ExcelJS 可 load；关键发现：光删 sheet 引用不够，drawing part 本身按文件名模式被扫描解析，必须一并移除）。
- **体积**：主包不变（638.9KB gzip），Univer 隔离在懒加载 adapter chunk（1,670KB gzip）——legacy 用户零影响。
- **测试**：17 单测（vitest）+ 6 fixture e2e。fixtures 生成脚本 `test/fixtures/make_m1_fixtures.py`。

**M1 已知边界**（按计划属后续里程碑）：Univer 引擎恒为只读（M2 编辑+增量导出）；批注/数据验证/工作表保护/图片浮层未映射（M3）；外链点击在 webview 内的拦截桥未接（M3）；legacy 引擎仍静态打包在主包（M3 默认切换时再拆）。F5 手测清单同 Phase 0 遗留项。
