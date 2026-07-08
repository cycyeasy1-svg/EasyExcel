# Change Log

EasyExcel 的变更历史。格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循 [Semantic Versioning](https://semver.org/lang/zh-CN/)。

## [1.0.0] - 2026-07-07

首次发布。

### Added

- `.xlsx` / `.xlsm` / `.xls` / `.csv` / `.tsv` / `.ods` 的预览与编辑
- 基于 Univer 的电子表格引擎（多工作表、样式、合并单元格、公式、超链接、数据验证、工作表保护信息、图片、条件格式的展示）
- 查找与替换
- `.xlsx` / `.xlsm` 覆盖保存采用 XML 差分补丁方式的字节级无损保存——未编辑的部件（图形、图表、透视表、宏、打印设置等）一个字节都不会改变
- 为守住字节级无损保证而禁用结构性操作（行/列/工作表的插入删除、合并、行高列宽、冻结、数据验证与条件格式的编辑）
- 保存 / 另存为按钮与保存完成提示
- CSV / TSV 编码自动检测
- 界面语言切换（简体中文 / English，默认中文）
