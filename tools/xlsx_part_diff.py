#!/usr/bin/env python3
# 部件级 xlsx 对比：验证「除被编辑的内容外一个字节都没变」的独立审计工具。
#
# 用法：
#   python tools/xlsx_part_diff.py 原文件.xlsx 保存后.xlsx
#
# 典型流程：编辑前先复制一份原文件，在 EasyExcel 里编辑并保存，
# 然后用本工具对比副本与保存后的文件。预期输出：
#   - 只有被编辑 sheet 的 xl/worksheets/sheetN.xml 变化
#   - 改过样式时多一个 xl/styles.xml（append-only）
#   - 有公式且改过值时 calcChain 被删除（Excel 打开自动重建并重算）
#   - 其余部件（图形/图表/透视表/宏/打印设置/主题……）全部 identical
#
# 退出码：两文件部件完全一致时 0，有差异时 1（差异本身可能是预期的）。
import sys
import zipfile


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__ or "usage: xlsx_part_diff.py a.xlsx b.xlsx")
        print("usage: python tools/xlsx_part_diff.py <before.xlsx> <after.xlsx>")
        return 2

    za, zb = (zipfile.ZipFile(p) for p in sys.argv[1:3])
    # 目录条目（尾部 /）不承载内容，忽略
    names_a = {n for n in za.namelist() if not n.endswith("/")}
    names_b = {n for n in zb.namelist() if not n.endswith("/")}

    identical, changed = [], []
    for name in sorted(names_a & names_b):
        a, b = za.read(name), zb.read(name)
        if a == b:
            identical.append(name)
        else:
            changed.append((name, len(a), len(b)))

    removed = sorted(names_a - names_b)
    added = sorted(names_b - names_a)

    print(f"identical parts : {len(identical)}")
    for name in identical:
        print(f"  = {name}")
    if changed:
        print(f"CHANGED parts   : {len(changed)}")
        for name, la, lb in changed:
            print(f"  ~ {name}  ({la} -> {lb} bytes)")
    if removed:
        print(f"REMOVED parts   : {len(removed)}")
        for name in removed:
            print(f"  - {name}")
    if added:
        print(f"ADDED parts     : {len(added)}")
        for name in added:
            print(f"  + {name}")

    if not changed and not removed and not added:
        print("RESULT: files are part-level identical")
        return 0
    print("RESULT: differences found (check they match what you edited)")
    return 1


if __name__ == "__main__":
    sys.exit(main())
