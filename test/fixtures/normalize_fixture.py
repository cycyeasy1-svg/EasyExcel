# Rewrites openpyxl-generated xlsx part names to Excel-standard naming so the
# fixtures behave like real-Excel files when read by ExcelJS (which resolves
# comment/vml parts by hardcoded path patterns, not by rels).
#   xl/comments/comment1.xml        -> xl/comments1.xml
#   xl/drawings/commentsDrawing1.vml -> xl/drawings/vmlDrawing1.vml
import sys
import zipfile

RENAMES = {
    "xl/comments/comment1.xml": "xl/comments1.xml",
    "xl/drawings/commentsDrawing1.vml": "xl/drawings/vmlDrawing1.vml",
}
# (old absolute target, new target relative to xl/worksheets/)
REL_FIXES = {
    "/xl/comments/comment1.xml": "../comments1.xml",
    "/xl/drawings/commentsDrawing1.vml": "../drawings/vmlDrawing1.vml",
    "/xl/drawings/drawing1.xml": "../drawings/drawing1.xml",
}
CT_FIXES = {
    "/xl/comments/comment1.xml": "/xl/comments1.xml",
}


def normalize(path: str) -> None:
    src = zipfile.ZipFile(path)
    items = {n: src.read(n) for n in src.namelist()}
    src.close()

    out = {}
    for name, data in items.items():
        new_name = RENAMES.get(name, name)
        if name.endswith(".rels") or name == "[Content_Types].xml":
            text = data.decode("utf-8")
            table = CT_FIXES if name == "[Content_Types].xml" else REL_FIXES
            for old, new in table.items():
                text = text.replace(old, new)
            data = text.encode("utf-8")
        out[new_name] = data

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in out.items():
            z.writestr(name, data)
    print(f"normalized {path}")


for p in sys.argv[1:]:
    normalize(p)
