# M5 fixtures for XML-patch (byte-safe save) tests.
#
# m5-shapes.xlsx  (xlsxwriter: Excel-like XML output)
#   - Sheet1 "Data": values, formula, custom numFmt, a textbox shape
#     (xdr:sp drawing — the exact thing ExcelJS drops on rewrite)
#   - Sheet2 "Notes & <Specs>": XML-escapable sheet name, untouched by tests
#
# m5-openpyxl.xlsx (openpyxl: different XML shapes — no calcChain, minimal attrs)
#   - values + formula + styled cell, exercises the second producer form
import openpyxl
import xlsxwriter

# ---------------------------------------------------------------- xlsxwriter
wb = xlsxwriter.Workbook("test/fixtures/m5-shapes.xlsx")
fmt_money = wb.add_format({"num_format": '0.00"m"', "bold": True})
fmt_head = wb.add_format({"bold": True, "font_color": "#C0392B"})

ws = wb.add_worksheet("Data")
ws.write("A1", "Month", fmt_head)
ws.write("B1", "Sales", fmt_head)
ws.write("A2", "Jan")
ws.write("B2", 10.5, fmt_money)
ws.write("A3", "Feb")
ws.write("B3", 25, fmt_money)
ws.write("A4", "Total")
ws.write_formula("B4", "=SUM(B2:B3)", fmt_money, 35.5)
ws.write("C1", True)
ws.write("D1", "a & b < c")

ws.insert_textbox("D3", "design note", {
    "width": 160, "height": 60,
    "fill": {"color": "#FFF3CD"},
})

ws2 = wb.add_worksheet("Notes & <Specs>")
ws2.write("A1", "this sheet is never edited")
wb.close()
print("wrote m5-shapes.xlsx")

# ------------------------------------------------------------------ openpyxl
owb = openpyxl.Workbook()
ows = owb.active
ows.title = "Plain"
ows["A1"] = "hello"
ows["B1"] = 42
ows["A2"] = "=B1*2"
ows["B2"] = "keep me"
ows["B2"].font = openpyxl.styles.Font(bold=True)
owb.save("test/fixtures/m5-openpyxl.xlsx")
print("wrote m5-openpyxl.xlsx")

# ------------------------------------------------- calcChain variant (Excel 形态)
# 真实 Excel 会写 calcChain.xml + Content_Types Override + workbook rels 条目；
# xlsxwriter/openpyxl 都不产出，这里按 Excel 形态注入到 m5-shapes 的副本。
import zipfile

src = zipfile.ZipFile("test/fixtures/m5-shapes.xlsx")
with zipfile.ZipFile("test/fixtures/m5-calcchain.xlsx", "w", zipfile.ZIP_DEFLATED) as out:
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == "[Content_Types].xml":
            data = data.replace(
                b"</Types>",
                b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
            )
        elif item.filename == "xl/_rels/workbook.xml.rels":
            data = data.replace(
                b"</Relationships>",
                b'<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
            )
        out.writestr(item, data)
    out.writestr(
        "xl/calcChain.xml",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B4" i="1"/></calcChain>',
    )
src.close()
print("wrote m5-calcchain.xlsx")
