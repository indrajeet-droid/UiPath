Dim excel As Microsoft.Office.Interop.Excel.Application
 Dim wb As Microsoft.Office.Interop.Excel.Workbook
 Dim ws As Microsoft.Office.Interop.Excel.Worksheet
 Dim rng As Microsoft.Office.Interop.Excel.Range
excel = New Microsoft.Office.Interop.Excel.ApplicationClass
wb = excel.Workbooks.Open("D:\insert table\test.xlsx")
excel.Visible=True

ws=CType(wb.Sheets("Sheet2"),Microsoft.Office.Interop.Excel.Worksheet)
ws.Delete()

ws=CType(wb.Sheets("Sheet3"),Microsoft.Office.Interop.Excel.Worksheet)
ws.Delete()

ws=CType(wb.Sheets("Sheet4"),Microsoft.Office.Interop.Excel.Worksheet)
ws.Delete()
