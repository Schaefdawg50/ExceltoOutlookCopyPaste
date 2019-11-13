Option Explicit

Sub MLF_Eric()

Dim wb As ThisWorkbook, p As String, ws As Worksheet, rng As Range, new_wb As Workbook, rng2 As Range, LastRow As Long, LastColumn As Long

Set wb = ThisWorkbook
Set ws = wb.Sheets("Input")

Set rng = Range("A1")
    Worksheets("Input").UsedRange
    LastRow = rng.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = rng.SpecialCells(xlCellTypeLastCell).Column
    ws.Range(rng, rng.Cells(LastRow, LastColumn)).Select

p = "C:\Users\" & Environ("username") & "\Desktop\MLFSpreadsheet.html"

Workbooks.Add
Set new_wb = ActiveWorkbook

ThisWorkbook.Activate
rng.Copy (rng)
Workbooks("Input").Worksheets("Sheet1").Range("A1").Copy
    Workbooks("Book2").Worksheets("Sheet1").Range ("A1")
'rng.SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("Input").Range("A1")
'new_wb.Activate

'ActiveCell.PasteSpecial xlPasteAll
ActiveCell.PasteSpecial xlPasteValues
ActiveCell.PasteSpecial xlPasteFormats
ActiveCell.PasteSpecial xlPasteColumnWidths

new_wb.PublishObjects.Add(xlSourceRange, p, new_wb.Sheets(1).Name, new_wb.Sheets(1).UsedRange.Address, xlHtmlStatic).Publish (True)

Dim readme As Variant

Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
Dim final_file As Scripting.TextStream

Set final_file = fso.OpenTextFile(p, ForReading)
readme = final_file.ReadAll
Dim o As Outlook.Application
Set o = New Outlook.Application
Dim omail As Outlook.MailItem
Set omail = o.CreateItem(olMailItem)
'
With omail

    .To = "josh.schaefer@mavtechglobal.com"
    .Subject = "Test"
    .HTMLBody = "Newest spreadsheet" & "Please see below" & "<br>" & "<table align = left>" & readme & "</table>"
    '.Send
    .Display
    
End With




End Sub





