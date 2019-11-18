Dim args, objExcel

Set args = WScript.Arguments
Set objExcel = OpenObject("Excel.Application")

objExcel.workbooks.Open args(0)
objExcel.Visible = True

    objExcel.Run "your module name"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
