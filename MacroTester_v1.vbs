Dim inputCSV, runExcel

Set inputCSV = WScript.Arguments
Wscript.echo "Processing " & inputCSV(0) & " ..."
Set runExcel = CreateObject("Excel.Application")

runExcel.Workbooks.Open inputCSV(0)
runExcel.Visible = False
runExcel.Run "TestMacro"

'NOTE: Will have security warning in Excel if not run in a trusted location
runExcel.ActiveWorkbook.Save
runExcel.ActiveWorkbook.Close(0)
runExcel.Quit