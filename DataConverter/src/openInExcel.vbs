
' Open a file in excel and set columns C and D to show two decimal places.

if WScript.Arguments.Count < 1 Then
    WScript.Echo "Error! Please specify the source path and the destination. Usage: openInExcel SourcePath.xls"
    Wscript.Quit
End If

Dim oExcel

' Try to get an existing Excel application object
On Error resume next
Set oExcel = getobject(,"Excel.Application")
'WScript.Echo "Debug: got existing excel object"

' If there wasn't one, create a new one
If Err.Number <> 0 Then
    Set oExcel = CreateObject("Excel.Application")
    'WScript.Echo "Debug: created new excel object"
End If
On Error Goto 0

' Open the file
Dim oBook
Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))

' Set columns c and d to show 2 decimal places 
oExcel.Columns("C:D").Select
oExcel.Selection.NumberFormat = "0.00_);(0.00)"
oExcel.Range("A1").Select

oExcel.ScreenUpdating = True
oExcel.Visible = True
