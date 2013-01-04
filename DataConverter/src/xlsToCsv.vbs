

'This file is part of Data Converter.

'Data Converter is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by 
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful, 
'but without any warranty; without even the implied warranty of 
'merchantability or fitness for a particular purpose. See the 
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with Data Covnerter. If not, see <http://www.gnu.org/licenses/>.


' Convert any excel file to .csv
' Source: http://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line

if WScript.Arguments.Count < 2 Then
    WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToCsv SourcePath.xls Destination.csv"
    Wscript.Quit
End If
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
oExcel.DisplayAlerts = False
Dim oBook
Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))
oBook.SaveAs WScript.Arguments.Item(1), 6
oBook.Close False
'oExcel.Quit -- leave excel open because next we'll call openInExcel.vbs
'WScript.Echo "Debug: Converted xls to csv"
