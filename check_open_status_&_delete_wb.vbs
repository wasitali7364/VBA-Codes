' Declare variables
Dim wb, xlApp

' Create Excel application object
Set xlApp = CreateObject("Excel.Application")

' Try to open the workbook
On Error Resume Next
Set wb = xlApp.Workbooks.Open("C:\MyFolder\MyWorkbook.xlsx")
On Error GoTo 0

If wb Is Nothing Then
  ' Workbook is closed, just delete it
  Set wb = Nothing
  Set xlApp = Nothing
  Kill "C:\MyFolder\MyWorkbook.xlsx"
Else
  ' Workbook is open, close it and delete it
  wb.Close SaveChanges:=False
  Set wb = Nothing
  Set xlApp = Nothing
  Kill "C:\MyFolder\MyWorkbook.xlsx"
End If
