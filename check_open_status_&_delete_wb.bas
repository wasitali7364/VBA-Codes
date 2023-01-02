Sub CheckAndDeleteWorkbook()
  Dim wb As Workbook

  ' Try to open the workbook
  On Error Resume Next
  Set wb = Workbooks.Open("C:\MyFolder\MyWorkbook.xlsx")
  On Error GoTo 0

  If wb Is Nothing Then
    ' Workbook is closed, just delete it
    Kill "C:\MyFolder\MyWorkbook.xlsx"
  Else
    ' Workbook is open, close it and delete it
    wb.Close SaveChanges:=False
    Kill "C:\MyFolder\MyWorkbook.xlsx"
  End If
End Sub
