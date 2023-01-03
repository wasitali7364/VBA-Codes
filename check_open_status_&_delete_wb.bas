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

'How to use this code in Production?
'    'check if path exists
'    If Len(Dir("C:\Users\" & Environ("username") & "\Downloads\Missing_ORGID.xlsx")) = 0 Then
'        'path does not exist so create new file
'        wb.SaveAs "C:\Users\" & Environ("username") & "\Downloads\Missing_ORGID.xlsx"
'    Else
'        'path does not exist so check open status of file then close it and delete it
'        Call CheckAndDeleteWorkbook
'        'create new file
'        wb.SaveAs "C:\Users\" & Environ("username") & "\Downloads\Missing_ORGID.xlsx"
'    End If
