Sub add_sheet_with_name()
    ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)).name = "Test"
End Sub