Function pre_processing(column)
'Select the worksheet first before calling this macro
Dim i, last_row As Long
Dim element
Dim cell As Range
Dim replace_text() As Variant

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

last_row = Range(column & Rows.Count).End(xlUp).Row
replace_text = Array(" co.", " inc.", " llp", " pvt.", " ltd.", " lte.", " pte.", "india", " organisation", " usa", ",", ".", " ltd", " limited", " pte", " private", " lte", _
                                " corporation", " corpration", " corp", " pvt", ")", "(", "-", "_")
On Error GoTo errHandle
For Each element In replace_text
    Range(column & "2", column & last_row).Replace element, "", , , False
Next

For Each cell In Range(column & "2", column & last_row)
    cell.Value = Application.WorksheetFunction.Trim(cell)
Next

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Done"
Exit Function

errHandle:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Function

Sub test()
    pre_processing ("A")
End Sub
