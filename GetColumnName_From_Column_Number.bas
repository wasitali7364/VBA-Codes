Function GetColumnName(colNum As Integer) As String
    Dim columnName As String
    Dim modulo As Integer
    Dim quotient As Integer
    If colNum <= 0 Then
        GetColumnName = ""
        Exit Function
    End If
    Do While colNum > 0
        modulo = (colNum - 1) Mod 26
        columnName = Chr(modulo + 65) & columnName
        colNum = (colNum - modulo) \ 26
    Loop
    GetColumnName = columnName
End Function

'To use this function, simply call it with the column number as the argument:

Sub Test()
    MsgBox GetColumnName(1) ' Output: A
    MsgBox GetColumnName(27) ' Output: AA
End Sub
