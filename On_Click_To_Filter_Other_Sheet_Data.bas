Option Explicit


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim last_rn As Long
last_rn = Range("A" & Rows.Count).End(xlUp).Row
    If Not Application.Intersect(Target, Range("A5:A" & last_rn)) Is Nothing And Target.Font.Color = 255 Then
        Dim target_val As String
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        On Error GoTo errHandler
    
        target_val = Range(Target.Address).Value
        Sheet1.Range("A1").AutoFilter 2, target_val
        Sheet1.Select
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Exit Sub
errHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    End If
    
End Sub
