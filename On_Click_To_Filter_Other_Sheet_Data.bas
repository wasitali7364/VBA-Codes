Option Explicit


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim last_rn As Long
    Dim target_val As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo errHandler
    
    last_rn = Range("B" & Rows.Count).End(xlUp).Row
    
        If Not Application.Intersect(Target, Range("B2:B" & last_rn)) Is Nothing Then
            If Range(Target.Address).Rows.Count = 1 Then
                target_val = Range(Target.Address).Value
                Sheet3.Range("A1").AutoFilter 1, target_val
                Sheet3.Select
            End If
        End If
        
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
errHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub