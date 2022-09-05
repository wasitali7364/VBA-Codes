Attribute VB_Name = "Trim_Selection"
Option Explicit

Sub TrimSelection()

    Dim cell As Range
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    For Each cell In Selection
    
        If Not cell.HasFormula Then
        
            cell.Value = Application.WorksheetFunction.Substitute(cell.Value, Chr(160), " ")
            cell.Value = Application.WorksheetFunction.Substitute(cell.Value, Chr(150), "-")
            cell.Value = Trim(cell.Value)
        
        End If
            
    Next cell
    
    VBA.MsgBox "Done", vbOKOnly, "Trim the Selected Cells"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
errHandler:
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
VBA.MsgBox (Err.Description)

End Sub
