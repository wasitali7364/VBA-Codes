Sub Disable_Background_Refresh()
Dim dbr As Long
    With ThisWorkbook
        For dbr = 1 To .Connections.Count
          If .Connections(dbr).Type = xlConnectionTypeOLEDB Then
            On Error Resume Next
            .Connections(dbr).OLEDBConnection.BackgroundQuery = False
          End If
        Next dbr
    End With
End Sub