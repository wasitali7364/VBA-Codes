Attribute VB_Name = "GetOpenFilePath"
Option Explicit

Sub test_get_file_open()

Dim sfullpath As String

    sfullpath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx),*.xlsx", _
                          Title:="Please select an Excel file")
        
        If sfullpath <> "False" Then
        
            If Len(sfullpath) = 0 Then
               MsgBox ("No file selected")
               Exit Sub
            Else
                MsgBox (sfullpath)
                Exit Sub
            End If
            
        Else
            
            MsgBox ("No file Selected")
            Exit Sub
            
        End If

End Sub
