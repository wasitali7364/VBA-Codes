Attribute VB_Name = "iterate_files_in_folder"
Option Explicit

Sub LoopThroughFiles()
    Dim StrFile As String
    StrFile = Dir("C:\Users\wasit.ali.CORP\Downloads\itereate_folder_test\files\*.xls*")
    Do While Len(StrFile) > 0
        Debug.Print StrFile
        StrFile = Dir
    Loop
End Sub
