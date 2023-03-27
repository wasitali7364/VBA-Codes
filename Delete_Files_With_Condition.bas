Function DeleteSmallFiles(folder_path As String, delete_condition_in_bytes As Long)
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strPath As String
    
    'Replace with the path of the folder containing files to be deleted
    strPath = folder_path
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strPath)
    
    For Each objFile In objFolder.Files
        If objFile.Size < delete_condition_in_bytes Then
            Debug.Print objFile.Name
            'Delete the file if its size is less than condition bytes
            objFile.Delete
        End If
    Next objFile
    
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
End Function

        'Call This Function
        'DeleteSmallFiles("C:\Users\wasit.ali.CORP\Desktop\vba_delete_file_test\data\",80)
