Function GetThisWorkbookAbsolutePath()
Dim fso
Dim GetFolder
Set fso = CreateObject("Scripting.FileSystemObject")
GetFolder = fso.GetAbsolutePathName(".")
GetThisWorkbookAbsolutePath = GetFolder
End Function
