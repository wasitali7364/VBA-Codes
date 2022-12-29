dim fso, folder, Subfolder, files_count
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetAbsolutePathName(".") & "\Data"
files_count = fso.GetFolder(folder).Files.Count
msgbox folder
msgbox files_count
