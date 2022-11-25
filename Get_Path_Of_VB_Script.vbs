Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
file_path = dir & "\Dabur.xlsm"
msgbox file_path

'Output is Dynamic
'C:\Users\wasit.ali.CORP\OneDrive - C2FO\Documents\DSF_Reconciliation\Dabur\Dabur.xlsm