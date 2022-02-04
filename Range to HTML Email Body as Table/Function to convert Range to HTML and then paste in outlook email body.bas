Attribute VB_Name = "Module1"
Option Explicit


Sub testEmail()

Dim rng As Range
Dim emailApplication As Object
Dim emailItem As Object
Dim sendTo As String
Dim workFile As String
Dim newFile As String
Dim delFile As String

workFile = ActiveWorkbook.Name


sendTo = VBA.InputBox("Enter a valid Email Address", "Receiver's Email ID")
On Error GoTo errHandle
If sendTo <> "" Then
    
    Workbooks.Add.SaveAs ("C:\Users\wasit.ali.CORP\Desktop\test_" & Format(Now, "dd-mm-yy h-mm-ss"))
    newFile = ActiveWorkbook.Name
    
    Workbooks(workFile).Activate
    Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
    Workbooks(newFile).Activate
    ActiveWorkbook.ActiveSheet.Paste
    ActiveWorkbook.Save
    delFile = ActiveWorkbook.FullName
    ActiveWorkbook.Close
    
    Workbooks(workFile).Activate
    Set rng = Range("C21").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Application.CutCopyMode = False
    Set emailApplication = CreateObject("Outlook.Application")
    Set emailItem = emailApplication.CreateItem(0)
    
    With emailItem
        .to = sendTo
        .Subject = "Interested Supplier Data"
        .Attachments.Add delFile
        .HTMLBody = "Hi Team," & "<br>" & "<br>" & _
        "Please find attached the list of supplier and their invoices, who are interested to avail early payment facility." _
        & "<br>" & "Request you to kindly book these invoices for early payment." _
        & RangetoHTML(rng) _
        & "<br>" & "Regards,"
        emailItem.display
    End With
    
    Set emailItem = Nothing
    Set emailApplication = Nothing
    Kill (delFile)
    Exit Sub

End If

errHandle:
Set emailItem = Nothing
Set emailApplication = Nothing
Application.CutCopyMode = False
MsgBox "There seems to be an error" & vbCrLf & Err.Description
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2021
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
