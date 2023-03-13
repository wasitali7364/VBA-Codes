Dim folderName
Dim csv_file_path
Dim pay_date
Dim mail_to
Dim carbon_copy
Dim CountFiles
Dim first_row
Dim last_row
Dim income_mismatch_count

income_mismatch_count = 0

mail_to = "wasit.ali@c2fo.com"
carbon_copy = "wasit.ali@c2fo.com"

'Create File System Object
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  GetCurrentFolder = objFSO.GetAbsolutePathName(".")
  folderName = GetCurrentFolder & "\Files\"

'Count the number of files in a folder
  Set objFiles = objFSO.GetFolder(folderName).Files
  If Err.Number <> 0 Then
      CountFiles = 0
  Else
      CountFiles = objFiles.Count
  End If

'msgbox CountFiles & " Files In The Folder: " & folderName

'Forward Only if Files Are Downloaded From S3 else Stop This Macro
  If CountFiles > 0 Then
  
  'Create an instance of Excel
    Set ExcelApp = CreateObject("Excel.Application")

  'Do you want this Excel instance to be visible?
    ExcelApp.Visible = True 'or False

  'Prevent any App Launch Alerts (ie Update External Links)
    ExcelApp.DisplayAlerts = False

  'Use For Each loop to loop through each file in the folder
    For Each objFile In objFSO.GetFolder(folderName).Files
        if instr(objFile.Name,"merged_award") then
          csv_file_path = objFile.Path
          'Open Excel File
            Set wb = ExcelApp.Workbooks.Open(csv_file_path)
          'extract pay_date value
            pay_date = wb.Worksheets(1).Range("P2").value

          'find last row
            'last_row = wb.Worksheets(1).Range("D1048576").End(xlUp).Row
            last_row = 5000

          'create calculated column
            wb.Worksheets(1).Range("Q:Q").EntireColumn.Insert
            wb.Worksheets(1).Range("Q:Q").EntireColumn.NumberFormat = "General"
            wb.Worksheets(1).Range("Q1").Value = "DPE"

            wb.Worksheets(1).Range("R:R").EntireColumn.Insert
            wb.Worksheets(1).Range("R:R").EntireColumn.NumberFormat = "General"
            wb.Worksheets(1).Range("R1").Value = "Income"

            wb.Worksheets(1).Range("S:S").EntireColumn.Insert
            wb.Worksheets(1).Range("S:S").EntireColumn.NumberFormat = "General"
            wb.Worksheets(1).Range("S1").Value = "Income Check"

            for first_row = 2 to last_row
              if wb.Worksheets(1).Range("D" & first_row) <> "" then
                 wb.Worksheets(1).Range("Q" & first_row) = wb.Worksheets(1).Range("G" & first_row) - wb.Worksheets(1).Range("P" & first_row)
                 wb.Worksheets(1).Range("R" & first_row) = Round((wb.Worksheets(1).Range("T" & first_row) * wb.Worksheets(1).Range("V" & first_row) * wb.Worksheets(1).Range("Q" & first_row))/36000,2)
                 wb.Worksheets(1).Range("S" & first_row) = Round(wb.Worksheets(1).Range("R" & first_row), 2) = Round(wb.Worksheets(1).Range("N" & first_row),2)
                 if wb.Worksheets(1).Range("S" & first_row) = "FALSE" then
                    income_mismatch_count = income_mismatch_count + 1
                 end if
              end if
            next

            wb.Worksheets(1).Range("A1").AutoFilter
            wb.Worksheets(1).Cells.EntireColumn.AutoFit

            wb.Worksheets(1).Range("Q:Q").interior.Color = RGB(169, 229, 187) 'Green
            wb.Worksheets(1).Range("R:R").interior.Color = RGB(252, 246, 177) 'Yellow
            wb.Worksheets(1).Range("S:S").interior.Color = RGB(216, 207, 175) 'Dutch White

            wb.SaveAs folderName & "Emami Award File Calculation.xlsx", 51

          'Reset Display Alerts Before Closing
            ExcelApp.DisplayAlerts = True

            'wb.Save
          'Close Excel File
            wb.Close
        end if
    Next

  'End instance of Excel
    ExcelApp.Application.Quit

  'Create an instance of Outlook
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)

  'Create Email Content
    objMail.To = mail_to
    objMail.Cc = carbon_copy
    objMail.Subject = "Award File_C2FO_" & pay_date
    objMail.Body = "Dear Ashu g," & vbnewline & vbnewline & "Hope you're doing good !" & vbnewline & _
    "PFA the award file for pay date of " & formatdatetime(cdate(pay_date),1) & "." & vbnewline & _
    "The Amount is matching as per APR Calculation" & "." & vbnewline & vbnewline & _
    "Regards," & vbnewline &"Data & Business Intelligence Team India," & vbnewline & "Wasit Ali | Rajat Pandey" _
    & vbnewline & vbnewline & "Income Mismatch Invoice Count : " & income_mismatch_count

    For Each objFile In objFSO.GetFolder(folderName).Files
      if instr(objFile.Name, "Emami Award File Calculation") then
        objMail.Attachments.Add(objFile.Path)
      end if
    Next

    objMail.Display

  'End instance of Outlook
  ' objOutlook.Application.Quit

  'set instances to Nothing
    Set objMail = Nothing
    Set objOutlook = Nothing
    Set objFSO = Nothing
  'set instance of Excel to Nothing
    Set ExcelApp = Nothing

  'exit if Condition
  End If


'Leaves an onscreen message!
'MsgBox "Your Automated Task successfully ran at " & TimeValue(Now), vbInformation
