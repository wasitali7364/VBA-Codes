'Input Excel File's Full Path
  PrimaryExcelFilePath = "C:\Users\wasit.ali.CORP\Downloads\demo\Apple Supplier Stats-test.xlsm"

'Input Module/Macro name within the Excel File
  MacroPath = "Module1.refresh_connection_and_pivot_table"

'Create a function to get Date Format in Specific Format
  Function myDateFormat(myDate)
      d = TwoDigits(Day(myDate))
      m = TwoDigits(Month(myDate))    
      y = Year(myDate)
      myDateFormat= m & "-" & d & "-" & y
  End Function

  Function TwoDigits(num)
      If(Len(num)=1) Then
          TwoDigits="0"&num
      Else
          TwoDigits=num
      End If
  End Function

'Create a new file name
  new_file_name = "C:\Users\wasit.ali.CORP\Downloads\demo\Apple Monthly Report " & myDateFormat(now) & ".xlsx"

'Create an instance of Excel
  Set ExcelApp = CreateObject("Excel.Application")

'Do you want this Excel instance to be visible?
  ExcelApp.Visible = True 'or "False"

'Prevent any App Launch Alerts (ie Update External Links)
  ExcelApp.DisplayAlerts = False

'Open Excel File
  Set wb = ExcelApp.Workbooks.Open(PrimaryExcelFilePath)

'Execute Macro Code
  ExcelApp.Run MacroPath

'Save Excel File (if applicable)
  'wb.Save

'Copy to new workbook
  ExcelApp.ActiveWorkbook.SaveAs new_file_name, 51

'Reset Display Alerts Before Closing
  ExcelApp.DisplayAlerts = True

'Close Excel File
  ExcelApp.ActiveWorkbook.Close

'End instance of Excel
  ExcelApp.Application.Quit


'set instance of Excel to Nothing
  Set ExcelApp = Nothing


'Leaves an onscreen message!
 ' MsgBox "Your Automated Task successfully ran at " & TimeValue(Now), vbInformation