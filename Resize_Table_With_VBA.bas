Attribute VB_Name = "Module2"
Option Explicit

Sub resize_table_range()

 'Declare Variables
    Dim oSheetName As Worksheet
    Dim sTableName As String
    Dim loTable As ListObject
    
    'Define Variable
    sTableName = "test_table"
    
    'Define WorkSheet object
    Set oSheetName = Sheets("Sheet1")
    
    'Define Table Object
    Set loTable = oSheetName.ListObjects(sTableName)
        
    'Resize the table
    loTable.Resize Range("D15:L17")
    
    
End Sub
