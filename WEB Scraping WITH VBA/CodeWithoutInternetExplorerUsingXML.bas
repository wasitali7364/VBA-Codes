Attribute VB_Name = "Module1"
Option Explicit

'Import This Object Library
' Tools -> Refrences -> Microsoft HTML Object Library
' Tools -> Refrences -> Microsoft XML

Sub LoadWebsite()
Attribute LoadWebsite.VB_ProcData.VB_Invoke_Func = " \n14"

Dim XMLPage As New MSXML2.XMLHTTP60

'Create a variable which is going to hold the HTML Document We are going to Retrieve
Dim HTMLDoc As New MSHTML.HTMLDocument

'IHTMLElement holds refrence to any individual item on an entire HTMLDocument
Dim HTMLInput As MSHTML.IHTMLElement

'Get Collection of Tags
Dim HTMLTags As MSHTML.IHTMLElementCollection
Dim HTMLTag As MSHTML.IHTMLElement

'Get Collection of List
Dim HTMLLists As MSHTML.IHTMLElementCollection
Dim HTMLList As MSHTML.IHTMLElement


'Start Row and column
Dim RowNum As Long, ColNum As Integer


XMLPage.Open "GET", "https://www.x-rates.com/average/?from=USD&to=INR&amount=1&year=2021", False
XMLPage.send


'Set Refrence to the HTMLDocument in the Variable
HTMLDoc.body.innerHTML = XMLPage.responseText

Set HTMLTags = HTMLDoc.getElementsByClassName("OutputLinksAvg")

For Each HTMLTag In HTMLTags
    
    Worksheets.Add
    
    Set HTMLLists = HTMLDoc.getElementsByClassName("avgMonth")
    RowNum = 1
    For Each HTMLList In HTMLLists
        ColNum = 1
        Cells(RowNum, ColNum).Value = HTMLList.innerHTML
        RowNum = RowNum + 1
    Next HTMLList

    Set HTMLLists = HTMLDoc.getElementsByClassName("avgRate")
    RowNum = 1
    For Each HTMLList In HTMLLists
        ColNum = 2
        Cells(RowNum, ColNum).Value = HTMLList.innerHTML
        RowNum = RowNum + 1
    Next HTMLList
Next HTMLTag
End Sub
