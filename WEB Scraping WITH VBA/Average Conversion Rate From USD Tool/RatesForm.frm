VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RatesForm 
   Caption         =   "Choose Year & Currency"
   ClientHeight    =   2690
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   2910
   OleObjectBlob   =   "RatesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RatesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub GetAvgRate_Click()

    If Year.Value = "" Then
        MsgBox "You must pick Year"
        Exit Sub
    End If
    
    If ToCurrency.Value = "" Then
        MsgBox "You must pick Currency"
        Exit Sub
    End If
    
    If Year.Value < 2012 Then
        MsgBox "You must pick Year after 2011"
        Exit Sub
    End If
    
    GetRates ToCurrency.Value, Year.Value
    

End Sub

Private Sub UserForm_Initialize()

    With Year
        .AddItem "2022"
        .AddItem "2021"
        .AddItem "2020"
        .AddItem "2019"
        .AddItem "2018"
        .AddItem "2017"
        .AddItem "2016"
        .AddItem "2015"
        .AddItem "2014"
        .AddItem "2013"
        .AddItem "2012"
    End With

    With ToCurrency
        .AddItem "INR"
        .AddItem "GBP"
        .AddItem "CNY"
        .AddItem "AUD"
        .AddItem "EUR"
    End With

End Sub

