'Goto api.slack.com and create an App (bot). Then Webhook your channel with the App. You will get webhook url there. Use that url in this code.
Attribute VB_Name = "Slack_Alerts"
Option Explicit

Sub send_slack_message()
    Dim req As Object
    Dim url As String
    Dim msg As String
    msg = "Test Another Alert Message!"
    
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    'changed the original url for privacy
    url = "https://hooks.slack.com/services/T4f5dfd/fds2fdsfa/fdsaf52af5asdfa5dsfasdf"
    
    req.Open "POST", url, False
    
    req.setRequestHeader "Content-Type", "application/json"
    
    req.send "{""text"":""" & msg & """}"
    
'    Debug.Print (req.Status)
'    Debug.Print (req.statusText)
    
    Set req = Nothing
End Sub
