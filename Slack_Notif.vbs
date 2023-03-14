Dim req
Dim url
Dim msg

msg = "Run Successfull at " & formatdatetime(now)

'Send Slack Notification of Successfull Run

  Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  
  url = "https://hooks.slack.com/services/T01A678ESG/B023ADVH9F3/JSFDsdfdsfKlfaFKotK"
  
  req.Open "POST", url, False
  
  req.setRequestHeader "Content-Type", "application/json"
  
  req.send "{""text"":""" & msg & """}"

Set fso = Nothing
Set req = Nothing
