Dim objWshShell,IE

Set objWshShell = Wscript.CreateObject("Wscript.Shell")
Set IE = CreateObject("InternetExplorer.Application")

With IE
  .Visible = True
  .Navigate "https://www.google.com/"

'Wait for Browser
  Do While .Busy
    WScript.Sleep 1000
  Loop
  .documents.getElementsByID("playpause").Click()
End With