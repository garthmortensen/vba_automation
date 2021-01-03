Option Explicit

Public Sub launchBrowser()

Dim IE As Object

'launch IE for upload
Set IE = CreateObject("InternetExplorer.Application")
IE.Navigate ("https://ftp.website.com/")
IE.Visible = True

End Sub
