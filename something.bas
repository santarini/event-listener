Sub eventListenerIE()
    
    'pulls using IE

    Dim IE As Object
    Dim adrs As String
    
    
    'create a "temp" sheet
    'Sheets.Add.Name = "temp"
    
    'define adrs as URL
    adrs = "https://mail.google.com/mail/u/0/#inbox"
    
    Set IE = CreateObject("InternetExplorer.Application")
    
    IE.Top = 0
    IE.Left = 0
    IE.Width = 800
    IE.Height = 600
    IE.AddressBar = 0
    IE.StatusBar = 0
    IE.Toolbar = 0
    IE.Visible = True
    
    IE.Navigate ("https://mail.google.com/mail/u/0/#inbox")
    Do
    DoEvents
    Loop Until IE.ReadyState = 4




End Sub
