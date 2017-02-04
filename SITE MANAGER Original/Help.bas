Attribute VB_Name = "Help"
Public Sub HelpUsingEnterprise()
On Error GoTo err
With Screen.ActiveForm

    .HelpCommonDialog.DialogTitle = "Using the Main System"
    .HelpCommonDialog.HelpFile = App.HelpFile
    .HelpCommonDialog.HelpContext = 1
    .HelpCommonDialog.HelpCommand = cdlHelpContext
    .HelpCommonDialog.ShowHelp

End With
Exit Sub
err:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description
End Sub

