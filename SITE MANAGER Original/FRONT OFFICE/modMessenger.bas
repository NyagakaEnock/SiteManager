Attribute VB_Name = "modMessenger"
Option Explicit

Public Sub ErrorMessage()
    If err.Number = 7 Then
        MsgBox "Your System is Short of Memory!" & vbCrLf & "You have too many programs running!" & vbCrLf & "Close some of the programs or Upgrade your computer!", vbInformation + vbOKOnly, "Memory Manager"
    ElseIf err.Number = 3021 Then
    
    MsgBox " There is some required data missing in the Database"
       
    
    Else
        MsgBox err.Number & vbCrLf & err.Description, vbInformation, "System Error"
    End If
End Sub

Public Sub UpdateErrorMessage()
    If err.Number = -2147217873 Then
        MsgBox "The Record Cannot be saved due to Database Primary Key Violation! A similar record already exists in the database!", vbCritical + vbOKOnly, "Canceling Update"
    ElseIf err.Number = -2147467259 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical + vbOKOnly, "Canceling Update"
    ElseIf err.Number = -2147352571 Then
        MsgBox "Update Cancelled! You cannot save this record because some required fields are blank or have invalid data!", vbCritical + vbOKOnly, "Canceling Update"
    Else
        ErrorMessage
    End If
End Sub

