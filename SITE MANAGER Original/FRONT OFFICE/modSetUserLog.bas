Attribute VB_Name = "modSetUserLog"
Option Explicit

'update the user log every time a user logs on
'update the user log every time a user logs out

Public rsLoginDate As ADODB.Recordset, LoginDate As Date, LoginTime As Date
Public rsLogoutDate As ADODB.Recordset, LogOutDate As Date, LogOutTime As Date
Public LogUserName As String, LoginID, GenerateLoginID As String
Public OutUserName As String, OutDate As Variant, OutTime As Variant
Public OutLoginID As String, InLoginID As Variant
Public DateToday As Date
Public LoggedOut As Boolean

Public Sub UpdateLogin()
'write login information at login time
Set rsLoginDate = New ADODB.Recordset
rsLoginDate.Open "SELECT sysuserlog.* FROM sysuserlog;", cnCOMMON, adOpenKeyset, adLockOptimistic

LoginDate = Format(Now, "dd-mmm-yyyy")
LoginTime = Format(Now, "hh:mm:ss AMPM")
LogUserName = frmLogin.txtUserName.Text
LoginID = GenerateLoginID

With rsLoginDate
    .AddNew
        !LoginID = LoginID
        !UserName = LogUserName
        !LoginDate = LoginDate
        !LoginTime = LoginTime
    .Update
    .Requery
End With

GoTo SetOuts

SetOuts:
    CurrentUserName = frmLogin.txtUserName.Text
    InLoginID = rsLoginDate!LoginID
    DateToday = rsLoginDate!LoginDate
    
'rsLoginDate.Close
    
End Sub

Public Sub UpdateLogout()
On Error GoTo err
'pick the current/active user
'check to ensure the current date is same as today
'check to ensure the logintime is the same
On Error GoTo err
Set rsLogoutDate = New ADODB.Recordset
Dim strLogoutDate As String

DateToday = Format(Now, "dd mmmm yyyy")
strLogoutDate = "SELECT sysuserlog.* from sysuserlog WHERE UserName LIKE '" & CurrentUserName & "';"

rsLogoutDate.Open strLogoutDate, cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsLogoutDate

    If .RecordCount < 1 Then Exit Sub
        .MoveLast
        OutUserName = !UserName
        OutLoginID = !LoginID
        OutDate = !LoginDate
        OutTime = !LoginTime
        
    End With

If OutUserName = CurrentUserName And OutLoginID = InLoginID Then
   GoTo Updates
End If

Updates:
With rsLogoutDate
    !LogOutDate = Format(Now, "dd-mmm-yyyy")
    !LogOutTime = Format(Now, "hh:mm:ss AMPM")
    .Update
    .Resync
    .Requery
End With
'rsLoginDate.Close
'rsLogoutDate.Close
Exit Sub
err:
    ErrorMessage
End Sub
