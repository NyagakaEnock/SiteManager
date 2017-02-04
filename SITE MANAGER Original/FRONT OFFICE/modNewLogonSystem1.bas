Attribute VB_Name = "modNewLogonSystem1"
Option Explicit

Private rsLOGIN As ADODB.Recordset
Private rsLOGOUT As ADODB.Recordset

Public MyLoginID  As String, MyCompName As String
Public FormUnload As Form, MyControl As Control, MaForm As Form
Public ActiveTrue As Boolean

Public AllowAccess As Boolean, ThisUser As String, ThisSystem As String
Private rsTestAdmin As ADODB.Recordset
Public CountLogin As Long

Public Function LoginAccessAllowed() As Boolean
On Error GoTo err
With Screen.ActiveForm
    ThisUser = .txtUserName.Text
    Set rsFindRecord = New ADODB.Recordset
    
    rsFindRecord.Open "SELECT * FROM AdminIndividualRights WHERE UserName='" & ThisUser & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GoTo TestForAdministration
    Else
        GoTo ExamineRights
    End If
    
TestForAdministration:

    Set rsTestAdmin = New ADODB.Recordset
    
    rsTestAdmin.Open "SELECT UserGroup FROM AdminUserRegister WHERE UserName='" & ThisUser & "';", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    If rsTestAdmin.EOF And rsTestAdmin.BOF Then
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
        MsgBox "There are No Data Access Rights Assigned to Your Profile! Contact System Administrator!!", vbOKOnly + vbCritical, "Access Denied": GoTo OUTS
    ElseIf rsTestAdmin!UserGroup = "ADMIN" Then
        LoginAccessAllowed = True: GoTo OUTS
    Else
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
        MsgBox "There are No Data Access Rights Assigned to Your Profile! Contact System Administrator!!", vbOKOnly + vbCritical, "Access Denied": GoTo OUTS
    End If
    
ExamineRights:
    'determine the rights due to the user to access current system
    ThisSystem = Trim(App.EXEName)
    
    Select Case ThisSystem
    Case "SYSMGR"
        If rsFindRecord!SYSMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "JBRIEF"
        If rsFindRecord!JBRIEF = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "SITEMGR"
        If rsFindRecord!SITEMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "PURMGR"
        If rsFindRecord!PURMGR = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "TRIPMST"
        If rsFindRecord!TRIPMST = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case "FOFFICE"
        If rsFindRecord!FOFFICE = 1 Then
            LoginAccessAllowed = True
        Else
            LoginAccessAllowed = False: CountLogin = CountLogin + 1
        End If
    Case Else
        LoginAccessAllowed = False: CountLogin = CountLogin + 1
    End Select
    
OUTS:

    If Not LoginAccessAllowed Then
        MsgBox "SORRY!! You do not have the Right to Log-in-To or use this Module!!! Consult Your System Administrator!!!", vbCritical + vbOKOnly, "Module Access Denied"
    End If
    
    Set rsFindRecord = Nothing
    
End With
Exit Function
err:
ErrorMessage
End Function

Public Sub SaveLoginRecord()
On Error GoTo err
    Set rsLOGIN = New ADODB.Recordset
    rsLOGIN.Open "SELECT * FROM AdminUserLog ORDER BY LoginID;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsLOGIN
        .AddNew
        
        !UserName = CurrentUserName
        !LoginDate = Date
        !LoginTime = Format(Now, "hh:mm:ss AMPM")
        !CompName = MyCompName
        !systemused = App.EXEName
        
        .Update
    End With
    
    Exit Sub
err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then

rsLOGIN.CancelUpdate
rsLOGIN.Requery
End If
    ErrorMessage
End Sub

Public Sub UpdateLogoutRecord()
On Error GoTo err
    
    Set rsLOGOUT = New ADODB.Recordset
    rsLOGOUT.Open "SELECT * FROM AdminUserLog WHERE loginid='" & CLng(MyLoginID) & "' ORDER BY LoginID;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsLOGOUT
        !LogOutDate = Date
        !LogOutTime = Format(Now, "hh:mm:ss AMPM")
        
        .Update
        .Requery
    End With
    
    Exit Sub
err:
    If err.Number = -2147217873 Or err.Number = -2147467259 Or err.Number = -2147352571 Then

rsLOGOUT.CancelUpdate
rsLOGOUT.Requery
End If
If err.Number = 13 Or err.Number = 3704 Then Resume Next
    ErrorMessage
End Sub

Public Function GetMyLoginID() As String
On Error GoTo err
    Dim rsFindRecord As ADODB.Recordset
    Set rsFindRecord = New ADODB.Recordset
    Dim DTG As String
    DTG = Format(Date, "MMMM dd,yyyy")
    
    rsFindRecord.Open "SELECT * FROM AdminUserLog WHERE UserName='" & CurrentUserName & "' AND LoginDate='" & DTG & "' ORDER BY LoginID;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsFindRecord
    If .EOF And .BOF Then GetMyLoginID = "": GoTo OUTS
        .MoveLast
        GetMyLoginID = !LoginID
    End With
    
OUTS:
    Set rsFindRecord = Nothing
Exit Function
err:
    ErrorMessage
End Function

Public Sub CheckActiveProcess()
On Error GoTo err
    If NewRecord = True Or beditRECORD = True Then
        MsgBox "Requested Operation Cancelled due to Incomplete Transaction! You can Click Refresh to Cancel Ongoing Transactions!", vbCritical + vbOKOnly, "Transaction Monitor"
        ActiveTrue = True
    Else
        ActiveTrue = False
    End If
    Exit Sub
err:
    ErrorMessage
End Sub
