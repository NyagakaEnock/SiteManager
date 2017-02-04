VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log On SECURITY Services"
   ClientHeight    =   1665
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   983.737
   ScaleMode       =   0  'User
   ScaleWidth      =   4535.108
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3600
      TabIndex        =   6
      Top             =   1140
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   1200
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1050
      TabIndex        =   1
      Top             =   135
      Width           =   3645
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1050
      TabIndex        =   4
      Top             =   1140
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2340
      TabIndex        =   5
      Top             =   1140
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   637
      Width           =   3645
   End
   Begin MSComDlg.CommonDialog HelpCommonDialog 
      Left            =   4320
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpCommand     =   11
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   637
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean, i As Integer

Private Sub ShowTopMost()
'    SetWindowPos hwnd, conHwndTopmost, 150, 250, 380, 190, conSwpActivate Or conSwpShowWindow
End Sub

Private Sub ShowNoTopMost()
'    SetWindowPos hwnd, conHwndNoTopmost, 150, 250, 380, 190, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub cmdCancel_Click()
If MsgBox("This action is going to terminate the application. Click OK to Terminate!", vbOKCancel + vbCritical + vbDefaultButton2, "System Shut Down") = vbOK Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub cmdHelp_Click()
On Error GoTo err
With Me
    .HelpCommonDialog.HelpFile = App.HelpFile
    .HelpCommonDialog.HelpContext = 43
    .HelpCommonDialog.HelpCommand = cdlHelpContext
    .HelpCommonDialog.ShowHelp
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdOk_Click()
On Error GoTo err
With Me
    If .txtUserName.Text = Empty Then
        .txtUserName.SetFocus: Beep
    ElseIf .txtPassword.Text = Empty Then
        .txtPassword.SetFocus: Beep
    Else
        If ValidUserName Then
            If ActiveUserProfile Then
                If HasModuleRights Then
                    .txtPassword.Text = GetFullEncryption
                    If ValidPassword Then
                        Call SaveLoginRecord
                        SchedulingMain.StatusBar1.Panels(2).Text = "CURRENT USER: " & " " & UCase(frmLogin.txtUserName.Text): SchedulingMain.StatusBar1.Panels(3).Text = "LOG ON TIME: " & " " & LTime: SchedulingMain.StatusBar1.Panels(4).Text = "Computer Name: " & " " & GetCompName
                        Load SchedulingMain: Unload Me
                        SchedulingMain.Show vbModal
                    Else
                        .txtPassword.Text = Empty: .txtPassword.SetFocus
                        i = i + 1: If i = 3 Then MsgBox "Too many failed logon attempts. The system shuts down...!", vbCritical + vbOKOnly, "Forced Shut Down...!": End
                    End If
                Else
                    i = i + 1: If i = 3 Then MsgBox "Too many failed logon attempts. The system shuts down...!", vbCritical + vbOKOnly, "Forced Shut Down...!": End
                End If
            Else
                i = i + 1: If i = 3 Then MsgBox "Too many failed logon attempts. The system shuts down...!", vbCritical + vbOKOnly, "Forced Shut Down...!": End
            End If
        Else
            i = i + 1: If i = 3 Then MsgBox "Too many failed logon attempts. The system shuts down...!", vbCritical + vbOKOnly, "Forced Shut Down...!": End
        End If
    End If
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Function ActiveUserProfile() As Boolean
On Error GoTo err
With Me
    Set rsFindRecord = cnSECURE.Execute("SELECT UserMaster.Active FROM UserMaster WHERE UserName='" & Trim(.txtUserName.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        ActiveUserProfile = False
    ElseIf IsNull(rsFindRecord!Active) = True Or rsFindRecord!Active = "" Then
        ActiveUserProfile = False
    ElseIf rsFindRecord!Active = "I" Or rsFindRecord!Active = "i" Or rsFindRecord!Active = "N" Then
        ActiveUserProfile = False
    ElseIf rsFindRecord!Active = "A" Or rsFindRecord!Active = "a" Or rsFindRecord!Active = "Y" Then
        ActiveUserProfile = True
    Else
        ActiveUserProfile = False
    End If
    Set rsFindRecord = Nothing
    
    If Not ActiveUserProfile Then
        MsgBox "Access Denied. The specified user profile is disabled. Consult your system administrator for assistance...!", vbCritical + vbOKOnly, "Account is Disabled...!"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Public Function HasModuleRights() As Boolean
On Error GoTo err
With Me
    If UserIsAdministrator Then
        HasModuleRights = True: SchedulingMain.mnuFreeSites.Enabled = True
    Else
        SchedulingMain.mnuFreeSites.Enabled = False
        Set rsFindRecord = cnSECURE.Execute("SELECT UserModules.allow FROM UserModules WHERE UserName='" & Trim(.txtUserName.Text) & "' AND exeName ='" & App.EXEName & "';")
        If rsFindRecord.EOF And rsFindRecord.BOF Then
            HasModuleRights = False
        ElseIf IsNull(rsFindRecord!Allow) = True Or rsFindRecord!Allow = "" Then
            HasModuleRights = False
        ElseIf rsFindRecord!Allow = "0" Or rsFindRecord!Allow = 0 Or rsFindRecord!Allow = "N" Then
            HasModuleRights = False
        ElseIf rsFindRecord!Allow = 1 Or rsFindRecord!Allow = "1" Or rsFindRecord!Allow = "Y" Then
            HasModuleRights = True
        Else
            HasModuleRights = False
        End If
        Set rsFindRecord = Nothing
    End If
    If Not HasModuleRights Then
        MsgBox "Access Denied. You Do NOT have rights to log into this module. Consult your system administrator for assistance...!", vbCritical + vbOKOnly, "Account is Disabled...!"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Function ValidUserName() As Boolean
On Error GoTo err
With Me
    Set rsFindRecord = cnSECURE.Execute("SELECT UserMaster.UserName FROM UserMaster WHERE UserName='" & Trim(.txtUserName.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        ValidUserName = False
    Else
        ValidUserName = True
    End If
    Set rsFindRecord = Nothing
    If Not ValidUserName Then
        MsgBox "Invalid Username. The specified username does not exist in the system. Please check the spelling or consult your system administrator for assistance...!", vbCritical + vbOKOnly, "Invalid Username...!"
    End If
End With
Exit Function
err:
    ErrorMessage

End Function

Private Sub SaveLoginRecord()
On Error GoTo err
With Me
    CLoginID = GetNextLoginID: CurrentUserName = Trim(.txtUserName.Text)
    Set rsNewRecord = New ADODB.Recordset
    LTime = Now
    rsNewRecord.Open "INSERT INTO UserLog(LoginID,UserName,LoginDate,LoginTime,CompName,SystemUsed)VALUES(" & CLng(CLoginID) & ",'" & Trim(.txtUserName.Text) & "','" & Format(Date, "MMMM dd,yyyy") & "','" & FormatDateTime(Now, vbLongTime) & "','" & GetCompName & "','" & App.EXEName & "');", cnSECURE, adOpenKeyset, adLockOptimistic
    Set rsNewRecord = Nothing
End With
Exit Sub
err:
    ErrorMessage
End Sub

Private Function GetNextLoginID() As Long
On Error GoTo err
    Set rsFindRecord = cnSECURE.Execute("SELECT MAX(LoginID) AS LastID FROM UserLog WHERE LoginID IS NOT NULL;")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        GetNextLoginID = 1
    ElseIf IsNull(rsFindRecord!lastid) = True Then
        GetNextLoginID = 1
    Else
        GetNextLoginID = CLng(rsFindRecord!lastid + 1)
    End If
    Set rsFindRecord = Nothing
Exit Function
err:
    ErrorMessage
End Function

Private Function ValidPassword() As Boolean
On Error GoTo err
With Me
    Set rsFindRecord = cnSECURE.Execute("SELECT UserMaster.password FROM UserMaster WHERE UserName='" & Trim(.txtUserName.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        ValidPassword = False
    ElseIf IsNull(rsFindRecord!Password) Or rsFindRecord!Password = "" Then
        ValidPassword = False
    ElseIf rsFindRecord!Password <> Trim(.txtPassword.Text) Then
        ValidPassword = False
    ElseIf rsFindRecord!Password = Trim(.txtPassword.Text) Then
        ValidPassword = True
    Else
        ValidPassword = False
    End If
    Set rsFindRecord = Nothing
    If Not ValidPassword Then
        MsgBox "Invalid Password. Your username and password cannot be resolved. Passwords should be Lower Case alphabetical characters or number or a combination of both...!", vbCritical + vbOKOnly, "Invalid Password...!"
    End If
End With
Exit Function
err:
    ErrorMessage
End Function

Private Function UserIsAdministrator() As Boolean
On Error GoTo err
With Me
    Set rsFindRecord = cnSECURE.Execute("SELECT UserMaster.GroupNo FROM UserMaster WHERE UserName='" & Trim(.txtUserName.Text) & "';")
    If rsFindRecord.EOF And rsFindRecord.BOF Then
        UserIsAdministrator = False
    Else
        If IsNull(rsFindRecord!GroupNo) = True Or rsFindRecord!GroupNo = "" Then
            UserIsAdministrator = False
        ElseIf rsFindRecord!GroupNo = "A" Then
            UserIsAdministrator = True
        Else
            UserIsAdministrator = False
        End If
    End If
    Set rsFindRecord = Nothing
End With
Exit Function
err:
    ErrorMessage
End Function

Private Sub Form_Activate()
    'If Not SystemActivated Then End
With Me
    If .txtPassword.Text = Empty Then
        .txtUserName.SetFocus: .txtUserName.SelStart = 0: .txtUserName.SelLength = Len(.txtUserName.Text)
    End If
End With
End Sub

Private Sub Form_Deactivate()
    Call ShowNoTopMost
End Sub

Private Sub Form_GotFocus()
    Call ShowTopMost
End Sub

Private Sub Form_Initialize()
'    If Not SystemActivated Then End
    
End Sub

Private Sub Form_Load()
With Me
 DateExpiry
End With
End Sub

Private Sub Timer1_Timer()
With Me
'    Dim q, x
'    q = Timer1.Interval
'    If Timer1.Interval = 10000 Then
'        If Not NoNewMessages Then
'            Call Main
'        End If
'    End If
End With
End Sub
Private Sub DateExpiry()
  With Me
    Dim expirydate As String
        expirydate = "4/04/2008"
        expirydate = Format(expirydate, "yyyy/mm/dd")
         If Format(Date, "yyyy/mm/dd") <= expirydate Then Exit Sub
           OpenSECUREConnection
           OpenODBCConnection
           .txtUserName.Text = UserNameGet
           i = 0
  End With
End Sub

