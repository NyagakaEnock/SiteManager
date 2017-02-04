VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLogin1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Front Office System (ODAS Plus)"
   ClientHeight    =   2895
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5895
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLogin1.frx":0442
   ScaleHeight     =   1710.463
   ScaleMode       =   0  'User
   ScaleWidth      =   5535.086
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin MSComDlg.CommonDialog HelpCommonDialog 
         Left            =   240
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1545
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1110
         Width           =   3885
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000000&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1860
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H80000000&
         Caption         =   "O&K"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Administrators Only"
         Top             =   1800
         Width           =   1860
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1560
         TabIndex        =   0
         ToolTipText     =   "Administrators Only"
         Top             =   360
         Width           =   3885
      End
      Begin VB.Label lblLabels 
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1125
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         Caption         =   "USER NAME:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   375
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLOGIN As ADODB.Recordset

Dim rsLOGOUT As ADODB.Recordset

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
        .HelpCommonDialog.DialogTitle = "Using the Main System"
        .HelpCommonDialog.HelpFile = App.HelpFile
        .HelpCommonDialog.HelpContext = 43
        .HelpCommonDialog.HelpCommand = cdlHelpContext
        .HelpCommonDialog.ShowHelp
    End With
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub cmdOK_Click()
On Error GoTo err
If Me.txtPassword.Text = Empty Or Me.txtUserName.Text = Empty Then Exit Sub
If CountLogin >= 3 Then MsgBox "TOO MANY FAILED LOGON ATTEMPTS. The System Will Now SHUT DOWN!! Consult Your System Administrator for Assistance!!!", vbCritical + vbOKOnly, "FORCED SYSTEM SHUTDOWN": End
    Set rsLOGIN = New ADODB.Recordset
    rsLOGIN.Open "SELECT * FROM AdminUserRegister WHERE Username = '" & txtUserName.Text & "';", cnCOMMON, adOpenStatic, adLockOptimistic
    
    With rsLOGIN
    If .RecordCount = 0 Then
        If MsgBox("Unknown Username or User Not Registred! Contact System Administrator! Choose Retry to try again!", vbRetryCancel + vbCritical, "Login Security") = vbCancel Then
            Unload Me: Exit Sub
        Else
            CountLogin = CountLogin + 1
            Me.txtPassword = Empty
            Me.txtUserName = Empty
            Me.txtUserName.SetFocus
        End If
    ElseIf .RecordCount > 0 Then
        If !PEnabled = "N" Then
            If MsgBox("This User Account is Currently Disabled! You Cannot log on to the system!" & vbCrLf & "Consult the System Administrator!;", vbInformation + vbRetryCancel + vbDefaultButton1, "Disabled Account") = vbCancel Then Unload Me
            CountLogin = CountLogin + 1
            Me.txtPassword = Empty
            Me.txtUserName.SetFocus
            Me.txtUserName.SelStart = 0
            Me.txtUserName.SelLength = Len(Me.txtUserName.Text)
            
            MsgBox "15-Seconds Timeout in Progress!", vbOKOnly + vbCritical, "Timing Out"
            Timer1.Enabled = True
            Exit Sub
        End If
    End If
       
    Me.txtUserName.SetFocus
 
    If .BOF And .EOF Then Exit Sub
    If LoginAccessAllowed = True Then
    
    frmLogin.txtPassword.Text = GetFullEncryption
    If LCase(!UserName) = frmLogin.txtUserName.Text And LCase(!Password) = frmLogin.txtPassword.Text Then
        CurrentUserName = Trim(frmLogin.txtUserName.Text)
        MyCompName = Trim(GetCompName)
        
        Call SaveLoginRecord
        
        Load ALISFOManager: Unload Me
        ALISFOManager.Show 1: CountLogin = 0
        ALISFOManager.Enabled = True
        ALISFOManager.StatusBar1.Panels(2).Text = "CURRENT USER: " & " " & UCase(frmLogin.txtUserName.Text)
            
    Else
        If MsgBox("Your User Name and Password cannot be resolved. Consult the system administrator...or Re-enter the values!;", vbRetryCancel + vbExclamation, "Approval Validation") = vbRetry Then
            CountLogin = CountLogin + 1
            txtPassword.SetFocus
            txtPassword.SelStart = 0
            txtPassword.SelLength = Len(txtPassword.Text)
        Else
            MsgBox "Logon Terminated!", vbOKOnly + vbCritical, "Terminating"
            Unload Me: End
        End If
    End If
    End If
    End With
        
    Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Form_Load()
On Error GoTo err
    Call OpenConnection
    'Call openConnectionInvent
'    If Not SystemActivated Then
'        Unload Me
'        End
'    Else
        txtUserName.Text = UserNameGet
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
'    End If
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Timer1_Timer()
''On Error Resume Next
    With Timer1
    If .Interval = 15000 Then
        MsgBox "Logon Timeout! The application will now shut down!", vbOKOnly + vbCritical, "Logon Timeout"
        Timer1.Enabled = False
        Unload Me
    End If
    End With
End Sub

Private Sub txtUserName_GotFocus()
''On Error Resume Next
    Me.txtUserName.SetFocus: Me.txtUserName.SelStart = 0
    Me.txtUserName.SelLength = Len(Me.txtUserName.Text)
End Sub

Private Sub txtUserName_LostFocus()
    Dim CCase As String
    CCase = LCase(Me.txtUserName.Text)
    txtUserName.Text = CCase
End Sub

