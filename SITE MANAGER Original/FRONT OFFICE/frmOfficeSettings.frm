VERSION 5.00
Begin VB.Form frmOfficeSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SYSTEM MANAGEMENT (Microsoft Office Setup)"
   ClientHeight    =   4815
   ClientLeft      =   2850
   ClientTop       =   1860
   ClientWidth     =   8175
   Icon            =   "frmOfficeSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8175
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame2 
         Caption         =   "Choose an Application Program to Set Up"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7935
         Begin VB.OptionButton optAPP 
            BackColor       =   &H80000002&
            Caption         =   "MS Outlook"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   5925
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optAPP 
            BackColor       =   &H80000002&
            Caption         =   "MS PowerPoint"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   3990
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optAPP 
            BackColor       =   &H80000002&
            Caption         =   "MS Ecxel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton optAPP 
            BackColor       =   &H80000002&
            Caption         =   "MS Word"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.FileListBox filPath 
         BackColor       =   &H00FFC0C0&
         Height          =   1455
         Left            =   4320
         Pattern         =   "*.exe"
         TabIndex        =   8
         ToolTipText     =   "Double-Click on an Executable Program to select it."
         Top             =   2400
         Width           =   3735
      End
      Begin VB.DirListBox dirPath 
         BackColor       =   &H00FFC0C0&
         Height          =   1440
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   3855
      End
      Begin VB.DriveListBox drvPath 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   3855
      End
      Begin VB.CommandButton cmdExecute 
         BackColor       =   &H80000001&
         Caption         =   "&Save Settings"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click Here to Save the New Settings"
         Top             =   4080
         Width           =   3735
      End
      Begin VB.TextBox txtFullPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1560
         Width           =   7935
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   4320
         X2              =   8040
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label4 
         Caption         =   "Files/Programs:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Directory/Folder:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Disk Drive:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   8040
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Full Path of the Application"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmOfficeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExecute_Click()
On Error GoTo err
    Dim rsWord As ADODB.Recordset, rsExcel As ADODB.Recordset, rsPowerPoint As ADODB.Recordset, rsOutlook As ADODB.Recordset
    Dim ApplicationsPath As String
    If Me.txtFullPath = Empty Then
        MsgBox "Invalid Update! SELECT an Appropriate Executable Application Program from the Drive, Directory and File Listbox Provided Here, then Retry Saving! ", vbInformation + vbOKOnly, "Required Application"
        Me.drvPath.SetFocus
        Exit Sub
    Else
        ApplicationsPath = Trim(Me.txtFullPath.Text)
        
        If optAPP(0) Then
            Set rsWord = New ADODB.Recordset
            rsWord.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSWord" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            With rsWord
                If .RecordCount = 0 Then
                    .AddNew
                    !CODE = "AP01"
                    !CodeNumber = "MSWord"
                    !ApplicationPath = ApplicationsPath
                    .Update
                    .Requery
                ElseIf .RecordCount > 0 Then
                    If MsgBox("A Microsoft Word System Path has been set! Do you want to Edit!", vbQuestion + vbYesNo + vbDefaultButton1, "Applications Settings") = vbNo Then Exit Sub
                    Call EditMSWord
                End If
            End With
            
        ElseIf optAPP(1) Then
            Set rsExcel = New ADODB.Recordset
            rsExcel.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSExcel" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            With rsExcel
            If .RecordCount = 0 Then
                .AddNew
                !CODE = "AP02"
                !CodeNumber = "MSExcel"
                !ApplicationPath = ApplicationsPath
                .Update
                .Requery
            ElseIf .RecordCount > 0 Then
                If MsgBox("A Microsoft Excel System Path has been set! Do you want to Edit!", vbQuestion + vbYesNo + vbDefaultButton1, "Applications Settings") = vbNo Then Exit Sub
                Call EditMSExcel
            End If
            End With
            
        ElseIf optAPP(2) Then
            Set rsPowerPoint = New ADODB.Recordset
            rsPowerPoint.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSPowerPoint" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            With rsPowerPoint
            If .RecordCount = 0 Then
                .AddNew
                !CODE = "AP03"
                !CodeNumber = "MSPowerPoint"
                !ApplicationPath = ApplicationsPath
                .Update
                .Requery
            ElseIf .RecordCount > 0 Then
                If MsgBox("A Microsoft PowerPoint System Path has been set! Do you want to Edit!", vbQuestion + vbYesNo + vbDefaultButton1, "Applications Settings") = vbNo Then Exit Sub
                Call EditMSPowerPoint
            End If
            End With
            
        ElseIf optAPP(3) Then
            Set rsOutlook = New ADODB.Recordset
            rsOutlook.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSOutLook" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
            With rsOutlook
            If .RecordCount = 0 Then
                .AddNew
                !CODE = "AP04"
                !CodeNumber = "MSOutLook"
                !ApplicationPath = ApplicationsPath
                .Update
                .Requery
            ElseIf .RecordCount > 0 Then
                If MsgBox("A Microsoft Outlook System Path has been set! Do you want to Edit!", vbQuestion + vbYesNo + vbDefaultButton1, "Applications Settings") = vbNo Then Exit Sub
                Call EditOPAC
            End If
            End With
            
        Else
            MsgBox "No Valid Current Application To Reset! No Work Will Be Done!", vbCritical, "Illegal Command Execution"
            Exit Sub
        End If
    End If
        
    Me.cmdExecute.Enabled = False
    Exit Sub
err:
rsWord.CancelUpdate
rsWord.Requery
    MsgBox err.Description, vbInformation, "IDS SYSTEM"
End Sub

Private Sub EditMSWord()
On Error GoTo err
    Dim rsEdit As ADODB.Recordset
    Set rsEdit = New ADODB.Recordset
    
    rsEdit.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSWord" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsEdit
        !ApplicationPath = Me.txtFullPath.Text
        .Update
        .Requery
    End With
    
    Exit Sub
err:
rsEdit.CancelUpdate
rsEdit.Requery
    MsgBox err.Description, vbInformation, "MSOffice Settings"
End Sub

Private Sub EditMSExcel()
On Error GoTo err
    Dim rsEdit As ADODB.Recordset
    Set rsEdit = New ADODB.Recordset
    
    rsEdit.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSExcel" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsEdit
        !ApplicationPath = Me.txtFullPath.Text
        .Update
        .Requery
    End With
    
    Exit Sub
err:
rsEdit.CancelUpdate
rsEdit.Requery
    MsgBox err.Description, vbInformation, "MSOffice Settings"
End Sub

Private Sub EditOPAC()
On Error GoTo err
    Dim rsEdit As ADODB.Recordset
    Set rsEdit = New ADODB.Recordset
    
    rsEdit.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSOutLook" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsEdit
        !ApplicationPath = Me.txtFullPath.Text
        .Update
        .Requery
    End With
    
    Exit Sub
err:
rsEdit.CancelUpdate
rsEdit.Requery
    MsgBox err.Description, vbInformation, "MSOffice Settings"
End Sub

Private Sub EditMSPowerPoint()
On Error GoTo err
    Dim rsEdit As ADODB.Recordset
    Set rsEdit = New ADODB.Recordset
    
    rsEdit.Open "SELECT * FROM ParamMSOffice WHERE CodeNumber='" & "MSPowerPoint" & "' ;", cnCOMMON, adOpenKeyset, adLockOptimistic
    
    With rsEdit
        !ApplicationPath = Me.txtFullPath.Text
        .Update
        .Requery
    End With
    
    Exit Sub
err:
rsEdit.CancelUpdate
rsEdit.Requery

    MsgBox err.Description, vbInformation, "MSOffice Settings"
End Sub

Private Sub dirPath_Change()
On Error GoTo err
    filPath.FileName = dirPath.Path
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub drvPath_Change()
On Error GoTo err
    dirPath.Path = drvPath.Drive
    Exit Sub
err:
If err.Number = 68 Then
    MsgBox "The requested device is not available! Insert the Appropriate Device in the Drive!", vbCritical, "Directory"
    drvPath.Drive = "C:\"
    Exit Sub
Else
    ErrorMessage
End If
End Sub

Private Sub filPath_DblClick()
On Error GoTo err
If Me.filPath.Path = "C:\" Or Me.filPath.Path = "a:\" Or Me.filPath.Path = "d:\" Or Me.filPath.Path = "e:\" Or Me.filPath.Path = "f:\" Or Me.filPath.Path = "g:\" Or Me.filPath.Path = "h:\" Then
    Me.txtFullPath = Me.filPath.Path & Me.filPath.FileName
Else
    Me.txtFullPath.Text = Me.filPath.Path & "\" & Me.filPath.FileName
End If
    Me.cmdExecute.Enabled = True
    Exit Sub
err:
    ErrorMessage
End Sub

