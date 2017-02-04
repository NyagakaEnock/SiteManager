VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASMSiteSchedule 
   Caption         =   "SITE SCHEDULE"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTo 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdFree 
      BackColor       =   &H00C0C000&
      Caption         =   "FREE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9763
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmODASMSiteSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFree_Click()
On Error GoTo err
    With Me
        If .txtFrom = Empty Then
            MsgBox "Please specify the date to free site FROM... "
            .txtFrom.SetFocus
        ElseIf .txtTo = Empty Then
            MsgBox "Please specify the date to free site TO... "
            .txtTo.SetFocus
        Else
            FreeSites
        End If
    End With
    Exit Sub
err:
    ErrorMessage
End Sub
Private Sub FreeSites()
On Error GoTo err
With Me
        Dim SDate, EDate As Date
        SDate = Format(.txtFrom, "MMMM dd,YYYY"): EDate = Format(.txtTo, "MMMM dd,YYYY")
    If MsgBox("System is going to Free the site in question from: " & SDate & " to: " & EDate, vbOKCancel) = vbOK Then
        
        If bBillBoard = True Then
            
            Do While SDate <= EDate
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "Select * From ODASMBillBoardSchedule Where MastNo = '" & CurrentRecord & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,YYYY") & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub
                    rsFindRecord!JobBriefItemNo = ""
                    rsFindRecord!Reserved = "N"
                    rsFindRecord!Allocated = "N"
                    
                    rsFindRecord.Update
                    SDate = DateAdd(d, 1, SDate)
            Loop
            ShowBBSchedule
        ElseIf bBillBoardFace = True Then
            Do While SDate <= EDate
                Set rsFindRecord = New ADODB.Recordset
                rsFindRecord.Open "Select * From ODASMSiteSchedule Where SiteNo = '" & CurrentRecord & "' and ScheduleDate = '" & Format(SDate, "MMMM dd,YYYY") & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                    If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub
                    rsFindRecord!JobBriefItemNo = ""
                    rsFindRecord!Reserved = "N"
                    rsFindRecord!Allocated = "N"
                    
                    rsFindRecord.Update
                    SDate = DateAdd("d", 1, SDate)
            Loop
            ShowBSSchedule
        End If
    End If
End With
Exit Sub
err:
ErrorMessage
End Sub
Private Sub Form_Load()
    If bBillBoardFace = True Then
        Me.cmdFree.Enabled = True: Me.txtFrom.Locked = False: Me.txtTo.Locked = False
        ShowBSSchedule
    ElseIf bBillBoard = True Then
        Me.cmdFree.Enabled = True: Me.txtFrom.Locked = False: Me.txtTo.Locked = False
        ShowBBSchedule
    Else
        Me.cmdFree.Enabled = False: Me.txtFrom.Locked = True: Me.txtTo.Locked = True
        ShowSiteSchedule
    End If
End Sub
Private Sub Form_Resize()
        With frmODASMSiteSchedule
                
                .ListView1.Height = .ScaleHeight
                .ListView1.Width = .ScaleWidth
        End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

