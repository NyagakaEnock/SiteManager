VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOpenJobCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Job Card"
   ClientHeight    =   6315
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   0
      TabIndex        =   38
      Top             =   5040
      Width           =   10575
      Begin VB.TextBox Text21 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpEnvisiageDate 
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   38289
      End
      Begin VB.TextBox txtEnvisage 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   2640
         TabIndex        =   44
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSupervisor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   6840
         TabIndex        =   43
         Top             =   240
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpCommenceDate 
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   38289
      End
      Begin VB.TextBox txtCommence 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   2640
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   5280
         TabIndex        =   47
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Envisiaged Date Of Completion"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Supervisor"
         Height          =   255
         Left            =   5280
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Date Of Commencement"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   10575
      Begin VB.Frame Frame4 
         Caption         =   "Job Brief Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3255
         Left            =   5280
         TabIndex        =   13
         Top             =   120
         Width           =   5175
         Begin VB.TextBox txtJBAuthorizedBy 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   37
            Top             =   2760
            Width           =   3255
         End
         Begin VB.TextBox txtJBDateAuthorized 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   36
            Top             =   2280
            Width           =   3255
         End
         Begin VB.TextBox txtJBApprovedBy 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   35
            Top             =   1800
            Width           =   3255
         End
         Begin VB.TextBox txtJBDateApproved 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   34
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtJBDateCreated 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   33
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtJobBriefNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   32
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label18 
            Caption         =   "Athorized By"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Date Authorized"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Approved By"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Date Approved"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Date Created"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Job Brief No"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Quotation Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   5055
         Begin VB.TextBox txtQTAuthorizedBy 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   25
            Top             =   2880
            Width           =   3255
         End
         Begin VB.TextBox txtQTDateAuthorized 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   24
            Top             =   2400
            Width           =   3255
         End
         Begin VB.TextBox txtQTApprovedBy 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox txtQTDateApproved 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   19
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtQTDate 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox txtQuotationNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label12 
            Caption         =   "Authorized By"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Date Authorized"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Approved By"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Date Approved"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Quotation Date"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Quotation No"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   15
         Left            =   720
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtDeptCode 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   4080
         TabIndex        =   56
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtContactName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   55
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtExpectedDOC 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         TabIndex        =   53
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDeadlineDate 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   51
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtJobCardNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5880
         TabIndex        =   49
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtClientName 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtLpono 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   8880
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDateAuthorized 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtDepartment 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label25 
         Caption         =   "Contact"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Expected D.O.C"
         Height          =   255
         Left            =   6960
         TabIndex        =   52
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Deadline Date"
         Height          =   255
         Left            =   3600
         TabIndex        =   50
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "L.P.O No"
         Height          =   255
         Left            =   7920
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Job Card No"
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Date Authorized"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenJobCard 
         Caption         =   "Open JobCard"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnukljHJJKNkl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseJobCard 
         Caption         =   "Close Job Card"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmOpenJobCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyJobCardNo

Private Sub dtpCommenceDate_CloseUp()
On Error GoTo Err
With Me
  If .dtpCommenceDate.Value < Date Then
    MsgBox "Commence Date must be today or in the future", vbCritical, "Invalid Date"
    .txtCommence.SetFocus
    Else
    .txtCommence.Text = .dtpCommenceDate.Value
    End If
End With
Exit Sub
Err:
   ErrorMessage
End Sub

Private Sub dtpEnvisiageDate_CloseUp()
On Error GoTo Err
With Me
  If .dtpEnvisiageDate.Value < Date Then
    MsgBox "Envisiaged Date of Completion must be today or in the future", vbCritical, "Invalid Date"
    .txtEnvisage.SetFocus
    Else
    .txtEnvisage.Text = .dtpEnvisiageDate.Value
    End If
End With
Exit Sub
Err:
   ErrorMessage
End Sub

Private Sub Form_Load()
On Error GoTo Err
With SchedulingMain
If .ListView1.ListItems.Count = 0 Or .ListView1.View <> lvwReport Then Exit Sub
    
    Dim i, j, k
    j = .ListView1.ListItems.Count
    
    If j = 0 Then Exit Sub
    
    For i = 1 To j
        If .ListView1.ListItems(i).Checked = False Then
            .ListView1.ListItems(i).Checked = False
        End If
   
    
      If .ListView1.ListItems(i).Checked = True Then
    
        
        frmOpenJobCard.txtJobCardNo.Text = .ListView1.ListItems(i).Text
        frmOpenJobCard.txtDeptCode.Text = .ListView1.ListItems(i).SubItems(1)
        frmOpenJobCard.txtDepartment.Text = .ListView1.ListItems(i).SubItems(2)
        frmOpenJobCard.txtClientName.Text = .ListView1.ListItems(i).SubItems(8)
        frmOpenJobCard.txtContactName.Text = .ListView1.ListItems(i).SubItems(9)
        frmOpenJobCard.txtDateAuthorized.Text = .ListView1.ListItems(i).SubItems(3)
        frmOpenJobCard.txtDeadLineDate.Text = .ListView1.ListItems(i).SubItems(7)
        frmOpenJobCard.txtExpectedDOC = .ListView1.ListItems(i).SubItems(6)
        frmOpenJobCard.txtSupervisor.Text = CurrentUserName
        
        Call LoadJobCardDetails
    ElseIf .ListView1.ListItems(i).Checked = False Then
        
    End If
    Next i
End With
Exit Sub
Err:
    ErrorMessage
End Sub
Private Sub LoadJobCardDetails()
'On Error GoTo Err
With Me
   Set rsFindRecord = New ADODB.Recordset
     rsFindRecord.Open "SELECT A.QuotationNo As QT,A.QuotationDate As QTDate,A.DateApproved As QTDateApproved,A.ApprovedBy As QTApprovedBy,A.DateAuthorized As QTDateAuthorized,A.AuthorizedBy As QTAuthorizedBy,B.JobBriefNo As JBBriefNo,B.DateCreated As JBDateCreated,B.DateApproved As JBDateApproved,B.ApprovedBy As JBApprovedBy,B.DateAuthorized As JbDateAuthorized,B.AuthorizedBy As JBAuthorizedBy FROM AdvertQuotation A,AdvertJobBrief B,AdvertJobCard C WHERE A.QuotationNo = B.QuotationNumber AND B.JobBriefNo = C.JobCardNo AND B.JobBriefNo = '" & JobCardNo & "' AND C.JobCardno = '" & JobCardNo & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
     If rsFindRecord.BOF And rsFindRecord.EOF Then Exit Sub

         .txtQuotationNo.Text = rsFindRecord!QT
         .txtQTDate.Text = rsFindRecord!QTDate
         .txtQTDateApproved.Text = rsFindRecord!QTDateApproved
         .txtQTApprovedBy.Text = rsFindRecord!QTApprovedBy
         .txtQTDateAuthorized.Text = rsFindRecord!QTDateAuthorized
         .txtQTAuthorizedBy.Text = rsFindRecord!QTAuthorizedBy
         .txtJobBriefNo.Text = rsFindRecord!JBBriefNo
         .txtJBDateCreated.Text = rsFindRecord!JBDateCreated
         .txtJBDateApproved.Text = rsFindRecord!JBDateApproved
         .txtJBApprovedBy.Text = rsFindRecord!JBApprovedBy
         .txtJBDateAuthorized.Text = rsFindRecord!JBDateAuthorized
         .txtJBAuthorizedBy.Text = rsFindRecord!JBAuthorizedBy

End With
Exit Sub
Err:
   ErrorMessage
End Sub
Private Sub mnuOpenJobCard_Click()
'On Error GoTo Err
Dim CommenceDate, EnvisiageDate As Variant
With Me
CommenceDate = Format(.txtCommence.Text, "MMMM dd,yyyy")
EnvisiageDate = Format(.txtEnvisage.Text, "MMMM dd,yyyy")
   
   If .txtJobCardNo.Text = "" Then Exit Sub
     If Not ValidOpen Then Exit Sub
     If CDate(.txtDeadLineDate) < Date Then
       MsgBox "The job card has expired on   " & .txtDeadLineDate.Text & "  ......!! The System is going to abort the procedure!! ", vbCritical, "Expired Job Brief"
       Unload frmOpenJobCard
         ElseIf CDate(.dtpEnvisiageDate.Value) < CDate(.dtpCommenceDate.Value) Then
            MsgBox "The Envisiaged Date of Completion Can not be earlier than the commence date", vbExclamation, "Invalid Date"
            .txtEnvisage.SetFocus: Exit Sub
              ElseIf CDate(.dtpEnvisiageDate) > CDate(.txtDeadLineDate.Text) Then
                MsgBox "The envisiage date of completion can not be later than the deadline date", vbExclamation, "Invalid Date"
                  .txtEnvisage.SetFocus: Exit Sub
                    ElseIf CDate(.txtDeadLineDate.Text) = Date Then
                      If MsgBox("This Job Card expires today...!!! Are you sure you want to proceed with the job", vbExclamation + vbYesNo, "Job Brief Expiry") = vbYes Then
                       GoTo OpenCard
                      
                       Else
                       Unload frmOpenJobCard
                       End If
                  Else
OpenCard:
               Set rsLineUpdate = New ADODB.Recordset
                  rsLineUpdate.Open "UPDATE AdvertJobCard SET EnvisiagedDateOfCompletion = '" & EnvisiageDate & "',Dateofcommencement = '" & CommenceDate & "',Status =  '" & "OPEN" & "' ,Opened = '" & "Y" & "',OpenedBy = '" & CurrentUserName & "',DateOpened = '" & MyCurrentDate & "' WHERE JobcardNo = '" & Trim(.txtJobCardNo.Text) & "' AND DeptCode = '" & Trim(.txtDeptCode.Text) & "',SupervisedBy = '" & Trim(.txtSupervisor.Text) & "'", cnCOMMON, adOpenKeyset, adLockOptimistic
                Set rsLineUpdate = Nothing
            End If
       
          MsgBox "Job Card number " & .txtJobCardNo.Text & " has successfully been opened", vbInformation, "Open Job Card"
          Unload frmOpenJobCard
          RemoveCurrentListItem
End With
Exit Sub
Err:
  ErrorMessage
End Sub
Public Sub RemoveCurrentListItem()
'On Error GoTo Err
With SchedulingMain
Dim i, j, k
   j = .ListView1.ListItems.Count: i = 1
     If j = 0 Then Exit Sub
     
     For i = 1 To j
      If .ListView1.ListItems(i).Checked = True Then
         .ListView1.ListItems.Remove (i): Exit Sub
      End If
    Next i
End With
Exit Sub
Err:
   ErrorMessage
End Sub

Private Function ValidOpen() As Boolean
On Error GoTo Err
  With Me
     If .txtCommence.Text = "" Then
      MsgBox "Job commence Date Required"
      .txtCommence.SetFocus
       ValidOpen = False
        ElseIf .txtEnvisage.Text = "" Then
        MsgBox "Envisiaged Date of Completion Required"
        .txtEnvisage.SetFocus
        ValidOpen = False
        ElseIf .txtJobCardNo.Text = "" Then
        MsgBox "Job Card No Required"
        .txtJobCardNo.SetFocus
        ValidOpen = False
        Else
        ValidOpen = True
        End If
  End With
Exit Function
Err:
   ErrorMessage
End Function
