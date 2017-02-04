VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASSearchGeneral 
   Caption         =   "Searching For General Data  - Specified Duration"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Search List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8415
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtStartDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtLastDate 
         Height          =   285
         Left            =   5760
         TabIndex        =   3
         Text            =   " "
         Top             =   240
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPickerLastDate 
         Height          =   285
         Left            =   7920
         TabIndex        =   2
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39357
      End
      Begin MSComCtl2.DTPicker DTPickerStartDate 
         Height          =   255
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39357
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Date"
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Print Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4080
      Width           =   3015
   End
End
Attribute VB_Name = "frmODASSearchGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveReportNo()
With frmODASYearsearch
        Set rsSAVE = New ADODB.Recordset
        'strSQL = "Select * From ODASPReport where ReportNo like'" & .txtReportNo & "' "
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF Or rsSAVE.BOF Then
         rsSAVE.AddNew
         'rsSAVE!StartDate = .txtStartDate
         'rsSAVE!Enddate = .txtEndDate
         'rsSAVE!PreparedBy = CurrentUserName
         'rsSAVE!DatePrepared = Date
        End If
         rsSAVE.Update
        '.txtReportNo.Text = rsSAVE!ReportNO
        'CurrentRecord = .txtReportNo.Text
         rsSAVE.Requery
        Set rsSAVE = Nothing
End With
End Sub

Private Sub updateReportNo()
With frmODASYearsearch
        
        Set rsSAVE = New ADODB.Recordset
        
       ' strSQL = "Select * From ODASMInstallment where PaymentDueDate >= '" & Format(.txtStartDate, "YYYY/MM/DD") & "' and PaymentDueDate <= '" & Format(.txtEndDate, "YYYY/MM/DD") & "' and PaymentFlag = 'Y' "
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
        Dim DF As Integer
        DF = rsSAVE.RecordCount
        
        Do While Not rsSAVE.EOF
                
                Set rsCONTROL = New ADODB.Recordset
                strCONTROL = "Select * From ODASMInstallment where ContractNo = '" & rsSAVE!ContractNo & "' "
                rsCONTROL.Open strCONTROL, cnCOMMON, adOpenKeyset, adLockOptimistic
                If rsCONTROL.EOF Or rsCONTROL.BOF Then
                Else
                       ' rsCONTROL!ReportNO = .txtReportNo
        
                        rsCONTROL.Update
                End If
                
                rsSAVE.MoveNext
        Loop
        
        Set rsSAVE = Nothing
End With
End Sub

Private Sub cboendperiod_Click()
'If Me.cbostartperiod.Text = "" Then
'MsgBox "Enter the start period first"
'Me.cbostartperiod.SetFocus
'Else
''showALLRentPAIDThisPeriod
' 'computeTOTALRentThisPeriod
'
'End If
End Sub

Private Sub cbostartperiod_Click()
With Me
  SelectDescription
  
End With
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
                Load frmRptGeneralReport
                frmRptGeneralReport.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASSearchSiteNewSites
                
                If .txtStartDate.Text <= " " Then
                    MsgBox " The start Period is Required ............"
                    .txtStartDate.SetFocus
                ElseIf .txtLastDate.Text <= " " Then
                    MsgBox "The End Period is Required ..............."
                    .txtLastDate.SetFocus
               
                Else
                        bSaveRECORD = True
                End If
                
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub

Private Sub Command1_Click()
With Me
  ValidateRECORD
  If bSaveRECORD = True Then
     showALLRentNOTPAIDThisPeriod
  End If
End With
End Sub

Private Sub DTPickerLastDate_CloseUp()
   Me.txtLastDate.Text = Me.DTPickerLastDate.Value
End Sub

Private Sub DTPickerStartDate_CloseUp()
    Me.txtStartDate.Text = Me.DTPickerStartDate.Value
End Sub

Private Sub Form_Load()
With Me
  .txtLastDate.Text = Date
  .txtStartDate.Text = Date
End With
End Sub
