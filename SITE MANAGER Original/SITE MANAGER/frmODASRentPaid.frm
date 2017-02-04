VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmODASRentPaid 
   Caption         =   "Rent Paid"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2778
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.TextBox txtTotalRent 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Text            =   " "
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cbostartperiod 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1710
   End
   Begin VB.ComboBox cboendperiod 
      Height          =   315
      Left            =   4200
      TabIndex        =   0
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Start Period"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "End Period"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   480
      Width           =   780
   End
End
Attribute VB_Name = "frmODASRentPaid"
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
If Me.cbostartperiod.Text = "" Then
MsgBox "Enter the start period first"
Me.cbostartperiod.SetFocus
Else
'showALLRentPAIDThisPeriod
 'computeTOTALRentThisPeriod
 
End If
End Sub

Private Sub cbostartperiod_Click()
With Me
'SelectAccountingperiod
  SelectDescription
  
End With
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
                'SaveReportNo
                'updateReportNo
                'Unload frmODASYearsearch
                showALLRentPAIDThisPeriod
                Load frmODASRRentPayments
                frmODASRRentPayments.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASRentPaid
                
                If .cbostartperiod.Text <= Empty Then
                    MsgBox " The start Period is Required ............"
                    .cbostartperiod.SetFocus
                ElseIf .cboendperiod.Text <= Empty Then
                    MsgBox "The End Period is Required ..............."
                    .cboendperiod.SetFocus
               
                Else
                        bSaveRECORD = True
                End If
                
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub Form_Activate()
    SelectAccountingperiod
    updateCurrentPayment
End Sub


