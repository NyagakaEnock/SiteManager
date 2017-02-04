VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASSearchSitesNotPaid 
   Caption         =   "Outstanding Rent as at"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   11460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Picture         =   "frmODASSearchSitesNotPaid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Export to Excel"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Search List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11295
      Begin MSComCtl2.DTPicker txtLastDate 
         Height          =   330
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   117637123
         CurrentDate     =   39357
      End
      Begin MSComCtl2.DTPicker txtStartDate 
         Height          =   330
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   117637123
         CurrentDate     =   39357
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Date:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00808000&
      Caption         =   "&Print Record"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   3015
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8916
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
End
Attribute VB_Name = "frmODASSearchSitesNotPaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strReport As String

Private Sub SaveReportNo()
With frmODASSearchSitesNotPaid
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
With frmODASSearchSitesNotPaid
        
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

Private Sub cmdExport_Click()
 bexportRECORD = True
 On Error GoTo err
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
                frmODASRRentDue.Caption = Me.Caption
                Load frmODASRRentDue
                frmODASRRentDue.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage

End Sub

Private Sub cmdSearch_Click()
 bexportRECORD = False
On Error GoTo err
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
                frmODASRRentDue.Caption = Me.Caption
                Load frmODASRRentDue
                frmODASRRentDue.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASSearchSitesNotPaid
                
                If .txtLastDate.Value <= " " Then
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
    Select Case strReport
        Case "PendingPayment"
            'showALLRentPendingPayment
            showALLRentPendingPaymentAsAtASingleDate
        Case "PendingPaymentAsAtASingleDate"
            showALLRentPendingPaymentAsAtASingleDate
        Case "VouchersPrepared"
            showALLRentVouchersPrepared
        Case "PendingConfirmation"
            showALLRentPendingConfirmation
        Case "PaymentsConfirmed"
            showALLRentWithPaymentsConfirmed
        Case Else
            showALLRentNOTPAIDThisPeriod
    End Select
     .cmdSearch.Enabled = True
     .cmdExport.Enabled = True
  End If
End With
End Sub


Private Sub Form_Activate()
With Me
    Select Case strReport
        Case "PendingPayment"
            .txtStartDate.Visible = False
            .Caption = "Rents Pending payment As At"
        Case "PendingPaymentAsAtASingleDate"
            .txtStartDate.Visible = False
            .Caption = "Rents Pending payment As At"
        Case "VouchersPrepared"
            .txtStartDate.Visible = True
            .Caption = "Payment Vouchers Prepared Between "
        Case "PendingConfirmation"
                .txtStartDate.Visible = True
            .Caption = "Payments Pending Confirmation Between"
        Case "PaymentsConfirmed"
            .txtStartDate.Visible = True
            .Caption = "Payments Confirmed Between"
        Case Else
             .txtStartDate.Visible = True
            .Caption = "Outstanding Payments Between "
    End Select
End With

End Sub

Private Sub Form_Load()
With Me
  .txtStartDate.Value = Date
  .txtLastDate.Value = Date
  .cmdSearch.Enabled = False
  .cmdExport.Enabled = False


End With
End Sub
