VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmODASPaySerach 
   Caption         =   "Rent Paid"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCurrentPeriod 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtEndDate 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtStartDate 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPickerStartDate 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Format          =   3801089
      CurrentDate     =   38298
   End
   Begin MSComCtl2.DTPicker DTPickerEndDate 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Format          =   3801089
      CurrentDate     =   38298
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "End Date"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmODASPaySerach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveReportNo()
With frmODASYearsearch
        Set rsSAVE = New ADODB.Recordset
        strSQL = "Select * From ODASPReport where ReportNo like'" & .txtReportNo & "' "
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF Or rsSAVE.BOF Then
         rsSAVE.AddNew
         rsSAVE!StartDate = .txtStartDate
         rsSAVE!Enddate = .txtEndDate
         rsSAVE!PreparedBy = CurrentUserName
         rsSAVE!DatePrepared = Date
        End If
         rsSAVE.Update
        .txtReportNo.Text = rsSAVE!ReportNO
        CurrentRecord = .txtReportNo.Text
         rsSAVE.Requery
        Set rsSAVE = Nothing
End With
End Sub

Private Sub updateReportNo()
With frmODASYearsearch
        
        Set rsSAVE = New ADODB.Recordset
        
        strSQL = "Select * From ODASMInstallment where PaymentDueDate >= '" & Format(.txtStartDate, "YYYY/MM/DD") & "' and PaymentDueDate <= '" & Format(.txtEndDate, "YYYY/MM/DD") & "' and PaymentFlag = 'Y' "
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
                        rsCONTROL!ReportNO = .txtReportNo
        
                        rsCONTROL.Update
                End If
                
                rsSAVE.MoveNext
        Loop
        
        Set rsSAVE = Nothing
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
                Load frmODASRRentPayments
                frmODASRRentPayments.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASYearsearch
                
                If .txtStartDate.Text <= Empty Then
                    MsgBox " The start Date is Required ............"
                    .txtStartDate.SetFocus
                ElseIf .txtEndDate.Text <= Empty Then
                    MsgBox "The Last Date is Required ..............."
                    .txtEndDate.SetFocus
                ElseIf CDate(.txtStartDate) > CDate(.txtEndDate) Then
                    MsgBox "The Start Date Cannot be After the Last Date....."
                Else
                        bSaveRECORD = True
                End If
                
        End With
        
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub DTPickerEndDate_CloseUp()
With Me
Me.txtEndDate.Text = Me.DTPickerEndDate
CurrentEndDate = .txtEndDate

End With
End Sub

Private Sub DTPickerStartDate_CloseUp()
With Me
Me.txtStartDate.Text = Me.DTPickerStartDate
CurrentStartDate = .txtStartDate
End With
End Sub

Private Sub Form_Activate()
  Me.txtCurrentPeriod.Text = CurrentPeriod
End Sub

Private Sub Form_Load()
Me.txtEndDate = Date
Me.txtStartDate = Date
Me.DTPickerEndDate.Value = Date
Me.DTPickerStartDate.Value = Date
End Sub

