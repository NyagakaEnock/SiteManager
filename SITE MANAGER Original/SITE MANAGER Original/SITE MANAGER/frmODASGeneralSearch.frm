VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmODASGeneralsearch 
   Caption         =   "Rent Due"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   " "
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   10095
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Left            =   7080
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " "
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   8160
         TabIndex        =   7
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   39357
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   6000
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   39357
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "End Period"
         Height          =   195
         Left            =   5040
         TabIndex        =   3
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Start Period"
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   825
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
      Left            =   3000
      TabIndex        =   0
      Top             =   4560
      Width           =   3495
   End
End
Attribute VB_Name = "frmODASGeneralsearch"
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
'         rsSAVE!Enddate = .txtEndDate
'         rsSAVE!PreparedBy = CurrentUserName
'         rsSAVE!DatePrepared = Date
        End If
         rsSAVE.Update
'        .txtReportNo.Text = rsSAVE!ReportNO
'        CurrentRecord = .txtReportNo.Text
         rsSAVE.Requery
        Set rsSAVE = Nothing
End With
End Sub

Private Sub updateReportNo()
With frmODASYearsearch
        
        Set rsSAVE = New ADODB.Recordset
        
        'strSQL = "Select * From ODASMInstallment where PaymentDueDate >= '" & Format(.txtStartDate, "YYYY/MM/DD") & "' and PaymentDueDate <= '" & Format(.txtEndDate, "YYYY/MM/DD") & "' and PaymentFlag = 'N' "
        rsSAVE.Open strSQL, cnCOMMON, adOpenKeyset, adLockOptimistic
        If rsSAVE.EOF Or rsSAVE.BOF Then Exit Sub
        Dim DF As Integer
        DF = rsSAVE.RecordCount
        
        Do While Not rsSAVE.EOF
                
                Set rsCONTROL = New ADODB.Recordset
                'strCONTROL = "Select * From ODASMInstallment where ContractNo = '" & rsSAVE!ContractNo & "'"
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
Private Sub cmdSearch_Click()
On Error GoTo err
        bSaveRECORD = False
        ValidateRECORD
        If bSaveRECORD = True Then
'                SaveReportNo
'                updateReportNo
'                Unload frmODASYearsearch
                Load frmODASRRentDue
                frmODASRRentDue.Show vbModal
        End If
Exit Sub
err:
    ErrorMessage
End Sub
Private Sub ValidateRECORD()
On Error GoTo err
        With frmODASYearsearch
                
                If .cbostartperiod.Text <= Empty Then
                    MsgBox " The start Period is Required ............"
                    .cbostartperiod.SetFocus
                ElseIf .cboendperiod.Text <= Empty Then
                    MsgBox "The End period is Required ..............."
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
 updatecurrentperiod
End Sub
Private Sub cboendperiod_Click()


 
End Sub

Private Sub cbostartperiod_Click()
With Me
  SelectDescription1
  End With
End Sub

